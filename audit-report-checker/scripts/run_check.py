#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_check.py — 审计报告核查编排脚本（代码侧检查）

职责定位：
  Claude 做语义（定位/口径判断/终审），代码做确定性算术。三层分工：
  - 定位：Claude Step2 一次性生成 note_map.json（科目→附注表id+取数口径），准且低token
  - 取数：DeepSeek locate_note_values 按 note_map 精确取附注数（只发该科目候选小表，省token）
  - 算术：calculator 比较（报表数 vs 附注数、期初+增-减=期末），确定性
  执行的检查：L1 报表间勾稽 + 表注勾稽算术 + 附注内reconcile + 横加竖加scan + 文本/格式。

输入：
  python3 run_check.py <parse输出目录> --statements <statements.json> --scope standard -o results.json

输出：
  results.json：统一的检查结果结构，export_report.py 也读这个

执行的检查（按 scope）：
  必做（所有 scope）：
    1. L1 报表间勾稽（有 statements 时）
    2. 格式-页码连续性
    3. 格式-金额单位一致性
    4. 文本-错别字词库

  standard/deep 额外：
    5. 格式-公司名称一致性
    6. 格式-附注编号连续性
    7. 格式-事务所/文号一致性

Severity 分级：
  - error：calculator MISMATCH（确定性差异，如勾稽不平、错别字）
  - warning：可疑（页码缺失、公司名变体、单位不一致）
  - info：通过或提示
"""
import argparse
import json
import os
import re
import sys
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Optional

# 同目录 import calculator_rules
_script_dir = Path(__file__).parent
sys.path.insert(0, str(_script_dir))

from calculator import check_equal, evaluate_expression, parse_number, reconcile, sum_numbers
from calculator_rules import (
    CheckItem,
    check_balance_sheet,
    check_cash_flow_statement,
    check_equity_change_statement,
    check_income_statement,
)

# 尝试导入 ai_worker（可选，用于 DeepSeek 并发加速）
try:
    import ai_worker
    AI_WORKER_AVAILABLE = True
except ImportError:
    AI_WORKER_AVAILABLE = False


# ──────────────────────────────────────────────────────────────────
# 结果结构
# ──────────────────────────────────────────────────────────────────
def create_result(
    check_type: str,
    rule_name: str,
    severity: str,
    passed: bool,
    description: str,
    expected: Optional[str] = None,
    actual: Optional[str] = None,
    difference: Optional[float] = None,
    source_location: str = "",
    target_location: str = "",
    evidence: str = "",
    page: Optional[str] = None,
    context: str = "",
    confidence: Optional[float] = None,
) -> dict:
    """创建统一的结果对象。

    page/context 用于在 Excel 报告中定位原文：
    - page: 页码（纯数字或含"p"前缀），审计师据此翻原文
    - context: 原文摘录（约50-100字），审计师不用翻原文即可核对
    source_file 由 main() 统一回填（所有结果指向同一份原报告）。
    confidence: 置信度（0-1），AI 检查结果使用，代码验算默认 1.0
    """
    result = {
        "check_type": check_type,
        "rule_name": rule_name,
        "severity": severity,
        "passed": passed,
        "description": description,
        "source_location": source_location,
        "target_location": target_location,
        "evidence": evidence,
        "confidence": confidence if confidence is not None else 1.0,  # 默认 1.0（代码验算确定性）
    }
    if expected is not None:
        result["expected"] = expected
    if actual is not None:
        result["actual"] = actual
    if difference is not None:
        result["difference"] = f"{difference:.2f}" if isinstance(difference, (int, float)) else str(difference)
    if page is not None:
        result["page"] = str(page)
    if context:
        result["context"] = context
    return result


def convert_check_item(
    check_item: CheckItem,
    check_type: str,
    source_location: str,
    page: Optional[str] = None,
    context: str = "",
) -> dict:
    """将 calculator_rules 的 CheckItem 转换为结果对象。"""
    severity = check_item.severity
    passed = check_item.passed
    # 如果 passed=True，severity 降级为 info（通过）
    if passed and severity == "error":
        severity = "info"
    # 如果 passed=False，保持 error
    return create_result(
        check_type=check_type,
        rule_name=check_item.rule_name,
        severity=severity,
        passed=passed,
        description=check_item.description,
        expected=check_item.expected,
        actual=check_item.actual,
        difference=check_item.difference,
        source_location=source_location,
        target_location="",
        evidence=f"calculator_rules.{check_item.rule_name}",
        page=page,
        context=context,
    )


# ──────────────────────────────────────────────────────────────────
# 检查函数
# ──────────────────────────────────────────────────────────────────
def run_l1_checks(statements: dict, results: list, extracted_tables_path: Optional[Path] = None) -> None:
    """执行 L1 报表间勾稽检查。"""
    if not statements:
        return

    # 读取 extracted_tables.json 用于回填页码
    tables_data = None
    if extracted_tables_path and extracted_tables_path.exists():
        try:
            tables_data = json.loads(extracted_tables_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    # 表名映射到检查函数
    check_funcs = {
        "资产负债表": check_balance_sheet,
        "利润表": check_income_statement,
        "现金流量表": check_cash_flow_statement,
        "所有者权益变动表": check_equity_change_statement,
        # 合并报表变体
        "合并资产负债表": check_balance_sheet,
        "合并利润表": check_income_statement,
        "合并现金流量表": check_cash_flow_statement,
        "合并所有者权益变动表": check_equity_change_statement,
        # 母公司报表变体
        "母公司资产负债表": check_balance_sheet,
        "母公司利润表": check_income_statement,
        "母公司现金流量表": check_cash_flow_statement,
        "母公司所有者权益变动表": check_equity_change_statement,
    }

    # 报表特征词用于从 extracted_tables 识别报表页码
    table_keywords = {
        "资产负债表": ["流动资产", "流动负债", "资产总计"],
        "利润表": ["营业收入", "营业成本"],
        "现金流量表": ["经营活动产生的现金流量", "销售商品"],
        "所有者权益变动表": ["所有者权益变动表"],
    }

    for table_name, statement in statements.items():
        # 查找匹配的检查函数
        check_func = None
        for key, func in check_funcs.items():
            if key in table_name:
                check_func = func
                break

        if check_func is None:
            continue

        # 从 extracted_tables 查找报表页码
        page = None
        if tables_data:
            # 确定报表类型
            table_type = None
            for key in table_keywords.keys():
                if key in table_name:
                    table_type = key
                    break

            if table_type:
                # 在 tables 中查找包含特征词的表
                for tbl in tables_data.get("tables", []):
                    if tbl.get("page") is None:
                        continue
                    # 检查 headers 和 rows 内容
                    all_text = " ".join(tbl.get("headers", []))
                    for row in tbl.get("rows", []):
                        all_text += " " + " ".join(row)
                    # 检查是否包含任一特征词
                    for kw in table_keywords.get(table_type, []):
                        if kw in all_text:
                            page = str(tbl["page"])
                            break
                    if page:
                        break

        try:
            check_items = check_func(statement)
            for item in check_items:
                result = convert_check_item(
                    item,
                    "报表间",
                    table_name,
                    page=page,
                    context=table_name,
                )
                results.append(result)
        except Exception as e:
            results.append(
                create_result(
                    check_type="报表间",
                    rule_name=f"{table_name}检查",
                    severity="error",
                    passed=False,
                    description=f"检查执行失败: {e}",
                    source_location=table_name,
                    evidence="calculator_rules",
                )
            )


def check_page_continuity(report_md: Path, results: list) -> None:
    """检查页码连续性。"""
    if not report_md.exists():
        return

    content = report_md.read_text(encoding="utf-8")
    # 提取所有 page=N
    page_numbers = re.findall(r"page=(\d+)", content)
    if not page_numbers:
        return

    pages = sorted(set(int(p) for p in page_numbers))
    if not pages:
        return

    # 检测缺失页码
    missing = []
    for i in range(min(pages), max(pages) + 1):
        if i not in pages:
            missing.append(i)

    if missing:
        # 判断是否在合理范围内（≤20页且≤10%）
        total_pages = len(pages)
        missing_count = len(missing)
        allowed_missing = min(20, int(total_pages * 0.1))

        if missing_count <= allowed_missing:
            severity = "info"  # 少量缺失，可能是 pdfplumber 漏页
        else:
            severity = "warning"  # 过多缺失，可疑

        results.append(
            create_result(
                check_type="格式",
                rule_name="页码连续性",
                severity=severity,
                passed=False,
                description=f"缺失页码: {missing[:10]}{'...' if len(missing) > 10 else ''}",
                source_location="report.md",
                evidence="页码提取正则",
            )
        )
    else:
        results.append(
            create_result(
                check_type="格式",
                rule_name="页码连续性",
                severity="info",
                passed=True,
                description=f"页码连续，共 {len(pages)} 页",
                source_location="report.md",
                evidence="页码提取正则",
            )
        )


def check_unit_consistency(report_md: Path, results: list) -> None:
    """检查金额单位一致性。"""
    if not report_md.exists():
        return

    content = report_md.read_text(encoding="utf-8")
    # 移除 HTML 注释，避免干扰
    content = re.sub(r"<!--.*?-->", "", content, flags=re.DOTALL)

    # 查找所有金额单位声明
    unit_pattern = r"单位[：:]\s*(元|万元|千元|百万元)"
    units = re.findall(unit_pattern, content)

    if not units:
        results.append(
            create_result(
                check_type="格式",
                rule_name="金额单位声明",
                severity="info",
                passed=True,
                description="未找到金额单位声明",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
        return

    # 检查是否一致
    unique_units = set(units)
    if len(unique_units) == 1:
        results.append(
            create_result(
                check_type="格式",
                rule_name="金额单位一致性",
                severity="info",
                passed=True,
                description=f"金额单位一致: {list(unique_units)[0]}",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
    else:
        results.append(
            create_result(
                check_type="格式",
                rule_name="金额单位一致性",
                severity="warning",
                passed=False,
                description=f"金额单位不一致: {', '.join(unique_units)}",
                source_location="report.md",
                evidence="正则匹配",
            )
        )


# 错别字词库
TYPO_WORDS = {
    "先讲先出": "先进先出",
    "其本": "基本",
    "员债": "负债",
    "所有都权益": "所有者权益",
    "润表": "利润表",
    "计报告": "审计报告",
    "资负债表": "资产负债表",
    "现金流理表": "现金流量表",
    "财国报表": "财务报表",
    "合计师事务所": "会计师事务所",
    "特殊普通合饮": "特殊普通合伙",
    "财务报麦": "财务报告",
    "审诉报告": "审计报告",
}

# 完整词白名单（用于消除误报：若错误模式是某个完整词的子串且原文包含该完整词，则跳过）
FULL_WORD_WHITELIST = [
    "利润表",  # 对应"润表"
    "审计报告",  # 对应"计报告"/"审诉报告"
    "资产负债表",  # 对应"资负债表"
    "现金流量表",  # 对应"现金流理表"
    "财务报表",  # 对应"财国报表"
    "会计师事务所",  # 对应"合计师事务所"
    "特殊普通合伙",  # 对应"特殊普通合饮"
    "财务报告",  # 对应"财务报麦"
    "所有者权益",  # 对应"所有都权益"
    "先进先出",  # 对应"先讲先出"
]


def check_typos(report_md: Path, results: list) -> None:
    """检查错别字。"""
    if not report_md.exists():
        return

    content_raw = report_md.read_text(encoding="utf-8")
    # 移除 HTML 注释和空白，解决跨行词问题
    content = re.sub(r"<!--.*?-->", "", content_raw, flags=re.DOTALL)
    content = re.sub(r"\s+", "", content)  # 移除所有空白

    found = []
    for typo, correct in TYPO_WORDS.items():
        if typo in content:
            # 白名单过滤：若错误模式是某个完整词的子串且原文包含该完整词，则跳过
            is_false_positive = False
            for full_word in FULL_WORD_WHITELIST:
                if typo in full_word and full_word in content:
                    # 这是一个误报（原文用了正确的完整词，只是匹配到了截断片段）
                    is_false_positive = True
                    break

            if is_false_positive:
                continue

            # 查找上下文（用清洗后的文本，跨行词在原文中被换行断开）
            pos = content.find(typo)
            if pos != -1:
                context = content[max(0, pos - 20): pos + len(typo) + 20]

                # 从原始 content_raw 提取页码
                page = None
                # 构建跨行匹配的正则（用 \s* 连接每个字符）
                typo_pattern = r"\s*".join(re.escape(c) for c in typo)
                match = re.search(typo_pattern, content_raw)
                if match:
                    # 从匹配位置往前找最近的 SOURCE 注释
                    before_text = content_raw[:match.start()]
                    # 在 before_text 中查找最后一个 SOURCE 注释
                    last_source_match = None
                    for m in re.finditer(r'<!-- SOURCE[^>]*page=(\d+)[^>]*-->', before_text):
                        last_source_match = m
                    if last_source_match:
                        page = last_source_match.group(1)

                found.append((typo, correct, context, page))
            else:
                found.append((typo, correct, "(跨行匹配成功)", None))

    if found:
        for typo, correct, context, page in found:
            results.append(
                create_result(
                    check_type="文本",
                    rule_name="错别字检查",
                    severity="error",
                    passed=False,
                    description=f"疑似错别字: '{typo}' 应为 '{correct}'",
                    source_location="report.md",
                    evidence=f"内置词库: ...{context}...",
                    page=page,
                    context=context,
                )
            )
    else:
        results.append(
            create_result(
                check_type="文本",
                rule_name="错别字检查",
                severity="info",
                passed=True,
                description=f"未发现错别字（检查了 {len(TYPO_WORDS)} 个常见词）",
                source_location="report.md",
                evidence="内置词库",
            )
        )


def check_company_consistency(extracted_tables: Path, report_md: Path, results: list) -> None:
    """检查公司名称一致性（standard/deep）。

    增强逻辑：
    1. 从 report.md 全文提取所有公司名（正则匹配含"有限公司/股份有限公司/集团/公司"的连续短语）
    2. 统计各核心词出现次数，次数最多=主体公司名
    3. 审计报告抬头（首个"XX公司"）与编制单位重点比对
    4. 全文公司名核心词唯一 → info 通过；存在明显不同主体 → warning
    """
    if not extracted_tables.exists() or not report_md.exists():
        return

    tables_data = json.loads(extracted_tables.read_text(encoding="utf-8"))
    company_from_json = tables_data.get("company", "")
    report_content = report_md.read_text(encoding="utf-8")

    # 从 report.md 提取公司名（编制单位）
    company_pattern = r"编制单位\s*[:：]\s*([^\n]+)"
    match = re.search(company_pattern, report_content)
    company_from_report = match.group(1).strip() if match else ""

    # 按法定后缀截断公司名
    def truncate_company_name(name: str) -> str:
        """按法定后缀截断公司名，去除杂质。"""
        if not name:
            return ""
        # 法定后缀列表（按长度降序，优先匹配长的）
        legal_suffixes = [
            "集团有限公司",
            "股份有限公司",
            "有限责任公司",
            "有限公司",
            "合伙企业",
            "集团",
        ]
        for suffix in legal_suffixes:
            if suffix in name:
                # 找到后缀位置，截取到后缀末尾
                idx = name.find(suffix)
                return name[: idx + len(suffix)].strip()
        # 没有匹配后缀，保留原样
        return name.strip()

    company_from_report = truncate_company_name(company_from_report)
    company_from_json = truncate_company_name(company_from_json)

    # 从页眉提取（简单正则，可能不准确）
    header_pattern = r"^[^\n]*(?:有限公司|股份有限公司|集团|公司)[^\n]*$"
    headers = re.findall(header_pattern, report_content, re.MULTILINE)
    company_from_header = headers[0] if headers else ""

    # 核心词提取（去掉常见后缀）
    def extract_core(name: str) -> str:
        """提取公司名核心词。"""
        if not name:
            return ""
        # 移除常见后缀
        suffixes = [
            "有限公司",
            "股份有限公司",
            "集团有限公司",
            "股份公司",
            "公司",
            "集团",
        ]
        core = name
        for suffix in suffixes:
            if core.endswith(suffix):
                core = core[: -len(suffix)]
                break
        return core.strip()

    # ========== 增强逻辑：从全文提取所有公司名并统计核心词 ==========
    # 正则匹配公司名：后缀必须完整（有限公司/股份有限公司等），不匹配单独"公司"/"集团"
    # （单独"公司"会误匹配"本公司/该公司/贵公司"等代词；前缀只允许汉字/字母数字，排除括号引号）
    company_name_pattern = r'([\u4e00-\u9fa5A-Za-z0-9]{2,30}(?:有限公司|股份有限公司|有限责任公司|集团有限公司|合伙企业))'
    all_company_matches = re.findall(company_name_pattern, report_content)

    # 去重
    unique_companies = list(set(all_company_matches))

    # 统计各核心词出现次数（排除代词前缀：本公司/该公司/贵公司/各公司/其公司 等）
    core_count = {}
    _PRONOUN_PREFIXES = ('本', '该', '贵', '各', '其', '乃', '我')
    for company in all_company_matches:
        core = extract_core(company)
        # 过滤：太短、代词前缀
        if core and len(core) >= 2 and not core.startswith(_PRONOUN_PREFIXES):
            core_count[core] = core_count.get(core, 0) + 1

    if not core_count:
        # 未提取到公司名，走旧逻辑
        if company_from_json and company_from_report:
            core_json = extract_core(company_from_json)
            core_report = extract_core(company_from_report)
            if core_json and core_report and core_json != core_report:
                results.append(
                    create_result(
                        check_type="格式",
                        rule_name="公司名称一致性",
                        severity="warning",
                        passed=False,
                        description=f"公司名不一致: extracted_tables='{company_from_json}' vs report.md='{company_from_report}'",
                        source_location="extracted_tables.json + report.md",
                        evidence="核心词匹配",
                    )
                )
            else:
                results.append(
                    create_result(
                        check_type="格式",
                        rule_name="公司名称一致性",
                        severity="info",
                        passed=True,
                        description=f"公司名一致: {company_from_json}",
                        source_location="extracted_tables.json + report.md",
                        evidence="核心词匹配",
                    )
                )
        else:
            results.append(
                create_result(
                    check_type="格式",
                    rule_name="公司名称一致性",
                    severity="info",
                    passed=True,
                    description="公司名信息不完整，跳过检查",
                    source_location="extracted_tables.json + report.md",
                    evidence="核心词匹配",
                )
            )
        return

    # 次数最多的核心词=主体公司名
    main_core = max(core_count, key=core_count.get)
    main_core_count = core_count[main_core]
    # 主体公司名（保留原始后缀，从已匹配的公司名中找）
    main_company_name = main_core
    for c in unique_companies:
        if extract_core(c) == main_core:
            main_company_name = c
            break

    # 审计报告抬头（report.md 开头首个"XX公司"）
    header_company_match = re.search(r'^([^\n]*(?:有限公司|股份有限公司|集团|公司)[^\n]*)', report_content, re.MULTILINE)
    header_company = header_company_match.group(1).strip() if header_company_match else ""
    header_company = truncate_company_name(header_company)

    # ========== 只检查结构化位置（编制单位/抬头/json）与主体一致性 ==========
    # 全文统计公司名只用于确定"主体"（出现最多的）。正文叙述里提到母公司/前身/关联方
    # 是正常的，不是"用错公司名"。真正的"公司主体用错"只发生在结构化身份位置：
    # 编制单位行、审计报告抬头、页眉——这些位置公司名必须与主体一致。
    inconsistencies = []

    # 1. 编制单位 vs 主体
    if company_from_report:
        report_core = extract_core(company_from_report)
        if report_core and report_core != main_core and report_core not in main_core and main_core not in report_core:
            inconsistencies.append(f"编制单位'{company_from_report}'与主体'{main_company_name}'核心词不同")

    # 2. extracted_tables company vs 主体
    if company_from_json:
        json_core = extract_core(company_from_json)
        if json_core and json_core != main_core and json_core not in main_core and main_core not in json_core:
            inconsistencies.append(f"表头公司'{company_from_json}'与主体'{main_company_name}'核心词不同")

    # 3. 审计报告抬头 vs 编制单位（两者都存在且核心词不同 → 严重不一致）
    if header_company and company_from_report:
        header_core = extract_core(header_company)
        report_core = extract_core(company_from_report)
        if header_core and report_core and header_core != report_core and header_core not in report_core and report_core not in header_core:
            inconsistencies.append(f"审计报告抬头'{header_company}'与编制单位'{company_from_report}'核心词不同")

    if inconsistencies:
        results.append(
            create_result(
                check_type="格式",
                rule_name="公司名称一致性",
                severity="warning",
                passed=False,
                description="公司主体不一致: " + "；".join(inconsistencies),
                source_location="report.md",
                evidence=f"主体公司名 '{main_company_name}'（全文 {main_core_count} 次）",
            )
        )
    else:
        # 结构化位置一致（或无可比对的结构化位置，如附注-only 无编制单位/表头）
        loc_note = ""
        if not company_from_report and not company_from_json:
            loc_note = "（附注-only：无编制单位/表头可比对，仅以正文主体为准）"
        results.append(
            create_result(
                check_type="格式",
                rule_name="公司名称一致性",
                severity="info",
                passed=True,
                description=f"公司主体一致: {main_company_name}{loc_note}",
                source_location="report.md",
                evidence=f"主体核心词 '{main_core}' 全文出现 {main_core_count} 次（共 {len(all_company_matches)} 次公司名提及）",
            )
        )


def check_note_number_continuity(report_md: Path, results: list) -> None:
    """检查附注编号连续性（standard/deep）。"""
    if not report_md.exists():
        return

    content = report_md.read_text(encoding="utf-8")
    # 移除 HTML 注释
    content = re.sub(r"<!--.*?-->", "", content, flags=re.DOTALL)

    # 查找附注编号：附注四、（X）
    note_pattern = r"附注[一二三四五六七八九十百千万]+[、,.]\s*[\(（](\d+)[\)）]"
    matches = re.findall(note_pattern, content)

    if not matches:
        results.append(
            create_result(
                check_type="格式",
                rule_name="附注编号连续性",
                severity="info",
                passed=True,
                description="未找到附注编号",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
        return

    numbers = sorted(set(int(m) for m in matches))
    if not numbers:
        return

    # 检查连续性
    missing = []
    for i in range(min(numbers), max(numbers) + 1):
        if i not in numbers:
            missing.append(i)

    if missing:
        results.append(
            create_result(
                check_type="格式",
                rule_name="附注编号连续性",
                severity="warning",
                passed=False,
                description=f"附注编号不连续，缺失: {missing[:10]}{'...' if len(missing) > 10 else ''}",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
    else:
        results.append(
            create_result(
                check_type="格式",
                rule_name="附注编号连续性",
                severity="info",
                passed=True,
                description=f"附注编号连续，共 {len(numbers)} 个",
                source_location="report.md",
                evidence="正则匹配",
            )
        )


def check_firm_consistency(report_md: Path, results: list) -> None:
    """检查事务所/文号一致性（standard/deep）。"""
    if not report_md.exists():
        return

    content = report_md.read_text(encoding="utf-8")
    # 移除 HTML 注释
    content = re.sub(r"<!--.*?-->", "", content, flags=re.DOTALL)

    # 查找事务所名（常见模式）
    firm_patterns = [
        r"([^\s,，。；;\n]+会计师事务所(?:有限责任)?公司)",
        r"([^\s,，。；;\n]+会计师事务所(?:特殊普通合伙)?)",
    ]
    firms = []
    for pattern in firm_patterns:
        firms.extend(re.findall(pattern, content))

    # 查找文号（审计报告文号）
    report_number_pattern = r"审计报告\s*(?:文号|编号)?[：:]\s*([A-Z0-9\-]+)"
    report_numbers = re.findall(report_number_pattern, content)

    if not firms:
        results.append(
            create_result(
                check_type="格式",
                rule_name="事务所一致性",
                severity="info",
                passed=True,
                description="未找到事务所信息",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
        return

    unique_firms = set(firms)
    if len(unique_firms) > 1:
        results.append(
            create_result(
                check_type="格式",
                rule_name="事务所一致性",
                severity="warning",
                passed=False,
                description=f"事务所名不一致: {', '.join(list(unique_firms)[:3])}{'...' if len(unique_firms) > 3 else ''}",
                source_location="report.md",
                evidence="正则匹配",
            )
        )
    else:
        results.append(
            create_result(
                check_type="格式",
                rule_name="事务所一致性",
                severity="info",
                passed=True,
                description=f"事务所名一致: {list(unique_firms)[0]}",
                source_location="report.md",
                evidence="正则匹配",
            )
        )

    if report_numbers:
        unique_numbers = set(report_numbers)
        if len(unique_numbers) > 1:
            results.append(
                create_result(
                    check_type="格式",
                    rule_name="文号一致性",
                    severity="warning",
                    passed=False,
                    description=f"文号不一致: {', '.join(list(unique_numbers)[:3])}{'...' if len(unique_numbers) > 3 else ''}",
                    source_location="report.md",
                    evidence="正则匹配",
                )
            )
        else:
            results.append(
                create_result(
                    check_type="格式",
                    rule_name="文号一致性",
                    severity="info",
                    passed=True,
                    description=f"文号一致: {list(unique_numbers)[0]}",
                    source_location="report.md",
                    evidence="正则匹配",
                )
            )


# ──────────────────────────────────────────────────────────────────
# Scan 功能：代码分担粗活（竖加/横加验算）
# ──────────────────────────────────────────────────────────────────
# 41科目列表（用于表注定位）
SUBJECTS_LIST = [
    "货币资金", "应收票据", "应收账款", "预付款项", "其他应收款", "存货", "合同资产",
    "持有待售资产", "一年内到期非流动资产", "其他流动资产", "长期股权投资",
    "其他权益工具投资", "其他非流动金融资产", "投资性房地产", "固定资产", "在建工程",
    "使用权资产", "无形资产", "商誉", "长期待摊费用", "递延所得税资产",
    "其他非流动资产", "短期借款", "应付账款", "预收款项", "合同负债",
    "应付职工薪酬", "应交税费", "其他应付款", "一年内到期非流动负债",
    "其他流动负债", "长期借款", "租赁负债", "长期应付款", "预计负债", "递延收益",
    "股本", "资本公积", "盈余公积", "未分配利润", "营业收入"
]

# 四表特征词（扩大范围以涵盖更多四表模式）
FOUR_STATEMENTS_KEYWORDS = [
    "资产总计", "负债合计", "所有者权益合计", "营业收入", "营业利润",
    "流动资产合计", "非流动资产合计", "流动负债合计", "非流动负债合计",
    "现金流量", "利润总额", "净利润"
]


def scan_tables(tables: list) -> dict:
    """扫描 extracted_tables 的所有 table，分类：
    - 四表：含特征词（跳过，L1已处理）
    - 附注明细：有金额列、行数≥2
    - 标题表：无金额或行数<2（跳过）

    返回：{
        "four_statements": [...],
        "note_detail": [...],
        "title": [...]
    }
    """
    categories = {
        "four_statements": [],
        "note_detail": [],
        "title": []
    }

    for table in tables:
        # 检查是否为四表
        is_four_statement = False
        all_text = " ".join(table.get("headers", []))
        for row in table.get("rows", []):
            all_text += " " + " ".join(row)

        for keyword in FOUR_STATEMENTS_KEYWORDS:
            if keyword in all_text:
                is_four_statement = True
                break

        if is_four_statement:
            categories["four_statements"].append(table)
            continue

        # 检查是否有金额列和行数
        rows = table.get("rows", [])
        has_amount = False
        if rows:
            # 检查是否有金额（数字、千分位逗号、负号等）
            for row in rows:
                row_text = " ".join(row)
                if re.search(r'[-\d,]+\.\d+|[-\d,]+', row_text):
                    has_amount = True
                    break

        if has_amount and len(rows) >= 2:
            categories["note_detail"].append(table)
        else:
            categories["title"].append(table)

    return categories


def _is_multicurrency_table(table: dict) -> bool:
    """检测多币种明细表（如货币资金-美元/日元/林吉特）。

    这类表的明细按币种拆分，合计行通常是人民币合计，
    明细求和≠合计，竖加/横加会系统性误报，应跳过。
    """
    headers = table.get("headers", [])
    rows = table.get("rows", [])
    currencies = ["美元", "人民币", "欧元", "日元", "英镑", "港币", "林吉特", "瑞士法郎",
                  "加元", "澳元", "韩元", "新台币", "新加坡元", "泰铢", "越南盾", "卢布"]
    text = " ".join(str(h) for h in headers)
    for row in rows[:30]:
        text += " " + " ".join(str(c) for c in row)
    found = [c for c in currencies if c in text]
    # 至少出现两种币种，且存在"合计"行
    if len(found) >= 2 and any("合计" in str(r[0]) for r in rows if r and len(r) > 0):
        return True
    return False


def code_vertical_check(table: dict) -> list:
    """竖加代码验算（D1）：含"合计/小计/总计"行的附注明细表。
    - 找合计行（项目名列含"合计/小计/总计"）
    - 明细行 = 项目名非合计、非"其中:"开头、该行有数字
    - "减:"开头行：数字取负
    - 用 calculator sum 明细数字，check_equal 和合计行（容差0.01）
    - 返回 results 条目
    - 跳过四表（L1已处理）
    - **修复误报**：只对金额列竖加（跳过百分比列）、排除sub_detail行、合计行空值跳过
    """
    results = []
    headers = table.get("headers", [])
    rows = table.get("rows", [])
    page = table.get("page")
    table_id = table.get("id")

    if not rows:
        return results

    # 跳过多币种明细表（明细和≠人民币合计）
    if _is_multicurrency_table(table):
        return results

    # 检查是否为四表（跳过）
    all_text = " ".join(headers)
    for row in rows:
        all_text += " " + " ".join(row)
    for keyword in FOUR_STATEMENTS_KEYWORDS:
        if keyword in all_text:
            return results

    # 找项目名列（通常是第一列）
    item_col = 0

    # **修复0：处理多行表头** - 合并表头行，确保百分比关键词能被检测到
    merged_headers = headers.copy()
    header_rows_to_skip = 0

    # 检查前两行是否都是表头
    if rows and len(rows) >= 2:
        first_row = rows[0]
        second_row = rows[1]

        # 检查第一行是否是表头行
        is_first_header = False
        if len(first_row) == len(headers):
            # 第一列重复或包含表头关键词
            if first_row and first_row[0] and (first_row[0] == headers[0] or headers[0] in str(first_row[0])):
                is_first_header = True
            elif any(kw in " ".join(first_row).lower() for kw in ["金额", "比例", "余额", "价值", "准备", "类别"]):
                is_first_header = True

        # 检查第二行是否是表头行
        is_second_header = False
        if len(second_row) == len(headers):
            # 包含表头关键词（如'金额'|'比例'）
            if any(kw in " ".join(second_row).lower() for kw in ["金额", "比例", "余额", "价值", "准备", "类别"]):
                is_second_header = True

        # 合并表头（优先合并两行，如果第二行是表头）
        if is_second_header:
            # 合并第一行和第二行
            merged_headers = [
                f"{headers[i]} {first_row[i]} {second_row[i]}" if i < len(first_row) and i < len(second_row)
                else f"{headers[i]} {first_row[i]}" if i < len(first_row)
                else headers[i]
                for i in range(len(headers))
            ]
            header_rows_to_skip = 2
        elif is_first_header:
            # 只合并第一行
            merged_headers = [f"{h} {first_row[i]}" if i < len(first_row) else h for i, h in enumerate(headers)]
            header_rows_to_skip = 1

    # 跳过表头行，只保留数据行
    if header_rows_to_skip > 0:
        rows = rows[header_rows_to_skip:]

    # **修复2：找金额列（跳过百分比列）** - 保留列名特征（%等关键词），不依赖启发式数值范围判断
    amount_cols = []
    percent_keywords = ["%", "比例", "百分比", "rate", "损失率", "计提比例", "占比"]
    for col_idx, header in enumerate(merged_headers):
        if col_idx == item_col:
            continue
        header_lower = str(header).lower()
        # 检查列名是否为百分比列（保留明显特征）
        is_percent_col = any(kw in header_lower for kw in percent_keywords)
        if is_percent_col:
            continue  # 跳过百分比列

        # 检查该列是否有数字
        has_number = False
        for row in rows:
            if col_idx < len(row) and row[col_idx]:
                if re.search(r'[-\d,]+\.\d+|[-\d,]+', row[col_idx]):
                    has_number = True
                    break
        if has_number:
            amount_cols.append(col_idx)

    if not amount_cols:
        return results

    # 找所有合计/总计行（总合计）和小计行（中间汇总），分别记录
    total_row_indices = []
    subtotal_row_indices = []
    for idx, row in enumerate(rows):
        if len(row) > item_col and row[item_col]:
            item_name = row[item_col]
            if "小计" in item_name:
                subtotal_row_indices.append(idx)
            elif any(k in item_name for k in ["合计", "总计"]):
                total_row_indices.append(idx)

    if not total_row_indices and not subtotal_row_indices:
        return results

    # 对比目标：优先最后一个总合计，否则最后一个小计
    if total_row_indices:
        total_row_idx = total_row_indices[-1]
    else:
        total_row_idx = subtotal_row_indices[-1]
    total_item_name = rows[total_row_idx][item_col] if len(rows[total_row_idx]) > item_col else "合计"

    # 所有需排除的汇总行（合计+小计+总计），避免被当明细重复加
    all_summary_indices = set(total_row_indices + subtotal_row_indices)

    # 提取表名（从首行或 headers）
    table_name = table.get("name") or (headers[0] if headers else "")
    if not table_name and rows:
        table_name = " ".join(rows[0]) if rows[0] else ""

    # 对每个金额列进行竖加验算
    for amount_col in amount_cols:
        detail_numbers = []
        for idx, row in enumerate(rows):
            # 排除所有合计/小计/总计行
            if idx in all_summary_indices:
                continue
            if len(row) <= item_col:
                continue
            item_name = row[item_col] if len(row) > item_col else ""

            if len(row) > amount_col and row[amount_col]:
                try:
                    value = parse_number(row[amount_col])
                    # "减:"/"减："开头行，数字取负
                    if item_name and "减" in item_name and (":" in item_name or "：" in item_name):
                        value = -abs(value)
                    detail_numbers.append(value)
                except (ValueError, ZeroDivisionError):
                    pass

        if not detail_numbers:
            continue

        # 计算 sum
        calculated_sum = sum(detail_numbers)

        # 获取合计行的值
        total_value = None
        total_cell = ""
        if total_row_idx < len(rows) and len(rows[total_row_idx]) > amount_col:
            total_cell = rows[total_row_idx][amount_col] if len(rows[total_row_idx]) > amount_col else ""
            if total_cell:
                try:
                    total_value = parse_number(total_cell)
                except (ValueError, ZeroDivisionError):
                    pass

        # **修复4：合计行空值跳过**（空、0、横线等无效值）
        if total_value is None:
            continue
        # 检查合计行是否为空值或占位符（横线、破折号等）
        if total_cell and str(total_cell).strip() in ["-", "—", "–", "——", "", "0", "0.00"]:
            # 合计行无有效数字，无法验算，跳过该列
            continue

        # 检查是否相等（容差0.01）
        check_result = check_equal(str(calculated_sum), str(total_value), tolerance=0.01)
        passed = check_result["match"]
        diff = check_result["diff"]
        severity = "info" if passed else "warning"

        # 确定 rule_name
        rule_name = f"{table_name if table_name else '表格'}竖加"
        if total_item_name and total_item_name != "合计":
            rule_name = f"{total_item_name}竖加"

        results.append(
            create_result(
                check_type="竖加",
                rule_name=rule_name,
                severity=severity,
                passed=passed,
                description=f"{rule_name}: 明细行求和 = {calculated_sum:.2f}, 合计行 = {total_value:.2f}" +
                            (f", 差异 = {diff:.2f}" if diff > 0.01 else ""),
                expected=f"{calculated_sum:.2f}",
                actual=f"{total_value:.2f}",
                difference=diff if diff > 0.01 else 0.0,
                source_location=f"表格ID {table_id}",
                page=page,
                context=f"{table_name} - {total_item_name}",
                evidence="code_vertical_scan_fix"  # 标记为修复版
            )
        )

    # 通用：result 继承 table 的 source_file + chapter（Word 章节定位）
    for r in results:
        r.setdefault("source_file", table.get("source_file", ""))
        r.setdefault("chapter", table.get("chapter", ""))
    return results


def code_horizontal_check(table: dict) -> list:
    """横加代码验算（D2）：已知横加模式代码验算。
    - 模式A：账面余额+坏账准备（或减值准备）+账面价值 → 余额-准备=价值
    - 模式B：期初余额+本期增加+本期减少+期末余额 → 期初+增加-减少=期末
    - 模式C：原价+累计折旧+减值准备+账面价值 → 原价-折旧-减值=价值
    - 列名匹配（包含匹配，容忍变体）
    """
    results = []

    # 跳过多币种明细表
    if _is_multicurrency_table(table):
        return results

    headers = table.get("headers", [])
    rows = table.get("rows", [])
    page = table.get("page")
    table_id = table.get("id")

    # 检查是否为四表（跳过）
    all_text = " ".join(headers)
    for row in rows:
        all_text += " " + " ".join(row)
    for keyword in FOUR_STATEMENTS_KEYWORDS:
        if keyword in all_text:
            return results

    if len(headers) < 3 or not rows:
        return results

    # 转换 headers 为小写用于匹配
    headers_lower = [h.lower() for h in headers]

    # 模式A：账面余额+坏账准备+账面价值
    pattern_a = {
        "name": "模式A",
        "balance": None,
        "provision": None,
        "net": None,
        "check": lambda b, p, n: b - p - n  # 余额 - 准备 = 价值
    }

    # 模式B：期初余额+本期增加+本期减少+期末余额
    pattern_b = {
        "name": "模式B",
        "beginning": None,
        "increase": None,
        "decrease": None,
        "ending": None,
        "check": lambda beg, inc, dec, end: beg + inc - dec - end  # 期初 + 增加 - 减少 = 期末
    }

    # 模式C：原价+累计折旧+减值准备+账面价值
    pattern_c = {
        "name": "模式C",
        "original": None,
        "depreciation": None,
        "impairment": None,
        "net": None,
        "check": lambda orig, dep, imp, net: orig - dep - imp - net  # 原价 - 折旧 - 减值 = 价值
    }

    # 识别列
    for idx, h_lower in enumerate(headers_lower):
        # 模式A识别
        if "账面余额" in h_lower or "余额" in h_lower:
            pattern_a["balance"] = idx
        elif ("坏账准备" in h_lower or "减值准备" in h_lower or
              "准备" in h_lower and "价值" not in h_lower):
            pattern_a["provision"] = idx
        elif "账面价值" in h_lower or "价值" in h_lower:
            pattern_a["net"] = idx

        # 模式B识别
        elif "期初" in h_lower:
            pattern_b["beginning"] = idx
        elif "增加" in h_lower:
            pattern_b["increase"] = idx
        elif "减少" in h_lower:
            pattern_b["decrease"] = idx
        elif "期末" in h_lower:
            pattern_b["ending"] = idx

        # 模式C识别
        elif "原价" in h_lower or "原值" in h_lower:
            pattern_c["original"] = idx
        elif "累计折旧" in h_lower or "折旧" in h_lower:
            pattern_c["depreciation"] = idx
        elif "减值" in h_lower:
            pattern_c["impairment"] = idx
        elif "账面价值" in h_lower or "净值" in h_lower:
            pattern_c["net"] = idx

    # 验证哪个模式匹配
    matched_pattern = None
    if (pattern_a["balance"] is not None and
        pattern_a["provision"] is not None and
        pattern_a["net"] is not None):
        matched_pattern = pattern_a

    elif (pattern_b["beginning"] is not None and
          pattern_b["increase"] is not None and
          pattern_b["decrease"] is not None and
          pattern_b["ending"] is not None):
        matched_pattern = pattern_b

    elif (pattern_c["original"] is not None and
          pattern_c["depreciation"] is not None and
          pattern_c["impairment"] is not None and
          pattern_c["net"] is not None):
        matched_pattern = pattern_c

    if not matched_pattern:
        return results

    # 对每行进行验算
    table_name = table.get("name") or (headers[0] if headers else "")
    for row in rows:
        if not row:
            continue

        try:
            if matched_pattern["name"] == "模式A":
                b_idx, p_idx, n_idx = matched_pattern["balance"], matched_pattern["provision"], matched_pattern["net"]
                if len(row) <= max(b_idx, p_idx, n_idx):
                    continue

                balance = parse_number(row[b_idx]) if row[b_idx] else 0.0
                provision = parse_number(row[p_idx]) if row[p_idx] else 0.0
                net = parse_number(row[n_idx]) if row[n_idx] else 0.0

                diff = matched_pattern["check"](balance, provision, net)
                passed = abs(diff) <= 0.01
                severity = "info" if passed else "warning"
                item_name = row[0] if row[0] else "未命名行"

                results.append(
                    create_result(
                        check_type="横加",
                        rule_name=f"{table_name}横加",
                        severity=severity,
                        passed=passed,
                        description=f"{item_name}: 余额({balance:.2f}) - 准备({provision:.2f}) = 价值({net:.2f})" +
                                    (f", 差异={diff:.2f}" if abs(diff) > 0.01 else ""),
                        expected=f"{balance - provision:.2f}",
                        actual=f"{net:.2f}",
                        difference=abs(diff),
                        source_location=f"表格ID {table_id}",
                        page=page,
                        context=f"{table_name} - {item_name}",
                        evidence="code_scan_模式A"
                    )
                )

            elif matched_pattern["name"] == "模式B":
                beg_idx, inc_idx, dec_idx, end_idx = (matched_pattern["beginning"],
                                                      matched_pattern["increase"],
                                                      matched_pattern["decrease"],
                                                      matched_pattern["ending"])
                if len(row) <= max(beg_idx, inc_idx, dec_idx, end_idx):
                    continue

                beginning = parse_number(row[beg_idx]) if row[beg_idx] else 0.0
                increase = parse_number(row[inc_idx]) if row[inc_idx] else 0.0
                decrease = parse_number(row[dec_idx]) if row[dec_idx] else 0.0
                ending = parse_number(row[end_idx]) if row[end_idx] else 0.0

                diff = matched_pattern["check"](beginning, increase, decrease, ending)
                passed = abs(diff) <= 0.01
                severity = "info" if passed else "warning"
                item_name = row[0] if row[0] else "未命名行"

                results.append(
                    create_result(
                        check_type="横加",
                        rule_name=f"{table_name}横加",
                        severity=severity,
                        passed=passed,
                        description=f"{item_name}: 期初({beginning:.2f}) + 增加({increase:.2f}) - 减少({decrease:.2f}) = 期末({ending:.2f})" +
                                    (f", 差异={diff:.2f}" if abs(diff) > 0.01 else ""),
                        expected=f"{beginning + increase - decrease:.2f}",
                        actual=f"{ending:.2f}",
                        difference=abs(diff),
                        source_location=f"表格ID {table_id}",
                        page=page,
                        context=f"{table_name} - {item_name}",
                        evidence="code_scan_模式B"
                    )
                )

            elif matched_pattern["name"] == "模式C":
                orig_idx, dep_idx, imp_idx, net_idx = (matched_pattern["original"],
                                                       matched_pattern["depreciation"],
                                                       matched_pattern["impairment"],
                                                       matched_pattern["net"])
                if len(row) <= max(orig_idx, dep_idx, imp_idx, net_idx):
                    continue

                original = parse_number(row[orig_idx]) if row[orig_idx] else 0.0
                depreciation = parse_number(row[dep_idx]) if row[dep_idx] else 0.0
                impairment = parse_number(row[imp_idx]) if row[imp_idx] else 0.0
                net = parse_number(row[net_idx]) if row[net_idx] else 0.0

                diff = matched_pattern["check"](original, depreciation, impairment, net)
                passed = abs(diff) <= 0.01
                severity = "info" if passed else "warning"
                item_name = row[0] if row[0] else "未命名行"

                results.append(
                    create_result(
                        check_type="横加",
                        rule_name=f"{table_name}横加",
                        severity=severity,
                        passed=passed,
                        description=f"{item_name}: 原价({original:.2f}) - 折旧({depreciation:.2f}) - 减值({impairment:.2f}) = 价值({net:.2f})" +
                                    (f", 差异={diff:.2f}" if abs(diff) > 0.01 else ""),
                        expected=f"{original - depreciation - impairment:.2f}",
                        actual=f"{net:.2f}",
                        difference=abs(diff),
                        source_location=f"表格ID {table_id}",
                        page=page,
                        context=f"{table_name} - {item_name}",
                        evidence="code_scan_模式C"
                    )
                )

        except (ValueError, ZeroDivisionError):
            continue

    # 通用：result 继承 table 的 source_file + chapter（Word 章节定位）
    for r in results:
        r.setdefault("source_file", table.get("source_file", ""))
        r.setdefault("chapter", table.get("chapter", ""))
    return results


def scan_candidates(tables: list) -> list:
    """候选预筛（D3）：非四表、有金额列、未匹配已知横加模式的附注明细表。
    - 识别变动表（列名含期初/期末/增加/减少，但非标准模式B完整匹配）
    - 其余有金额列的表 → 复杂横加候选
    - 输出 candidates.json
    """
    candidates = []

    # 先分类表
    categories = scan_tables(tables)
    note_detail_tables = categories["note_detail"]

    for table in note_detail_tables:
        # 跳过多币种明细表
        if _is_multicurrency_table(table):
            continue

        headers = table.get("headers", [])
        rows = table.get("rows", [])
        page = table.get("page")
        table_id = table.get("id")
        table_name = table.get("name") or (headers[0] if headers else "")

        # 检查是否匹配已知横加模式
        headers_lower = [h.lower() for h in headers]

        # 检查是否为变动表
        has_beginning = any("期初" in h for h in headers_lower)
        has_ending = any("期末" in h for h in headers_lower)
        has_increase = any("增加" in h for h in headers_lower)
        has_decrease = any("减少" in h for h in headers_lower)

        # 标准模式B：四列都有
        is_standard_pattern_b = (has_beginning and has_ending and has_increase and has_decrease)

        # 变动表候选：有期初/期末/增加/减少但不完整
        is_variation_table = False
        if (has_beginning or has_ending or has_increase or has_decrease) and not is_standard_pattern_b:
            is_variation_table = True

        if is_variation_table:
            candidates.append({
                "id": table_id,
                "page": page,
                "table_name": table_name,
                "reason": "变动表",
                "columns": headers,
                "rows_count": len(rows)
            })
        else:
            # 复杂横加候选
            candidates.append({
                "id": table_id,
                "page": page,
                "table_name": table_name,
                "reason": "复杂横加",
                "columns": headers,
                "rows_count": len(rows)
            })

    return candidates


def build_subjects_index(tables: list) -> dict:
    """表注定位（D4）：按41科目在 extracted_tables 预定位。
    - 对每科目，找 table 中项目名（首列值）或表名包含科目名的表
    - 输出 subjects_index.json
    """
    subjects_index = {subject: [] for subject in SUBJECTS_LIST}

    for table in tables:
        table_id = table.get("id")
        page = table.get("page")
        headers = table.get("headers", [])
        rows = table.get("rows", [])

        # 获取表名
        table_name = table.get("name") or (headers[0] if headers else "")

        for subject in SUBJECTS_LIST:
            # 检查表名是否包含科目名
            if subject in table_name:
                subjects_index[subject].append({
                    "table_id": table_id,
                    "page": page,
                    "table_name": table_name
                })
                continue

            # 检查行项目名是否包含科目名
            for row in rows:
                if row and row[0] and subject in row[0]:
                    subjects_index[subject].append({
                        "table_id": table_id,
                        "page": page,
                        "table_name": table_name
                    })
                    break  # 一张表对同一科目只记录一次

    # 移除空科目
    subjects_index = {k: v for k, v in subjects_index.items() if v}

    return subjects_index


# ──────────────────────────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────────────────────────
def _merge_located_values(located: list) -> dict:
    """合并 located（科目多张表的取数结果）的 values，非空优先。"""
    merged = {}
    for item in located:
        for k, v in (item.get("values") or {}).items():
            if v not in (None, "", "-", "—"):
                merged.setdefault(k, v)
    return merged


def _pick_value(merged: dict, keys: list) -> Optional[str]:
    """按候选 key 列表取第一个非空值（精确优先，再包含匹配）。"""
    for k in keys:
        if k in merged and merged[k] not in (None, "", "-", "—"):
            return merged[k]
    for mk, mv in merged.items():
        for k in keys:
            if k in mk and mv not in (None, "", "-", "—"):
                return mv
    return None


def _resolve_note_value(merged: dict, field: str, formula: str) -> Optional[str]:
    """从附注取数 values 解析目标值。

    若 formula 含运算符（如"账面余额-坏账准备"），优先按 formula 运算——
    比 field 直接取更可靠（避免 DeepSeek 把明细/账面余额当账面价值）。
    formula 无运算（如"期末余额"）或运算失败时，退回 field 直接取。
    """
    # 1. 若 formula 含运算符，优先运算（用合计数相减，比 DeepSeek 直取账面价值可靠）
    expr = ""
    if formula:
        expr = formula.split("=")[0].strip() if "=" in formula else formula.strip()
    has_op = bool(expr) and any(op in expr for op in ["+", "-"])
    if has_op:
        calc_expr = expr
        for k in sorted(merged.keys(), key=len, reverse=True):
            v = merged[k]
            if v in (None, "", "-", "—"):
                continue
            try:
                calc_expr = calc_expr.replace(k, f"({parse_number(str(v))})")
            except Exception:
                continue
        if not re.search(r"[\u4e00-\u9fa5]", calc_expr):
            try:
                return f"{evaluate_expression(calc_expr):.2f}"
            except Exception:
                pass
    # 2. field 直接取（formula 无运算或运算失败）
    if field:
        v = _pick_value(merged, [field])
        if v is not None:
            return v
        # field 同义词兜底：年末余额/期末余额/账面余额/账面价值/金额/本年发生额 互换
        FIELD_SYNONYMS = {
            "年末余额": ["期末余额", "账面余额", "账面价值", "金额", "余额"],
            "期末余额": ["年末余额", "账面余额", "账面价值", "金额", "余额"],
            "账面价值": ["账面余额", "年末余额", "期末余额"],
            "账面余额": ["账面价值", "年末余额", "期末余额"],
            "本年发生额": ["发生额", "本年金额", "金额", "本期金额"],
            "本期金额": ["本年发生额", "发生额", "金额"],
        }
        for syn in FIELD_SYNONYMS.get(field, []):
            v = _pick_value(merged, [syn])
            if v is not None:
                return v
    return None


def check_statement_to_note(
    statements: dict, note_map: dict, located_values: dict, results: list
) -> None:
    """表注勾稽：报表数 vs 附注数（层3算术）。

    定位（note_map）由 Claude Step2 语义生成，取数（located_values）由 DeepSeek locate 完成；
    本函数只做确定性算术：按科目 field/formula 取附注数，与报表 statements 比较。
    """
    if not statements or not note_map:
        return

    # 科目 → 报表 current 值（按报表名嵌套，取值优先合并报表，避免母公司同名覆盖）
    flat_by_stmt = {}
    for stmt_name, stmt in statements.items():
        if not isinstance(stmt, dict):
            continue
        for item, vals in stmt.items():
            if isinstance(vals, dict) and vals.get("current") not in (None, ""):
                flat_by_stmt.setdefault(item, {})[stmt_name] = vals["current"]

    def _pick_report_val(item):
        """优先取合并报表值（含'合并'的表名），避免母公司覆盖。"""
        candidates = flat_by_stmt.get(item, {})
        if not candidates:
            return None
        for sn in candidates:
            if "合并" in sn:
                return candidates[sn]
        return next(iter(candidates.values()))

    check_count = 0
    pass_count = 0
    for subject, info in note_map.items():
        if not isinstance(info, dict):
            continue
        field = info.get("field", "")
        formula = info.get("formula", "")
        pages = info.get("pages", [])
        chapter = info.get("chapter")
        page = str(pages[0]) if pages else chapter  # Word 无 pages 用 chapter 章节定位
        located = located_values.get(subject, [])
        merged = _merge_located_values(located)

        report_val = _pick_report_val(subject)
        if report_val is None:
            results.append(create_result(
                check_type="表注", rule_name=f"{subject}表注勾稽",
                severity="info", passed=True,
                description=f"{subject}：报表未列示该科目，跳过表注勾稽",
                page=page, context=subject, evidence="报表无对应项目"))
            continue

        note_val = _resolve_note_value(merged, field, formula)
        check_count += 1
        if note_val is None:
            results.append(create_result(
                check_type="表注", rule_name=f"{subject}表注勾稽",
                severity="warning", passed=False,
                description=f"{subject}：附注未能提取目标值（field={field}, formula={formula}），需人工核对",
                source_location=f"报表 {report_val}", target_location="附注(取数失败)",
                page=page, context=subject,
                evidence="Agent复核：存疑（取数失败，Claude Step5手动）"))
            continue

        chk = check_equal(str(report_val), str(note_val))
        desc = f"{subject}：报表 {report_val} vs 附注 {note_val}"
        if chk.get("match"):
            pass_count += 1
            results.append(create_result(
                check_type="表注", rule_name=f"{subject}表注勾稽",
                severity="info", passed=True, description=desc + "（一致）",
                expected=str(report_val), actual=str(note_val), difference=chk.get("diff"),
                page=page, context=f"{subject} 报表数 vs 附注数", evidence="Agent复核：确认一致"))
        else:
            diff = chk.get("diff")
            diff_str = f"{diff:.2f}" if diff is not None else "未知"
            results.append(create_result(
                check_type="表注", rule_name=f"{subject}表注勾稽",
                severity="warning", passed=False,
                description=desc + f"（差异 {diff_str}），需人工复核",
                expected=str(report_val), actual=str(note_val), difference=diff,
                page=page, context=f"{subject} 报表数 vs 附注数",
                evidence="Agent复核：存疑（数值不一致）"))

    print(f"[INFO] 表注勾稽: 检查 {check_count} 科目，通过 {pass_count}", file=sys.stderr)


def check_note_internal_reconcile(
    note_map: dict, located_values: dict, results: list
) -> None:
    """附注内变动表 reconcile：期初 + 增加 - 减少 = 期末（层3算术）。

    复用表注取数的 located_values（_LOCATE_NOTE_VALUES_PROMPT 同时提取期初/增/减/期末）。
    四要素齐全才视为变动表验算。
    """
    if not note_map:
        return

    recon_count = 0
    pass_count = 0
    for subject, info in note_map.items():
        if not isinstance(info, dict):
            continue
        located = located_values.get(subject, [])
        merged = _merge_located_values(located)
        beginning = _pick_value(merged, ["期初余额", "年初余额", "期初"])
        increase = _pick_value(merged, ["本期增加", "本年增加", "增加额", "增加"])
        decrease = _pick_value(merged, ["本期减少", "本年减少", "减少额", "减少"])
        ending = _pick_value(merged, ["期末余额", "年末余额", "期末"])
        if beginning is None or increase is None or decrease is None or ending is None:
            continue  # 非变动表，跳过
        pages = info.get("pages", [])
        chapter = info.get("chapter")
        page = str(pages[0]) if pages else chapter  # Word 无 pages 用 chapter 章节定位
        recon_count += 1
        recon = reconcile(beginning, increase, decrease, ending)
        desc = f"{subject}变动表：期初{beginning}+增加{increase}-减少{decrease}=期末{ending}"
        if recon.get("match"):
            pass_count += 1
            results.append(create_result(
                check_type="附注内", rule_name=f"{subject}变动表reconcile",
                severity="info", passed=True, description=desc + "（平衡）",
                expected=f"{recon.get('calculated_ending'):.2f}", actual=f"{recon.get('actual_ending'):.2f}",
                difference=recon.get("diff"), page=page,
                context=f"{subject} 期初+增-减=期末", evidence="Agent复核：确认平衡"))
        else:
            results.append(create_result(
                check_type="附注内", rule_name=f"{subject}变动表reconcile",
                severity="warning", passed=False,
                description=desc + f"（不平，差异 {recon.get('diff'):.2f}），需人工复核",
                expected=f"{recon.get('calculated_ending'):.2f}", actual=f"{recon.get('actual_ending'):.2f}",
                difference=recon.get("diff"), page=page,
                context=f"{subject} 期初+增-减=期末", evidence="Agent复核：存疑（不平）"))

    if recon_count:
        print(f"[INFO] 附注内reconcile: 检查 {recon_count} 科目，通过 {pass_count}", file=sys.stderr)


# ──────────────────────────────────────────────────────────────────
# AI 结构标注缓存（避免重复调 API）
# ──────────────────────────────────────────────────────────────────
def _table_cache_key(table: dict) -> str:
    """按表内容生成稳定 hash。"""
    content = json.dumps({
        "source_file": table.get("source_file", ""),
        "name": table.get("name", ""),
        "chapter": table.get("chapter", ""),
        "headers": table.get("headers", []),
        "rows": table.get("rows", []),
    }, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(content.encode("utf-8")).hexdigest()[:32]


def _annotation_cache_path(input_dir: Path) -> Path:
    return input_dir / "annotations_cache.json"


def _load_annotation_cache(input_dir: Path) -> dict:
    path = _annotation_cache_path(input_dir)
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_annotation_cache(input_dir: Path, cache: dict):
    path = _annotation_cache_path(input_dir)
    path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def _annotate_tables_with_cache(
    candidates: list,
    client,
    input_dir: Path,
) -> list:
    """带缓存的 annotate_tables：内容未变的表直接读本地缓存，只对新/变表调 DeepSeek。"""
    cache = _load_annotation_cache(input_dir)
    cached_version = cache.get("version", "")
    # 用 ai_worker 模块版本或模型名作为缓存版本，模型变化时失效
    expected_version = f"{getattr(client, 'model', '')}_v1"
    if cached_version != expected_version:
        cache = {"version": expected_version}

    missing = []
    for table in candidates:
        key = _table_cache_key(table)
        if key not in cache:
            missing.append(table)

    if missing:
        print(f"[INFO] AI 标注缓存命中 {len(candidates) - len(missing)}/{len(candidates)} 张，需调用 {len(missing)} 张", file=sys.stderr)
        new_annotations = ai_worker.annotate_tables(missing, client)
        for ann in new_annotations:
            table_id = ann.get("table_id")
            # 找到对应 table 算 key
            for t in candidates:
                if str(t.get("id")) == str(table_id):
                    cache[_table_cache_key(t)] = ann
                    break
    else:
        print(f"[INFO] AI 标注缓存全部命中 {len(candidates)} 张，跳过 DeepSeek 调用", file=sys.stderr)

    _save_annotation_cache(input_dir, cache)

    annotations = []
    for t in candidates:
        key = _table_cache_key(t)
        if key in cache:
            ann = cache[key]
            # 确保 annotation 里有 table_id（缓存可能来自旧格式）
            if "table_id" not in ann:
                ann["table_id"] = t.get("id")
            annotations.append(ann)
    return annotations


def main():
    parser = argparse.ArgumentParser(
        description="审计报告核查编排脚本（代码侧检查）"
    )
    parser.add_argument(
        "input_dir",
        help="parse 输出目录（包含 extracted_tables.json 和 report.md）"
    )
    parser.add_argument(
        "--statements",
        help="statements.json 文件（Claude 预处理的四表数据）"
    )
    parser.add_argument(
        "--note-map",
        help="note_map.json（Claude Step2 生成的科目→附注表定位映射，用于表注勾稽）；默认读 <input_dir>/note_map.json"
    )
    parser.add_argument(
        "--scope",
        choices=["quick", "standard", "deep"],
        default="standard",
        help="检查范围：quick（仅基础）/standard（默认）/deep（深度）"
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="输出文件路径（results.json）"
    )
    parser.add_argument(
        "--scan",
        action="store_true",
        help="启用 scan 模式：代码验算竖加/横加，输出 candidates.json 和 subjects_index.json"
    )
    parser.add_argument(
        "--use-ai",
        action="store_true",
        default=True,
        help="使用 DeepSeek AI 加速（默认启用，需配置 ~/.deepseek/config.json）"
    )
    parser.add_argument(
        "--no-ai",
        action="store_true",
        help="禁用 AI 加速，使用代码规则检查"
    )
    parser.add_argument(
        "--auto-manifest",
        action="store_true",
        help="无 manifest.json 时自动生成（DeepSeek 并发检测 xlsx 结构）"
    )
    parser.add_argument(
        "--auto-note-map",
        action="store_true",
        help="无 note_map.json 时自动生成（DeepSeek 并发 + 回原文校验）"
    )
    parser.add_argument(
        "--api-key",
        help="DeepSeek API Key（临时覆盖配置文件）"
    )
    parser.add_argument(
        "--model",
        help="DeepSeek 模型名（临时覆盖配置文件）"
    )
    parser.add_argument(
        "--skip-annotation",
        action="store_true",
        help="跳过 AI 结构标注（复用本地缓存或上一次标注结果）"
    )
    parser.add_argument(
        "--skip-text-ai",
        action="store_true",
        help="跳过 AI 文本错别字/病句检查"
    )
    parser.add_argument(
        "--skip-ai-review",
        action="store_true",
        help="跳过 DeepSeek 对 warning/error 的预复核（默认执行，留给 Claude Step5 做最终静默终审）"
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=60,
        help="DeepSeek API 并发数（默认 60，受账号限流限制）"
    )
    args = parser.parse_args()

    os.environ["AUDIT_AI_MAX_WORKERS"] = str(args.max_workers)

    # AI 配置加载
    ai_config = None
    use_ai = args.use_ai and not args.no_ai

    if use_ai:
        if not AI_WORKER_AVAILABLE:
            print("[WARN] ai_worker 模块不可用，跳过 AI 加速", file=sys.stderr)
            use_ai = False
        else:
            has_key, config = ai_worker.check_api_key(args.api_key)
            if not has_key:
                print("[INFO] 未配置 DeepSeek，将使用降级模式（Claude主控串行）", file=sys.stderr)
                print("[INFO] 如需加速请配置 ~/.deepseek/config.json 或设置 DEEPSEEK_API_KEY 环境变量", file=sys.stderr)
                use_ai = False
            else:
                ai_config = config
                # 覆盖模型参数
                if args.model:
                    ai_config["model"] = args.model
                print(f"[INFO] DeepSeek AI 加速已启用（模型: {ai_config.get('model')}）", file=sys.stderr)

    input_dir = Path(args.input_dir).resolve()
    if not input_dir.exists():
        print(f"[ERROR] 输入目录不存在: {input_dir}", file=sys.stderr)
        sys.exit(1)

    extracted_tables = input_dir / "extracted_tables.json"
    report_md = input_dir / "report.md"

    # ──────────────────────────────────────────────────────────────────
    # --auto-manifest：自动生成 manifest.json
    # ──────────────────────────────────────────────────────────────────
    if args.auto_manifest:
        manifest_path = input_dir / "manifest.json"
        if not manifest_path.exists():
            print("[INFO] --auto-manifest：自动生成 manifest.json", file=sys.stderr)
            if not use_ai or not AI_WORKER_AVAILABLE:
                print("[ERROR] --auto-manifest 需要 DeepSeek AI，请配置 ~/.deepseek/config.json 或 --api-key", file=sys.stderr)
                sys.exit(1)

            try:
                # 构建 DeepSeek client
                from ai_worker import DeepSeekClient, detect_layout, detect_docx_layout, detect_docx_layout_code
                client = DeepSeekClient(
                    api_key=ai_config["api_key"],
                    model=ai_config["model"],
                    base_url=ai_config["base_url"],
                )

                # 找 xlsx/xlsm 文件：优先 input_dir，其次 files_index.json，最后父目录
                xlsx_files = (
                    list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.XLSX")) +
                    list(input_dir.glob("*.xlsm")) + list(input_dir.glob("*.XLSM"))
                )
                if not xlsx_files:
                    files_index_path = input_dir / "files_index.json"
                    if files_index_path.exists():
                        try:
                            files_index = json.loads(files_index_path.read_text(encoding="utf-8"))
                            xlsx_files = [
                                input_dir / fname
                                for fname, info in files_index.items()
                                if info.get("loader") == "xlsx" and (input_dir / fname).exists()
                            ]
                        except Exception:
                            pass
                if not xlsx_files:
                    xlsx_files = (
                        list(input_dir.parent.glob("*.xlsx")) + list(input_dir.parent.glob("*.XLSX")) +
                        list(input_dir.parent.glob("*.xlsm")) + list(input_dir.parent.glob("*.XLSM"))
                    )

                if xlsx_files:
                    xlsx_path = xlsx_files[0]
                    print(f"[INFO] 使用 xlsx 文件: {xlsx_path.name}", file=sys.stderr)
                    manifest = detect_layout(str(xlsx_path), client)
                else:
                    # 无单独 Excel 报表：尝试从 extracted_tables.json 识别内嵌四表
                    extracted_tables_path = input_dir / "extracted_tables.json"
                    if extracted_tables_path.exists():
                        print("[INFO] 未找到 xlsx/xlsm，尝试代码识别 Word/PDF 内嵌四表...", file=sys.stderr)
                        manifest = detect_docx_layout_code(str(extracted_tables_path))
                        if manifest and manifest.get("statements_map"):
                            print(f"[INFO] 代码识别四表成功: {len(manifest['statements_map'])} 张", file=sys.stderr)
                        else:
                            print("[INFO] 代码识别未命中，改用 DeepSeek 识别...", file=sys.stderr)
                            manifest = detect_docx_layout(str(extracted_tables_path), client)
                        if not manifest or not manifest.get("statements_map"):
                            print("[ERROR] 未能从 extracted_tables.json 自动识别四表", file=sys.stderr)
                            sys.exit(1)
                    else:
                        print("[ERROR] 未找到 xlsx/xlsm 文件，也找不到 extracted_tables.json 用于自动生成 manifest", file=sys.stderr)
                        sys.exit(1)

                # 写入 manifest.json
                manifest_path.write_text(
                    json.dumps(manifest, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )
                print(f"[INFO] manifest.json 已生成: {manifest_path}", file=sys.stderr)

                # 自动调用 apply_manifest 生成 statements.json
                statements_path = input_dir / "statements.json"
                if not statements_path.exists():
                    print("[INFO] 自动调用 apply_manifest 生成 statements.json", file=sys.stderr)
                    try:
                        # 同进程调用 apply_manifest，避免 subprocess 丢失 stdout
                        import apply_manifest
                        orig_argv = sys.argv
                        sys.argv = [
                            str(Path(__file__).parent / "apply_manifest.py"),
                            str(manifest_path),
                            "-o",
                            str(statements_path),
                        ]
                        apply_manifest.main()
                    except SystemExit as se:
                        if se.code not in (0, None):
                            print(f"[WARN] apply_manifest 退出码: {se.code}", file=sys.stderr)
                    except Exception as e:
                        print(f"[WARN] apply_manifest 失败: {e}", file=sys.stderr)
                    finally:
                        sys.argv = orig_argv
                    if statements_path.exists():
                        print(f"[INFO] statements.json 已生成: {statements_path}", file=sys.stderr)

            except Exception as e:
                print(f"[ERROR] 自动生成 manifest 失败: {e}", file=sys.stderr)
                print("[HINT] 可手动编写 manifest.json 或检查 DeepSeek 配置", file=sys.stderr)
                sys.exit(1)
        else:
            print(f"[INFO] manifest.json 已存在，跳过自动生成", file=sys.stderr)

    def _resolve_input_path(path_arg: Optional[str], default_name: str) -> Path:
        """解析用户传入的相对路径：优先相对于 input_dir，其次当前工作目录。

        这避免用户在 parse 输出目录执行命令时，因当前工作目录不同而找不到
        statements.json / note_map.json。
        """
        if not path_arg:
            return input_dir / default_name
        p = Path(path_arg)
        if not p.is_absolute():
            # 相对路径：优先基于 input_dir 解析
            candidate = input_dir / p
            if candidate.exists():
                return candidate.resolve()
            # 其次尝试当前工作目录
            candidate_cwd = Path.cwd() / p
            if candidate_cwd.exists():
                return candidate_cwd.resolve()
            # 都不存在时返回 input_dir 下的候选（后续打印明确提示）
            return candidate.resolve()
        return p.resolve()

    # ──────────────────────────────────────────────────────────────────
    # 读取 statements.json
    # ──────────────────────────────────────────────────────────────────
    statements = {}
    if args.statements:
        statements_path = _resolve_input_path(args.statements, "statements.json")
        if statements_path.exists():
            try:
                statements = json.loads(statements_path.read_text(encoding="utf-8"))
                print(f"[INFO] 加载 statements.json: {statements_path}", file=sys.stderr)
            except Exception as e:
                print(f"[WARN] 读取 statements.json 失败: {e}", file=sys.stderr)
        else:
            print(f"[ERROR] statements.json 不存在: {statements_path}", file=sys.stderr)
            print(f"[HINT] 请确认文件路径，或在 {input_dir} 目录下生成 statements.json", file=sys.stderr)
    else:
        # --auto-manifest 可能已自动生成 statements.json，默认加载
        default_statements_path = input_dir / "statements.json"
        if default_statements_path.exists():
            try:
                statements = json.loads(default_statements_path.read_text(encoding="utf-8"))
                print(f"[INFO] 加载默认 statements.json: {default_statements_path}", file=sys.stderr)
            except Exception as e:
                print(f"[WARN] 读取 statements.json 失败: {e}", file=sys.stderr)

    # ──────────────────────────────────────────────────────────────────
    # --auto-note-map：自动生成 note_map.json
    # ──────────────────────────────────────────────────────────────────
    if args.auto_note_map:
        note_map_path = input_dir / "note_map.json"
        if not note_map_path.exists():
            print("[INFO] --auto-note-map：自动生成 note_map.json", file=sys.stderr)
            if not use_ai or not AI_WORKER_AVAILABLE:
                print("[ERROR] --auto-note-map 需要 DeepSeek AI，请配置 ~/.deepseek/config.json 或 --api-key", file=sys.stderr)
                sys.exit(1)

            if not statements:
                print("[ERROR] --auto-note-map 需要 statements.json（可配合 --auto-manifest 自动生成）", file=sys.stderr)
                sys.exit(1)

            if not extracted_tables.exists():
                print(f"[ERROR] extracted_tables.json 不存在: {extracted_tables}", file=sys.stderr)
                sys.exit(1)

            try:
                # 读取 extracted_tables.json
                tables_data = json.loads(extracted_tables.read_text(encoding="utf-8"))
                tables = tables_data.get("tables", [])

                # 调用 build_note_map
                from ai_worker import DeepSeekClient, build_note_map
                client = DeepSeekClient(
                    api_key=ai_config["api_key"],
                    model=ai_config["model"],
                    base_url=ai_config["base_url"],
                )
                note_map = build_note_map(tables, statements, client)

                # 写入 note_map.json
                note_map_path.write_text(
                    json.dumps(note_map, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )
                print(f"[INFO] note_map.json 已生成: {note_map_path}", file=sys.stderr)

            except Exception as e:
                print(f"[ERROR] 自动生成 note_map 失败: {e}", file=sys.stderr)
                print("[HINT] 可手动编写 note_map.json 或检查 DeepSeek 配置", file=sys.stderr)
                # 不退出，继续运行
        else:
            print(f"[INFO] note_map.json 已存在，跳过自动生成", file=sys.stderr)

    # ──────────────────────────────────────────────────────────────────
    # 读取 note_map.json
    # ──────────────────────────────────────────────────────────────────
    note_map = {}
    note_map_path = None
    if args.note_map:
        note_map_path = _resolve_input_path(args.note_map, "note_map.json")
    else:
        default_nm = input_dir / "note_map.json"
        if default_nm.exists():
            note_map_path = default_nm
    if note_map_path:
        if note_map_path.exists():
            try:
                note_map = json.loads(note_map_path.read_text(encoding="utf-8"))
                print(f"[INFO] 加载 note_map: {len(note_map)} 个科目 → {note_map_path}", file=sys.stderr)
            except Exception as e:
                print(f"[WARN] 读取 note_map.json 失败: {e}", file=sys.stderr)
        else:
            print(f"[ERROR] note_map.json 不存在: {note_map_path}", file=sys.stderr)
            print(f"[HINT] 请确认文件路径，或在 {input_dir} 目录下生成 note_map.json", file=sys.stderr)

    # 执行检查
    results = []

    # 必做（所有 scope）
    print("[INFO] 执行 L1 报表间勾稽检查...", file=sys.stderr)
    run_l1_checks(statements, results, extracted_tables)

    print("[INFO] 执行页码连续性检查...", file=sys.stderr)
    check_page_continuity(report_md, results)

    print("[INFO] 执行金额单位一致性检查...", file=sys.stderr)
    check_unit_consistency(report_md, results)

    print("[INFO] 执行错别字检查...", file=sys.stderr)
    check_typos(report_md, results)

    # standard/deep 额外
    if args.scope in ["standard", "deep"]:
        print("[INFO] 执行公司名称一致性检查...", file=sys.stderr)
        check_company_consistency(extracted_tables, report_md, results)

        print("[INFO] 执行附注编号连续性检查...", file=sys.stderr)
        check_note_number_continuity(report_md, results)

        print("[INFO] 执行事务所/文号一致性检查...", file=sys.stderr)
        check_firm_consistency(report_md, results)

    # deep 额外（深度检查的语义部分由 Claude 主控做，本脚本不额外加）

    # Scan 模式：代码分担粗活（竖加/横加验算）
    tables_data = None
    if args.scan and extracted_tables.exists():
        print("[INFO] 启用 scan 模式，执行代码验算...", file=sys.stderr)

        # 提前确定输出路径
        output_path = Path(args.output).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            tables_data = json.loads(extracted_tables.read_text(encoding="utf-8"))
            tables = tables_data.get("tables", [])

            # 表分类
            categories = scan_tables(tables)
            print(f"[INFO] 表分类: 四表 {len(categories['four_statements'])}，附注明细 {len(categories['note_detail'])}，标题表 {len(categories['title'])}", file=sys.stderr)

            # 对附注明细表执行竖加验算
            vertical_count = 0
            for table in categories["note_detail"]:
                vertical_results = code_vertical_check(table)
                if vertical_results:
                    results.extend(vertical_results)
                    vertical_count += 1
            print(f"[INFO] 竖加验算覆盖: {vertical_count} 表", file=sys.stderr)

            # 对附注明细表执行横加验算
            horizontal_count = 0
            for table in categories["note_detail"]:
                horizontal_results = code_horizontal_check(table)
                if horizontal_results:
                    results.extend(horizontal_results)
                    horizontal_count += 1
            print(f"[INFO] 横加验算覆盖: {horizontal_count} 表", file=sys.stderr)

            # 生成候选预筛（candidates.json）
            candidates = scan_candidates(tables)
            candidates_path = output_path.parent / "candidates.json"
            candidates_path.write_text(
                json.dumps(candidates, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            print(f"[INFO] 候选预筛: {len(candidates)} 表 → {candidates_path}", file=sys.stderr)

            # AI 结构标注（如果启用）
            if use_ai and ai_config and not args.skip_annotation:
                print("[INFO] DeepSeek AI 并发标注表格结构...", file=sys.stderr)
                try:
                    # 构建 DeepSeek client
                    client = ai_worker.DeepSeekClient(
                        api_key=ai_config["api_key"],
                        model=ai_config["model"],
                        base_url=ai_config["base_url"],
                    )

                    # scan_candidates 输出轻量摘要（只有 rows_count 无 rows/headers），annotate 需完整表 → 按 id 从 tables 补全
                    _candidate_ids = {c.get("id") for c in candidates}
                    _full_candidates = [t for t in tables if t.get("id") in _candidate_ids]
                    annotations = _annotate_tables_with_cache(_full_candidates, client, input_dir)
                    print(f"[INFO] AI 标注完成: {len(annotations)} 表", file=sys.stderr)

                    # 按 AI 标注进行竖加验算
                    ai_vertical_count = 0
                    for annotation in annotations:
                        table_id = annotation.get("table_id")
                        ann_data = annotation.get("annotation", {})

                        # 找到对应表格（table_id 可能是字符串/整数，统一比较）
                        table = None
                        for tbl in tables:
                            if str(tbl.get("id")) == str(table_id):
                                table = tbl
                                break

                        if not table:
                            continue

                        # 按 AI 标注验算
                        rows_annotation = ann_data.get("rows", [])
                        columns_annotation = ann_data.get("columns", [])
                        horizontal_annotation = ann_data.get("horizontal", [])

                        # 竖加验算：按标注找 total 行和 detail 行
                        total_rows = [r for r in rows_annotation if r.get("type") in ("total", "subtotal", "sum", "合计", "小计")]
                        # detail 行排除 is_subitem（其中子项不重复计入合计）和 sub_detail 类型
                        detail_rows = [r for r in rows_annotation if r.get("type") == "detail" and not r.get("is_subitem")]

                        for total_row in total_rows:
                            total_row_idx = total_row.get("row")
                            rows_data = table.get("rows", [])

                            if total_row_idx >= len(rows_data):
                                continue

                            # 对 AI 标注的金额列验算（排除百分比列，百分比不竖加）
                            amount_cols = [c for c in columns_annotation if c.get("type") == "amount" or (c.get("type") not in ("percentage", "percent", "ratio") and "amount" in str(c.get("type", "")).lower())]

                            for amount_col in amount_cols:
                                col_idx = amount_col.get("col")
                                headers = table.get("headers", [])
                                if col_idx >= len(headers):
                                    continue

                                # 计算 detail 行求和
                                detail_values = []
                                for detail_row in detail_rows:
                                    detail_row_idx = detail_row.get("row")
                                    if detail_row_idx >= len(rows_data):
                                        continue
                                    row_data = rows_data[detail_row_idx]
                                    if col_idx >= len(row_data):
                                        continue
                                    cell = row_data[col_idx]
                                    if not cell:
                                        continue

                                    try:
                                        value = parse_number(cell)
                                        op = detail_row.get("op", "+")
                                        if op == "-":
                                            value = -value
                                        detail_values.append(value)
                                    except (ValueError, ZeroDivisionError):
                                        pass

                                if not detail_values:
                                    continue

                                # 计算合计
                                calculated_sum = sum(detail_values)

                                # 获取 total 行的值
                                total_value = None
                                total_row_data = rows_data[total_row_idx]
                                if col_idx < len(total_row_data):
                                    total_cell = total_row_data[col_idx]
                                    if total_cell:
                                        try:
                                            total_value = parse_number(total_cell)
                                        except (ValueError, ZeroDivisionError):
                                            pass

                                # **修复：合计行空值跳过**
                                if total_value is None:
                                    continue
                                # 检查合计行是否为空值或占位符
                                if total_cell and str(total_cell).strip() in ["-", "—", "–", "——", "", "0", "0.00"]:
                                    continue

                                # 检查是否相等
                                check_result = check_equal(str(calculated_sum), str(total_value), tolerance=0.01)
                                passed = check_result["match"]
                                diff = check_result["diff"]
                                severity = "info" if passed else "warning"

                                table_name = table.get("name") or (headers[0] if headers else "")
                                results.append(
                                    create_result(
                                        check_type="竖加",
                                        rule_name=f"{table_name}AI标注竖加",
                                        severity=severity,
                                        passed=passed,
                                        description=f"{table_name} AI标注竖加: 明细求和 = {calculated_sum:.2f}, 合计 = {total_value:.2f}" +
                                                    (f", 差异 = {diff:.2f}" if diff > 0.01 else ""),
                                        expected=f"{calculated_sum:.2f}",
                                        actual=f"{total_value:.2f}",
                                        difference=diff if diff > 0.01 else 0.0,
                                        source_location=f"表格ID {table_id}",
                                        page=table.get("page"),
                                        context=f"{table_name} - AI标注",
                                        evidence="ai_worker_标注+code验算_fix"  # 标记为修复版
                                    )
                                )
                                ai_vertical_count += 1

                        # 横加验算：按标注关系
                        for horizontal in horizontal_annotation:
                            formula = horizontal.get("formula", "")
                            operands = horizontal.get("operands", [])
                            result_col = horizontal.get("result_col")

                            if not operands or result_col is None:
                                continue

                            rows_data = table.get("rows", [])
                            tbl_headers = table.get("headers", [])
                            table_name = table.get("name") or (tbl_headers[0] if tbl_headers else "")

                            for row_idx in range(len(rows_data)):
                                row_data = rows_data[row_idx]
                                if not row_data:
                                    continue

                                try:
                                    # 计算 operands
                                    calculated = 0.0
                                    for operand in operands:
                                        op_col = operand.get("col")
                                        op = operand.get("op", "+")

                                        if op_col >= len(row_data):
                                            continue
                                        cell = row_data[op_col]
                                        if not cell:
                                            continue

                                        value = parse_number(cell)
                                        if op == "+":
                                            calculated += value
                                        elif op == "-":
                                            calculated -= value

                                    # 获取 result_col 的值
                                    if result_col >= len(row_data):
                                        continue
                                    actual_value = parse_number(row_data[result_col])

                                    # 检查是否相等
                                    diff = abs(calculated - actual_value)
                                    passed = diff <= 0.01
                                    severity = "info" if passed else "warning"

                                    item_name = row_data[0] if row_data[0] else "未命名行"
                                    results.append(
                                        create_result(
                                            check_type="横加",
                                            rule_name=f"{table_name}AI标注横加",
                                            severity=severity,
                                            passed=passed,
                                            description=f"{item_name}: {formula} = {actual_value:.2f}" +
                                                        (f", 差异 = {diff:.2f}" if diff > 0.01 else ""),
                                            expected=f"{calculated:.2f}",
                                            actual=f"{actual_value:.2f}",
                                            difference=diff if diff > 0.01 else 0.0,
                                            source_location=f"表格ID {table_id}",
                                            page=table.get("page"),
                                            context=f"{table_name} - {item_name}",
                                            evidence="ai_worker_标注+code验算"
                                        )
                                    )
                                except (ValueError, ZeroDivisionError):
                                    continue

                    print(f"[INFO] AI 标注验算覆盖: {ai_vertical_count} 竖加 + {len(annotations)} 横加", file=sys.stderr)

                    # 去重：已被AI标注验算的表，移除 code_vertical 的纯代码竖加结果（避免重复+误报）
                    annotated_ids = {str(a.get("table_id")) for a in annotations}
                    if annotated_ids:
                        before = len(results)
                        results[:] = [
                            r for r in results
                            if not (
                                r.get("check_type") == "竖加"
                                and "code_vertical" in str(r.get("evidence", ""))
                                and str(r.get("source_location", "")).replace("表格ID ", "") in annotated_ids
                            )
                        ]
                        removed = before - len(results)
                        if removed:
                            print(f"[INFO] 去重：移除 {removed} 条 code_vertical 竖加（已被AI标注覆盖）", file=sys.stderr)

                except Exception as e:
                    print(f"[WARN] AI 标注失败: {e}", file=sys.stderr)
                    import traceback
                    traceback.print_exc()

            # 生成表注定位（subjects_index.json）
            subjects_index = build_subjects_index(tables)
            subjects_path = output_path.parent / "subjects_index.json"
            subjects_path.write_text(
                json.dumps(subjects_index, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            print(f"[INFO] 表注定位: {len(subjects_index)} 科目 → {subjects_path}", file=sys.stderr)

            # AI 加速：场景1 - 附注表结构标注（已在 scan 块内完成，此处无需重复）

        except Exception as e:
            print(f"[ERROR] Scan 执行失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    # AI 加速：场景2 - 文本错别字检查
    if use_ai and AI_WORKER_AVAILABLE and report_md.exists() and not args.skip_text_ai:
        print("[INFO] AI 加速：文本错别字检查...", file=sys.stderr)
        try:
            # 构建 DeepSeek client
            client = ai_worker.DeepSeekClient(
                api_key=ai_config["api_key"],
                model=ai_config["model"],
                base_url=ai_config["base_url"],
            )

            # 读取 report.md 并分段
            report_content = report_md.read_text(encoding="utf-8")

            # 按 SOURCE 注释分段（兼容 Word 路无 page 和 PDF 路 page=N）
            source_pattern = r'<!-- SOURCE[^>]*-->'
            source_matches = list(re.finditer(source_pattern, report_content))

            text_chunks = []
            for i, match in enumerate(source_matches):
                # 提取页码（PDF 路），Word 路无 page 为 None
                page_m = re.search(r'page=(\d+)', match.group(0))
                page = page_m.group(1) if page_m else None
                start = match.end()
                end = source_matches[i + 1].start() if i + 1 < len(source_matches) else len(report_content)
                segment = report_content[start:end].strip()

                # 移除 HTML 注释，只保留正文
                segment = re.sub(r'<!--.*?-->', '', segment, flags=re.DOTALL)

                if segment and len(segment) > 50:  # 过滤太短的段落
                    text_chunks.append({
                        "page": page,
                        "text": segment
                    })

            if text_chunks:
                # 只传文本列表
                text_list = [chunk["text"] for chunk in text_chunks]
                ai_check_result = ai_worker.check_typos_ai(text_list, client)

                # 处理错别字
                ai_typos = ai_check_result.get("typos", [])
                if ai_typos:
                    for typo in ai_typos:
                        # 容错：typo 可能不是 dict（DeepSeek 返回格式不稳定）
                        if not isinstance(typo, dict):
                            continue
                        orig = typo.get("original", "")
                        sug = typo.get("suggestion", "")
                        reason = typo.get("reason", "")
                        confidence = typo.get("confidence", 0.0)
                        severity = typo.get("severity", "warning")  # ai_worker 已根据 confidence 设置 severity
                        if not orig or not sug or orig == sug:
                            continue
                        # PDF 断字/漏字降级：original 是 suggestion 的缺字子序列
                        # （"资负债表"是"资产负债表"缺"产"、"税优惠税率"是"税收优惠税率"缺"收"）
                        # → pdfplumber 提取漏字非真错字，跳过；保留"先讲先出"等替换型真错
                        if len(orig) < len(sug):
                            _seq_it = iter(sug)
                            if all(c in _seq_it for c in orig):
                                continue
                        # 低 confidence（<0.85）的条目，ai_worker 已降为 warning，这里直接使用
                        results.append(
                            create_result(
                                check_type="文本",
                                rule_name="错别字检查_AI",
                                severity=severity,
                                passed=False,
                                description=f"AI 检测错别字: '{orig}' 应为 '{sug}' - {reason}",
                                source_location="report.md",
                                evidence=f"ai_worker_错别字: {reason}",
                                context=orig,
                                confidence=confidence
                            )
                        )
                    found_typos = sum(1 for t in ai_typos if isinstance(t, dict))
                    print(f"[INFO] AI 错别字检查：发现 {found_typos} 个错别字", file=sys.stderr)
                else:
                    print("[INFO] AI 错别字检查：未发现错别字", file=sys.stderr)

                # 处理病句（新增）
                ai_grammar_errors = ai_check_result.get("grammar_errors", [])
                if ai_grammar_errors:
                    for err in ai_grammar_errors:
                        # 容错：err 可能不是 dict
                        if not isinstance(err, dict):
                            continue
                        original = err.get("original", "")
                        suggestion = err.get("suggestion", "")
                        reason = err.get("reason", "")
                        confidence = err.get("confidence", 0.0)
                        if not original or not suggestion:
                            continue
                        # 病句统一为 warning（主观性强）
                        results.append(
                            create_result(
                                check_type="文本",
                                rule_name="病句检查",
                                severity="warning",
                                passed=False,
                                description=f"病句: {reason}",
                                source_location="report.md",
                                evidence=f"原句: {original}\n建议: {suggestion}",
                                context=original,
                                confidence=confidence
                            )
                        )
                    found_grammar = sum(1 for e in ai_grammar_errors if isinstance(e, dict))
                    print(f"[INFO] AI 病句检查：发现 {found_grammar} 个病句", file=sys.stderr)
                else:
                    print("[INFO] AI 病句检查：未发现病句", file=sys.stderr)
        except Exception as e:
            print(f"[WARN] AI 错别字检查失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    # AI 加速：场景4 - 表注勾稽 + 附注内 reconcile（定位 note_map → 取数 locate → 算术 calculator）
    if use_ai and AI_WORKER_AVAILABLE and note_map and statements and ai_config:
        print("[INFO] AI 加速：表注勾稽（定位 note_map → 取数 locate → 算术 calculator）...", file=sys.stderr)
        try:
            nm_tables = []
            if extracted_tables.exists():
                try:
                    nm_tables = json.loads(extracted_tables.read_text(encoding="utf-8")).get("tables", [])
                except Exception:
                    pass
            nm_client = ai_worker.DeepSeekClient(
                api_key=ai_config["api_key"], model=ai_config["model"], base_url=ai_config["base_url"]
            )
            located = ai_worker.locate_note_values(note_map, nm_tables, nm_client)
            print(f"[INFO] 表注取数完成: {len(located)} 个科目", file=sys.stderr)
            check_statement_to_note(statements, note_map, located, results)
            check_note_internal_reconcile(note_map, located, results)
        except Exception as e:
            print(f"[WARN] 表注勾稽失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            results.append(create_result(
                check_type="表注", rule_name="表注勾稽", severity="warning", passed=False,
                description=f"表注勾稽执行失败（API/取数异常），已降级为 Claude Step5 手动复核: {e}",
                evidence="Agent复核：降级Claude手动"))

    # AI 加速：场景3 - 二次复核（warning/error 条目）
    # 默认执行 DeepSeek 预复核；最终终审仍由 Claude Step5 回原文完成
    if use_ai and AI_WORKER_AVAILABLE and not args.skip_ai_review:
        print("[INFO] AI 加速：二次复核 warning/error 条目...", file=sys.stderr)
        try:
            # 构建 DeepSeek client
            client = ai_worker.DeepSeekClient(
                api_key=ai_config["api_key"],
                model=ai_config["model"],
                base_url=ai_config["base_url"],
            )

            # 筛选 warning/error 条目
            warnings_to_review = [
                result for result in results
                if result.get("severity") in ["warning", "error"] and not result.get("passed", True)
            ]

            if warnings_to_review:
                review_results = ai_worker.review_results(warnings_to_review, client)

                # 应用复核结果
                confirm_count = 0
                downgrade_count = 0
                needs_review_count = 0

                for review in review_results:
                    result_idx = review.get("original_index")
                    decision = review.get("decision")
                    reason = review.get("reason", "")

                    if result_idx is None or result_idx >= len(results):
                        continue

                    result = results[result_idx]

                    if decision == "confirm":
                        # 确认真错，保持 severity，添加 evidence
                        result["evidence"] = f"{result.get('evidence', '')} [DeepSeek复核确认: {reason}]"
                        confirm_count += 1

                    elif decision == "downgrade":
                        # 降为 info
                        result["severity"] = "info"
                        result["description"] = f"{result.get('description', '')} [经复核为特殊结构: {reason}]"
                        result["evidence"] = f"{result.get('evidence', '')} [DeepSeek降级: {reason}]"
                        downgrade_count += 1

                    elif decision == "needs_review":
                        # 保持 warning，添加 evidence
                        result["evidence"] = f"{result.get('evidence', '')} [需人工核实: {reason}]"
                        needs_review_count += 1

                print(f"[INFO] AI 二次复核：确认 {confirm_count}，降级 {downgrade_count}，需人工核实 {needs_review_count}", file=sys.stderr)
            else:
                print("[INFO] 无 warning/error 条目需要复核", file=sys.stderr)

        except Exception as e:
            print(f"[WARN] AI 二次复核失败: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    # 构建输出
    output_data = {
        "meta": {
            "company": "",
            "period": "",
            "scope": args.scope,
            "source": "被检查文件名",  # 简化，实际可从 extracted_tables 读取
            "checked_at": datetime.now().isoformat(),
        },
        "results": results,
    }

    # 从 extracted_tables 读取 company/period（如果之前没读取）
    source_file_path = ""
    if extracted_tables.exists():
        if tables_data is None:
            try:
                tables_data = json.loads(extracted_tables.read_text(encoding="utf-8"))
            except Exception:
                pass

        if tables_data:
            output_data["meta"]["company"] = tables_data.get("company", "")
            output_data["meta"]["period"] = tables_data.get("period", "")
            if tables_data.get("sources"):
                output_data["meta"]["source"] = tables_data["sources"][0]
            # 还原原始报告的绝对路径（用于 Excel 超链接定位原文）
            # parse_report 把 source_file 存为相对路径，这里结合 source_root 还原
            source_root = tables_data.get("source_root")
            for src in tables_data.get("sources", []):
                if source_root and not os.path.isabs(src):
                    # 相对路径：用 source_root 还原成绝对路径
                    source_file_path = str(Path(source_root) / src)
                else:
                    source_file_path = src
                break

    # 给结果回填 source_file（按检查类型指向实际所在文件，Excel 页码超链接用）
    # 旧逻辑取 sources[0]（常是封面）→ 报表间/表注超链接打不开；改为按 check_type 推断实际文件
    _sources = tables_data.get("sources", []) if tables_data else []
    source_root = tables_data.get("source_root") if tables_data else None

    # 通用 fallback：按文件名启发式找报表/附注文件（覆盖多种命名约定，不依赖特定报告）
    _REPORT_KW = ("财务报表", "资产负债表", "利润表", "现金流量表", "balance", "financial")
    _NOTE_KW = ("附注", "notes", "note", "annex")

    def _find_src(kws):
        # 优先 .pdf（有 page 页码，超链接能定位到具体页），fallback 任意匹配
        for s in _sources:
            if s.endswith('.pdf') and any(k in s.lower() for k in kws):
                # 还原绝对路径
                return str(Path(source_root) / s) if source_root and not os.path.isabs(s) else s
        for s in _sources:
            if any(k in s.lower() for k in kws):
                return str(Path(source_root) / s) if source_root and not os.path.isabs(s) else s
        return ""

    _report_src = _find_src(_REPORT_KW) or source_file_path
    _note_src = _find_src(_NOTE_KW) or source_file_path
    for r in results:
        if r.get("source_file"):
            # 如果是相对路径，还原成绝对路径
            src = r["source_file"]
            if source_root and not os.path.isabs(src):
                r["source_file"] = str(Path(source_root) / src)
            continue  # 已有（部分 result 从 table 继承了精确 source_file）
        ct = r.get("check_type", "")
        if ct == "报表间":
            r["source_file"] = _report_src
        elif ct in ("表注", "附注内", "横加", "竖加"):
            r["source_file"] = _note_src
        else:
            r["source_file"] = source_file_path

    # 写输出（如果之前没有定义 output_path，现在定义）
    if 'output_path' not in locals():
        output_path = Path(args.output).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

    # 写输出
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(output_data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    # 打印摘要
    total = len(results)
    passed = sum(1 for r in results if r["passed"])
    errors = sum(1 for r in results if r["severity"] == "error" and not r["passed"])
    warnings = sum(1 for r in results if r["severity"] == "warning")
    infos = sum(1 for r in results if r["severity"] == "info")

    print(f"[OK] 检查完成，结果已写入: {output_path}", file=sys.stderr)
    print(f"     总计: {total}，通过: {passed}，错误: {errors}，警告: {warnings}，提示: {infos}", file=sys.stderr)


if __name__ == "__main__":
    main()