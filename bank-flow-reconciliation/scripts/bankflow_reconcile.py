#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
"""
银行流水核对脚本 - Standalone Bank Flow Reconciliation Script

银行流水与财务序时账智能双向核对工具。
支持 9 层匹配算法，自动生成核对底稿。
"""

import os
import sys
import time
import json
import re
import math
import itertools
import traceback
import argparse
from datetime import datetime

import pandas as pd
import numpy as np

# 条件导入：rapidfuzz 优先，difflib 兜底
try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    import difflib

# 条件导入：openpyxl 用于 Excel 输出
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

import warnings
warnings.filterwarnings('ignore')


# ========================================================
# 1. 辅助函数
# ========================================================

def log(msg, indent=0):
    """简单日志输出"""
    prefix = "  " * indent
    print(f"{prefix}{msg}")


def safe_float(val):
    """安全转换为浮点数，失败返回 0.0"""
    if val is None or pd.isna(val):
        return 0.0
    try:
        str_val = str(val).replace(',', '').strip()
        return float(str_val) if str_val else 0.0
    except:
        return 0.0


def format_date(val):
    """
    将日期值转换为标准格式字符串 (YYYY-MM-DD)
    支持多种格式：Excel serial date, datetime, 字符串
    """
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass

    try:
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime('%Y-%m-%d')

        str_val = str(val).strip()
        if not str_val:
            return None

        dt = pd.to_datetime(str_val, errors='coerce', dayfirst=False)
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')

        # 中文日期: 2025年1月3日
        m = re.match(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?', str_val)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

        # 点分隔: 2025.01.03
        m = re.match(r'(\d{4})\.(\d{1,2})\.(\d{1,2})', str_val)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

        # 英文月份: Jan 3 2025, 3-Jan-2025
        m = re.match(r'([A-Za-z]+)\s+(\d{1,2}),?\s*(\d{4})', str_val)
        if m:
            dt2 = pd.to_datetime(str_val, errors='coerce', format='mixed')
            if pd.notna(dt2):
                return dt2.strftime('%Y-%m-%d')

        m = re.match(r'(\d{1,2})[\s\-]([A-Za-z]+)[\s\-](\d{4})', str_val)
        if m:
            dt2 = pd.to_datetime(str_val, errors='coerce', format='mixed')
            if pd.notna(dt2):
                return dt2.strftime('%Y-%m-%d')

        # 纯数字: 20250103
        m = re.match(r'^(\d{4})(\d{2})(\d{2})$', str_val)
        if m:
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

        return None
    except Exception:
        return None


def validate_dates(df, date_col, source_name, raw_date_col_name):
    """验证日期列的解析成功率"""
    if df.empty or date_col not in df.columns:
        return

    total_rows = len(df)
    failed_rows = df[date_col].isna().sum()

    if total_rows == 0:
        return

    fail_rate = failed_rows / total_rows
    if fail_rate > 0.5:
        log(f"❌ 日期解析失败！{source_name}中 {failed_rows}/{total_rows} 行（{fail_rate*100:.0f}%）的日期无法识别", 0)
        log(f"   配置的日期列名: [{raw_date_col_name}]", 0)

        failed_df = df[df[date_col].isna()]
        if not failed_df.empty:
            samples = []
            for _, row in failed_df.head(5).iterrows():
                raw_val = row.get(raw_date_col_name)
                if raw_val is not None and not (isinstance(raw_val, float) and pd.isna(raw_val)):
                    samples.append(str(raw_val))
            if samples:
                log(f"   无法识别的日期样本: {samples}", 0)

        log(f"   支持的日期格式: 2025-01-03, 2025/01/03, 2025.01.03, 20250103, 2025年1月3日, Jan 3 2025, 3-Jan-2025 等", 0)
        raise ValueError(f"{source_name}日期列 [{raw_date_col_name}] 中超过 {fail_rate*100:.0f}% 的日期无法解析，请检查日期格式是否正确。")


def read_data_file(file_path, **kwargs):
    """
    根据文件扩展名自动选择读取方式，支持 .xlsx/.xls 和 .csv
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        # CSV encoding fallback (Fix 4)
        encodings = ['utf-8', 'gbk', 'gb18030', 'utf-8-sig']
        for enc in encodings:
            try:
                return pd.read_csv(file_path, encoding=enc, **kwargs)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"无法解码 CSV 文件，已尝试编码: {encodings}")
    else:
        return pd.read_excel(file_path, **kwargs)


# ========================================================
# 2. 工具类
# ========================================================

class DataCleaner:
    """数据清洗工具类"""
    @staticmethod
    def clean_name(text):
        """清洗交易对手名称：去除括号内容、空白字符、换行"""
        if not isinstance(text, str):
            return ""
        text = text.strip().replace('　', ' ')
        text = re.sub(r'\(.*?\)', '', text)
        text = re.sub(r'（.*?）', '', text)
        text = text.replace("()", "").replace("（）", "")
        text = text.replace("\n", "").replace("\t", "")
        return text.strip()


class SimilarityEngine:
    """相似度计算引擎"""
    def __init__(self):
        if HAS_RAPIDFUZZ:
            log("✅ 已加载 RapidFuzz 引擎 (高性能模糊匹配)", 1)
        else:
            log("✅ 已加载 Difflib 引擎 (标准库模糊匹配)", 1)

    def compute_similarity(self, text1, text2):
        """计算两个文本的相似度，返回 0.0-1.0 之间的值"""
        if not text1 or not text2:
            return 0.0
        t1 = DataCleaner.clean_name(str(text1))
        t2 = DataCleaner.clean_name(str(text2))
        if not t1 or not t2:
            return 0.0
        # 包含关系直接返回 1.0
        if t1 in t2 or t2 in t1:
            return 1.0

        if HAS_RAPIDFUZZ:
            return fuzz.token_sort_ratio(t1, t2) / 100.0
        else:
            return difflib.SequenceMatcher(None, t1, t2).ratio()


class ConfigLoader:
    """配置加载器 - 仅支持 config_dict 模式"""
    def __init__(self, config_dict=None):
        self.gl_config = {}
        self.bank_config = {}

        if config_dict:
            # 支持嵌套的 gl_config 和 bank_config
            if 'gl_config' in config_dict:
                self.gl_config = config_dict['gl_config']
            else:
                # 兼容扁平的配置格式
                self.gl_config = {
                    '序时账路径': config_dict.get('序时账路径'),
                    '公司列名': config_dict.get('公司列名'),
                    '日期列名': config_dict.get('日期列名'),
                    '凭证号列名': config_dict.get('凭证号列名'),
                    '科目代码列名': config_dict.get('科目代码列名'),
                    '科目名称列名': config_dict.get('科目名称列名'),
                    '对方科目列名': config_dict.get('对方科目列名'),
                    '摘要列名': config_dict.get('摘要列名'),
                    '借方金额列名': config_dict.get('借方金额列名'),
                    '贷方金额列名': config_dict.get('贷方金额列名'),
                    '银行科目代码': config_dict.get('银行科目代码'),
                    '客商辅助项列名': config_dict.get('客商辅助项列名'),
                    '交易对手列名': config_dict.get('交易对手列名'),
                    '账户列名': config_dict.get('账户列名')
                }

            if 'bank_config' in config_dict:
                self.bank_config = config_dict['bank_config']
            else:
                # 兼容扁平的配置格式
                self.bank_config = {
                    '银行流水路径': config_dict.get('银行流水路径'),
                    '公司列名': config_dict.get('公司列名'),
                    '账户列名': config_dict.get('账户列名'),
                    '日期列名': config_dict.get('日期列名'),
                    '交易对手列名': config_dict.get('交易对手列名'),
                    '摘要列名': config_dict.get('摘要列名'),
                    '收入金额列名': config_dict.get('收入金额列名'),
                    '支出金额列名': config_dict.get('支出金额列名')
                }


# ========================================================
# 3. 数据处理类
# ========================================================

class GLProcessor:
    """序时账处理器"""
    def __init__(self, config):
        self.cfg = config
        self.cleaner = DataCleaner()
        self.MAX_SPLIT_LIMIT = 30

    def process(self, work_dir):
        """处理序时账文件"""
        rel_path = self.cfg.get('序时账路径')
        if not rel_path:
            return pd.DataFrame()
        file_path = rel_path if os.path.isabs(rel_path) else os.path.join(work_dir, rel_path)
        log(f"📄 读取序时账: {os.path.basename(file_path)}", 0)

        try:
            df_raw = read_data_file(file_path, dtype=str, keep_default_na=False)
            df_raw = df_raw.replace('', np.nan)
        except Exception as e:
            log(f"❌ 读取失败: {e}", 1)
            return pd.DataFrame()

        if df_raw.empty:
            log(f"⚠️ 序时账文件内容为空", 1)
            return pd.DataFrame()

        # Clean column names: strip whitespace, remove fullwidth spaces (Fix 7)
        df_raw.columns = df_raw.columns.astype(str).str.strip().str.replace('\u3000', '').str.replace('\xa0', ' ')

        # ── 日期范围过滤（可选，硬失败：列名错误时抛异常）──
        date_range = self.cfg.get('日期范围', '')
        if date_range and date_range.strip():
            parts = [p.strip() for p in date_range.split(',')]
            if len(parts) == 2 and parts[0] and parts[1]:
                date_col = self.cfg.get('日期列名', '')
                if not date_col:
                    raise ValueError(f"日期范围过滤已配置但未找到日期列名配置")
                if date_col not in df_raw.columns:
                    raise ValueError(f"日期范围过滤失败：列 [{date_col}] 不在序时账中。请检查配置。")
                before = len(df_raw)
                dates = pd.to_datetime(df_raw[date_col], errors='coerce')
                # 结束日含当天全天：normalize 抹掉时分秒
                end_date = pd.to_datetime(parts[1]).normalize() + pd.Timedelta(days=1)
                mask = (dates >= parts[0]) & (dates < end_date)
                df_raw = df_raw[mask].copy()
                log(f"📅 日期过滤: {parts[0]} ~ {parts[1]} → {before}→{len(df_raw)} 行", 0)

        # ── 主体过滤（可选，硬失败：列名错误时抛异常）──
        # 格式: "列名,值1|值2"，第一个逗号前是列名，后是值
        entity_filter = self.cfg.get('主体过滤', '')
        if entity_filter and entity_filter.strip():
            parts = [p.strip() for p in entity_filter.split(',', 1)]
            if len(parts) == 2 and parts[0] and parts[1]:
                ent_col, ent_vals = parts[0], parts[1]
                if ent_col not in df_raw.columns:
                    raise ValueError(f"主体过滤失败：列 [{ent_col}] 不在序时账中。请检查配置。")
                before = len(df_raw)
                allowed = [v.strip() for v in ent_vals.split('|')]
                df_raw = df_raw[df_raw[ent_col].astype(str).str.strip().isin(allowed)].copy()
                log(f"🏢 主体过滤: {ent_col} ∈ {allowed} → {before}→{len(df_raw)} 行", 0)

        # Validate required columns exist (Fix 2)
        required_cols = ['日期列名', '科目代码列名', '借方金额列名', '贷方金额列名']
        missing = []
        for key in required_cols:
            col_name = self.cfg.get(key)
            if not col_name:
                missing.append(f"{key}未配置")
            elif col_name not in df_raw.columns:
                missing.append(f"{key}=[{col_name}] 在序时账列中不存在")
        if missing:
            for m in missing:
                log(f"❌ {m}", 1)
            log(f"   序时账实际列: {list(df_raw.columns)[:15]}", 1)
            return pd.DataFrame()

        # 添加默认公司列
        comp_col = self.cfg.get('公司列名')
        acc_col = self.cfg.get('账户列名')
        if not comp_col or comp_col not in df_raw.columns:
            df_raw['__temp_comp__'] = '默认公司'
            comp_col = '__temp_comp__'
        if not acc_col or acc_col not in df_raw.columns:
            df_raw['__temp_acc__'] = '默认账户'
            acc_col = '__temp_acc__'

        # 判断模式：直接模式 vs 拆分模式
        partner_direct_col = self.cfg.get('交易对手列名')
        use_direct_mode = partner_direct_col and partner_direct_col.strip() and partner_direct_col in df_raw.columns

        if use_direct_mode:
            log(f"✅ 使用直接模式：从列 [{partner_direct_col}] 读取交易对手", 1)
        else:
            partner_cols_cfg = self.cfg.get('客商辅助项列名')
            if partner_cols_cfg and partner_cols_cfg.strip():
                log(f"✅ 使用拆分模式：从对方科目辅助项提取交易对手", 1)
            else:
                log("⚠️⚠️ 警告：序时账未配置交易对手信息！", 1)
                log("   L1/L3/L4/L5/L7 层将全部失效，预计匹配率从 ~97% 降至 ~40%", 1)
                log("   强烈建议配置 [客商辅助项列名] 或 [交易对手列名]", 1)

        # 银行科目过滤
        bank_subject = self.cfg.get('银行科目代码')
        subj_col = self.cfg.get('科目代码列名') or self.cfg.get('科目名称列名')

        if not subj_col or subj_col not in df_raw.columns:
            log(f"❌ 配置错误: 序时账中找不到科目列名 [{subj_col}]", 1)
            return pd.DataFrame()

        bank_subjects = [s.strip() for s in str(bank_subject).replace('，', ',').split(',') if s.strip()]

        def matches_any_subject(val):
            if pd.isna(val):
                return False
            val_str = str(val)
            return any(val_str.startswith(subj) for subj in bank_subjects)

        df_raw['__is_target__'] = df_raw[subj_col].apply(matches_any_subject)
        target_count = df_raw['__is_target__'].sum()
        if target_count == 0:
            subjects_str = '、'.join(bank_subjects)
            log(f"❌ 警告：在序时账中未找到以 '{subjects_str}' 开头的科目行！", 1)
            return pd.DataFrame()

        # 凭证分组：只用 [公司, 日期, 凭证号] 分组
        voucher_group_keys = [comp_col, self.cfg.get('日期列名'), self.cfg.get('凭证号列名')]
        valid_keys = [k for k in voucher_group_keys if k in df_raw.columns]

        # 填充分组键中的空值，避免 groupby 静默丢弃行（数据完整性保护）
        for k in valid_keys:
            if df_raw[k].isna().any():
                n_na = df_raw[k].isna().sum()
                log(f"  ⚠️ [{k}] 列有 {n_na} 个空值，填充为 '__missing__' 后参与分组", 0)
                df_raw[k] = df_raw[k].fillna('__missing__')

        splitted_records = []
        grouped = df_raw.groupby(valid_keys, sort=False)

        for _, group in grouped:
            res = self._process_voucher(group, comp_col, acc_col, use_direct_mode, partner_direct_col)
            splitted_records.extend(res)

        df_res = pd.DataFrame(splitted_records)
        if df_res.empty:
            return df_res

        date_col_name = self.cfg.get('日期列名')
        validate_dates(df_res, '__gl_date__', '序时账', date_col_name)

        counts = df_res['__orig_idx__'].value_counts()
        df_res['__split_status__'] = df_res['__orig_idx__'].apply(lambda x: "已拆分" if counts.get(x, 0) > 1 else None)
        return df_res

    def _process_voucher(self, group, comp_col, acc_col, use_direct_mode, partner_direct_col):
        """处理单个凭证"""
        target_rows = group[group['__is_target__']]
        if target_rows.empty:
            return []

        summary_col = self.cfg.get('摘要列名')
        debit_col = self.cfg.get('借方金额列名')
        credit_col = self.cfg.get('贷方金额列名')

        def make_record(orig_row, base_row, amount_override=None, partner_override=None, summary_override=None):
            """
            创建拆分后的记录
            orig_row: 原始行（用于保留原始列数据）
            base_row: 银行目标行（用于提取公司/账户/日期/凭证号等基础信息）
            """
            rec = orig_row.to_dict()
            d_val = safe_float(orig_row.get(debit_col))
            c_val = safe_float(orig_row.get(credit_col))
            net_amt = d_val - c_val

            if amount_override is not None:
                net_amt = amount_override
                if net_amt >= 0:
                    rec[debit_col] = abs(net_amt)
                    rec[credit_col] = 0
                else:
                    rec[debit_col] = 0
                    rec[credit_col] = abs(net_amt)

            if summary_override and summary_col:
                rec[summary_col] = summary_override

            rec['__link_comp__'] = base_row[comp_col]
            rec['__link_acc__'] = base_row[acc_col]
            rec['__gl_date__'] = format_date(base_row[self.cfg.get('日期列名')])
            rec['__gl_voucher__'] = base_row[self.cfg.get('凭证号列名')]
            rec['__gl_amount__'] = net_amt
            rec['__orig_idx__'] = orig_row.name

            summ = rec.get(summary_col)
            rec['__gl_summary__'] = str(summ) if summ and pd.notna(summ) else ''

            raw_p = partner_override
            rec['__gl_partner_raw__'] = raw_p
            rec['__gl_partner_clean__'] = self.cleaner.clean_name(raw_p)

            return rec

        # 【直接模式】：直接从指定列读取交易对手
        if use_direct_mode:
            results = []
            for _, r in target_rows.iterrows():
                if abs(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col))) < 0.005:
                    continue
                partner_val = r.get(partner_direct_col)
                partner_str = str(partner_val) if pd.notna(partner_val) and str(partner_val).strip() else None
                rec = make_record(r, r, partner_override=partner_str)
                results.append(rec)
            return results

        # 【拆分模式】：从对方科目提取客商
        # 检查是否存在多个不同账户的银行行
        if acc_col in target_rows.columns:
            unique_accounts = target_rows[acc_col].dropna().unique()
        else:
            unique_accounts = []

        if len(unique_accounts) > 1:
            # 多银行账户凭证：按账户分组，每组独立拆分
            return self._process_voucher_multi_bank(
                group, target_rows, comp_col, acc_col,
                summary_col, debit_col, credit_col, make_record
            )

        # 单银行账户凭证
        total_bank = sum(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col)) for _, r in target_rows.iterrows())

        # 处理内部转账 (总净额为0)
        if abs(total_bank) < 0.005:
            gross_amt = sum(abs(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col))) for _, r in target_rows.iterrows())
            if gross_amt > 0.005:
                return [make_record(r, r) for _, r in target_rows.iterrows()]
            else:
                return []

        counter_rows = group[~group['__is_target__']]
        if counter_rows.empty or len(counter_rows) > self.MAX_SPLIT_LIMIT:
            return [make_record(r, r) for _, r in target_rows.iterrows() if abs(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col))) > 0.005]

        partner_cols_cfg = self.cfg.get('客商辅助项列名')
        partner_cols = [c.strip() for c in str(partner_cols_cfg).replace('，', ',').split(',')] if partner_cols_cfg else []

        # 调用拆分匹配算法
        return self._split_by_matching(
            group, target_rows, counter_rows, comp_col, acc_col,
            summary_col, debit_col, credit_col, partner_cols, make_record
        )

    def _split_by_matching(self, group, target_rows, counter_rows, comp_col, acc_col,
                           summary_col, debit_col, credit_col, partner_cols, make_record):
        """
        借鉴 finance_viewer 匹配法的拆分逻辑。

        核心原则：
        1. 拆分的唯一目的是获取交易对手信息
        2. 只有对方科目有多个不同客商时才拆分
        3. 金额以银行存款行自身为准，按对方科目比例分配
        4. 借贷方发生额绝不改变，不产生残差
        """
        # 提取对方科目行的客商信息
        counter_info = []
        for _, c_row in counter_rows.iterrows():
            c_d = safe_float(c_row.get(debit_col))
            c_c = safe_float(c_row.get(credit_col))
            c_amt = c_d - c_c
            if abs(c_amt) < 0.005:
                continue
            p_val = None
            for col in partner_cols:
                val = c_row.get(col)
                if pd.notna(val) and str(val).strip() != '':
                    p_val = str(val).strip()
                    break
            counter_info.append({'amt': c_amt, 'abs_amt': abs(c_amt), 'partner': p_val, 'orig': c_row})

        # 判断是否需要拆分
        partners_with_value = [c['partner'] for c in counter_info if c['partner']]
        unique_partners = list(dict.fromkeys(partners_with_value))

        if len(unique_partners) <= 1:
            partner_str = unique_partners[0] if unique_partners else None
            results = []
            for _, r in target_rows.iterrows():
                if abs(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col))) < 0.005:
                    continue
                results.append(make_record(r, r, partner_override=partner_str))
            return results

        # 需要拆分：有多个不同客商
        all_lines = []
        for _, r in group.iterrows():
            d_val = safe_float(r.get(debit_col))
            c_val = safe_float(r.get(credit_col))
            net = d_val - c_val
            if abs(net) < 0.005:
                continue
            is_target = bool(r.get('__is_target__', False))
            p_val = None
            if not is_target:
                for col in partner_cols:
                    val = r.get(col)
                    if pd.notna(val) and str(val).strip() != '':
                        p_val = str(val).strip()
                        break
            direction = 'debit' if net > 0 else 'credit'
            all_lines.append({
                'row': r, 'idx': r.name,
                'd_val': d_val, 'c_val': c_val,
                'net': net, 'abs_amt': abs(net),
                'direction': direction,
                'is_target': is_target,
                'partner': p_val,
                'matched': False,
                'paired_counterparts': [],
            })

        debit_group = [l for l in all_lines if l['direction'] == 'debit']
        credit_group = [l for l in all_lines if l['direction'] == 'credit']

        if not debit_group or not credit_group:
            results = []
            for _, r in target_rows.iterrows():
                if abs(safe_float(r.get(debit_col)) - safe_float(r.get(credit_col))) < 0.005:
                    continue
                results.append(make_record(r, r))
            return results

        # 精确匹配（1:1）
        for d_line in debit_group:
            if d_line['matched']:
                continue
            for c_line in credit_group:
                if c_line['matched']:
                    continue
                if abs(d_line['abs_amt'] - c_line['abs_amt']) < 0.005:
                    d_line['matched'] = True
                    c_line['matched'] = True
                    d_line['paired_counterparts'].append(c_line)
                    c_line['paired_counterparts'].append(d_line)
                    break

        # 组合精确匹配（N:1）
        for d_line in debit_group:
            if d_line['matched']:
                continue
            candidates = [c for c in credit_group if not c['matched']]
            if len(candidates) < 2 or len(candidates) > 8:
                continue
            found = self._find_subset_sum_lines(candidates, d_line['abs_amt'])
            if found is not None:
                d_line['matched'] = True
                for c in found:
                    c['matched'] = True
                    d_line['paired_counterparts'].append(c)
                    c['paired_counterparts'].append(d_line)

        for c_line in credit_group:
            if c_line['matched']:
                continue
            candidates = [d for d in debit_group if not d['matched']]
            if len(candidates) < 2 or len(candidates) > 8:
                continue
            found = self._find_subset_sum_lines(candidates, c_line['abs_amt'])
            if found is not None:
                c_line['matched'] = True
                for d in found:
                    d['matched'] = True
                    c_line['paired_counterparts'].append(d)
                    d['paired_counterparts'].append(c_line)

        # 剩余未匹配行：比例分配
        remaining_debit = [d for d in debit_group if not d['matched']]
        remaining_credit = [c for c in credit_group if not c['matched']]
        if remaining_debit and remaining_credit:
            credit_total = sum(c['abs_amt'] for c in remaining_credit)
            debit_total = sum(d['abs_amt'] for d in remaining_debit)
            if credit_total > 0:
                for d_line in remaining_debit:
                    for c_line in remaining_credit:
                        d_line['paired_counterparts'].append(c_line)
                        c_line['paired_counterparts'].append(d_line)
                    d_line['matched'] = True
            if debit_total > 0:
                for c_line in remaining_credit:
                    c_line['matched'] = True

        # 从配对结果中为每个银行行提取交易对手并决定是否拆分
        results = []
        for line in all_lines:
            if not line['is_target']:
                continue
            row = line['row']
            d_val = line['d_val']
            c_val = line['c_val']
            if abs(d_val - c_val) < 0.005:
                continue

            counterparts = [cp for cp in line['paired_counterparts'] if not cp['is_target']]

            if not counterparts:
                results.append(make_record(row, row))
                continue

            cp_with_partner = [cp for cp in counterparts if cp['partner']]
            cp_unique_partners = list(dict.fromkeys(cp['partner'] for cp in cp_with_partner))

            if len(cp_unique_partners) <= 1:
                partner_str = cp_unique_partners[0] if cp_unique_partners else None
                results.append(make_record(row, row, partner_override=partner_str))
                continue

            # 多个不同客商 → 按客商分组，按比例拆分
            partner_amounts = {}
            for cp in counterparts:
                p = cp['partner'] or '__no_partner__'
                partner_amounts[p] = partner_amounts.get(p, 0.0) + cp['abs_amt']

            total_cp_amt = sum(partner_amounts.values())
            if total_cp_amt < 0.005:
                results.append(make_record(row, row))
                continue

            bank_net = d_val - c_val
            split_records = []
            allocated_sum = 0.0
            partner_list = list(partner_amounts.items())

            for i, (partner, cp_amt) in enumerate(partner_list):
                if i == len(partner_list) - 1:
                    split_amt = bank_net - allocated_sum
                else:
                    ratio = cp_amt / total_cp_amt
                    split_amt = round(bank_net * ratio, 2)
                    allocated_sum += split_amt

                p_override = partner if partner != '__no_partner__' else None
                split_records.append(make_record(row, row, amount_override=split_amt, partner_override=p_override))

            results.extend(split_records)

        return results

    @staticmethod
    def _find_subset_sum_lines(candidates, target):
        """在 candidates 中查找 abs_amt 之和等于 target 的子集"""
        for size in range(2, len(candidates) + 1):
            for subset in itertools.combinations(candidates, size):
                total = sum(s['abs_amt'] for s in subset)
                if abs(total - target) < 0.005:
                    return list(subset)
        return None

    def _process_voucher_multi_bank(self, group, target_rows, comp_col, acc_col,
                                     summary_col, debit_col, credit_col, make_record):
        """处理同一凭证包含多个银行账户的情况"""
        counter_rows = group[~group['__is_target__']]
        partner_cols_cfg = self.cfg.get('客商辅助项列名')
        partner_cols = [c.strip() for c in str(partner_cols_cfg).replace('，', ',').split(',')] if partner_cols_cfg else []

        bank_groups = {}
        for _, t_row in target_rows.iterrows():
            acc_val = str(t_row.get(acc_col, ''))
            amt = safe_float(t_row.get(debit_col)) - safe_float(t_row.get(credit_col))
            if abs(amt) < 0.005:
                continue
            if acc_val not in bank_groups:
                bank_groups[acc_val] = []
            bank_groups[acc_val].append(t_row)

        if not bank_groups:
            return []

        if len(bank_groups) == 1:
            return self._split_by_matching(
                group, target_rows, counter_rows, comp_col, acc_col,
                summary_col, debit_col, credit_col, partner_cols, make_record
            )

        results = []
        for acc_val, rows_list in bank_groups.items():
            sub_target = pd.DataFrame([r for r in rows_list])
            sub_group = pd.concat([sub_target, counter_rows], ignore_index=True)
            sub_results = self._split_by_matching(
                sub_group, sub_target, counter_rows, comp_col, acc_col,
                summary_col, debit_col, credit_col, partner_cols, make_record
            )
            results.extend(sub_results)

        return results


class BankProcessor:
    """银行流水处理器"""
    def __init__(self, config):
        self.cfg = config
        self.cleaner = DataCleaner()

    def process(self, work_dir):
        """处理银行流水文件"""
        rel_path = self.cfg.get('银行流水路径')
        if not rel_path:
            return pd.DataFrame()

        file_path = rel_path if os.path.isabs(rel_path) else os.path.join(work_dir, rel_path)
        log(f"📄 读取银行流水: {os.path.basename(file_path)}", 0)

        try:
            df_raw = read_data_file(file_path, dtype=str, keep_default_na=False)
            df_raw = df_raw.replace('', np.nan)
        except Exception as e:
            log(f"❌ 读取失败: {e}", 1)
            return pd.DataFrame()

        if df_raw.empty:
            log(f"⚠️ 银行流水文件内容为空", 1)
            return pd.DataFrame()

        # Clean column names: strip whitespace, remove fullwidth spaces (Fix 7)
        df_raw.columns = df_raw.columns.astype(str).str.strip().str.replace('\u3000', '').str.replace('\xa0', ' ')

        comp_col = self.cfg.get('公司列名')
        acc_col = self.cfg.get('账户列名')
        if not comp_col or comp_col not in df_raw.columns:
            df_raw['__temp_comp__'] = '默认公司'
            comp_col = '__temp_comp__'
        if not acc_col or acc_col not in df_raw.columns:
            df_raw['__temp_acc__'] = '默认账户'
            acc_col = '__temp_acc__'

        inc_col = self.cfg.get('收入金额列名')
        exp_col = self.cfg.get('支出金额列名')
        date_col = self.cfg.get('日期列名')

        # Validate required columns exist (Fix 2)
        required_bank_cols = ['日期列名', '收入金额列名', '支出金额列名']
        missing_bank = []
        for key in required_bank_cols:
            col_name = self.cfg.get(key)
            if not col_name:
                missing_bank.append(f"{key}未配置")
            elif col_name not in df_raw.columns:
                missing_bank.append(f"{key}=[{col_name}] 在银行流水列中不存在")
        if missing_bank:
            for m in missing_bank:
                log(f"❌ {m}", 1)
            log(f"   银行流水实际列: {list(df_raw.columns)[:15]}", 1)
            return pd.DataFrame()

        # Check trading partner column (Fix 3)
        partner_col = self.cfg.get('交易对手列名')
        if not partner_col or partner_col not in df_raw.columns:
            log("⚠️⚠️ 警告：银行流水未配置交易对手列！", 1)
            log("   L1/L3/L4/L5/L7 层将全部失效，预计匹配率大幅下降", 1)
        else:
            nonempty = df_raw[partner_col].astype(str).str.strip().replace('', pd.NA).dropna()
            fill_rate = len(nonempty) / len(df_raw) if len(df_raw) > 0 else 0
            if fill_rate < 0.5:
                log(f"⚠️ 警告：交易对手列 [{partner_col}] 非空率仅 {fill_rate*100:.1f}%", 1)
                log("   大量空值会显著影响匹配准确率", 1)

        records = []
        for _, row in df_raw.iterrows():
            inc = safe_float(row.get(inc_col))
            exp = safe_float(row.get(exp_col))
            amt = inc - exp
            if abs(amt) < 0.005:
                continue

            rec = row.to_dict()
            rec['__link_comp__'] = row[comp_col]
            rec['__link_acc__'] = row[acc_col]
            rec['__bank_date__'] = format_date(row.get(date_col))
            rec['__bank_amount__'] = amt

            p = row.get(self.cfg.get('交易对手列名'))
            rec['__bank_partner_clean__'] = self.cleaner.clean_name(p)

            records.append(rec)

        df_result = pd.DataFrame(records)

        if not df_result.empty:
            date_col_name = self.cfg.get('日期列名')
            validate_dates(df_result, '__bank_date__', '银行流水', date_col_name)

        return df_result


# ========================================================
# 4. 核心匹配引擎（9层算法）
# ========================================================

class ReconciliationEngine:
    """银行流水与序时账核对引擎 - 9层匹配算法"""
    def __init__(self, df_bank, df_gl, sim_engine, config, ai_review_dir=None, ai_review_limit=500):
        self.full_bank = df_bank
        self.full_gl = df_gl
        self.sim = sim_engine
        self.cfg = config

        self.ai_review_dir = ai_review_dir
        self.ai_review_limit = ai_review_limit
        self.ai_review_pending = False
        self.ai_applied = 0

        self.match_counter = 0
        self.full_bank['matched_id'] = None
        self.full_gl['matched_id'] = None
        self.full_bank['match_desc'] = None
        self.full_gl['match_desc'] = None

        self.l5_stats = {'高': 0, '中': 0, '低': 0}
        self._l7_total = 0
        self._l8_total = 0

        self._add_month_col(self.full_bank, '__bank_date__')
        self._add_month_col(self.full_gl, '__gl_date__')

    def _add_month_col(self, df, date_col):
        """添加 __dt__ 和 __month__ 列"""
        df['__dt__'] = pd.to_datetime(df[date_col], errors='coerce')
        df['__month__'] = df['__dt__'].dt.to_period('M')

    def run(self):
        """执行核对流程"""
        # 检测账户信息不对称（Fix 1）
        gl_accs = set(self.full_gl['__link_acc__'].unique()) if '__link_acc__' in self.full_gl.columns else set()
        bank_accs = set(self.full_bank['__link_acc__'].unique()) if '__link_acc__' in self.full_bank.columns else set()

        gl_has_real = gl_accs and gl_accs != {'默认账户'}
        bank_has_real = bank_accs and bank_accs != {'默认账户'}

        if (gl_has_real and not bank_has_real) or (bank_has_real and not gl_has_real):
            print("⚠️ 警告：检测到账户信息不对称！")
            print(f"   序时账账户分组: {list(gl_accs)[:5]}")
            print(f"   银行流水账户分组: {list(bank_accs)[:5]}")
            print("   两侧账户标识不交集，可能导致 0% 匹配率。")
            print("   建议：检查序时账是否有账户列，或取消银行流水的账户列名配置。")

        groups = set()
        if not self.full_bank.empty:
            groups.update(zip(self.full_bank['__link_comp__'], self.full_bank['__link_acc__']))
        if not self.full_gl.empty:
            groups.update(zip(self.full_gl['__link_comp__'], self.full_gl['__link_acc__']))
        groups = sorted(list(groups), key=lambda g: tuple(str(x) for x in g))

        if not groups:
            log("❌ 错误：未找到匹配的 [公司+账号] 分组！", 0)
            return pd.DataFrame(), pd.DataFrame()

        log(f"🚀 开始核对，共 {len(groups)} 个核对组", 0)

        for i, (comp, acc) in enumerate(groups):
            log(f"⏳ [进度: {i+1}/{len(groups)}] 处理: [{comp}]-[{acc}]", 1)

            b_idxs = self.full_bank[(self.full_bank['__link_comp__'] == comp) & (self.full_bank['__link_acc__'] == acc)].index
            g_idxs = self.full_gl[(self.full_gl['__link_comp__'] == comp) & (self.full_gl['__link_acc__'] == acc)].index

            if len(b_idxs) == 0 and len(g_idxs) == 0:
                continue

            self._match_group(b_idxs, g_idxs)

        if sum(self.l5_stats.values()) > 0:
            msg = ", ".join([f"{k}:{v}" for k, v in self.l5_stats.items() if v > 0])
            log(f"   -> L5 [模糊匹配]: {msg}", 0)

        # AI 复核（可选）：未匹配少时 dump 供 AI 语义判断；已有 AI 匹配则应用。
        # 必须在 summary 统计之前执行，确保匹配率/金额含 AI 复核匹配。
        if self.ai_review_dir:
            self._handle_ai_review()
            if self.ai_review_pending:
                return None, None

        # summary 统计（在 AI 复核之后，matched_id 已含 AI 标记）
        summary_data = []

        for i, (comp, acc) in enumerate(groups):
            b_idxs = self.full_bank[(self.full_bank['__link_comp__'] == comp) & (self.full_bank['__link_acc__'] == acc)].index
            g_idxs = self.full_gl[(self.full_gl['__link_comp__'] == comp) & (self.full_gl['__link_acc__'] == acc)].index

            if len(b_idxs) == 0 and len(g_idxs) == 0:
                continue

            grp_b = self.full_bank.loc[b_idxs]
            grp_g = self.full_gl.loc[g_idxs]

            inc_col = self.cfg.bank_config.get('收入金额列名')
            exp_col = self.cfg.bank_config.get('支出金额列名')
            b_inc = grp_b[inc_col].apply(safe_float).sum() if inc_col in grp_b.columns else 0
            b_exp = grp_b[exp_col].apply(safe_float).sum() if exp_col in grp_b.columns else 0

            deb_col = self.cfg.gl_config.get('借方金额列名')
            cre_col = self.cfg.gl_config.get('贷方金额列名')
            g_deb = grp_g[deb_col].apply(safe_float).sum() if deb_col in grp_g.columns else 0
            g_cre = grp_g[cre_col].apply(safe_float).sum() if cre_col in grp_g.columns else 0

            b_matched = grp_b[grp_b['matched_id'].notna()]
            b_total_abs = grp_b['__bank_amount__'].abs().sum()
            b_match_rate_amt = (b_matched['__bank_amount__'].abs().sum() / b_total_abs) if b_total_abs > 0 else 0

            g_matched = grp_g[grp_g['matched_id'].notna()]
            g_total_abs = grp_g['__gl_amount__'].abs().sum()
            g_match_rate_amt = (g_matched['__gl_amount__'].abs().sum() / g_total_abs) if g_total_abs > 0 else 0

            b_cnt = len(grp_b)
            g_cnt = len(grp_g)
            b_match_rate_cnt = (len(b_matched) / b_cnt) if b_cnt > 0 else 0
            g_match_rate_cnt = (len(g_matched) / g_cnt) if g_cnt > 0 else 0

            match_rate_amt = min(b_match_rate_amt, g_match_rate_amt)
            match_rate_cnt = min(b_match_rate_cnt, g_match_rate_cnt)

            summary_data.append({
                '公司': comp, '账号': acc,
                '银行流水收入金额': b_inc, '银行流水支出金额': b_exp, '银行流水行数': b_cnt,
                '序时账借方金额': g_deb, '序时账贷方金额': g_cre, '序时账行数': g_cnt,
                '银行流水匹配率(金额)': b_match_rate_amt, '序时账匹配率(金额)': g_match_rate_amt,
                '综合匹配率(金额)': match_rate_amt,
                '银行流水匹配率(条数)': b_match_rate_cnt, '序时账匹配率(条数)': g_match_rate_cnt,
                '综合匹配率(条数)': match_rate_cnt
            })

        df_detail = self._generate_detail_report()
        df_summary = pd.DataFrame(summary_data)

        return df_summary, df_detail

    # ========== AI 复核（可选，分层并行） ==========
    @staticmethod
    def _month_of(v):
        """取 YYYY-MM 月份（未匹配记录按月分批）"""
        if v is None:
            return '未知'
        try:
            if pd.isna(v):
                return '未知'
        except (TypeError, ValueError):
            pass
        s = str(v)
        return s[:7] if len(s) >= 7 else '未知'

    def _handle_ai_review(self):
        """AI 复核两阶段（分层并行）：
        阶段1: 按月分批导出全部未匹配记录供 AI/subagent 判断（不写 Excel）
        阶段2: AI/subagent 汇总写回 ai_matches.json 后，重跑应用匹配并正常输出
        """
        dump_path = os.path.join(self.ai_review_dir, 'ai_unmatched.json')
        match_path = os.path.join(self.ai_review_dir, 'ai_matches.json')
        os.makedirs(self.ai_review_dir, exist_ok=True)

        if os.path.exists(match_path):
            self._apply_ai_matches(match_path)
            return

        bank_um = self.full_bank[self.full_bank['matched_id'].isnull()]
        gl_um = self.full_gl[self.full_gl['matched_id'].isnull()]

        if len(bank_um) == 0 and len(gl_um) == 0:
            log("✅ 全部匹配，无需 AI 复核", 0)
            return

        # 进入门槛：原始未匹配超过阈值则跳过（AI 复核就是为了减少人工复核量，
        # 未匹配记录全量导出、不做任何筛选——数据不全会导致组合判断无法进行）
        if len(bank_um) > self.ai_review_limit or len(gl_um) > self.ai_review_limit:
            log(f"⏭️ 未匹配过多（银行 {len(bank_um)} / 序时账 {len(gl_um)}），"
                f"跳过 AI 复核（进入阈值 {self.ai_review_limit}，先检查范围对齐再重跑）", 0)
            return

        # ---- 全量未匹配记录 ----
        bank_records = []
        for idx, row in bank_um.iterrows():
            amt = row.get('__bank_amount__', 0)
            bank_records.append({
                'idx': int(idx),
                '日期': row.get('__bank_date__', ''),
                '月份': self._month_of(row.get('__bank_date__', '')),
                '金额': abs(float(amt)),
                '方向': '收入' if amt >= 0 else '支出',
                '对手': row.get('__bank_partner_clean__', '') or '',
                '摘要': row.get('__bank_summary__', '') if '__bank_summary__' in row.index else '',
            })

        gl_records = []
        for idx, row in gl_um.iterrows():
            amt = row.get('__gl_amount__', 0)
            gl_records.append({
                'idx': int(idx),
                '日期': row.get('__gl_date__', ''),
                '月份': self._month_of(row.get('__gl_date__', '')),
                '金额': abs(float(amt)),
                '方向': '借方' if amt >= 0 else '贷方',
                '对手': row.get('__gl_partner_clean__', '') or '',
                '凭证号': row.get('__gl_voucher__', '') if '__gl_voucher__' in row.index else '',
                '摘要': row.get('__gl_summary__', '') if '__gl_summary__' in row.index else '',
            })

        # ---- 按月分批导出（供 subagent 并行复核）----
        by_month = {}
        for rec in bank_records + gl_records:
            m = rec['月份']
            if m not in by_month:
                by_month[m] = {'bank': [], 'gl': []}
            by_month[m]['bank' if rec['方向'] in ('收入', '支出') else 'gl'].append(rec)

        base_note = ("AI 复核未匹配记录。idx 是全局行标识，日期+金额+对手+摘要供语义判断。"
                     "判断出可匹配的记录后写 ai_matches.json：{\"matches\": [{\"bank_idx\": <银行行idx> 或 [组合idx...], "
                     "\"gl_idx\": <GL行idx> 或 [组合idx...], \"type\": \"AI:组合推理-高\", \"note\": \"判断依据\"}]}。"
                     "支持 N:1 组合（多笔银行=一笔GL），引擎会校验两侧金额合计相等（容差1元），不相等自动跳过。"
                     "未匹配多或无法判断时可不写该文件直接重跑。")

        month_files = []
        for m in sorted(by_month):
            mb, mg = by_month[m]['bank'], by_month[m]['gl']
            if not mb and not mg:
                continue
            fname = os.path.join(self.ai_review_dir, f'ai_unmatched_{m}.json')
            with open(fname, 'w', encoding='utf-8') as f:
                json.dump({
                    "说明": f"{m} 批次（银行 {len(mb)} 笔 / 序时账 {len(mg)} 笔）。" + base_note,
                    "bank_unmatched": mb,
                    "gl_unmatched": mg,
                }, f, ensure_ascii=False, indent=2)
            month_files.append((m, len(mb), len(mg)))

        data = {
            "说明": "AI 复核未匹配记录总览。未匹配记录全量导出（不做任何规则预筛），按月拆分到 "
                    "ai_unmatched_YYYY-MM.json；批次多时用 subagent 并行复核（每批一个，输出统一合并写入 ai_matches.json）。"
                    + base_note,
            "bank_unmatched": bank_records,
            "gl_unmatched": gl_records,
        }
        with open(dump_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        log(f"📋 未匹配记录已导出供 AI 复核: {dump_path}", 0)
        log(f"   银行 {len(bank_records)} 笔 / 序时账 {len(gl_records)} 笔", 1)
        for m, nb, ng in month_files:
            log(f"   -> {m} 批次: ai_unmatched_{m}.json（银行 {nb} / 序时账 {ng}）", 1)
        log("   💡 批次多时用 subagent 并行复核（每批一个 subagent），结果统一合并写入 ai_matches.json", 1)
        log("   ✋ AI 复核前置确认（用户 gate）：必须先向用户展示核对结果（匹配率/未匹配数）"
            "并征得同意，才开始逐条语义判断、写 ai_matches.json 并重跑本命令", 1)
        self.ai_review_pending = True

    def _apply_ai_matches(self, match_path):
        """应用 AI 复核结果（ai_matches.json），标记新增匹配"""
        try:
            with open(match_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            log(f"⚠️ 读取 AI 匹配文件失败: {e}", 0)
            return

        for m in data.get('matches', []):
            b_idx = m.get('bank_idx')
            g_idx = m.get('gl_idx')
            if b_idx is None or g_idx is None:
                continue
            b_ids = b_idx if isinstance(b_idx, list) else [b_idx]
            g_ids = g_idx if isinstance(g_idx, list) else [g_idx]
            if not b_ids or not g_ids:
                continue
            b_ids = [int(x) for x in b_ids]
            g_ids = [int(x) for x in g_ids]

            # 所有行必须存在且均未匹配
            valid = all(x in self.full_bank.index for x in b_ids) and \
                    all(x in self.full_gl.index for x in g_ids)
            if not valid:
                log(f"⚠️ AI 匹配行索引无效，跳过: bank={b_ids} gl={g_ids}", 0)
                continue
            if any(pd.notna(self.full_bank.loc[x, 'matched_id']) for x in b_ids) or \
               any(pd.notna(self.full_gl.loc[x, 'matched_id']) for x in g_ids):
                log(f"⚠️ AI 匹配行已被引擎匹配，跳过: bank={b_ids} gl={g_ids}", 0)
                continue

            # 金额硬校验：银行侧合计 ≈ GL侧合计（容差 1 元）
            b_sum = sum(abs(float(self.full_bank.loc[x, '__bank_amount__'])) for x in b_ids)
            g_sum = sum(abs(float(self.full_gl.loc[x, '__gl_amount__'])) for x in g_ids)
            if abs(b_sum - g_sum) > 1.0:
                log(f"⚠️ AI 匹配金额不等（银行 {b_sum:.2f} vs GL {g_sum:.2f}），跳过: bank={b_ids} gl={g_ids}", 0)
                continue

            desc = m.get('type', 'AI:人工判断')
            self._mark(b_ids, g_ids, desc)
            self.ai_applied += 1

        log(f"✅ 已应用 AI 复核匹配 {self.ai_applied} 笔", 0)

    def _match_group(self, b_idxs, g_idxs):
        """执行所有9层匹配"""
        self._l1_partner_exact(b_idxs, g_idxs)
        self._l2_summary_exact(b_idxs, g_idxs)
        self._l3_smart_split(b_idxs, g_idxs)
        self._l4_partner_aggregation(b_idxs, g_idxs)
        self._l5_fuzzy_semantic(b_idxs, g_idxs)
        self._l6_same_month_amount(b_idxs, g_idxs)
        self._l7_partner_aggregation_to_single(b_idxs, g_idxs)
        self._l7_nm_aggregation(b_idxs, g_idxs)
        if self._l7_total > 0:
            log(f"   -> L7 [同月金额-定向凑数]: {self._l7_total} 组", 0)
            self._l7_total = 0
        self._l8_global_brute_force(b_idxs, g_idxs)
        self._l8_reverse_brute_force(b_idxs, g_idxs)
        if self._l8_total > 0:
            log(f"   -> L8 [同月金额-凑数]: {self._l8_total}", 0)
            self._l8_total = 0
        self._l9_cross_month_boundary(b_idxs, g_idxs)

    def _mark(self, b_ids, g_ids, desc):
        """标记匹配记录"""
        self.match_counter += 1
        mid = f"M{self.match_counter:05d}"
        if len(b_ids) > 0:
            self.full_bank.loc[b_ids, 'matched_id'] = mid
            self.full_bank.loc[b_ids, 'match_desc'] = desc
        if len(g_ids) > 0:
            self.full_gl.loc[g_ids, 'matched_id'] = mid
            self.full_gl.loc[g_ids, 'match_desc'] = desc

    # ========== L1: 交易对手精准匹配 ==========
    def _l1_partner_exact(self, b_idxs, g_idxs):
        valid_g = self.full_gl.loc[g_idxs]
        valid_g = valid_g[valid_g['matched_id'].isnull() & valid_g['__gl_partner_clean__'].notna() & (valid_g['__gl_partner_clean__'] != '')]
        if valid_g.empty:
            return

        # Q1: 全局排序贪心 — 收集所有候选对，按匹配质量排序后贪心分配
        # Q2: 双向匹配 — gl_partner in b_partner 或 b_partner in gl_partner
        all_pairs = []
        for g_idx, g_row in valid_g.iterrows():
            if pd.notna(self.full_gl.loc[g_idx, 'matched_id']):
                continue
            partner = g_row['__gl_partner_clean__']
            amt = g_row['__gl_amount__']
            mon = g_row['__month__']
            candidates = self.full_bank.loc[b_idxs]
            candidates = candidates[
                candidates['matched_id'].isnull() &
                (candidates['__month__'] == mon) &
                (abs(candidates['__bank_amount__'] - amt) < 0.005)
            ]
            for b_idx, b_row in candidates.iterrows():
                if pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                    continue
                b_partner = b_row['__bank_partner_clean__']
                if not b_partner or pd.isna(b_partner):
                    continue
                # Q2: 双向包含检查
                if partner in b_partner:
                    # 正向：GL对手是Bank对手的子串（原有逻辑）
                    if partner == b_partner:
                        score = 3.0  # 精确匹配优先级最高
                    else:
                        score = 2.0 + len(partner) / max(len(b_partner), 1)
                    all_pairs.append((g_idx, b_idx, score))
                elif len(b_partner) >= 2 and b_partner in partner:
                    # 反向：Bank对手是GL对手的子串（Q2新增，要求≥2字防误匹配）
                    score = 1.0 + len(b_partner) / max(len(partner), 1)
                    all_pairs.append((g_idx, b_idx, score))

        if not all_pairs:
            return

        # 按分数降序排序，同分按索引稳定排序
        all_pairs.sort(key=lambda x: (-x[2], x[0], x[1]))
        used_g = set()
        used_b = set()
        match_cnt = 0
        for g_idx, b_idx, score in all_pairs:
            if g_idx in used_g or b_idx in used_b:
                continue
            if pd.notna(self.full_gl.loc[g_idx, 'matched_id']) or pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                continue
            self._mark([b_idx], [g_idx], "L1:交易对手-精准")
            used_g.add(g_idx)
            used_b.add(b_idx)
            match_cnt += 1
        if match_cnt:
            log(f"   -> L1 [交易对手-精准]: {match_cnt}", 0)

    # ========== L2: 摘要精准匹配 ==========
    def _l2_summary_exact(self, b_idxs, g_idxs):
        valid_g = self.full_gl.loc[g_idxs]
        valid_g = valid_g[valid_g['matched_id'].isnull() & valid_g['__gl_summary__'].notna() & (valid_g['__gl_summary__'] != '')]
        group_used_bank = set()
        match_cnt = 0
        for g_idx, g_row in valid_g.iterrows():
            if pd.notna(self.full_gl.loc[g_idx, 'matched_id']):
                continue
            summary = g_row['__gl_summary__']
            amt = g_row['__gl_amount__']
            mon = g_row['__month__']
            candidates = self.full_bank.loc[b_idxs]
            candidates = candidates[
                candidates['matched_id'].isnull() &
                (~candidates.index.isin(group_used_bank)) &
                (candidates['__month__'] == mon) &
                (abs(candidates['__bank_amount__'] - amt) < 0.005)
            ]
            if candidates.empty:
                continue
            best_b_idx = None
            for b_idx, b_row in candidates.iterrows():
                b_name = b_row['__bank_partner_clean__']
                if b_name and len(b_name) > 1 and b_name in summary:
                    best_b_idx = b_idx
                    break
            if best_b_idx is not None:
                self._mark([best_b_idx], [g_idx], "L2:摘要-精准")
                group_used_bank.add(best_b_idx)
                match_cnt += 1
        if match_cnt:
            log(f"   -> L2 [摘要-精准]: {match_cnt}", 0)

    # ========== L3: 同客商定向凑数 ==========
    def _l3_smart_split(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()]
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()]
        if curr_b.empty or curr_g.empty:
            return
        curr_g = curr_g[curr_g['__gl_partner_clean__'] != '']
        if curr_g.empty:
            return
        match_cnt = 0
        used_b_local = set()
        used_g_local = set()

        def is_same_entity(b_row, partner, g_summ):
            b_name = b_row['__bank_partner_clean__']
            if not b_name:
                return False
            if partner in b_name or b_name in partner:
                return True
            if g_summ and b_name in str(g_summ):
                return True
            return False

        for m, g_grp_month in curr_g.groupby('__month__'):
            b_grp_month = curr_b[curr_b['__month__'] == m]
            if b_grp_month.empty:
                continue
            for g_idx, g_row in g_grp_month.iterrows():
                if g_idx in used_g_local or pd.notna(self.full_gl.loc[g_idx, 'matched_id']):
                    continue
                partner = g_row['__gl_partner_clean__']
                target_amt = g_row['__gl_amount__']
                g_summ = g_row['__gl_summary__']
                candidates = b_grp_month[(~b_grp_month.index.isin(used_b_local)) & (b_grp_month.apply(lambda r: is_same_entity(r, partner, g_summ), axis=1))]
                if candidates.empty:
                    continue
                matched_subset = self._find_subset_match(candidates, target_amt)
                if matched_subset is not None:
                    self._execute_time_match(matched_subset.index, [g_idx], "L3:同客商-定向凑数", used_b_local, used_g_local)
                    match_cnt += 1
            b_rem = b_grp_month[~b_grp_month.index.isin(used_b_local)]
            g_rem = g_grp_month[~g_grp_month.index.isin(used_g_local)]
            if b_rem.empty or g_rem.empty:
                continue
            for b_idx, b_row in b_rem.iterrows():
                b_p = b_row['__bank_partner_clean__']
                if not b_p:
                    continue
                target_amt = b_row['__bank_amount__']
                # NaN-safe partner matching (Fix 6)
                partner_series = g_rem['__gl_partner_clean__'].fillna('')
                candidates = g_rem[partner_series.str.contains(b_p, regex=False)]
                if candidates.empty:
                    continue
                matched_subset = self._find_subset_match(candidates, target_amt, col_name='__gl_amount__')
                if matched_subset is not None:
                    self._execute_time_match([b_idx], matched_subset.index, "L3:同客商-定向凑数", used_b_local, used_g_local)
                    match_cnt += 1
        if match_cnt:
            log(f"   -> L3 [同客商-定向凑数]: {match_cnt} 组", 0)

    def _find_subset_match(self, df, target, col_name='__bank_amount__'):
        """查找子集金额匹配"""
        if abs(df[col_name].sum() - target) < 0.005:
            return df
        for d, grp in df.groupby('__dt__'):
            if abs(grp[col_name].sum() - target) < 0.005:
                return grp
        if len(df) < 12:
            idxs = df.index.tolist()
            cnt = 0
            for r in range(2, min(len(df), 6)):
                for combo in itertools.combinations(idxs, r):
                    cnt += 1
                    if cnt > 8000:
                        break
                    sub = df.loc[list(combo)]
                    if abs(sub[col_name].sum() - target) < 0.005:
                        dates = sub['__dt__']
                        if (dates.max() - dates.min()).days <= 3:
                            return sub
                if cnt > 8000:
                    break
        return None

    def _execute_time_match(self, b_idxs, g_idxs, type_name, used_b_set, used_g_set):
        """执行带时间约束的匹配"""
        all_dates = pd.concat([self.full_gl.loc[g_idxs, '__dt__'], self.full_bank.loc[b_idxs, '__dt__']])
        min_dt, max_dt = all_dates.min(), all_dates.max()
        diff_days = (max_dt - min_dt).days if pd.notna(min_dt) else 99
        sub_type = f"{type_name}(同月)"
        if diff_days == 0:
            sub_type = f"{type_name}(同日)"
        elif diff_days <= 3:
            sub_type = f"{type_name}(3日内)"
        self._mark(b_idxs, g_idxs, sub_type)
        used_b_set.update(b_idxs)
        used_g_set.update(g_idxs)

    # ========== L4: 同客商多对多汇总 ==========
    def _l4_partner_aggregation(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()]
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()]
        if curr_b.empty or curr_g.empty:
            return
        curr_g = curr_g[curr_g['__gl_partner_clean__'] != '']
        if curr_g.empty:
            return
        match_cnt = 0
        used_b_local = set()

        def is_same_entity(b_row, partner, g_summs):
            b_name = b_row['__bank_partner_clean__']
            if not b_name:
                return False
            if partner in b_name or b_name in partner:
                return True
            for s in g_summs:
                if s and b_name in str(s):
                    return True
            return False

        all_dates = sorted(curr_g['__dt__'].unique())
        for d in all_dates:
            g_day = curr_g[curr_g['__dt__'] == d]
            b_day = curr_b[curr_b['__dt__'] == d]
            if g_day.empty or b_day.empty:
                continue
            unique_partners = g_day['__gl_partner_clean__'].unique()
            for partner in unique_partners:
                g_sub = g_day[g_day['__gl_partner_clean__'] == partner]
                if pd.notna(self.full_gl.loc[g_sub.index[0], 'matched_id']):
                    continue
                g_summs = g_sub['__gl_summary__'].tolist()
                b_candidates = b_day[(~b_day.index.isin(used_b_local)) & (b_day.apply(lambda r: is_same_entity(r, partner, g_summs), axis=1))]
                if b_candidates.empty:
                    continue
                if abs(g_sub['__gl_amount__'].sum() - b_candidates['__bank_amount__'].sum()) < 0.005:
                    self._execute_time_match(b_candidates.index, g_sub.index, "L4:同客商-汇总", used_b_local, set())
                    match_cnt += 1
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()]
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()]
        curr_g = curr_g[curr_g['__gl_partner_clean__'] != '']
        for m, g_grp_month in curr_g.groupby('__month__'):
            b_grp_month = curr_b[curr_b['__month__'] == m]
            if b_grp_month.empty:
                continue
            unique_partners = g_grp_month['__gl_partner_clean__'].unique()
            for partner in unique_partners:
                g_sub = g_grp_month[g_grp_month['__gl_partner_clean__'] == partner]
                if g_sub.empty or pd.notna(self.full_gl.loc[g_sub.index[0], 'matched_id']):
                    continue
                g_summs = g_sub['__gl_summary__'].tolist()
                b_candidates = b_grp_month[(~b_grp_month.index.isin(used_b_local)) & (b_grp_month.apply(lambda r: is_same_entity(r, partner, g_summs), axis=1))]
                if b_candidates.empty:
                    continue
                if abs(g_sub['__gl_amount__'].sum() - b_candidates['__bank_amount__'].sum()) < 0.005:
                    self._execute_time_match(b_candidates.index, g_sub.index, "L4:同客商-汇总", used_b_local, set())
                    match_cnt += 1
        if match_cnt:
            log(f"   -> L4 [同客商-多对多汇总]: {match_cnt} 组", 0)

    # ========== L5: 模糊语义匹配 ==========
    def _l5_fuzzy_semantic(self, b_idxs, g_idxs):
        curr_b_all = self.full_bank.loc[b_idxs]
        curr_g_all = self.full_gl.loc[g_idxs]
        for m, grp_b in curr_b_all.groupby('__month__'):
            grp_g = curr_g_all[curr_g_all['__month__'] == m]
            if grp_g.empty:
                continue
            unmatched_b = grp_b[grp_b['matched_id'].isnull()]
            unmatched_g = grp_g[grp_g['matched_id'].isnull()]
            if unmatched_b.empty or unmatched_g.empty:
                continue
            self._run_fuzzy_core(unmatched_b, unmatched_g)

    def _run_fuzzy_core(self, curr_b, curr_g):
        g_has_p = curr_g[curr_g['__gl_partner_clean__'].notna() & (curr_g['__gl_partner_clean__'] != '')]
        if not g_has_p.empty:
            self._fuzzy_matrix_match(curr_b, g_has_p, "交易对手-模糊")

        curr_g_rem = curr_g[curr_g['matched_id'].isnull()]
        if not curr_g_rem.empty:
            self._fuzzy_matrix_match(curr_b, curr_g_rem, "摘要-模糊", use_summary=True)

    def _fuzzy_matrix_match(self, df_b, df_g, label_prefix, use_summary=False):
        df_b = df_b[df_b['matched_id'].isnull()]
        df_g = df_g[df_g['matched_id'].isnull()]
        if df_b.empty or df_g.empty:
            return

        # Q1: 全局排序贪心 — 收集所有候选对的模糊分数，全局排序后贪心分配
        all_pairs = []
        for g_idx, g_row in df_g.iterrows():
            if pd.notna(self.full_gl.loc[g_idx, 'matched_id']):
                continue

            g_text = g_row['__gl_summary__'] if use_summary else g_row['__gl_partner_clean__']
            g_amt = g_row['__gl_amount__']

            candidates = df_b[abs(df_b['__bank_amount__'] - g_amt) < 0.005]
            if candidates.empty:
                continue

            for b_idx, b_row in candidates.iterrows():
                if pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                    continue
                b_text = b_row['__bank_partner_clean__']

                score = self.sim.compute_similarity(g_text, b_text)
                if score > 0.65:  # 最低阈值
                    all_pairs.append((g_idx, b_idx, score))

        if not all_pairs:
            return

        # 按分数降序排序，同分按索引稳定排序
        all_pairs.sort(key=lambda x: (-x[2], x[0], x[1]))
        used_g = set()
        used_b = set()
        for g_idx, b_idx, score in all_pairs:
            if g_idx in used_g or b_idx in used_b:
                continue
            if pd.notna(self.full_gl.loc[g_idx, 'matched_id']) or pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                continue

            level = ""
            if score > 0.92:
                level = "高"
            elif score > 0.80:
                level = "中"
            elif score > 0.65:
                level = "低"

            if level:
                label = f"L5:{label_prefix}({level})"
                self._mark([b_idx], [g_idx], label)
                self.l5_stats[level] += 1
                used_g.add(g_idx)
                used_b.add(b_idx)

    # ========== L6: 同月金额独一匹配 ==========
    def _l6_same_month_amount(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()].copy()
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()].copy()
        if len(curr_b) == 0 or len(curr_g) == 0:
            return
        match_cnt = 0
        used_g_local = set()
        for m, grp_b in curr_b.groupby('__month__'):
            grp_g = curr_g[curr_g['__month__'] == m]
            if grp_g.empty:
                continue
            for b_idx, b_row in grp_b.iterrows():
                if pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                    continue
                target_amt = b_row['__bank_amount__']
                candidates = grp_g[(grp_g['matched_id'].isnull()) & (~grp_g.index.isin(used_g_local)) & (abs(grp_g['__gl_amount__'] - target_amt) < 0.005)]
                if not candidates.empty:
                    best_g_idx = candidates.index[0]
                    self._mark([b_idx], [best_g_idx], "L6:同月金额-独一")
                    used_g_local.add(best_g_idx)
                    match_cnt += 1
        if match_cnt:
            log(f"   -> L6 [同月金额-独一]: {match_cnt}", 0)

    # ========== L7: 同月金额定向凑数（多对单） ==========
    def _l7_partner_aggregation_to_single(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()].copy()
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()].copy()

        if curr_b.empty or curr_g.empty:
            return

        curr_b_with_partner = curr_b[curr_b['__bank_partner_clean__'].notna() & (curr_b['__bank_partner_clean__'] != '')]
        if curr_b_with_partner.empty:
            return

        match_cnt = 0
        used_b_local = set()
        used_g_local = set()

        for m, grp_b_month in curr_b_with_partner.groupby('__month__'):
            grp_g_month = curr_g[curr_g['__month__'] == m]
            if grp_g_month.empty:
                continue

            unique_partners = grp_b_month[~grp_b_month.index.isin(used_b_local)]['__bank_partner_clean__'].unique()

            for partner in unique_partners:
                partner_flows = grp_b_month[
                    (grp_b_month['__bank_partner_clean__'] == partner) &
                    (~grp_b_month.index.isin(used_b_local))
                ]
                if partner_flows.empty:
                    continue

                # 策略1: 按相同日期汇总
                for date, date_flows in partner_flows.groupby('__dt__'):
                    if date_flows.index.isin(used_b_local).any():
                        continue
                    total_amt = date_flows['__bank_amount__'].sum()

                    candidates = grp_g_month[
                        (~grp_g_month.index.isin(used_g_local)) &
                        (abs(grp_g_month['__gl_amount__'] - total_amt) < 0.005)
                    ]
                    if not candidates.empty:
                        best_g_idx = candidates.index[0]
                        self._mark(date_flows.index.tolist(), [best_g_idx], f"L7:同月金额-定向凑数(同日)")
                        used_b_local.update(date_flows.index)
                        used_g_local.add(best_g_idx)
                        match_cnt += 1
                        continue

                # 策略2: 按近3天汇总
                partner_flows_remaining = partner_flows[~partner_flows.index.isin(used_b_local)]
                if not partner_flows_remaining.empty:
                    dates_sorted = sorted(partner_flows_remaining['__dt__'].unique())
                    for i, start_date in enumerate(dates_sorted):
                        end_date = start_date + pd.Timedelta(days=2)
                        window_flows = partner_flows_remaining[
                            (partner_flows_remaining['__dt__'] >= start_date) &
                            (partner_flows_remaining['__dt__'] <= end_date) &
                            (~partner_flows_remaining.index.isin(used_b_local))
                        ]
                        if len(window_flows) < 2:
                            continue

                        total_amt = window_flows['__bank_amount__'].sum()
                        candidates = grp_g_month[
                            (~grp_g_month.index.isin(used_g_local)) &
                            (abs(grp_g_month['__gl_amount__'] - total_amt) < 0.005)
                        ]
                        if not candidates.empty:
                            best_g_idx = candidates.index[0]
                            self._mark(window_flows.index.tolist(), [best_g_idx], f"L7:同月金额-定向凑数(3日内)")
                            used_b_local.update(window_flows.index)
                            used_g_local.add(best_g_idx)
                            match_cnt += 1

                # 策略3: 按整月汇总
                partner_flows_remaining = partner_flows[~partner_flows.index.isin(used_b_local)]
                if len(partner_flows_remaining) >= 2:
                    total_amt = partner_flows_remaining['__bank_amount__'].sum()
                    candidates = grp_g_month[
                        (~grp_g_month.index.isin(used_g_local)) &
                        (abs(grp_g_month['__gl_amount__'] - total_amt) < 0.005)
                    ]
                    if not candidates.empty:
                        best_g_idx = candidates.index[0]
                        self._mark(partner_flows_remaining.index.tolist(), [best_g_idx], f"L7:同月金额-定向凑数(同月)")
                        used_b_local.update(partner_flows_remaining.index)
                        used_g_local.add(best_g_idx)
                        match_cnt += 1

        if match_cnt:
            self._l7_total += match_cnt

    # ========== L7: N:M 汇总匹配 ==========
    def _l7_nm_aggregation(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()].copy()
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()].copy()
        if curr_b.empty or curr_g.empty:
            return

        curr_b_with_partner = curr_b[curr_b['__bank_partner_clean__'].notna() & (curr_b['__bank_partner_clean__'] != '')]
        if curr_b_with_partner.empty:
            return

        match_cnt = 0
        used_b_local = set()
        used_g_local = set()

        def find_g_combo(b_subset, g_available):
            """在序时账中找金额总和等于 b_subset 总额的组合"""
            b_total = b_subset['__bank_amount__'].sum()
            g_pool = g_available.to_dict('records')
            g_pool_idxs = g_available.index.tolist()
            if not g_pool_idxs:
                return None
            if len(g_pool_idxs) > 20:
                return None
            cnt = 0
            for r in range(1, min(len(g_pool_idxs) + 1, 8)):
                for combo in itertools.combinations(range(len(g_pool_idxs)), r):
                    cnt += 1
                    if cnt > 8000:
                        return None
                    if abs(sum(g_pool[i]['__gl_amount__'] for i in combo) - b_total) < 0.005:
                        return [g_pool_idxs[i] for i in combo]
            return None

        for m, grp_b_month in curr_b_with_partner.groupby('__month__'):
            grp_g_month = curr_g[curr_g['__month__'] == m]
            if grp_g_month.empty:
                continue

            unique_partners = grp_b_month[~grp_b_month.index.isin(used_b_local)]['__bank_partner_clean__'].unique()

            for partner in unique_partners:
                partner_flows = grp_b_month[
                    (grp_b_month['__bank_partner_clean__'] == partner) &
                    (~grp_b_month.index.isin(used_b_local))
                ]
                if len(partner_flows) < 2:
                    continue

                g_available = grp_g_month[~grp_g_month.index.isin(used_g_local)]
                if g_available.empty:
                    continue

                # 策略1: 按同日子组匹配
                for dt, day_flows in partner_flows.groupby('__dt__'):
                    if day_flows.index.isin(used_b_local).any():
                        continue
                    if len(day_flows) < 2:
                        continue
                    g_avail = g_available[~g_available.index.isin(used_g_local)]
                    matched_g = find_g_combo(day_flows, g_avail)
                    if matched_g is not None:
                        self._execute_time_match(day_flows.index.tolist(), matched_g, "L7:同月金额-定向凑数", used_b_local, used_g_local)
                        match_cnt += 1

                # 策略2: 按3日窗口子组匹配
                remaining = partner_flows[~partner_flows.index.isin(used_b_local)]
                if len(remaining) >= 2:
                    dates_sorted = sorted(remaining['__dt__'].unique())
                    for start_date in dates_sorted:
                        end_date = start_date + pd.Timedelta(days=2)
                        window = remaining[
                            (remaining['__dt__'] >= start_date) &
                            (remaining['__dt__'] <= end_date) &
                            (~remaining.index.isin(used_b_local))
                        ]
                        if len(window) < 2:
                            continue
                        g_avail = grp_g_month[~grp_g_month.index.isin(used_g_local)]
                        matched_g = find_g_combo(window, g_avail)
                        if matched_g is not None:
                            self._execute_time_match(window.index.tolist(), matched_g, "L7:同月金额-定向凑数", used_b_local, used_g_local)
                            match_cnt += 1

                # 策略3: 剩余全量兜底
                remaining = partner_flows[~partner_flows.index.isin(used_b_local)]
                if len(remaining) >= 2:
                    g_avail = grp_g_month[~grp_g_month.index.isin(used_g_local)]
                    matched_g = find_g_combo(remaining, g_avail)
                    if matched_g is not None:
                        self._execute_time_match(remaining.index.tolist(), matched_g, "L7:同月金额-定向凑数", used_b_local, used_g_local)
                        match_cnt += 1

        if match_cnt:
            self._l7_total += match_cnt

    # ========== L8: 全局暴力凑数 ==========
    def _l8_global_brute_force(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()].copy()
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()].copy()
        if len(curr_b) == 0 or len(curr_g) == 0:
            return
        match_cnt = 0
        for m, grp_b in curr_b.groupby('__month__'):
            grp_g = curr_g[curr_g['__month__'] == m]
            if grp_g.empty:
                continue
            g_pool = grp_g.to_dict('records')
            g_pool_idxs = grp_g.index.tolist()
            for b_idx, b_row in grp_b.iterrows():
                if pd.notna(self.full_bank.loc[b_idx, 'matched_id']):
                    continue
                target = b_row['__bank_amount__']
                candidates = []
                for i, g in enumerate(g_pool):
                    if pd.notna(self.full_gl.loc[g_pool_idxs[i], 'matched_id']):
                        continue
                    candidates.append(i)
                candidates.sort(key=lambda i: abs((g_pool[i]['__dt__'] - b_row['__dt__']).days))
                candidates = candidates[:40]
                found = False
                cnt = 0
                for r in range(1, 6):
                    if found:
                        break
                    for combo in itertools.combinations(candidates, r):
                        cnt += 1
                        if cnt > 8000:
                            break
                        if abs(sum(g_pool[i]['__gl_amount__'] for i in combo) - target) < 0.005:
                            real_g_idxs = [g_pool_idxs[i] for i in combo]
                            self._execute_time_match([b_idx], real_g_idxs, "L8:同月金额-凑数", set(), set())
                            found = True
                            match_cnt += 1
                            break
                    if cnt > 8000:
                        break
        if match_cnt:
            self._l8_total += match_cnt

    # ========== L8 反向: 序时账单笔凑数 ==========
    def _l8_reverse_brute_force(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()].copy()
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()].copy()
        if curr_b.empty or curr_g.empty:
            return
        match_cnt = 0
        for m, grp_g in curr_g.groupby('__month__'):
            grp_b = curr_b[curr_b['__month__'] == m]
            if grp_b.empty:
                continue
            b_pool = grp_b.to_dict('records')
            b_pool_idxs = grp_b.index.tolist()
            for g_idx, g_row in grp_g.iterrows():
                if pd.notna(self.full_gl.loc[g_idx, 'matched_id']):
                    continue
                target = g_row['__gl_amount__']
                candidates = []
                for i, b in enumerate(b_pool):
                    if pd.notna(self.full_bank.loc[b_pool_idxs[i], 'matched_id']):
                        continue
                    candidates.append(i)
                candidates.sort(key=lambda i: abs((b_pool[i]['__dt__'] - g_row['__dt__']).days))
                candidates = candidates[:40]
                found = False
                cnt = 0
                for r in range(2, 6):
                    if found:
                        break
                    for combo in itertools.combinations(candidates, r):
                        cnt += 1
                        if cnt > 8000:
                            break
                        if abs(sum(b_pool[i]['__bank_amount__'] for i in combo) - target) < 0.005:
                            real_b_idxs = [b_pool_idxs[i] for i in combo]
                            self._execute_time_match(real_b_idxs, [g_idx], "L8:同月金额-凑数", set(), set())
                            found = True
                            match_cnt += 1
                            break
                    if cnt > 8000:
                        break
        if match_cnt:
            self._l8_total += match_cnt

    # ========== L9: 跨月边界匹配 ==========
    def _l9_cross_month_boundary(self, b_idxs, g_idxs):
        curr_b = self.full_bank.loc[b_idxs]
        curr_b = curr_b[curr_b['matched_id'].isnull()]
        curr_g = self.full_gl.loc[g_idxs]
        curr_g = curr_g[curr_g['matched_id'].isnull()]
        if curr_b.empty or curr_g.empty:
            return

        b_day = curr_b['__dt__'].dt.day
        g_day = curr_g['__dt__'].dt.day
        b_boundary_idxs = curr_b[(b_day <= 3) | (b_day >= 28)].index
        g_boundary_idxs = curr_g[(g_day <= 3) | (g_day >= 28)].index
        if len(b_boundary_idxs) == 0 or len(g_boundary_idxs) == 0:
            return

        counter_before = self.match_counter

        # 轮1: 银行流水月初移到上月
        b_early_idxs = self.full_bank.loc[b_boundary_idxs]
        b_early_idxs = b_early_idxs[b_early_idxs['__dt__'].dt.day <= 3].index
        if len(b_early_idxs) > 0:
            b_orig = self.full_bank.loc[b_early_idxs, '__month__'].copy()
            for idx in b_early_idxs:
                dt = self.full_bank.loc[idx, '__dt__']
                if pd.notna(dt):
                    self.full_bank.loc[idx, '__month__'] = (dt - pd.DateOffset(months=1)).to_period('M')
            self._run_l1_l8(b_boundary_idxs, g_boundary_idxs)
            self.full_bank.loc[b_orig.index, '__month__'] = b_orig

        # 轮2: 序时账月初移到上月
        g_early_idxs = self.full_gl.loc[g_boundary_idxs]
        g_early_idxs = g_early_idxs[g_early_idxs['__dt__'].dt.day <= 3].index
        if len(g_early_idxs) > 0:
            g_orig = self.full_gl.loc[g_early_idxs, '__month__'].copy()
            for idx in g_early_idxs:
                dt = self.full_gl.loc[idx, '__dt__']
                if pd.notna(dt):
                    self.full_gl.loc[idx, '__month__'] = (dt - pd.DateOffset(months=1)).to_period('M')
            self._run_l1_l8(b_boundary_idxs, g_boundary_idxs)
            self.full_gl.loc[g_orig.index, '__month__'] = g_orig

        # 将新匹配的 match_desc 加上 "L9:" 前缀
        new_match_cnt = self.match_counter - counter_before
        if new_match_cnt > 0:
            for i in range(counter_before + 1, self.match_counter + 1):
                mid = f"M{i:05d}"
                b_mask = self.full_bank['matched_id'] == mid
                g_mask = self.full_gl['matched_id'] == mid
                if b_mask.any():
                    old_desc = self.full_bank.loc[b_mask, 'match_desc'].iloc[0]
                    self.full_bank.loc[b_mask, 'match_desc'] = f"L9:跨月-{old_desc}"
                if g_mask.any():
                    old_desc = self.full_gl.loc[g_mask, 'match_desc'].iloc[0]
                    self.full_gl.loc[g_mask, 'match_desc'] = f"L9:跨月-{old_desc}"
            log(f"   -> L9 [跨月边界匹配]: {new_match_cnt} (复用L1-L8)", 0)

    def _run_l1_l8(self, b_idxs, g_idxs):
        """复用 L1-L8 匹配逻辑的辅助方法"""
        self._l1_partner_exact(b_idxs, g_idxs)
        self._l2_summary_exact(b_idxs, g_idxs)
        self._l3_smart_split(b_idxs, g_idxs)
        self._l4_partner_aggregation(b_idxs, g_idxs)
        self._l5_fuzzy_semantic(b_idxs, g_idxs)
        self._l6_same_month_amount(b_idxs, g_idxs)
        self._l7_partner_aggregation_to_single(b_idxs, g_idxs)
        self._l7_nm_aggregation(b_idxs, g_idxs)
        self._l8_global_brute_force(b_idxs, g_idxs)
        self._l8_reverse_brute_force(b_idxs, g_idxs)

    def get_partner_check_status(self, match_type, b_partner, g_partner):
        """获取交易对手核对状态"""
        if not match_type or match_type == '未匹配':
            return "未匹配"
        b_p = str(b_partner).strip() if pd.notna(b_partner) else ""
        g_p = str(g_partner).strip() if pd.notna(g_partner) else ""
        if b_p and g_p:
            if b_p == g_p or b_p in g_p or g_p in b_p:
                return "一致"
            sim_score = self.sim.compute_similarity(b_p, g_p)
            return "一致" if sim_score > 0.65 else "不一致"
        level = match_type.split(':')[0]
        # L1/L3/L4/L5 匹配以交易对手为核心条件，匹配成功即说明对手一致
        if level in ['L1', 'L3', 'L4', 'L5']:
            return "一致"
        if not b_p and not g_p:
            return "无交易对手"
        # 单面无对手（含 AI 复核/摘要匹配）：对手未参与匹配，无法核对
        return "无交易对手"

    def _generate_detail_report(self):
        log("📋 正在生成[核对明细]...", 0)
        bank_raw_cols = [c for c in self.full_bank.columns if not str(c).startswith('__') and c not in ['matched_id', 'match_desc']]
        gl_raw_cols = [c for c in self.full_gl.columns if not str(c).startswith('__') and c not in ['matched_id', 'match_desc']]

        output_rows = []
        all_mids = set(self.full_bank['matched_id'].dropna()) | set(self.full_gl['matched_id'].dropna())

        for mid in all_mids:
            b_sub = self.full_bank[self.full_bank['matched_id'] == mid]
            g_sub = self.full_gl[self.full_gl['matched_id'] == mid]

            dates = []
            if not b_sub.empty:
                dates.append(b_sub.iloc[0]['__dt__'])
            if not g_sub.empty:
                dates.append(g_sub.iloc[0]['__dt__'])
            sort_date = min(dates) if dates else pd.Timestamp.max

            gl_voucher_sort = g_sub.iloc[0]['__gl_voucher__'] if not g_sub.empty else ''
            base = b_sub.iloc[0] if not b_sub.empty else g_sub.iloc[0]
            max_len = max(len(b_sub), len(g_sub))
            m_type = base['match_desc']

            b_p_ref = next((x for x in b_sub['__bank_partner_clean__'] if x), "") if not b_sub.empty else ""
            g_p_ref = next((x for x in g_sub['__gl_partner_clean__'] if x), "") if not g_sub.empty else ""
            check_status = self.get_partner_check_status(m_type, b_p_ref, g_p_ref)

            for i in range(max_len):
                row = {
                    '__sort_comp__': base['__link_comp__'],
                    '__sort_acc__': base['__link_acc__'],
                    '__sort_date__': sort_date,
                    '__sort_vouch__': gl_voucher_sort,
                    '__is_matched__': 0,
                    '基础信息_公司': base['__link_comp__'],
                    '基础信息_账号': base['__link_acc__'],
                    '核对结果_核对状态': '已匹配',
                    '核对结果_匹配编号': mid,
                    '核对结果_匹配类型': m_type,
                    '核对结果_交易对手核对': check_status
                }
                if i < len(b_sub):
                    for c in bank_raw_cols:
                        row[f'银行流水_{c}'] = b_sub.iloc[i][c]
                if i < len(g_sub):
                    for c in gl_raw_cols:
                        row[f'序时账_{c}'] = g_sub.iloc[i][c]
                    row['辅助信息_计算交易对手(原始)'] = g_sub.iloc[i]['__gl_partner_raw__']
                    row['辅助信息_拆分说明'] = g_sub.iloc[i]['__split_status__']
                output_rows.append(row)

        un_b = self.full_bank[self.full_bank['matched_id'].isnull()]
        for _, r in un_b.iterrows():
            row = {
                '__sort_comp__': r['__link_comp__'],
                '__sort_acc__': r['__link_acc__'],
                '__sort_date__': r['__dt__'],
                '__sort_vouch__': '',
                '__is_matched__': 1,
                '基础信息_公司': r['__link_comp__'],
                '基础信息_账号': r['__link_acc__'],
                '核对结果_核对状态': '未匹配',
                '核对结果_匹配编号': None,
                '核对结果_匹配类型': '未匹配',
                '核对结果_交易对手核对': '未匹配'
            }
            for c in bank_raw_cols:
                row[f'银行流水_{c}'] = r[c]
            output_rows.append(row)

        un_g = self.full_gl[self.full_gl['matched_id'].isnull()]
        for _, r in un_g.iterrows():
            row = {
                '__sort_comp__': r['__link_comp__'],
                '__sort_acc__': r['__link_acc__'],
                '__sort_date__': r['__dt__'],
                '__sort_vouch__': r['__gl_voucher__'],
                '__is_matched__': 1,
                '基础信息_公司': r['__link_comp__'],
                '基础信息_账号': r['__link_acc__'],
                '核对结果_核对状态': '未匹配',
                '核对结果_匹配编号': None,
                '核对结果_匹配类型': '未匹配',
                '核对结果_交易对手核对': '未匹配'
            }
            for c in gl_raw_cols:
                row[f'序时账_{c}'] = r[c]
            row['辅助信息_计算交易对手(原始)'] = r['__gl_partner_raw__']
            row['辅助信息_拆分说明'] = r['__split_status__']
            output_rows.append(row)

        df_out = pd.DataFrame(output_rows)
        if df_out.empty:
            return df_out

        amount_col_names = []
        if hasattr(self.cfg, 'gl_config'):
            debit_col = self.cfg.gl_config.get('借方金额列名')
            credit_col = self.cfg.gl_config.get('贷方金额列名')
            if debit_col:
                amount_col_names.append(debit_col)
            if credit_col:
                amount_col_names.append(credit_col)
        if hasattr(self.cfg, 'bank_config'):
            income_col = self.cfg.bank_config.get('收入金额列名')
            expense_col = self.cfg.bank_config.get('支出金额列名')
            if income_col:
                amount_col_names.append(income_col)
            if expense_col:
                amount_col_names.append(expense_col)

        for col in df_out.columns:
            for amount_col in amount_col_names:
                if amount_col in str(col):
                    try:
                        df_out[col] = pd.to_numeric(df_out[col], errors='ignore')
                        break
                    except:
                        pass

        df_out = df_out.sort_values(
            by=['__sort_comp__', '__sort_acc__', '__sort_date__', '__sort_vouch__', '__is_matched__', '核对结果_匹配编号']
        ).drop(columns=['__sort_comp__', '__sort_acc__', '__sort_date__', '__sort_vouch__', '__is_matched__'])

        return df_out

    def save_formatted_excel(self, df_summary, df_detail, path):
        """保存格式化的 Excel 文件"""
        log("💾 正在保存 Excel 文件...", 0)
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='核对摘要', index=False)

            cols = df_detail.columns
            header_map = []
            prefixes = ['基础信息_', '核对结果_', '银行流水_', '序时账_', '辅助信息_']
            for prefix in prefixes:
                for c in cols:
                    if c.startswith(prefix):
                        header_map.append((prefix[:-1], c.replace(prefix, ''), c))
            for c in cols:
                if not any(c.startswith(p) for p in prefixes):
                    header_map.append(('其他', c, c))

            ordered_keys = [x[2] for x in header_map]
            df_final = df_detail[ordered_keys]

            df_final.to_excel(writer, sheet_name='核对明细', index=False, header=False, startrow=2)

        wb = load_workbook(path)
        ws_sum = wb['核对摘要']
        for cell in ws_sum[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')

        # 核对摘要数字格式：金额千分位，比率百分率
        amount_cols = {'银行流水收入金额', '银行流水支出金额', '序时账借方金额', '序时账贷方金额'}
        rate_cols = {'银行流水匹配率(金额)', '序时账匹配率(金额)', '综合匹配率(金额)',
                     '银行流水匹配率(条数)', '序时账匹配率(条数)', '综合匹配率(条数)'}
        for cell in ws_sum[1]:
            col_name = cell.value
            if col_name in amount_cols:
                for row in ws_sum.iter_rows(min_row=2, min_col=cell.column, max_col=cell.column):
                    row[0].number_format = '#,#0.00'
            elif col_name in rate_cols:
                for row in ws_sum.iter_rows(min_row=2, min_col=cell.column, max_col=cell.column):
                    row[0].number_format = '0.0%'

        ws_det = wb['核对明细']
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal='center', vertical='center')
        fill_gray = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')
        font_bold = Font(bold=True)

        amount_col_names = []
        if hasattr(self.cfg, 'gl_config'):
            debit_col = self.cfg.gl_config.get('借方金额列名')
            credit_col = self.cfg.gl_config.get('贷方金额列名')
            if debit_col:
                amount_col_names.append(debit_col)
            if credit_col:
                amount_col_names.append(credit_col)
        if hasattr(self.cfg, 'bank_config'):
            income_col = self.cfg.bank_config.get('收入金额列名')
            expense_col = self.cfg.bank_config.get('支出金额列名')
            if income_col:
                amount_col_names.append(income_col)
            if expense_col:
                amount_col_names.append(expense_col)

        amount_col_indices = []
        for col_idx, (top, sub, full_col) in enumerate(header_map, 1):
            for amount_col in amount_col_names:
                if amount_col in sub:
                    amount_col_indices.append(col_idx)
                    break

        for col_idx, (top, sub, _) in enumerate(header_map, 1):
            c1 = ws_det.cell(row=1, column=col_idx)
            c1.value = top
            c1.alignment = align_center
            c1.font = font_bold
            c1.fill = fill_gray
            c1.border = border

            c2 = ws_det.cell(row=2, column=col_idx)
            c2.value = sub
            c2.alignment = align_center
            c2.font = font_bold
            c2.border = border

        if amount_col_indices:
            for row_idx in range(3, ws_det.max_row + 1):
                for col_idx in amount_col_indices:
                    cell = ws_det.cell(row=row_idx, column=col_idx)
                    if cell.value is not None and cell.value != '':
                        try:
                            cell.number_format = '#,##0.00'
                        except:
                            pass

        start_col = 1
        curr_top = header_map[0][0]
        for i in range(1, len(header_map)):
            next_top = header_map[i][0]
            if next_top != curr_top:
                ws_det.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=i)
                start_col = i + 1
                curr_top = next_top
            if i == len(header_map) - 1:
                ws_det.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=i+1)

        wb.save(path)


# ========================================================
# 5. 主函数和 CLI
# ========================================================

def main():
    """主函数：命令行入口"""
    parser = argparse.ArgumentParser(
        description='银行流水与财务序时账智能双向核对工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  python bankflow_reconcile.py config.json
  python bankflow_reconcile.py config.json --output result.xlsx
  python bankflow_reconcile.py config.json --work-dir /path/to/data

配置文件格式 (JSON):
{
  "gl_config": {
    "序时账路径": "序时账.xls",
    "公司列名": "公司",
    "日期列名": "日期",
    "凭证号列名": "凭证号码",
    "科目代码列名": "科目代码",
    "科目名称列名": "科目名称",
    "对方科目列名": "对方科目",
    "摘要列名": "摘要",
    "借方金额列名": "借方金额",
    "贷方金额列名": "贷方金额",
    "银行科目代码": "1002,1004",
    "客商辅助项列名": "客户,供应商",
    "交易对手列名": "客户",
    "账户列名": "账户"
  },
  "bank_config": {
    "银行流水路径": "银行流水.xls",
    "公司列名": "公司",
    "账户列名": "账号",
    "日期列名": "交易时间",
    "交易对手列名": "对方户名",
    "摘要列名": "摘要",
    "收入金额列名": "收入金额",
    "支出金额列名": "支出金额"
  }
}
        """
    )
    parser.add_argument('config', help='配置文件路径 (JSON格式)')
    parser.add_argument('--output', '-o', help='输出文件路径 (默认: 银行流水核对底稿_<时间戳>.xlsx)')
    parser.add_argument('--work-dir', '-w', help='工作目录 (数据文件所在目录，默认: 当前目录)')
    parser.add_argument('--ai-review', help='AI 复核目录：未匹配少时导出供 AI 判断，AI 写回 ai_matches.json 后重跑同一命令即生成含 AI 复核的底稿')
    parser.add_argument('--ai-review-limit', type=int, default=500, help='AI 复核进入阈值（原始未匹配任一侧超过则跳过，先修范围对齐），默认 500')

    args = parser.parse_args()

    print("🚀 启动银行流水智能双向核对工具...")
    start_time = time.time()

    work_dir = args.work_dir if args.work_dir else os.getcwd()
    if not os.path.isabs(work_dir):
        work_dir = os.path.abspath(work_dir)

    config_path = args.config
    if not os.path.isabs(config_path):
        config_path = os.path.abspath(config_path)

    log(f"📂 工作目录: {work_dir}", 0)
    log(f"📋 配置文件: {config_path}", 0)

    try:
        if not os.path.exists(config_path):
            log(f"❌ 配置文件不存在: {config_path}", 0)
            sys.exit(1)

        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        config = ConfigLoader(config_dict=config_data)

        sim_engine = SimilarityEngine()

        gl_proc = GLProcessor(config.gl_config)
        df_gl = gl_proc.process(work_dir)

        bank_proc = BankProcessor(config.bank_config)
        df_bank = bank_proc.process(work_dir)

        if df_gl.empty:
            log(f"❌ 序时账为空或配置错误", 0)
            sys.exit(1)
        if df_bank.empty:
            log(f"❌ 银行流水为空或配置错误", 0)
            sys.exit(1)

        engine = ReconciliationEngine(df_bank, df_gl, sim_engine, config,
                                      ai_review_dir=args.ai_review,
                                      ai_review_limit=args.ai_review_limit)
        df_summary, df_detail = engine.run()

        if engine.ai_review_pending:
            ai_dir = os.path.abspath(args.ai_review)
            print("\n" + "=" * 60)
            print("📋 AI 复核待处理（未写核对底稿）:")
            print(f"   1. 阅读总览（全部未匹配记录，未做规则预筛）: {os.path.join(ai_dir, 'ai_unmatched.json')}")
            for fname in sorted(os.listdir(ai_dir)):
                if fname.startswith('ai_unmatched_') and fname.endswith('.json') and fname != 'ai_unmatched.json':
                    print(f"      - {fname}（可交 subagent 并行判断，每批一个）")
            print(f"   2. 汇总各批判断结果写入: {os.path.join(ai_dir, 'ai_matches.json')}")
            print(f"   3. 重新运行同一命令 → 生成含 AI 复核的核对底稿")
            print("=" * 60 + "\n")
            sys.exit(0)

        if engine.ai_applied > 0:
            log(f"✨ 核对底稿已包含 {engine.ai_applied} 笔 AI 复核匹配（匹配类型以 AI: 开头）", 0)

        output_path = args.output
        if not output_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(work_dir, f"银行流水核对底稿_{timestamp}.xlsx")
        elif not os.path.isabs(output_path):
            output_path = os.path.abspath(output_path)

        engine.save_formatted_excel(df_summary, df_detail, output_path)

        print("\n" + "=" * 60)
        print(f"📊 执行结束 | 总耗时: {time.time() - start_time:.2f}s")
        print("-" * 60)
        print(f"{'账号':<25} | {'金额匹配率':<10} | {'条数匹配率':<10}")
        print("-" * 60)

        for _, row in df_summary.iterrows():
            acc = str(row['账号'])
            if len(acc) > 20:
                acc = acc[:17] + "..."
            rate_amt = f"{row['综合匹配率(金额)'] * 100:.1f}%"
            rate_cnt = f"{row['综合匹配率(条数)'] * 100:.1f}%"
            print(f"{acc:<25} | {rate_amt:<10} | {rate_cnt:<10}")

        print("-" * 60)
        print(f"📂 结果文件: {os.path.basename(output_path)}")
        print("=" * 60 + "\n")

    except KeyboardInterrupt:
        log("\n🛑 用户中断", 0)
        sys.exit(130)
    except Exception as e:
        traceback.print_exc()
        log(f"❌ 运行出错: {str(e)}", 0)
        sys.exit(1)


if __name__ == "__main__":
    main()