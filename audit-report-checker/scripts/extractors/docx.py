"""extractors/docx.py — docx/pdf 提取后的 extracted_tables.json → 统一行迭代器。"""
import json
import sys
from pathlib import Path


def load_extracted_tables(manifest_path: Path) -> dict:
    """找 extracted_tables.json：manifest 同目录、parse_output 子目录、*_output/子目录。"""
    base = manifest_path.parent
    candidates = [
        base / "extracted_tables.json",
        base / "parse_output" / "extracted_tables.json",
    ]
    for sub in base.iterdir():
        if sub.is_dir() and (sub.name in ("parse_output",) or sub.name.endswith("_output")):
            candidates.append(sub / "extracted_tables.json")
    for cand in candidates:
        if cand.exists():
            with cand.open(encoding="utf-8") as f:
                return json.load(f)
    return None


def find_table_in_extracted(extracted_data, ref):
    """按 id(int) 或 name(str) 在 extracted_tables 中定位表。"""
    tables = extracted_data.get("tables", []) if isinstance(extracted_data, dict) else extracted_data

    if isinstance(ref, int):
        for t in tables:
            if t.get("id") == ref:
                return t
        print(f"[WARN] 表 id={ref} 未找到", file=sys.stderr)
        return None

    # 字符串：按 name 精确/模糊匹配
    for t in tables:
        name = t.get("name", "")
        if name and (ref in name or name in ref):
            return t

    # Fallback：按内容关键词识别
    return _identify_table_by_content(tables, str(ref))


def _identify_table_by_content(tables, ref: str):
    """按表内容关键词识别四表。"""
    keywords_map = {
        "合并资产负债表": ["资产总计", "流动资产合计", "非流动资产合计", "负债合计", "所有者权益合计"],
        "母公司资产负债表": ["资产总计", "流动资产合计", "非流动资产合计", "负债合计", "所有者权益合计"],
        "合并利润表": ["营业收入", "营业成本", "营业利润", "利润总额", "净利润", "归属于母公司所有者的净利润"],
        "母公司利润表": ["营业收入", "营业成本", "营业利润", "利润总额", "净利润"],
        "合并现金流量表": ["经营活动现金流入小计", "经营活动现金流出小计", "投资活动现金流入小计", "筹资活动现金流入小计"],
        "母公司现金流量表": ["经营活动现金流入小计", "经营活动现金流出小计", "投资活动现金流入小计", "筹资活动现金流入小计"],
    }

    for t in tables:
        rows = t.get("rows", [])
        if not rows:
            continue
        row_text = " ".join(" ".join(str(c) for c in row) for row in rows[:30])

        for table_type, keywords in keywords_map.items():
            if table_type in ref or ref in table_type:
                if all(kw in row_text for kw in keywords[:3]):
                    return t
    return None


def iter_docx_rows(extracted_data: dict, refs: list):
    """按 refs（表 id 或 name 列表）顺序产生行。"""
    for ref in refs:
        table = find_table_in_extracted(extracted_data, ref)
        if not table:
            print(f"[WARN] 找不到表: {ref}", file=sys.stderr)
            continue
        for row in table.get("rows", []):
            yield row
