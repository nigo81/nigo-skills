"""extractors/base.py — 通用表格行抽取（与原始格式无关）。

输入是统一后的「行」迭代器：每行是单元格值列表。
输出是 statements.json 中使用的 {项目名: {current, prior}} 字典。
"""
import re
import sys


def clean_name(s) -> str:
    """清洗项目名：去序号/中文编号/数字编号/其中/加减前缀/无数字括号注释/空格分散。"""
    s = str(s).strip()
    s = re.sub(r'\s+', '', s)  # 去所有空白（处理"负 债 合 计"→"负债合计"）
    s = re.sub(r'^[一二三四五六七八九十]+[、.．]*', '', s)
    s = re.sub(r'^[(（][一二三四五六七八九十]+[)）]', '', s)
    s = re.sub(r'^[(（]?\d+[)）]?[、.．]*', '', s)
    s = re.sub(r'^其中[：:]*', '', s)
    s = re.sub(r'^[加减][：:]*', '', s)
    s = re.sub(r'[（(]([^()（）]*?)[)）]', lambda m: '' if not re.search(r'\d', m.group(1)) else m.group(0), s)
    return s.strip()


# 通用项目名标准化：把各种报告变体映射到 calculator_rules 使用的规范名
_CANONICAL_ALIASES = {
    # 资产负债表总计
    "负债和所有者权益合计": "负债和所有者权益总计",
    "负债及所有者权益合计": "负债和所有者权益总计",
    "负债和股东权益总计": "负债和所有者权益总计",
    "负债及股东权益总计": "负债和所有者权益总计",
    "负债和股东权益合计": "负债和所有者权益总计",
    "负债及股东权益合计": "负债和所有者权益总计",
    "股东权益合计": "所有者权益合计",
    "归属于母公司股东的权益合计": "归属于母公司所有者权益合计",
    # 利润表
    "营业总收入": "营业收入",
    "归属于母公司股东的净利润": "归属于母公司所有者的净利润",
    # 现金流量表
    "年初现金及现金等价物余额": "期初现金及现金等价物余额",
    "年末现金及现金等价物余额": "期末现金及现金等价物余额",
    "现金及现金等价物净增加额": "现金及现金等价物净增加额",
    # 权益变动表（矩阵行的期初期末）
    "上期期末余额": "上年年末余额",
    "上期末余额": "上年年末余额",
    "上年期末余额": "上年年末余额",
    "期初余额": "上年年末余额",
    "年初余额": "上年年末余额",
    "本年期初余额": "上年年末余额",
    "本期期初余额": "上年年末余额",
    "本期期末余额": "本年年末余额",
    "本年期末余额": "本年年末余额",
    "期末余额": "本年年末余额",
    "本年余额": "本年年末余额",
}


def canonical_name(name: str) -> str:
    """把常见变体统一成 calculator_rules 能识别的规范名。"""
    return _CANONICAL_ALIASES.get(name, name)


def fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, (int, float)):
        return f"{v:.2f}"
    s = str(v).strip()
    if s in ("-", "—", "–"):
        return ""
    return s


def _should_skip_name(name: str, raw: str) -> bool:
    if not name or name in ("项目", "项 目", "项  目"):
        return True
    if any(k in str(raw) for k in ("编制单位", "法定代表人", "金额单位", "校验", "主管会计", "会计机构")):
        return True
    if "其中" in name:
        return True
    return False


def extract_linear(rows, item_col: int, current_col: int, prior_col: int) -> dict:
    """从行迭代器中按列角色抽取线性报表（资产负债表/利润表/现金流量表）。"""
    result = {}
    for cells in rows:
        if not cells or len(cells) <= max(item_col, current_col, prior_col):
            continue
        raw = cells[item_col]
        if raw is None:
            continue
        name = canonical_name(clean_name(raw))
        if _should_skip_name(name, raw):
            continue
        cur = fmt(cells[current_col])
        pri = fmt(cells[prior_col])
        if not cur and not pri:
            if "合计" in name or "总计" in name:
                cur, pri = "0.00", "0.00"
            else:
                continue
        if name in result:
            prev = result[name]
            if prev["current"] not in ("", "0", "0.00") or prev["prior"] not in ("", "0", "0.00"):
                if cur in ("", "0", "0.00") and pri in ("", "0", "0.00"):
                    continue
        result[name] = {"current": cur, "prior": pri}
    return result


def extract_matrix(main_rows, prior_rows, value_col: int) -> dict:
    """权益变动表矩阵：行项目在第 0 列，金额在 value_col 列。

    main_rows 为本期/本年，prior_rows 为上期/上年。
    """
    def scan(rows):
        out = {}
        for cells in rows:
            if not cells or len(cells) <= value_col:
                continue
            raw = cells[0]
            if raw is None:
                continue
            name = canonical_name(clean_name(raw))
            if not name or name in ("项目", "项 目"):
                continue
            if "其中" in name and len(name) <= 4:
                continue
            v = fmt(cells[value_col])
            if not v:
                continue
            if name in ("行次",):
                continue
            if name not in out:
                out[name] = v
        return out

    main = scan(main_rows)
    prior = scan(prior_rows)
    result = {}
    for k in set(main) | set(prior):
        result[k] = {"current": main.get(k, ""), "prior": prior.get(k, "")}
    return result
