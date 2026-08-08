#!/usr/bin/env python3
"""
bankflow_precheck.py — 银行流水核对数据探查工具

只做两件事：
1. 读文件，输出列名 + 样本行 + 基本统计
2. 如果指定 --gl-subject-col + --bank-subject，过滤银行存款行并输出统计

不做任何列名猜测。AI 看输出后自己判断列映射，直接写 reconcile_config.json。
"""
import argparse
import sys
from pathlib import Path

import pandas as pd


def read_excel_any(path: str) -> pd.DataFrame:
    """读取 .xlsx/.xls/.csv，返回 DataFrame（header=0）。"""
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str)
    elif ext == ".xls":
        return pd.read_excel(path, dtype=str)
    else:
        return pd.read_excel(path, dtype=str, engine="openpyxl")


def show_header_preview(df: pd.DataFrame, n: int = 3) -> None:
    """输出前 n 行数据（转置显示，方便看完整列）。"""
    print(f"  前 {n} 行数据（转置）:")
    for i in range(min(n, len(df))):
        print(f"  --- 第 {i+1} 行 ---")
        for col in df.columns:
            val = str(df.iloc[i][col])
            if val and val != "nan":
                print(f"    {col}: {val[:80]}")


def show_col_stats(df: pd.DataFrame, label: str = "") -> None:
    """输出所有列的非空率 + 数值检测 + 样本值。"""
    print(f"\n  各列非空统计{('（' + label + '）') if label else ''}:")
    for col in df.columns:
        nonempty = df[col].dropna()
        nonempty = nonempty[nonempty.astype(str).str.strip() != ""]
        rate = len(nonempty) / len(df) * 100 if len(df) > 0 else 0
        # 检测是否含数字
        sample_nums = []
        for v in nonempty.head(10):
            try:
                float(str(v).replace(",", ""))
                sample_nums.append(v)
            except:
                pass
        num_hint = f" [数值列, 样本: {sample_nums[:3]}]" if len(sample_nums) >= 3 else ""
        if rate > 0 or num_hint:
            print(f"    {col}: {rate:.0f}% 非空{num_hint}")


def show_tail(df: pd.DataFrame, n: int = 5) -> None:
    """输出后 n 行（检测尾部非交易行）。"""
    print(f"\n  后 {n} 行（检测尾部汇总/统计行）:")
    for i in range(max(0, len(df) - n), len(df)):
        vals = {}
        for col in df.columns:
            v = str(df.iloc[i][col])
            if v and v != "nan":
                vals[col] = v[:60]
        # 只显示非空值
        summary = " | ".join(f"{k}={v}" for k, v in list(vals.items())[:6])
        print(f"    行{i+1}: {summary}")


def profile_bank(path: str, header_row: int | None = None) -> None:
    """探查银行流水文件。"""
    p = Path(path)
    print("=" * 60)
    print(f"📋 银行流水: {p.name}")
    print("=" * 60)

    df = read_excel_any(path)
    print(f"  文件格式: {p.suffix}")
    print(f"  总行数: {len(df)}")
    print(f"  列数: {len(df.columns)}")

    # 表头偏移检测
    unnamed_count = sum(1 for c in df.columns if "Unnamed" in str(c) or str(c).startswith("nan"))
    if unnamed_count > len(df.columns) * 0.3:
        print(f"\n  ⚠️ 表头偏移检测: {unnamed_count}/{len(df.columns)} 列名为 'Unnamed'")
        print("     → 表头可能不在第1行！")
        if header_row:
            df = pd.read_excel(path, dtype=str, header=header_row - 1, engine="openpyxl" if p.suffix == ".xlsx" else None)
            print(f"     已按 --bank-header-row={header_row} 重新读取")
        else:
            # 自动预览前3行帮助判断
            print("     前3行预览（header=None）:")
            df_raw = pd.read_excel(path, header=None, nrows=3, dtype=str,
                                   engine="openpyxl" if p.suffix == ".xlsx" else None)
            for i in range(len(df_raw)):
                vals = [str(v)[:40] for v in df_raw.iloc[i].tolist() if str(v) != "nan"]
                print(f"       行{i+1}: {' | '.join(vals[:6])}")
            print("     → 用 --bank-header-row N 指定表头所在行号（从1开始）重跑")

    print(f"  列名: {list(df.columns)}")
    print()

    show_header_preview(df)
    show_col_stats(df, "全量")
    show_tail(df)

    # 日期范围（找含'日期'或'时间'的列并输出范围）
    print("\n  日期列检测（仅供参考，AI 判断）:")
    for col in df.columns:
        if any(kw in col for kw in ["日期", "时间", "交易日"]):
            try:
                dates = pd.to_datetime(df[col], errors="coerce").dropna()
                if len(dates) > 0:
                    print(f"    [{col}]: {dates.min().strftime('%Y-%m-%d')} ~ {dates.max().strftime('%Y-%m-%d')}")
            except:
                pass

    # 金额合计提示（找借/贷/收/支列）
    print("\n  金额列合计（仅供参考）:")
    for col in df.columns:
        lower = col.lower() if isinstance(col, str) else ""
        if any(kw in str(col) for kw in ["借方", "贷方", "收入", "支出", "转入", "转出"]):
            try:
                vals = pd.to_numeric(df[col].astype(str).str.replace(",", "", regex=False), errors="coerce")
                total = vals.sum()
                if total != 0:
                    print(f"    {col}: {total:,.2f}")
            except:
                pass

    return df


def profile_gl(
    path: str,
    bank_subject: str | None = None,
    subject_col: str | None = None,
    subject_split_char: str | None = None,
    bank_subject_name: str | None = None,
) -> pd.DataFrame | None:
    """探查序时账文件。"""
    p = Path(path)
    print("\n" + "=" * 60)
    print(f"📋 序时账: {p.name}")
    print("=" * 60)

    df = read_excel_any(path)
    print(f"  文件格式: {p.suffix}")
    print(f"  总行数: {len(df)}")
    print(f"  列数: {len(df.columns)}")
    print(f"  列名: {list(df.columns)}")
    print()

    show_header_preview(df)
    show_col_stats(df, "全量")
    show_tail(df)

    # 银行存款过滤
    if not bank_subject and not bank_subject_name:
        print("\n  ℹ️ 未指定 --bank-subject，跳过银行存款过滤。")
        print("     AI 查看列名后，用 --bank-subject + --gl-subject-col 重跑以获取银行存款统计。")
        return None

    if not subject_col and not bank_subject_name:
        print("\n  ⚠️ 指定了 --bank-subject 但未指定 --gl-subject-col（科目代码所在列名）。")
        print("     AI 查看列名后，用 --gl-subject-col '列名' 重跑。")
        return None

    # 过滤银行存款行
    if bank_subject_name and subject_col:
        # 按科目名称过滤（如天路：科目名称="银行存款"）
        bank_mask = df[subject_col].astype(str).str.contains(bank_subject_name, na=False)
        filter_desc = f"{subject_col} 含 '{bank_subject_name}'"
    elif subject_col and subject_split_char:
        # 品智：编码列="1002/xxx"，取split第一部分
        code_series = df[subject_col].fillna("").astype(str).str.split(subject_split_char).str[0]
        bank_mask = code_series.str.startswith(bank_subject)
        filter_desc = f"{subject_col}.split('{subject_split_char}')[0] startswith '{bank_subject}'"
    elif subject_col:
        # 正常：科目代码列 startswith 前缀
        bank_mask = df[subject_col].fillna("").astype(str).str.startswith(bank_subject)
        filter_desc = f"{subject_col} startswith '{bank_subject}'"
    else:
        print("\n  ⚠️ 无法过滤银行存款行。")
        return None

    bank_rows = df[bank_mask].copy()
    print(f"\n  🏦 银行存款过滤: {filter_desc}")
    print(f"     银行存款行数: {len(bank_rows)}")

    if len(bank_rows) == 0:
        print("     ⚠️ 银行存款行数为0！请检查科目代码列名和前缀。")
        return None

    # 凭证数
    voucher_col = None
    for col in df.columns:
        if any(kw in str(col) for kw in ["凭证号", "凭证编", "凭证唯一"]):
            voucher_col = col
            break
    if voucher_col:
        voucher_count = bank_rows[voucher_col].nunique()
        print(f"     涉及凭证数（{voucher_col} 去重）: {voucher_count}")

    # 金额统计
    for col in df.columns:
        if "借方" in str(col) or "贷方" in str(col):
            try:
                vals = pd.to_numeric(
                    bank_rows[col].astype(str).str.replace(",", "", regex=False),
                    errors="coerce",
                )
                total = vals.sum()
                if total != 0:
                    print(f"     {col}合计: {total:,.2f}")
            except:
                pass

    # 所有列在银行存款行上的诊断（AI 看这些信息做判断）
    print("\n  📊 银行存款行上各列诊断（AI 判断客商/账户/金额列映射）:")
    for col in bank_rows.columns:
        nonempty = bank_rows[col].dropna()
        nonempty = nonempty[nonempty.astype(str).str.strip() != ""]
        rate = len(nonempty) / len(bank_rows) * 100 if len(bank_rows) > 0 else 0
        uniques = nonempty.astype(str).nunique()
        samples = list(nonempty.astype(str).unique()[:3])
        # 截断样本值
        samples = [s[:60] for s in samples]
        print(f"    [{col}]: 非空率 {rate:.0f}%, {uniques}个唯一值, 样本: {samples}")

    # 对方科目行上的客商列诊断（帮助区分直接模式/拆分模式）
    counter_rows = df[~bank_mask].copy()
    if len(counter_rows) > 0:
        print("\n  📊 对方科目行上各列诊断（判断客商信息是否在对方行）:")
        for col in counter_rows.columns:
            nonempty = counter_rows[col].dropna()
            nonempty = nonempty[nonempty.astype(str).str.strip() != ""]
            rate = len(nonempty) / len(counter_rows) * 100 if len(counter_rows) > 0 else 0
            if rate > 10:  # 只显示有实质内容的列
                samples = list(nonempty.astype(str).unique()[:3])
                samples = [s[:60] for s in samples]
                print(f"    [{col}]: 非空率 {rate:.0f}%, 样本: {samples}")

    # 主体列分布检测（用于判断是否需要主体过滤）
    entity_keywords = ["主体", "公司", "单位", "组织", "实体"]
    entity_cols = [col for col in df.columns if any(kw in col for kw in entity_keywords)]
    if entity_cols:
        for col in entity_cols:
            print(f"\n  📊 主体列分布 [{col}]:")
            for val, cnt in df[col].value_counts().head(10).items():
                print(f"     {val}: {cnt} 行")

    return bank_rows


def main():
    parser = argparse.ArgumentParser(description="银行流水核对数据探查")
    parser.add_argument("--gl", required=True, help="序时账文件路径")
    parser.add_argument("--bank", required=True, help="银行流水文件路径")
    parser.add_argument("--bank-subject", default=None, help="银行存款科目代码前缀（如1002）")
    parser.add_argument("--gl-subject-col", default=None, help="科目代码所在列名（AI指定）")
    parser.add_argument("--subject-split-char", default=None, help="科目代码列的分隔符（如编码='1002/xxx'时用'/'）")
    parser.add_argument("--bank-subject-name", default=None, help="按科目名称过滤（如'银行存款'），替代代码前缀")
    parser.add_argument("--bank-header-row", type=int, default=None, help="银行流水表头所在行号（从1开始，用于跳过前置元数据行）")
    args = parser.parse_args()

    # Multi-sheet detection
    for label, path in [("序时账", args.gl), ("银行流水", args.bank)]:
        try:
            xl = pd.ExcelFile(path)
            if len(xl.sheet_names) > 1:
                print(f"\n  ⚠️ {label}有多个Sheet: {xl.sheet_names}")
                print(f"     当前读取第一个: [{xl.sheet_names[0]}]")
                print(f"     如果数据不在第一个Sheet，请先提取到单独文件再运行预检查")
        except Exception:
            pass  # Not an Excel file or can't read sheets

    profile_bank(args.bank, header_row=args.bank_header_row)
    profile_gl(
        args.gl,
        bank_subject=args.bank_subject,
        subject_col=args.gl_subject_col,
        subject_split_char=args.subject_split_char,
        bank_subject_name=args.bank_subject_name,
    )

    print("\n" + "=" * 60)
    print("✅ 探查完成。请根据以上信息判断列映射，编写 reconcile_config.json")
    print("=" * 60)


if __name__ == "__main__":
    main()
