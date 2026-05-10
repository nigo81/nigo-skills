"""A股上市公司财务报表查询脚本。

Usage:
    python fetch_reports.py <symbol> [--type TYPE] [--indicator INDICATOR] [--period PERIOD] [--last N] [--output FORMAT]

Examples:
    # 查所有报表，按报告期，最近4期
    python fetch_reports.py 600326

    # 只查利润表，按年度，最近5年
    python fetch_reports.py 600326 --type income --indicator 按年度 --last 5

    # 查资产负债表，指定报告期
    python fetch_reports.py 600326 --type balance --period 2025-12-31

    # 查财务摘要
    python fetch_reports.py 600326 --type abstract
"""

import argparse
import json
import sys

import akshare as ak
import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 25)

REPORT_TYPES = {
    "balance": {
        "name": "资产负债表",
        "func": "stock_financial_debt_ths",
    },
    "income": {
        "name": "利润表",
        "func": "stock_financial_benefit_ths",
    },
    "cashflow": {
        "name": "现金流量表",
        "func": "stock_financial_cash_ths",
    },
    "abstract": {
        "name": "财务摘要",
        "func": "stock_financial_abstract_ths",
    },
}

INDICATORS = ["按报告期", "按年度", "按单季度"]


def fetch_report(symbol: str, report_type: str, indicator: str) -> pd.DataFrame:
    info = REPORT_TYPES[report_type]
    func = getattr(ak, info["func"])

    if report_type == "abstract":
        return func(symbol=symbol)
    else:
        return func(symbol=symbol, indicator=indicator)


def filter_data(
    df: pd.DataFrame,
    period: str | None = None,
    last: int = 4,
    indicator: str = "按报告期",
) -> pd.DataFrame:
    if "报告期" in df.columns:
        df = df.sort_values("报告期", ascending=False)
    if period:
        df = df[df["报告期"].astype(str) == period]
    elif indicator == "按年度":
        df = df[df["报告期"].astype(str).str.endswith("-12-31")].head(last)
    else:
        df = df.head(last)
    return df


def main():
    parser = argparse.ArgumentParser(description="A股上市公司财务报表查询")
    parser.add_argument("symbol", help="6位股票代码，如 600326")
    parser.add_argument(
        "--type",
        choices=list(REPORT_TYPES.keys()) + ["all"],
        default="all",
        help="报表类型 (默认: all)",
    )
    parser.add_argument(
        "--indicator",
        choices=INDICATORS,
        default="按报告期",
        help="时间维度 (默认: 按报告期)",
    )
    parser.add_argument("--period", default=None, help="指定报告期，如 2025-12-31")
    parser.add_argument("--last", type=int, default=4, help="显示最近N期 (默认: 4)")
    parser.add_argument(
        "--output",
        choices=["table", "json", "csv"],
        default="table",
        help="输出格式 (默认: table)",
    )
    args = parser.parse_args()

    types = list(REPORT_TYPES.keys()) if args.type == "all" else [args.type]

    for t in types:
        info = REPORT_TYPES[t]
        print(f"\n{'='*70}")
        print(f"  {info['name']} ({args.symbol})")
        print(f"{'='*70}")

        try:
            df = fetch_report(args.symbol, t, args.indicator)
            df = filter_data(df, args.period, args.last, args.indicator)

            if df.empty:
                print(f"  未找到数据 (报告期={args.period})")
                continue

            if args.output == "table":
                print(df.to_string(index=False))
            elif args.output == "json":
                print(df.to_json(orient="records", force_ascii=False, indent=2))
            elif args.output == "csv":
                print(df.to_csv(index=False))

        except Exception as e:
            print(f"  查询失败: {e}")
            continue


if __name__ == "__main__":
    main()
