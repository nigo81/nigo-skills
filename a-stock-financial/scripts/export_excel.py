"""A股上市公司财务报表导出 Excel。

Usage:
    python export_excel.py <symbol> <company_name> [--indicator INDICATOR] [--period PERIOD] [--last N] [--output PATH]

Examples:
    # 导出所有报告期（默认最近20期）
    python export_excel.py 600326 西藏天路

    # 只导出2025年报
    python export_excel.py 600326 西藏天路 --period 2025-12-31

    # 导出最近3期
    python export_excel.py 600326 西藏天路 --last 3

    # 按年度导出，最近5年
    python export_excel.py 600326 西藏天路 --indicator 按年度 --last 5
"""

import argparse
import os

import akshare as ak
import pandas as pd


def export_excel(
    symbol: str,
    company_name: str,
    indicator: str = "按报告期",
    period: str | None = None,
    last: int = 20,
    output_path: str | None = None,
) -> str:
    if not output_path:
        output_path = f"/Users/nigo/claude/{company_name}_{symbol}_财报.xlsx"

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    sheets = {
        "资产负债表": ("stock_financial_debt_ths", {"symbol": symbol, "indicator": indicator}),
        "利润表": ("stock_financial_benefit_ths", {"symbol": symbol, "indicator": indicator}),
        "现金流量表": ("stock_financial_cash_ths", {"symbol": symbol, "indicator": indicator}),
    }

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, (func_name, kwargs) in sheets.items():
            try:
                func = getattr(ak, func_name)
                df = func(**kwargs)
                if "报告期" in df.columns:
                    df = df.sort_values("报告期", ascending=False)
                if period:
                    df = df[df["报告期"].astype(str) == period]
                else:
                    df = df.head(last)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                pd.DataFrame({"错误": [f"查询失败: {e}"]}).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )

        try:
            abstract = ak.stock_financial_abstract_ths(symbol=symbol)
            abstract = abstract.sort_values("报告期", ascending=False)
            if period:
                abstract = abstract[abstract["报告期"].astype(str) == period]
            elif indicator == "按年度":
                abstract = abstract[
                    abstract["报告期"].astype(str).str.endswith("-12-31")
                ].head(last)
            else:
                abstract = abstract.head(last)
            abstract.to_excel(writer, sheet_name="财务摘要", index=False)
        except Exception as e:
            pd.DataFrame({"错误": [f"查询失败: {e}"]}).to_excel(
                writer, sheet_name="财务摘要", index=False
            )

    return output_path


def main():
    parser = argparse.ArgumentParser(description="A股财务报表导出 Excel")
    parser.add_argument("symbol", help="6位股票代码")
    parser.add_argument("company_name", help="公司名称，用于文件名")
    parser.add_argument("--indicator", default="按报告期", help="时间维度 (默认: 按报告期)")
    parser.add_argument("--period", default=None, help="指定报告期，如 2025-12-31")
    parser.add_argument("--last", type=int, default=20, help="导出最近N期 (默认: 20)")
    parser.add_argument("--output", default=None, help="输出文件路径")
    args = parser.parse_args()

    path = export_excel(
        args.symbol, args.company_name, args.indicator,
        args.period, args.last, args.output,
    )
    print(f"已导出: {path}")


if __name__ == "__main__":
    main()
