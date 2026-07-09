#!/usr/bin/env python3
"""
审计报告核查算术工具

LLM 直接做数值加减会算错且不可复现。本工具是审计核查的算术生命线——所有数值计算都走它。
设计原则：AI 决定传哪些数，代码算出结果。

支持的功能：
- sum: 求和（验证竖加/横加）
- expr: 表达式计算（验证勾稽推导）
- check: 核对相等（含容差，处理四舍五入）
- reconcile: 验证变动表（期初+增加-减少=期末）

支持数字格式：
- 千分位逗号：1,234,567.89
- 欧洲格式：1.234,56
- 括号负数：(1,234.56) → -1234.56
- 横线/空：- — – 空字符串 → 0
- 全角字符：， （）
- 万元单位：含"万"字 → ×10000
"""

import argparse
import ast
import json
import re
import sys
from typing import Union


def parse_number(value: str) -> float:
    """
    解析财务报告数字，支持多种格式。

    Args:
        value: 数字字符串

    Returns:
        float: 解析后的数字

    Raises:
        ValueError: 无法解析时抛出
    """
    # 容错：locate/AI 返回的数字可能是 float/int（JSON number），直接转 float
    if isinstance(value, (int, float)):
        return float(value)

    if not value or value.strip() == "":
        return 0.0

    # 转换全角字符
    value = value.replace("，", ",").replace("（", "(").replace("）", ")")

    # 处理横线表示零的情况
    if value.strip() in ["-", "—", "–", "－", "_"]:
        return 0.0

    # 处理括号负数
    if value.strip().startswith("(") and value.strip().endswith(")"):
        value = value.strip()[1:-1]
        negative = True
    else:
        negative = False

    # 移除千分位逗号（欧洲格式先处理）
    # 判断是欧洲格式还是美式格式
    # 欧洲格式：1.234,56 或 1.234.567,89（.是千分位，,是小数）
    # 美式格式：1,234.56（,是千分位，.是小数）
    cleaned = value.replace(" ", "")

    # 判断欧洲格式 vs 美式格式：以"最后一个分隔符"为准
    # 美式：1,234,567.89（逗号=千分位，点=小数）→ 最后分隔符是点
    # 欧洲：1.234.567,89（点=千分位，逗号=小数）→ 最后分隔符是逗号
    last_comma = cleaned.rfind(",")
    last_dot = cleaned.rfind(".")
    if last_comma != -1 and last_dot != -1:
        if last_comma > last_dot:
            # 欧洲格式：点全部是千分位（移除），最后一个逗号是小数点
            integer = cleaned[:last_comma].replace(".", "").replace(",", "")
            decimal = cleaned[last_comma + 1:]
            cleaned = integer + "." + decimal
        else:
            # 美式格式：逗号全部是千分位（移除），点是小数点
            cleaned = cleaned.replace(",", "")
    elif last_comma != -1:
        # 只有逗号：可能是千分位（1,234）或欧洲小数（1,23）
        parts = cleaned.split(",")
        if len(parts) == 2 and len(parts[1]) == 3 and parts[0].isdigit() and parts[1].isdigit():
            # 逗号后恰好3位且都是数字 → 视为千分位（财务报告中最常见）
            cleaned = cleaned.replace(",", "")
        elif len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit() and len(parts[1]) != 3:
            # 逗号后非3位 → 欧洲小数（如 1,23 → 1.23）
            cleaned = parts[0] + "." + parts[1]
        else:
            # 多个逗号的纯整数千分位（如 1,234,567）或其他情况
            cleaned = cleaned.replace(",", "")
    # 只有点的保持原样

    # 提取数字部分
    num_match = re.search(r"[-+]?\d*\.?\d+", cleaned)
    if not num_match:
        raise ValueError(f"无法解析数字: {value}")

    num_str = num_match.group()
    try:
        result = float(num_str)
    except ValueError:
        raise ValueError(f"无法解析数字: {value}")

    # 处理负数
    if negative:
        result = -result

    # 处理万元单位
    if "万" in value:
        result = result * 10000

    return result


def sum_numbers(numbers_str: str) -> float:
    """
    求和。多数字用空格分隔，正则匹配数字保留千分位逗号避免拆碎。

    Args:
        numbers_str: 空格分隔的数字字符串

    Returns:
        float: 求和结果
    """
    # 贪婪匹配数字，避免千分位逗号被拆碎
    # \d[\d.,]*\d 匹配以数字开头和结尾，中间可以有数字、逗号、点的字符串
    # 或 \d 单个数字
    pattern = r"\d[\d.,]*\d|\d"
    matches = re.findall(pattern, numbers_str)

    total = 0.0
    for num_str in matches:
        try:
            num = parse_number(num_str)
            total += num
        except ValueError:
            # 跳过无法解析的
            continue

    return total


def evaluate_expression(expr: str) -> float:
    """
    表达式计算。用 ast 安全 eval（禁用危险函数）。

    Args:
        expr: 数学表达式字符串

    Returns:
        float: 计算结果
    """
    # 预处理：移除千分位逗号（数字不带千分位逗号，但为了保险）
    expr = expr.replace(",", "")

    # 创建安全的表达式评估器
    def safe_eval(node):
        if isinstance(node, ast.Expression):
            return safe_eval(node.body)
        elif isinstance(node, ast.BinOp):
            left = safe_eval(node.left)
            right = safe_eval(node.right)
            if isinstance(node.op, ast.Add):
                return left + right
            elif isinstance(node.op, ast.Sub):
                return left - right
            elif isinstance(node.op, ast.Mult):
                return left * right
            elif isinstance(node.op, ast.Div):
                return left / right
            elif isinstance(node.op, ast.Mod):
                return left % right
            elif isinstance(node.op, ast.Pow):
                return left ** right
            elif isinstance(node.op, ast.FloorDiv):
                return left // right
            else:
                raise TypeError(f"不支持的操作符: {type(node.op).__name__}")
        elif isinstance(node, ast.UnaryOp):
            operand = safe_eval(node.operand)
            if isinstance(node.op, ast.UAdd):
                return +operand
            elif isinstance(node.op, ast.USub):
                return -operand
            else:
                raise TypeError(f"不支持的一元操作符: {type(node.op).__name__}")
        elif isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)):
                return node.value
            else:
                raise TypeError(f"不支持的常量类型: {type(node.value).__name__}")
        else:
            raise TypeError(f"不支持的 AST 节点: {type(node).__name__}")

    try:
        tree = ast.parse(expr, mode="eval")
        result = safe_eval(tree)
        return float(result)
    except (SyntaxError, TypeError) as e:
        raise ValueError(f"表达式解析失败: {expr}, 错误: {e}")


def check_equal(value1: str, value2: str, tolerance: float = 0.01) -> dict:
    """
    核对两个值是否相等（含容差）。

    Args:
        value1: 第一个值字符串
        value2: 第二个值字符串
        tolerance: 容差，默认 0.01

    Returns:
        dict: 包含 match, diff, value1_parsed, value2_parsed
    """
    try:
        num1 = value1 if isinstance(value1, (int, float)) else parse_number(value1)
        num2 = value2 if isinstance(value2, (int, float)) else parse_number(value2)
    except ValueError as e:
        return {
            "match": False,
            "diff": None,
            "value1_parsed": None,
            "value2_parsed": None,
            "error": str(e),
        }

    diff = abs(num1 - num2)
    # 使用 round(diff, 6) 避免浮点边界误判
    is_match = round(diff, 6) <= tolerance

    return {
        "match": is_match,
        "diff": diff,
        "value1_parsed": num1,
        "value2_parsed": num2,
    }


def reconcile(
    beginning: str, increase: str, decrease: str, ending: str, tolerance: float = 0.01
) -> dict:
    """
    验证变动表：期初 + 增加 - 减少 = 期末。

    Args:
        beginning: 期初值
        increase: 增加值
        decrease: 减少值
        ending: 期末值
        tolerance: 容差，默认 0.01

    Returns:
        dict: 包含 match, diff, calculated_ending, actual_ending
    """
    try:
        beginning_num = parse_number(beginning)
        increase_num = parse_number(increase)
        decrease_num = parse_number(decrease)
        ending_num = parse_number(ending)
    except ValueError as e:
        return {
            "match": False,
            "diff": None,
            "calculated_ending": None,
            "actual_ending": None,
            "error": str(e),
        }

    calculated_ending = beginning_num + increase_num - decrease_num
    diff = abs(calculated_ending - ending_num)

    is_match = round(diff, 6) <= tolerance

    return {
        "match": is_match,
        "diff": diff,
        "calculated_ending": calculated_ending,
        "actual_ending": ending_num,
    }


def format_output(result: dict, format_type: str = "text") -> str:
    """
    格式化输出结果。

    Args:
        result: 结果字典
        format_type: 输出格式（text/json）

    Returns:
        str: 格式化后的输出
    """
    if format_type == "json":
        return json.dumps(result, ensure_ascii=False, indent=2)
    else:
        return result.get("output", str(result))


def main():
    parser = argparse.ArgumentParser(
        description="审计报告核查算术工具", formatter_class=argparse.RawTextHelpFormatter
    )
    # 扩展负数识别器：argparse 默认只认 ^-\d+$，不认带千分位逗号/小数的负数（如 -1,807,177.07），
    # 会误判为选项。财务报告负数常见（信用减值损失/资产减值损失/亏损/转回），必须兼容。
    parser._negative_number_matcher = re.compile(r"^-\d[\d,]*\.?\d*$|^-\.\d+$")
    parser.add_argument(
        "-f", "--format", choices=["text", "json"], default="text", help="输出格式"
    )

    subparsers = parser.add_subparsers(dest="command", required=True, help="子命令")

    # sum 子命令
    sum_parser = subparsers.add_parser("sum", help="求和（验证竖加/横加）")
    sum_parser.add_argument(
        "numbers", help="空格分隔的数字字符串，如 '1,234.56 5,678.90 100.00'"
    )

    # expr 子命令
    expr_parser = subparsers.add_parser("expr", help="表达式计算（验证勾稽推导）")
    expr_parser.add_argument("expression", help="数学表达式，如 '1000 - 300 + 50'")

    # check 子命令
    check_parser = subparsers.add_parser("check", help="核对相等（含容差）")
    # 子 parser 也要扩展负数识别（主 parser 的设置不自动继承到子 parser）
    check_parser._negative_number_matcher = re.compile(r"^-\d[\d,]*\.?\d*$|^-\.\d+$")
    check_parser.add_argument("value1", help="第一个值")
    check_parser.add_argument("value2", help="第二个值")
    check_parser.add_argument(
        "-t", "--tolerance", type=float, default=0.01, help="容差，默认 0.01"
    )

    # reconcile 子命令
    reconcile_parser = subparsers.add_parser(
        "reconcile", help="验证变动表（期初+增加-减少=期末）"
    )
    reconcile_parser._negative_number_matcher = re.compile(r"^-\d[\d,]*\.?\d*$|^-\.\d+$")
    reconcile_parser.add_argument("beginning", help="期初值")
    reconcile_parser.add_argument("increase", help="增加值")
    reconcile_parser.add_argument("decrease", help="减少值")
    reconcile_parser.add_argument("ending", help="期末值")
    reconcile_parser.add_argument(
        "-t", "--tolerance", type=float, default=0.01, help="容差，默认 0.01"
    )

    args = parser.parse_args()

    result = {}

    if args.command == "sum":
        total = sum_numbers(args.numbers)
        result = {
            "total": round(total, 2),
            "output": f"{total:.2f}" if total != int(total) else f"{int(total)}",
        }

    elif args.command == "expr":
        try:
            calculated = evaluate_expression(args.expression)
            result = {
                "result": round(calculated, 6),
                "output": f"{calculated:.2f}"
                if calculated != int(calculated)
                else f"{int(calculated)}",
            }
        except ValueError as e:
            result = {"error": str(e), "output": f"ERROR: {e}"}

    elif args.command == "check":
        check_result = check_equal(args.value1, args.value2, args.tolerance)
        if "error" in check_result:
            result = {
                "error": check_result["error"],
                "output": f"ERROR: {check_result['error']}",
            }
        else:
            status = "MATCH" if check_result["match"] else "MISMATCH"
            diff_str = f"{check_result['diff']:.2f}"
            result = {
                "match": check_result["match"],
                "diff": check_result["diff"],
                "output": f"{status} (diff={diff_str})",
            }

    elif args.command == "reconcile":
        recon_result = reconcile(
            args.beginning, args.increase, args.decrease, args.ending, args.tolerance
        )
        if "error" in recon_result:
            result = {
                "error": recon_result["error"],
                "output": f"ERROR: {recon_result['error']}",
            }
        else:
            status = "MATCH" if recon_result["match"] else "MISMATCH"
            diff_str = f"{recon_result['diff']:.2f}"
            result = {
                "match": recon_result["match"],
                "diff": recon_result["diff"],
                "calculated_ending": recon_result["calculated_ending"],
                "actual_ending": recon_result["actual_ending"],
                "output": f"{status} (diff={diff_str}, 计算期末={recon_result['calculated_ending']:.2f}, 实际期末={recon_result['actual_ending']:.2f})",
            }

    print(format_output(result, args.format))


if __name__ == "__main__":
    main()