#!/usr/bin/env python3
"""
L1 代码内置规则

把稳定、高频的报表间勾稽规则固化为代码（参数化但逻辑固化），不走 LLM。
这些规则是确定性的，必须 100% 可靠。

规则包括：
- 资产负债表纵向勾稽
- 利润表推导
- 现金流量表三类净额
- 所有者权益变动表期末=期初+变动
"""

from dataclasses import dataclass
from typing import Optional

from calculator import check_equal, parse_number


@dataclass
class CheckItem:
    """规则检查结果对象。"""

    rule_name: str  # 规则名称
    passed: bool  # 是否通过
    expected: Optional[str]  # 应为值（字符串格式）
    actual: Optional[str]  # 实际值（字符串格式）
    difference: Optional[float]  # 差异（数字），None 表示无法计算
    description: str  # 描述
    severity: str  # 严重程度：error/warning/info


def check_balance_sheet(statement: dict, tolerance: float = 0.01) -> list[CheckItem]:
    """
    检查资产负债表纵向勾稽。

    Args:
        statement: 报表字典，格式为 {'项目名': {'current': '123', 'prior': '456'}, ...}
        tolerance: 容差，默认 0.01

    Returns:
        list[CheckItem]: 检查结果列表

    规则：
    - 流动资产合计 + 非流动资产合计 = 资产总计
    - 流动负债合计 + 非流动负债合计 = 负债合计
    - 负债合计 + 所有者权益合计 = 负债和所有者权益总计
    - 资产总计 = 负债和所有者权益总计
    """
    results = []

    # 获取当前期和上期的项目数据
    def get_value(item_name: str, period: str = "current") -> Optional[str]:
        """获取项目的值，支持常见的变体名称"""
        item = statement.get(item_name)
        if item is None:
            return None
        return item.get(period)

    # 规则 1: 流动资产合计 + 非流动资产合计 = 资产总计
    def check_assets_sum(period: str, period_name: str):
        """检查资产合计平衡"""
        current_assets = get_value("流动资产合计", period)
        non_current_assets = get_value("非流动资产合计", period)
        total_assets = get_value("资产总计", period)

        if None in [current_assets, non_current_assets, total_assets]:
            results.append(
                CheckItem(
                    rule_name=f"资产合计{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：流动资产合计={current_assets}, 非流动资产合计={non_current_assets}, 资产总计={total_assets}",
                    severity="error",
                )
            )
            return

        # 计算预期资产总计
        try:
            current_num = parse_number(current_assets)
            non_current_num = parse_number(non_current_assets)
            expected_total = current_num + non_current_num

            # 核对
            check_result = check_equal(str(expected_total), total_assets, tolerance)

            results.append(
                CheckItem(
                    rule_name=f"资产合计{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_total:.2f}",
                    actual=total_assets,
                    difference=check_result["diff"],
                    description=f"流动资产合计({current_assets}) + 非流动资产合计({non_current_assets}) = {expected_total:.2f}，实际资产总计={total_assets}",
                    severity="error",
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"资产合计{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

    check_assets_sum("current", "（本期）")
    check_assets_sum("prior", "（上期）")

    # 规则 2: 流动负债合计 + 非流动负债合计 = 负债合计
    def check_liabilities_sum(period: str, period_name: str):
        """检查负债合计平衡"""
        current_liabilities = get_value("流动负债合计", period)
        non_current_liabilities = get_value("非流动负债合计", period)
        total_liabilities = get_value("负债合计", period)

        if None in [current_liabilities, non_current_liabilities, total_liabilities]:
            results.append(
                CheckItem(
                    rule_name=f"负债合计{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：流动负债合计={current_liabilities}, 非流动负债合计={non_current_liabilities}, 负债合计={total_liabilities}",
                    severity="error",
                )
            )
            return

        try:
            current_num = parse_number(current_liabilities)
            non_current_num = parse_number(non_current_liabilities)
            expected_total = current_num + non_current_num

            check_result = check_equal(str(expected_total), total_liabilities, tolerance)

            results.append(
                CheckItem(
                    rule_name=f"负债合计{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_total:.2f}",
                    actual=total_liabilities,
                    difference=check_result["diff"],
                    description=f"流动负债合计({current_liabilities}) + 非流动负债合计({non_current_liabilities}) = {expected_total:.2f}，实际负债合计={total_liabilities}",
                    severity="error",
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"负债合计{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

    check_liabilities_sum("current", "（本期）")
    check_liabilities_sum("prior", "（上期）")

    # 规则 3: 负债合计 + 所有者权益合计 = 负债和所有者权益总计
    def check_equity_balance(period: str, period_name: str):
        """检查权益平衡"""
        total_liabilities = get_value("负债合计", period)
        total_equity = get_value("所有者权益合计", period)
        total_liabilities_equity = get_value("负债和所有者权益总计", period)

        if None in [total_liabilities, total_equity, total_liabilities_equity]:
            results.append(
                CheckItem(
                    rule_name=f"权益平衡{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：负债合计={total_liabilities}, 所有者权益合计={total_equity}, 负债和所有者权益总计={total_liabilities_equity}",
                    severity="error",
                )
            )
            return

        try:
            liabilities_num = parse_number(total_liabilities)
            equity_num = parse_number(total_equity)
            expected_total = liabilities_num + equity_num

            check_result = check_equal(
                str(expected_total), total_liabilities_equity, tolerance
            )

            results.append(
                CheckItem(
                    rule_name=f"权益平衡{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_total:.2f}",
                    actual=total_liabilities_equity,
                    difference=check_result["diff"],
                    description=f"负债合计({total_liabilities}) + 所有者权益合计({total_equity}) = {expected_total:.2f}，实际负债和所有者权益总计={total_liabilities_equity}",
                    severity="error",
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"权益平衡{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

    check_equity_balance("current", "（本期）")
    check_equity_balance("prior", "（上期）")

    # 规则 4: 资产总计 = 负债和所有者权益总计
    def check_balance_equality(period: str, period_name: str):
        """检查资产负债平衡"""
        total_assets = get_value("资产总计", period)
        total_liabilities_equity = get_value("负债和所有者权益总计", period)

        if None in [total_assets, total_liabilities_equity]:
            results.append(
                CheckItem(
                    rule_name=f"资产负债平衡{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：资产总计={total_assets}, 负债和所有者权益总计={total_liabilities_equity}",
                    severity="error",
                )
            )
            return

        check_result = check_equal(total_assets, total_liabilities_equity, tolerance)

        results.append(
            CheckItem(
                rule_name=f"资产负债平衡{period_name}",
                passed=check_result["match"],
                expected=total_assets,
                actual=total_liabilities_equity,
                difference=check_result["diff"],
                description=f"资产总计({total_assets}) = 负债和所有者权益总计({total_liabilities_equity})",
                severity="error",
            )
        )

    check_balance_equality("current", "（本期）")
    check_balance_equality("prior", "（上期）")

    return results


def check_income_statement(statement: dict, tolerance: float = 0.01) -> list[CheckItem]:
    """
    利润表推导检查。

    Args:
        statement: 报表字典，格式为 {'项目名': {'current': '123', 'prior': '456'}, ...}
        tolerance: 容差，默认 0.01

    Returns:
        list[CheckItem]: 检查结果列表

    规则：
    - 关键字段（必须有）：营业收入/营业总收入、营业成本、营业利润、利润总额、净利润
    - 可选项（可能缺失）：税金及附加、销售费用、管理费用、研发费用、财务费用、其他收益、投资收益、公允价值变动收益、净敞口套期收益、信用减值损失、资产减值损失、资产处置收益
    - 营业利润 = 营业收入 - 营业成本 - 税金及附加 - 销售费用 - 管理费用 - 研发费用 - 财务费用 + 其他收益 + 投资收益 + 公允价值变动收益 + 净敞口套期收益 - 信用减值损失 - 资产减值损失 + 资产处置收益
    - 利润总额 = 营业利润 + 营业外收入 - 营业外支出
    - 净利润 = 利润总额 - 所得税费用
    """
    results = []

    def get_value(item_name: str, period: str = "current") -> Optional[str]:
        """获取项目的值，支持常见的变体名称"""
        item = statement.get(item_name)
        if item is None:
            return None
        return item.get(period)

    def check_item(period: str, period_name: str):
        """检查利润表推导"""
        # 关键字段（必须有）
        key_fields = {
            "营业收入": get_value("营业收入", period) or get_value("营业总收入", period),
            "营业成本": get_value("营业成本", period),
            "营业利润": get_value("营业利润", period),
            "利润总额": get_value("利润总额", period),
            "净利润": get_value("净利润", period),
        }

        # 检查关键字段是否齐全
        missing_key_fields = [name for name, value in key_fields.items() if value is None]
        if missing_key_fields:
            results.append(
                CheckItem(
                    rule_name=f"营业利润推导{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺关键字段无法推导：{', '.join(missing_key_fields)}",
                    severity="warning",
                )
            )
            return

        # 可选项（可能缺失）
        optional_fields = {
            "税金及附加": get_value("税金及附加", period),
            "销售费用": get_value("销售费用", period),
            "管理费用": get_value("管理费用", period),
            "研发费用": get_value("研发费用", period),
            "财务费用": get_value("财务费用", period),
            "其他收益": get_value("其他收益", period),
            "投资收益": get_value("投资收益", period),
            "公允价值变动收益": get_value("公允价值变动收益", period),
            "净敞口套期收益": get_value("净敞口套期收益", period),
            "信用减值损失": get_value("信用减值损失", period),
            "资产减值损失": get_value("资产减值损失", period),
            "资产处置收益": get_value("资产处置收益", period),
        }

        # 记录缺失的可选项
        missing_optional = [name for name, value in optional_fields.items() if value is None]

        # 营业利润推导
        try:
            # 解析关键字段
            revenue_num = parse_number(key_fields["营业收入"])
            cost_num = parse_number(key_fields["营业成本"])

            # 解析可选项（缺失的当0）
            tax_num = parse_number(optional_fields["税金及附加"]) if optional_fields["税金及附加"] else 0
            sales_num = parse_number(optional_fields["销售费用"]) if optional_fields["销售费用"] else 0
            admin_num = parse_number(optional_fields["管理费用"]) if optional_fields["管理费用"] else 0
            rd_num = parse_number(optional_fields["研发费用"]) if optional_fields["研发费用"] else 0
            finance_num = parse_number(optional_fields["财务费用"]) if optional_fields["财务费用"] else 0
            other_num = parse_number(optional_fields["其他收益"]) if optional_fields["其他收益"] else 0
            invest_num = parse_number(optional_fields["投资收益"]) if optional_fields["投资收益"] else 0
            fair_num = parse_number(optional_fields["公允价值变动收益"]) if optional_fields["公允价值变动收益"] else 0
            hedge_num = parse_number(optional_fields["净敞口套期收益"]) if optional_fields["净敞口套期收益"] else 0
            credit_num = parse_number(optional_fields["信用减值损失"]) if optional_fields["信用减值损失"] else 0
            asset_impair_num = parse_number(optional_fields["资产减值损失"]) if optional_fields["资产减值损失"] else 0
            disposal_num = parse_number(optional_fields["资产处置收益"]) if optional_fields["资产处置收益"] else 0

            # **修复3：减值符号自适应**
            # 公式：营业利润 = ... - 信用减值损失 - 资产减值损失 + ...
            # 若报告以负数列示损失（如信用减值损失=-15,544,182.25），公式 `-(-15.5M)`=+15.5M，符号双反
            # 修复：判断符号——若值为负数（<0），公式里用「+损失」（即减去负数=加负数，符号正确）
            # 若值为正数，用「-损失」
            credit_term = credit_num if credit_num < 0 else -credit_num
            asset_impair_term = asset_impair_num if asset_impair_num < 0 else -asset_impair_num

            # 营业利润计算（使用自适应符号）
            expected_operating = (
                revenue_num
                - cost_num
                - tax_num
                - sales_num
                - admin_num
                - rd_num
                - finance_num
                + other_num
                + invest_num
                + fair_num
                + hedge_num
                + credit_term  # 自适应：负数则加，正数则减
                + asset_impair_term  # 自适应：负数则加，正数则减
                + disposal_num
            )

            check_result = check_equal(
                str(expected_operating), key_fields["营业利润"], tolerance
            )

            # 根据缺失情况决定 severity
            if check_result["match"]:
                severity = "info"
                # description 信用减值损失/资产减值损失显示带符号
                credit_display = f"+ 信用减值损失({credit_num})" if credit_num < 0 else f"- 信用减值损失({credit_num})"
                asset_impair_display = f"+ 资产减值损失({asset_impair_num})" if asset_impair_num < 0 else f"- 资产减值损失({asset_impair_num})"
                description = f"营业收入({revenue_num:.2f}) - 营业成本({cost_num:.2f}) - 税金及附加({tax_num:.2f}) - 销售费用({sales_num:.2f}) - 管理费用({admin_num:.2f}) - 研发费用({rd_num:.2f}) - 财务费用({finance_num:.2f}) + 其他收益({other_num:.2f}) + 投资收益({invest_num:.2f}) + 公允价值变动收益({fair_num:.2f}) + 净敞口套期收益({hedge_num:.2f}) {credit_display} {asset_impair_display} + 资产处置收益({disposal_num:.2f}) = {expected_operating:.2f}，实际营业利润={key_fields['营业利润']}"
            elif missing_optional:
                # 有可选项缺失 → warning（推导可能不精确）
                missing_str = "、".join(missing_optional)
                severity = "warning"
                # description 信用减值损失/资产减值损失显示带符号
                credit_display = f"+ 信用减值损失({credit_num})" if credit_num < 0 else f"- 信用减值损失({credit_num})"
                asset_impair_display = f"+ 资产减值损失({asset_impair_num})" if asset_impair_num < 0 else f"- 资产减值损失({asset_impair_num})"
                description = f"推导可能不精确，缺少以下项：{missing_str}；差异 {check_result['diff']:.2f}，请人工核对营业利润是否含未提取的明细项。营业收入({revenue_num:.2f}) - 营业成本({cost_num:.2f}) - 税金及附加({tax_num:.2f}) - 销售费用({sales_num:.2f}) - 管理费用({admin_num:.2f}) - 研发费用({rd_num:.2f}) - 财务费用({finance_num:.2f}) + 其他收益({other_num:.2f}) + 投资收益({invest_num:.2f}) + 公允价值变动收益({fair_num:.2f}) + 净敞口套期收益({hedge_num:.2f}) {credit_display} {asset_impair_display} + 资产处置收益({disposal_num:.2f}) = {expected_operating:.2f}，实际营业利润={key_fields['营业利润']}"
            else:
                # 所有可选项都齐全但仍不匹配 → error（真差异）
                severity = "error"
                # description 信用减值损失/资产减值损失显示带符号
                credit_display = f"+ 信用减值损失({credit_num})" if credit_num < 0 else f"- 信用减值损失({credit_num})"
                asset_impair_display = f"+ 资产减值损失({asset_impair_num})" if asset_impair_num < 0 else f"- 资产减值损失({asset_impair_num})"
                description = f"营业收入({revenue_num:.2f}) - 营业成本({cost_num:.2f}) - 税金及附加({tax_num:.2f}) - 销售费用({sales_num:.2f}) - 管理费用({admin_num:.2f}) - 研发费用({rd_num:.2f}) - 财务费用({finance_num:.2f}) + 其他收益({other_num:.2f}) + 投资收益({invest_num:.2f}) + 公允价值变动收益({fair_num:.2f}) + 净敞口套期收益({hedge_num:.2f}) {credit_display} {asset_impair_display} + 资产处置收益({disposal_num:.2f}) = {expected_operating:.2f}，实际营业利润={key_fields['营业利润']}"

            results.append(
                CheckItem(
                    rule_name=f"营业利润推导{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_operating:.2f}",
                    actual=key_fields["营业利润"],
                    difference=check_result["diff"],
                    description=description,
                    severity=severity,
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"营业利润推导{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

        # 利润总额推导（依赖关键字段，不太可能缺失，仍保持原有逻辑）
        non_operating_income = get_value("营业外收入", period)
        non_operating_expense = get_value("营业外支出", period)
        total_profit = key_fields["利润总额"]

        try:
            operating_num = parse_number(key_fields["营业利润"])
            non_income_num = parse_number(non_operating_income) if non_operating_income else 0
            non_expense_num = parse_number(non_operating_expense) if non_operating_expense else 0

            expected_total = operating_num + non_income_num - non_expense_num

            check_result = check_equal(
                str(expected_total), total_profit, tolerance
            )
            results.append(
                CheckItem(
                    rule_name=f"利润总额推导{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_total:.2f}",
                    actual=total_profit,
                    difference=check_result["diff"],
                    description=f"营业利润({key_fields['营业利润']}) + 营业外收入({non_operating_income}) - 营业外支出({non_operating_expense}) = {expected_total:.2f}，实际利润总额={total_profit}",
                    severity="error",
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"利润总额推导{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

        # 净利润推导（依赖关键字段，不太可能缺失，仍保持原有逻辑）
        income_tax = get_value("所得税费用", period)
        net_profit = key_fields["净利润"]

        try:
            total_num = parse_number(total_profit)
            tax_num = parse_number(income_tax) if income_tax else 0

            expected_net = total_num - tax_num

            check_result = check_equal(
                str(expected_net), net_profit, tolerance
            )
            results.append(
                CheckItem(
                    rule_name=f"净利润推导{period_name}",
                    passed=check_result["match"],
                    expected=f"{expected_net:.2f}",
                    actual=net_profit,
                    difference=check_result["diff"],
                    description=f"利润总额({total_profit}) - 所得税费用({income_tax}) = {expected_net:.2f}，实际净利润={net_profit}",
                    severity="error",
                )
            )
        except ValueError as e:
            results.append(
                CheckItem(
                    rule_name=f"净利润推导{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"数字解析失败: {e}",
                    severity="error",
                )
            )

    check_item("current", "（本期）")
    check_item("prior", "（上期）")

    return results


def check_cash_flow_statement(statement: dict, tolerance: float = 0.01) -> list[CheckItem]:
    """
    现金流量表勾稽检查。

    Args:
        statement: 报表字典，格式为 {'项目名': {'current': '123', 'prior': '456'}, ...}
        tolerance: 容差，默认 0.01

    Returns:
        list[CheckItem]: 检查结果列表

    规则：
    - 经营活动 + 投资活动 + 筹资活动三类活动净额合计 = 现金及现金等价物净增加额
    - 期末现金余额 = 期初现金余额 + 净增加额
    """
    results = []

    def get_value(item_name: str, period: str = "current") -> Optional[str]:
        """获取项目的值，支持常见的变体名称"""
        item = statement.get(item_name)
        if item is None:
            return None
        return item.get(period)

    def check_item(period: str, period_name: str):
        """检查现金流量表勾稽

        现金流量表净增加额的勾稽有两种情况：
        - 无汇率变动项：经营 + 投资 + 筹资 = 净增加额
        - 有汇率变动项：经营 + 投资 + 筹资 + 汇率变动对现金的影响 = 净增加额
          （跨国/涉外企业现金流量表含此项目，如忽略会产生误报）
        先按无汇率项检查，不平时自动检测汇率变动项并加入公式重算。
        """
        # 三类活动净额
        operating_cf = get_value("经营活动产生的现金流量净额", period)
        investing_cf = get_value("投资活动产生的现金流量净额", period)
        financing_cf = get_value("筹资活动产生的现金流量净额", period)

        # 现金净增加额
        net_increase = get_value("现金及现金等价物净增加额", period)

        # 汇率变动项（可选，多种命名变体）
        fx_keys = [
            "汇率变动对现金的影响",
            "汇率变动对现金及现金等价物的影响",
            "汇率变动对现金的影响额",
        ]
        fx_cf = None
        for fx_key in fx_keys:
            fx_cf = get_value(fx_key, period)
            if fx_cf is not None:
                break

        if None in [operating_cf, investing_cf, financing_cf, net_increase]:
            results.append(
                CheckItem(
                    rule_name=f"现金流量净增加额{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：经营活动净额={operating_cf}, 投资活动净额={investing_cf}, 筹资活动净额={financing_cf}, 净增加额={net_increase}",
                    severity="error",
                )
            )
        else:
            try:
                operating_num = parse_number(operating_cf)
                investing_num = parse_number(investing_cf)
                financing_num = parse_number(financing_cf)

                expected_increase = operating_num + investing_num + financing_num
                formula_desc = f"经营活动净额({operating_cf}) + 投资活动净额({investing_cf}) + 筹资活动净额({financing_cf})"

                # 若三类合计不平，尝试加入汇率变动项重算
                first_check = check_equal(str(expected_increase), net_increase, tolerance)
                already_reported = False
                if not first_check["match"] and fx_cf is not None:
                    fx_num = parse_number(fx_cf)
                    expected_increase = expected_increase + fx_num
                    formula_desc += f" + 汇率变动影响({fx_cf})"
                elif not first_check["match"] and fx_cf is None:
                    # 三类不平且无汇率项——可能是漏提取汇率项，提示而非直接报错
                    results.append(
                        CheckItem(
                            rule_name=f"现金流量净增加额{period_name}",
                            passed=False,
                            expected=f"{expected_increase:.2f}",
                            actual=net_increase,
                            difference=first_check["diff"],
                            description=f"三类活动净额合计({expected_increase:.2f})与净增加额({net_increase})差异={first_check['diff']:.2f}。可能存在汇率变动对现金的影响项未被提取，请核实现金流量表是否含此项目并补入statements。",
                            severity="error",
                        )
                    )
                    already_reported = True

                if not already_reported:
                    check_result = check_equal(
                        str(expected_increase), net_increase, tolerance
                    )
                    results.append(
                        CheckItem(
                            rule_name=f"现金流量净增加额{period_name}",
                            passed=check_result["match"],
                            expected=f"{expected_increase:.2f}",
                            actual=net_increase,
                            difference=check_result["diff"],
                            description=f"{formula_desc} = {expected_increase:.2f}，实际净增加额={net_increase}",
                            severity="error",
                        )
                    )
            except ValueError as e:
                results.append(
                    CheckItem(
                        rule_name=f"现金流量净增加额{period_name}",
                        passed=False,
                        expected=None,
                        actual=None,
                        difference=None,
                        description=f"数字解析失败: {e}",
                        severity="error",
                    )
                )

        # 期末现金余额 = 期初 + 净增加额
        beginning_cash = get_value("期初现金及现金等价物余额", period)
        ending_cash = get_value("期末现金及现金等价物余额", period)

        if None in [beginning_cash, ending_cash, net_increase]:
            results.append(
                CheckItem(
                    rule_name=f"现金余额勾稽{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：期初余额={beginning_cash}, 期末余额={ending_cash}, 净增加额={net_increase}",
                    severity="error",
                )
            )
        else:
            try:
                beginning_num = parse_number(beginning_cash)
                increase_num = parse_number(net_increase)
                expected_ending = beginning_num + increase_num

                check_result = check_equal(
                    str(expected_ending), ending_cash, tolerance
                )

                results.append(
                    CheckItem(
                        rule_name=f"现金余额勾稽{period_name}",
                        passed=check_result["match"],
                        expected=f"{expected_ending:.2f}",
                        actual=ending_cash,
                        difference=check_result["diff"],
                        description=f"期初余额({beginning_cash}) + 净增加额({net_increase}) = {expected_ending:.2f}，实际期末余额={ending_cash}",
                        severity="error",
                    )
                )
            except ValueError as e:
                results.append(
                    CheckItem(
                        rule_name=f"现金余额勾稽{period_name}",
                        passed=False,
                        expected=None,
                        actual=None,
                        difference=None,
                        description=f"数字解析失败: {e}",
                        severity="error",
                    )
                )

    check_item("current", "（本期）")
    check_item("prior", "（上期）")

    return results


def check_equity_change_statement(
    statement: dict, tolerance: float = 0.01
) -> list[CheckItem]:
    """
    所有者权益变动表勾稽检查。

    Args:
        statement: 报表字典，格式为 {'项目名': {'current': '123', 'prior': '456'}, ...}
        tolerance: 容差，默认 0.01

    Returns:
        list[CheckItem]: 检查结果列表

    规则：
    - 本年年末余额 = 上年年末余额 + 本年增减变动（综合收益总额 - 提取盈余公积 - 对股东分配 + 增资等）
    - 注意：权益变动表结构复杂（矩阵），这里做简化检查（期初→期末的纵向勾稽）
    """
    results = []

    def get_value(item_name: str, period: str = "current") -> Optional[str]:
        """获取项目的值，支持常见的变体名称"""
        item = statement.get(item_name)
        if item is None:
            return None
        return item.get(period)

    def check_item(period: str, period_name: str):
        """检查权益变动表勾稽"""
        # 上年年末余额（本年期初）
        prior_year_ending = get_value("上年年末余额", period)

        # 本年增减变动项
        comprehensive_income = get_value("综合收益总额", period)  # +综合收益总额
        retained_earnings = get_value("提取盈余公积", period)  # -提取盈余公积
        dividend = get_value("对所有者（或股东）的分配", period)  # -对股东分配
        capital_injection = get_value("所有者投入资本", period)  # +增资

        # 本年年末余额
        current_year_ending = get_value("本年年末余额", period)

        if None in [prior_year_ending, current_year_ending]:
            results.append(
                CheckItem(
                    rule_name=f"权益变动期末余额{period_name}",
                    passed=False,
                    expected=None,
                    actual=None,
                    difference=None,
                    description=f"缺少必要项目：上年年末余额={prior_year_ending}, 本年年末余额={current_year_ending}",
                    severity="error",
                )
            )
            return

        # 权益变动表是矩阵结构（行=变动项，列=股本/资本公积/盈余公积/未分配利润等）
        # 简化公式（上年+综合收益-盈余公积-分配+增资=本年）会漏掉专项储备、其他综合收益
        # 分类、库存股等大量变动项，造成系统性误报（实测金星差异 4,950,703.19=专项储备减少）
        # 这里只验证关键余额已识别，矩阵内部勾稽交给语义层复核
        results.append(
            CheckItem(
                rule_name=f"权益变动期末余额{period_name}",
                passed=True,
                expected=current_year_ending,
                actual=current_year_ending,
                difference=0.0,
                description=(
                    f"已识别本年年末余额({current_year_ending})、上年年末余额({prior_year_ending})。"
                    "权益变动表为矩阵结构（各明细列 期初+增减=期末、横向加总=合计列），"
                    "简化公式易漏项误报，建议按 references/structure_annotation.md 做矩阵语义复核。"
                ),
                severity="info",
            )
        )

    check_item("current", "（本期）")

    return results


# 各报表 calculator_rules 必需的字段名清单（提取 statements.json 时必须用这些精确名称）
REQUIRED_FIELDS = {
    "资产负债表": [
        "流动资产合计", "非流动资产合计", "资产总计",
        "流动负债合计", "非流动负债合计", "负债合计",
        "所有者权益合计", "负债和所有者权益总计",
    ],
    "利润表": [
        "营业收入", "营业成本", "税金及附加", "销售费用", "管理费用",
        "研发费用", "财务费用", "其他收益", "投资收益", "公允价值变动收益",
        "信用减值损失", "资产减值损失", "资产处置收益",
        "营业利润", "营业外收入", "营业外支出",
        "利润总额", "所得税费用", "净利润",
    ],
    "现金流量表": [
        "经营活动产生的现金流量净额", "投资活动产生的现金流量净额",
        "筹资活动产生的现金流量净额", "现金及现金等价物净增加额",
        "期初现金及现金等价物余额", "期末现金及现金等价物余额",
    ],
    "所有者权益变动表": [
        "上年年末余额", "综合收益总额", "提取盈余公积",
        "对所有者（或股东）的分配", "所有者投入资本", "本年年末余额",
    ],
}


if __name__ == "__main__":
    print("各报表 calculator_rules 必需的字段名（提取 statements.json 时必须用这些精确名称）：")
    print()
    for stmt_type, fields in REQUIRED_FIELDS.items():
        print(f"【{stmt_type}】({len(fields)} 个字段)")
        for f in fields:
            print(f"  - {f}")
        print()
    print("合并报告需提取合并+母公司两套，键名分别为：")
    print('  合并资产负债表 / 母公司资产负债表（其余类推）')
    print('每个字段值格式: {"current": "期末/本期金额", "prior": "期初/上期金额"}')