#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
"""
银行流水合并脚本 — 银行流水核对 Skill
从 digit-audit/legacy/modules/bankflowmerge.py 移植核心逻辑，
移除框架依赖(core)，改为独立可运行的命令行工具。

用法:
  python bankflow_merge.py --config config.json
  python bankflow_merge.py --files f1.xlsx f2.xls --header-rows 5 2 --header-cols A A

合并脚本接收 AI 识别的参数，智能读取各种格式的银行流水，
输出统一的12列标准化Excel。
"""
import os
import sys
import json
import argparse
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import numbers
import re
import warnings
import time
from typing import List, Any, Union, Dict
from collections import Counter
from datetime import datetime
import traceback

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')


def get_file_ext(file_path):
    """获取文件扩展名（小写），兼容 Windows 大小写不敏感"""
    return os.path.splitext(file_path)[1].lower()


# ============================================================================
# 列字母工具
# ============================================================================
def increment_column(col: str) -> str:
    """列字母递增（支持 AA, AB 等）"""
    if len(col) == 1:
        return 'AA' if col == 'Z' else chr(ord(col) + 1)
    else:
        prefix = col[:-1]
        last_char = col[-1]
        if last_char == 'Z':
            return increment_column(prefix) + 'A'
        return prefix + chr(ord(last_char) + 1)


COL_TO_IDX: Dict[str, int] = {}
IDX_TO_COL: Dict[int, str] = {}
def _build_col_maps():
    global COL_TO_IDX, IDX_TO_COL
    col = 'A'
    for i in range(702):
        COL_TO_IDX[col] = i
        IDX_TO_COL[i] = col
        col = increment_column(col)
_build_col_maps()


# ============================================================================
# 智能时间解析
# ============================================================================
def smart_parse_datetime(date_val, time_val=None) -> str:
    """
    智能解析日期时间，支持 datetime / Excel序列号 / 多种字符串格式。
    返回 "YYYY-MM-DD HH:MM:SS"。
    """
    try:
        date_str = ""
        time_str = "00:00:00"

        if pd.isna(date_val) or date_val == "":
            return ""

        # datetime 对象（可能包含时间）
        if isinstance(date_val, (pd.Timestamp, datetime)):
            date_str = date_val.strftime("%Y-%m-%d")
            if date_val.hour != 0 or date_val.minute != 0 or date_val.second != 0:
                time_str = date_val.strftime("%H:%M:%S")

        # 数字（Excel 日期序列号）
        elif isinstance(date_val, (int, float)):
            try:
                date_obj = pd.to_datetime('1899-12-30') + pd.Timedelta(days=float(date_val))
                date_str = date_obj.strftime("%Y-%m-%d")
                if date_obj.hour != 0 or date_obj.minute != 0 or date_obj.second != 0:
                    time_str = date_obj.strftime("%H:%M:%S")
            except Exception:
                date_str = str(int(date_val))

        # 字符串
        else:
            date_val_str = str(date_val).strip()
            try:
                date_obj = pd.to_datetime(date_val_str)
                date_str = date_obj.strftime("%Y-%m-%d")
                if date_obj.hour != 0 or date_obj.minute != 0 or date_obj.second != 0:
                    time_str = date_obj.strftime("%H:%M:%S")
            except Exception:
                for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y.%m.%d"]:
                    try:
                        date_obj = datetime.strptime(date_val_str, fmt)
                        date_str = date_obj.strftime("%Y-%m-%d")
                        break
                    except Exception:
                        continue
                if not date_str:
                    date_str = date_val_str

        # 单独的时间部分
        if time_val is not None and not pd.isna(time_val) and time_val != "":
            if isinstance(time_val, (int, float)):
                try:
                    hours = int(float(time_val) * 24)
                    minutes = int((float(time_val) * 24 * 60) % 60)
                    seconds = int((float(time_val) * 24 * 3600) % 60)
                    time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                except Exception:
                    pass
            else:
                time_val_str = str(time_val).strip()
                for fmt in ["%H:%M:%S", "%H:%M", "%H时%M分%S秒", "%H时%M分"]:
                    try:
                        time_obj = datetime.strptime(time_val_str, fmt)
                        time_str = time_obj.strftime("%H:%M:%S")
                        break
                    except Exception:
                        continue
                if not any(c in time_str for c in ":时分"):
                    try:
                        time_num = float(time_val_str)
                        if time_num < 1:
                            hours = int(time_num * 24)
                            minutes = int((time_num * 24 * 60) % 60)
                            seconds = int((time_num * 24 * 3600) % 60)
                            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    except Exception:
                        pass

        return f"{date_str} {time_str}"
    except Exception as e:
        print(f"    WARNING 时间解析失败: {date_val}, {time_val} - {e}")
        return str(date_val) if date_val else ""


# ============================================================================
# 银行名称提取
# ============================================================================
BANK_KEYWORDS = {
    "工商银行": "工商银行", "工行": "工商银行",
    "农业银行": "农业银行", "农行": "农业银行",
    "中国银行": "中国银行", "中行": "中国银行",
    "建设银行": "建设银行", "建行": "建设银行",
    "交通银行": "交通银行", "交行": "交通银行",
    "浦发银行": "浦发银行", "浦发": "浦发银行",
    "民生银行": "民生银行", "民生": "民生银行",
    "招商银行": "招商银行", "招行": "招商银行",
    "兴业银行": "兴业银行", "兴业": "兴业银行",
    "光大银行": "光大银行", "光大": "光大银行",
    "华夏银行": "华夏银行", "华夏": "华夏银行",
    "平安银行": "平安银行", "平安": "平安银行",
    "广发银行": "广发银行", "广发": "广发银行",
    "中信银行": "中信银行", "中信": "中信银行",
    "农商银行": "农商银行", "农商": "农商银行",
    "邮政储蓄": "邮政储蓄", "邮储": "邮政储蓄",
}

def extract_bank_name(text) -> str:
    """从文本提取银行规范简称"""
    if not text or pd.isna(text):
        return ""
    text = str(text)
    for keyword, standard_name in BANK_KEYWORDS.items():
        if keyword in text:
            return standard_name
    return text[:10]


# ============================================================================
# 重复列名处理
# ============================================================================
def _get_renamed_columns_list(original_cols: List[str]) -> List[str]:
    """重复列名自动加序号（如 名称→名称1,名称2）"""
    col_counts = Counter(original_cols)
    counts = {}
    new_cols = []
    duplicates = [name for name, count in col_counts.items() if count > 1]
    if duplicates:
        print(f"    WARNING 重复列名: {', '.join(duplicates[:5])}")
    for col in original_cols:
        if col_counts[col] > 1:
            counts[col] = counts.get(col, 0) + 1
            new_cols.append(f"{col}{counts[col]}")
        else:
            new_cols.append(col)
    return new_cols


# ============================================================================
# 合并单元格检测
# ============================================================================
def _get_max_merged_row(file_path: str, sheet_name: Union[int, str], target_row: int) -> int:
    """检测表头行跨合并单元格的最大行数"""
    max_row_idx = -1
    try:
        if get_file_ext(file_path) in ('.xlsx', '.xlsm'):
            workbook = load_workbook(file_path, read_only=False)
            sheet = workbook[sheet_name] if isinstance(sheet_name, str) else workbook.worksheets[sheet_name]
            for merged_cell_range in sheet.merged_cells.ranges:
                if merged_cell_range.min_row <= target_row <= merged_cell_range.max_row:
                    max_row_idx = max(max_row_idx, merged_cell_range.max_row - 1)
            workbook.close()
        elif get_file_ext(file_path) == '.xls':
            import xlrd
            workbook = xlrd.open_workbook(file_path, formatting_info=True)
            sheet = workbook.sheet_by_name(sheet_name) if isinstance(sheet_name, str) else workbook.sheet_by_index(sheet_name)
            target_row_idx = target_row - 1
            for rlo, rhi, clo, chi in sheet.merged_cells:
                if rlo <= target_row_idx < rhi:
                    max_row_idx = max(max_row_idx, rhi - 1)
    except Exception as e:
        print(f"    WARNING 合并单元格分析失败: {e}")
        return target_row - 1
    return max_row_idx if max_row_idx > -1 else target_row - 1


# ============================================================================
# 核心：智能表头读取
# ============================================================================
def read_data_with_smart_header(
    file_path: str,
    header_start_row: int,       # 1-based
    header_start_col: str,        # 'A', 'B', 'C'
    sheet_name_or_idx: Union[int, str] = 0
) -> pd.DataFrame:
    """
    智能读取 Excel，自动处理多行表头、合并单元格、重复列名。
    返回 dtype=str 的 DataFrame。
    """
    # 解析 sheet 名称
    sheet_name = sheet_name_or_idx
    if isinstance(sheet_name, int):
        try:
            wb = load_workbook(file_path, read_only=True)
            sheet_name = wb.sheetnames[sheet_name]
            wb.close()
        except Exception:
            try:
                import xlrd
                wb = xlrd.open_workbook(file_path)
                sheet_name = wb.sheet_names()[sheet_name]
            except Exception:
                sheet_name = 0

    header_start_row_0_idx = header_start_row - 1
    header_start_col_0_idx = COL_TO_IDX.get(header_start_col, 0)

    max_merged_row_idx = _get_max_merged_row(file_path, sheet_name, target_row=header_start_row)

    # 预览数据找到数据起始行
    try:
        df_preview = pd.read_excel(file_path, header=None, skiprows=header_start_row_0_idx,
                                   nrows=20, sheet_name=sheet_name, dtype=str)
    except Exception as e:
        print(f"    ERROR 读取预览失败: {e}")
        return pd.DataFrame()

    first_data_content_row_idx = -1
    if header_start_col_0_idx < len(df_preview.columns):
        for i in range(1, len(df_preview)):
            if pd.notna(df_preview.iloc[i, header_start_col_0_idx]):
                first_data_content_row_idx = header_start_row_0_idx + i
                break

    if first_data_content_row_idx == -1:
        first_data_row_final = max_merged_row_idx + 1 if max_merged_row_idx > -1 else header_start_row_0_idx + 1
    else:
        first_data_row_final = max(max_merged_row_idx + 1, first_data_content_row_idx)

    num_header_rows = first_data_row_final - header_start_row_0_idx
    if num_header_rows <= 0:
        num_header_rows = 1

    print(f"    表头: 起始行={header_start_row}, 共 {num_header_rows} 行")

    # 使用 openpyxl 读取 .xlsx 表头（避免 pandas 类型推断问题）
    try:
        if get_file_ext(file_path) in ('.xlsx', '.xlsm'):
            wb = load_workbook(file_path, data_only=True)
            ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]
            header_data = []
            max_col = ws.max_column
            for row_idx in range(header_start_row, header_start_row + num_header_rows):
                row_values = []
                for col_idx in range(header_start_col_0_idx, max_col):
                    cell_value = ws.cell(row_idx, col_idx + 1).value
                    row_values.append(str(cell_value) if cell_value is not None else '')
                header_data.append(row_values)
            wb.close()
            df_header = pd.DataFrame(header_data)
        else:
            df_header_full = pd.read_excel(file_path, header=None, skiprows=header_start_row_0_idx,
                                           nrows=num_header_rows, sheet_name=sheet_name, dtype=str)
            df_header = df_header_full.iloc[:, header_start_col_0_idx:].copy()
    except Exception as e:
        print(f"    ERROR 读取表头失败: {e}")
        traceback.print_exc()
        return pd.DataFrame()

    if df_header.empty:
        print("    WARNING 表头为空")
        return pd.DataFrame()

    # 填充合并单元格 (ffill 水平+垂直)
    df_header = df_header.replace('', pd.NA)
    df_header = df_header.apply(lambda x: x.ffill(), axis=1).ffill(axis=0)
    df_header = df_header.fillna('')

    # 合并多行表头
    if num_header_rows > 1:
        new_columns = [
            "".join(list(dict.fromkeys([
                str(part).strip().replace('\u3000', '').replace('\xa0', ' ').strip()
                for part in df_header[col].tolist() if pd.notna(part)
            ])))
            for col in df_header.columns
        ]
    else:
        if df_header.empty:
            df_data_preview = pd.read_excel(file_path, header=None,
                                           skiprows=header_start_row_0_idx + num_header_rows,
                                           nrows=5, sheet_name=sheet_name)
            df_data_preview = df_data_preview.iloc[:, header_start_col_0_idx:]
            num_cols = len(df_data_preview.columns)
            new_columns = [f"列{i+1}" for i in range(num_cols)]
        else:
            new_columns = [
                str(col).strip().replace('\u3000', '').replace('\xa0', ' ').strip()
                for col in df_header.iloc[0].tolist()
            ]

    # 读取数据
    data_start_row = header_start_row_0_idx + num_header_rows
    num_cols_to_read = len(new_columns)
    if num_cols_to_read == 0:
        print("    WARNING 未解析到表头列")
        return pd.DataFrame()

    try:
        df_data = pd.read_excel(file_path, header=None, skiprows=data_start_row,
                               sheet_name=sheet_name, dtype=str)
        if header_start_col_0_idx > 0 or len(df_data.columns) > num_cols_to_read:
            end_col_idx = header_start_col_0_idx + num_cols_to_read
            df_data = df_data.iloc[:, header_start_col_0_idx:end_col_idx]
        if len(df_data.columns) < num_cols_to_read:
            for i in range(len(df_data.columns), num_cols_to_read):
                df_data[i] = ''
        elif len(df_data.columns) > num_cols_to_read:
            df_data = df_data.iloc[:, :num_cols_to_read]
    except Exception as e:
        print(f"    ERROR 读取数据失败: {e}")
        return pd.DataFrame()

    unique_columns = _get_renamed_columns_list(new_columns)
    df_data.columns = unique_columns
    df_data.fillna('', inplace=True)
    try:
        df_data = df_data.map(lambda x: x.strip() if isinstance(x, str) else x)
    except AttributeError:
        # pandas < 2.1 compatibility
        df_data = df_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df_data.dropna(how='all', inplace=True)
    return df_data


# ============================================================================
# 工具函数
# ============================================================================
def clean_value(val) -> str:
    """清理无效值"""
    if not val or val in ['未知', 'unknown', 'Unknown', 'UNKNOWN', 'N/A', 'n/a']:
        return ""
    return str(val).strip()


def clean_account_number(val) -> str:
    """只保留纯数字银行账号（>= 5位），过滤掉含中文的账户名称"""
    if not val:
        return ""
    val_str = str(val).strip()
    if val_str in ['未知', 'unknown', 'Unknown', 'UNKNOWN', 'N/A', 'n/a']:
        return ""
    if any('\u4e00' <= char <= '\u9fff' for char in val_str):
        return ""
    digits_only = ''.join(filter(str.isdigit, val_str))
    if len(digits_only) < 5:
        return ""
    return digits_only


def ensure_list(param, default_value='', convert_to_int=False, num_files=1):
    """确保参数是列表并补齐长度"""
    if not isinstance(param, list):
        param = [param] if param else []
    if convert_to_int:
        param = [int(x) if str(x).strip() != '' else default_value for x in param]
    while len(param) < num_files:
        param.append(default_value)
    return param


# ============================================================================
# 标准化输出列定义（固定12列）
# ============================================================================
STANDARD_COLUMNS = [
    '来源目录', '文件名', '银行名称', '银行户名', '银行账号',
    '交易时间', '对方账户', '对方账号', '摘要',
    '收入金额', '支出金额', '余额'
]


# ============================================================================
# 尾部行跳过
# ============================================================================
# 银行流水文件末尾常有汇总行（如交通银行的"借方交易笔数/贷方交易金额"）。
# AI 在读取文件时应同时检查前几行和后几行，发现尾部非交易行后，
# 通过 config 的 skip_footer_rows 参数指定跳过行数。


def _skip_footer_rows(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """根据 config['skip_footer_rows'] 跳过尾部 N 行。"""
    if df.empty:
        return df

    skip_footer = config.get('skip_footer_rows', 0)
    if skip_footer and isinstance(skip_footer, int) and skip_footer > 0:
        if skip_footer >= len(df):
            print(f"  WARNING skip_footer_rows({skip_footer}) >= 数据行数({len(df)})，跳过全部")
            return pd.DataFrame()
        df = df.iloc[:-skip_footer].copy()
        print(f"  跳过尾部 {skip_footer} 行")

    return df


# ============================================================================
# 处理单个文件
# ============================================================================
def process_single_file(file_path: str, config: dict) -> pd.DataFrame:
    """
    处理单个银行流水文件，返回标准化12列 DataFrame。
    config 键:
      header_start_row, header_start_col (必填)
      bank_name, account_name, account_number (元数据)
      transaction_date_col, transaction_time_col (日期时间)
      counterparty_name_col, counterparty_account_col, summary_col
      income_col, expense_col, balance_col
    """
    header_row = config.get('header_start_row', 1)
    header_col = config.get('header_start_col', 'A')

    df = read_data_with_smart_header(file_path, header_row, header_col, 0)
    if df.empty:
        print("  WARNING 未读取到数据")
        return pd.DataFrame()

    print(f"  读取到列: {list(df.columns)}")

    # 跳过尾部汇总行/页脚行（AI 判断后通过 skip_footer_rows 指定）
    df = _skip_footer_rows(df, config)
    if df.empty:
        print("  WARNING 过滤后无有效数据")
        return pd.DataFrame()

    # 元数据
    bank_name = clean_value(config.get('bank_name', ''))
    account_name = clean_value(config.get('account_name', ''))
    account_number = clean_account_number(config.get('account_number', ''))
    if not bank_name:
        bank_name = extract_bank_name(os.path.basename(file_path))

    num_rows = len(df)
    result_df = pd.DataFrame({
        '来源目录': [os.path.dirname(file_path)] * num_rows,
        '文件名': [os.path.basename(file_path)] * num_rows,
        '银行名称': [bank_name] * num_rows,
        '银行户名': [account_name] * num_rows,
        '银行账号': [account_number] * num_rows,
    })

    # 交易时间（智能合并日期+时间）
    date_col = config.get('transaction_date_col', '')
    time_col = config.get('transaction_time_col', '')
    if date_col:
        if date_col in df.columns:
            if time_col and time_col in df.columns:
                result_df['交易时间'] = df.apply(
                    lambda row: smart_parse_datetime(row[date_col], row[time_col]), axis=1
                )
            else:
                result_df['交易时间'] = df[date_col].apply(lambda x: smart_parse_datetime(x))
        else:
            print(f"    ⚠️ 警告: 配置的日期列 [{date_col}] 不存在，可用列: {list(df.columns)}")
            result_df['交易时间'] = ""
    else:
        result_df['交易时间'] = ""

    # 对方账户
    col = config.get('counterparty_name_col', '')
    if col:
        if col in df.columns:
            result_df['对方账户'] = df[col].astype(str).values
        else:
            print(f"    ⚠️ 警告: 配置的对方账户列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['对方账户'] = ""
    else:
        result_df['对方账户'] = ""

    # 对方账号
    col = config.get('counterparty_account_col', '')
    if col:
        if col in df.columns:
            result_df['对方账号'] = df[col].astype(str).values
        else:
            print(f"    ⚠️ 警告: 配置的对方账号列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['对方账号'] = ""
    else:
        result_df['对方账号'] = ""

    # 摘要
    col = config.get('summary_col', '')
    if col:
        if col in df.columns:
            result_df['摘要'] = df[col].astype(str).values
        else:
            print(f"    ⚠️ 警告: 配置的摘要列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['摘要'] = ""
    else:
        result_df['摘要'] = ""

    # 收入金额
    col = config.get('income_col', '')
    if col:
        if col in df.columns:
            result_df['收入金额'] = df[col].values
        else:
            print(f"    ⚠️ 警告: 配置的收入金额列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['收入金额'] = ""
    else:
        result_df['收入金额'] = ""

    # 支出金额
    col = config.get('expense_col', '')
    if col:
        if col in df.columns:
            result_df['支出金额'] = df[col].values
        else:
            print(f"    ⚠️ 警告: 配置的支出金额列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['支出金额'] = ""
    else:
        result_df['支出金额'] = ""

    # 余额
    col = config.get('balance_col', '')
    if col:
        if col in df.columns:
            result_df['余额'] = df[col].values
        else:
            print(f"    ⚠️ 警告: 配置的余额列 [{col}] 不存在，可用列: {list(df.columns)}")
            result_df['余额'] = ""
    else:
        result_df['余额'] = ""

    return result_df[STANDARD_COLUMNS]


# ============================================================================
# 批量合并
# ============================================================================
def merge_bank_flows(config: dict, output_file: str = None) -> dict:
    """
    批量合并多个银行流水文件。

    config 键（均为列表，长度等于文件数）:
      file_paths (必填), header_start_rows, header_start_cols,
      bank_names, account_names, account_numbers,
      transaction_date_cols, transaction_time_cols,
      counterparty_name_cols, counterparty_account_cols, summary_cols,
      income_cols, expense_cols, balance_cols

    返回 {status, files_processed, files_failed, total_rows, output_file}
    """
    start_time = time.time()
    file_paths = config.get('file_paths', [])
    if not file_paths:
        return {"status": "error", "msg": "缺少 file_paths"}
    if not isinstance(file_paths, list):
        file_paths = [file_paths]

    num_files = len(file_paths)
    header_start_rows = ensure_list(config.get('header_start_rows', []), 1, convert_to_int=True, num_files=num_files)
    header_start_cols = ensure_list(config.get('header_start_cols', []), 'A', num_files=num_files)

    bank_names = [clean_value(x) for x in ensure_list(config.get('bank_names', []), '', num_files=num_files)]
    account_names = [clean_value(x) for x in ensure_list(config.get('account_names', []), '', num_files=num_files)]
    account_numbers = [clean_account_number(x) for x in ensure_list(config.get('account_numbers', []), '', num_files=num_files)]

    transaction_date_cols = ensure_list(config.get('transaction_date_cols', []), '', num_files=num_files)
    transaction_time_cols = ensure_list(config.get('transaction_time_cols', []), '', num_files=num_files)
    counterparty_name_cols = ensure_list(config.get('counterparty_name_cols', []), '', num_files=num_files)
    counterparty_account_cols = ensure_list(config.get('counterparty_account_cols', []), '', num_files=num_files)
    summary_cols = ensure_list(config.get('summary_cols', []), '', num_files=num_files)
    income_cols = ensure_list(config.get('income_cols', []), '', num_files=num_files)
    expense_cols = ensure_list(config.get('expense_cols', []), '', num_files=num_files)
    balance_cols = ensure_list(config.get('balance_cols', []), '', num_files=num_files)
    skip_footer_rows_list = ensure_list(config.get('skip_footer_rows', []), 0, convert_to_int=True, num_files=num_files)

    # 确定输出路径
    if output_file is None:
        out_dir = os.path.dirname(file_paths[0]) if file_paths else '.'
        output_file = os.path.join(out_dir, "银行流水_合并结果.xlsx")

    print(f"批量处理: {num_files} 个文件")
    print(f"输出文件: {output_file}")

    all_results = []
    success_count = 0
    fail_count = 0
    total_rows = 0

    for i in range(num_files):
        file_path = file_paths[i]
        print(f"\n[{i+1}/{num_files}] {os.path.basename(file_path)}")

        if not os.path.exists(file_path):
            print(f"  ERROR 文件不存在")
            fail_count += 1
            continue
        if get_file_ext(file_path) not in ('.xls', '.xlsx', '.xlsm'):
            print(f"  ERROR 不支持的格式")
            fail_count += 1
            continue

        file_config = {
            'header_start_row': header_start_rows[i],
            'header_start_col': header_start_cols[i],
            'bank_name': bank_names[i],
            'account_name': account_names[i],
            'account_number': account_numbers[i],
            'transaction_date_col': transaction_date_cols[i],
            'transaction_time_col': transaction_time_cols[i],
            'counterparty_name_col': counterparty_name_cols[i],
            'counterparty_account_col': counterparty_account_cols[i],
            'summary_col': summary_cols[i],
            'income_col': income_cols[i],
            'expense_col': expense_cols[i],
            'balance_col': balance_cols[i],
            'skip_footer_rows': skip_footer_rows_list[i],
        }

        try:
            file_result = process_single_file(file_path, file_config)
            if file_result is None or file_result.empty:
                print(f"  WARNING 未提取到数据")
                fail_count += 1
                continue
            all_results.append(file_result)
            success_count += 1
            total_rows += len(file_result)
            print(f"  OK 提取 {len(file_result)} 行")
        except Exception as e:
            print(f"  ERROR {e}")
            traceback.print_exc()
            fail_count += 1

    if not all_results:
        return {"status": "error", "msg": "所有文件处理失败"}

    final_df = pd.concat(all_results, ignore_index=True)

    # 写入 Excel（账号列文本格式，金额列数字格式）
    print(f"\n合并: 成功 {success_count}, 失败 {fail_count}, 共 {total_rows} 行")
    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        ws = writer.sheets['Sheet1']

        # 银行账号(5列) 文本格式
        for row in range(2, len(final_df) + 2):
            cell = ws.cell(row=row, column=5)
            cell.number_format = numbers.FORMAT_TEXT
            if cell.value and isinstance(cell.value, (int, float)):
                cell.value = str(int(cell.value))

        # 对方账号(8列) 文本格式
        for row in range(2, len(final_df) + 2):
            cell = ws.cell(row=row, column=8)
            cell.number_format = numbers.FORMAT_TEXT
            if cell.value and isinstance(cell.value, (int, float)):
                cell.value = str(int(cell.value))

        # 金额列(10,11,12) 数字格式
        for col_idx in [10, 11, 12]:
            for row in range(2, len(final_df) + 2):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value:
                    try:
                        val = str(cell.value).replace(',', '').strip()
                        if val and val != '':
                            cell.value = float(val)
                            cell.number_format = '#,##0.00'
                    except (ValueError, AttributeError):
                        pass

    duration = time.time() - start_time
    print(f"\n{'='*60}")
    print(f"耗时: {duration:.1f}s | 文件: {success_count} | 行数: {total_rows}")
    print(f"输出: {output_file}")
    print(f"{'='*60}")

    return {
        "status": "success",
        "files_processed": success_count,
        "files_failed": fail_count,
        "total_rows": total_rows,
        "output_file": output_file,
    }


# ============================================================================
# 命令行入口
# ============================================================================
def main():
    parser = argparse.ArgumentParser(description='银行流水批量合并')
    parser.add_argument('--config', type=str, help='JSON 配置文件路径')
    parser.add_argument('--output', type=str, default=None, help='输出文件路径')
    args = parser.parse_args()

    if not args.config:
        parser.print_help()
        sys.exit(1)

    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)

    result = merge_bank_flows(config, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
