#!/usr/bin/env python3
"""apply_manifest.py — 纯机械读 manifest.json 抽取四表 → statements.json。

零猜测：不猜列、不猜 sheet、不硬编码 ALIAS。所有结构信息来自 manifest.json。
manifest 由 AI 生成（首份手写/后续 DeepSeek 自动），本脚本只负责按契约抽数。

用法：python3 apply_manifest.py <manifest.json> [-o statements.json]
      python3 apply_manifest.py <报告目录> [-o statements.json]
"""
import json
import sys
import argparse
from pathlib import Path

from extractors.base import extract_linear, extract_matrix
from extractors.xlsx import load_xlsx_sheets, iter_xlsx_rows
from extractors.docx import load_extracted_tables, iter_docx_rows


def apply_aliases(data: dict, aliases: dict) -> dict:
    """对提取的数据套用 field_aliases 标准化项目名。"""
    if not aliases:
        return data
    return {aliases.get(k, k): v for k, v in data.items()}


def find_source_file(base_dir: Path, source: str) -> Path:
    """找源文件：base_dir、报告根目录(base_dir.parent)。

    不递归搜索所有子目录，避免误拿 backup/archive 等归档文件。
    """
    candidates = [base_dir / source, base_dir.parent / source]
    for cand in candidates:
        if cand.exists():
            return cand
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest", help="manifest.json 路径或报告目录")
    ap.add_argument("-o", "--output", default="statements.json")
    args = ap.parse_args()

    input_path = Path(args.manifest)
    if input_path.is_dir():
        manifest_path = input_path / "manifest.json"
        if not manifest_path.exists():
            for sub in input_path.iterdir():
                if sub.is_dir() and (sub.name in ("parse_output",) or sub.name.endswith("_output")):
                    cand = sub / "manifest.json"
                    if cand.exists():
                        manifest_path = cand
                        break
        if not manifest_path.exists():
            print("[ERROR] 找不到 manifest.json", file=sys.stderr)
            sys.exit(1)
    else:
        manifest_path = input_path

    mf = json.loads(manifest_path.read_text(encoding="utf-8"))
    base_dir = manifest_path.parent
    files_info = mf.get("files", {})

    xlsx_cache = {}
    extracted_data = None
    statements = {}
    merge_jobs = []  # (target, data) 延迟合并，避免顺序依赖

    for stmt_name, cfg in mf.get("statements_map", {}).items():
        source = cfg.get("source", "")
        if not source:
            print(f"[WARN] {stmt_name}: 缺 source", file=sys.stderr)
            continue

        source_path = find_source_file(base_dir, source)
        if not source_path:
            print(f"[ERROR] 找不到源文件: {source}", file=sys.stderr)
            continue

        loader = "xlsx"
        if source in files_info:
            loader = files_info[source].get("loader", "xlsx")

        aliases = cfg.get("field_aliases", {})
        data = None

        if loader in ("xlsx", "xlsm"):
            if str(source_path) not in xlsx_cache:
                xlsx_cache[str(source_path)] = load_xlsx_sheets(source_path)
            sheets = xlsx_cache[str(source_path)]

            if cfg.get("kind") == "matrix":
                main_rows = iter_xlsx_rows(sheets, cfg.get("sheets_main", []))
                prior_rows = iter_xlsx_rows(sheets, cfg.get("sheets_prior", []))
                data = extract_matrix(main_rows, prior_rows, cfg.get("value_col", 2))
            else:
                rows = iter_xlsx_rows(sheets, cfg.get("sheets", []))
                data = extract_linear(rows, cfg["item_col"], cfg["current_col"], cfg["prior_col"])

        elif loader in ("docx", "pdf"):
            if extracted_data is None:
                extracted_data = load_extracted_tables(manifest_path)
                if extracted_data is None:
                    print("[ERROR] 找不到 extracted_tables.json", file=sys.stderr)
                    sys.exit(1)

            if cfg.get("kind") == "matrix":
                main_rows = iter_docx_rows(extracted_data, cfg.get("sheets_main", []))
                prior_rows = iter_docx_rows(extracted_data, cfg.get("sheets_prior", []))
                data = extract_matrix(main_rows, prior_rows, cfg.get("value_col", 2))
            else:
                refs = cfg.get("sheets", [])
                if not refs:
                    # 兼容旧 manifest 的 source_table/source_table 列表
                    refs = cfg.get("source_table", [])
                    if not isinstance(refs, list):
                        refs = [refs]
                rows = iter_docx_rows(extracted_data, refs)
                data = extract_linear(rows, cfg["item_col"], cfg["current_col"], cfg["prior_col"])

        elif loader == "doc":
            print(f"[WARN] {stmt_name}: doc loader 不支持（报表数据通常不在 .doc）", file=sys.stderr)
            continue
        else:
            print(f"[WARN] {stmt_name}: 未知 loader {loader}", file=sys.stderr)
            continue

        if data:
            data = apply_aliases(data, aliases)

        target = cfg.get("merge_into")
        if target:
            merge_jobs.append((target, data))
            print(f"{stmt_name}: {len(data)}字段" + (f" →merge→ {target}" if target else ""), file=sys.stderr)
        else:
            statements[stmt_name] = data
            print(f"{stmt_name}: {len(data)}字段", file=sys.stderr)

    # 先确保所有 merge target 存在，再统一合并，消除对 statements_map 顺序的依赖
    for target, data in merge_jobs:
        if target not in statements:
            statements[target] = {}
        for k, v in data.items():
            if k not in statements[target]:
                statements[target][k] = v

    out = Path(args.output)
    out.write_text(json.dumps(statements, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"→ {out}", file=sys.stderr)


if __name__ == "__main__":
    main()
