#!/usr/bin/env python3
"""
中国国家法律法规数据库 (flk.npc.gov.cn) API 客户端
基于 2025年底改版后的新 API 接口逆向工程实现
"""

import argparse
import json
import re
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime

BASE_URL = "https://flk.npc.gov.cn/law-search"

# 法规分类编码
FLFG_CODES = {
    "宪法": [100],
    "法律": [102],
    "行政法规": [210],
    "监察法规": [220],
    "地方性法规": [230],
    "地方法规": [230],
    "司法解释": [320, 330, 340],
}

# 时效性编码
SXX_MAP = {1: "已废止", 2: "已修改", 3: "有效", 4: "尚未生效"}
SXX_REVERSE = {"已废止": 1, "已修改": 2, "有效": 3, "尚未生效": 4}


def _post_json(path: str, data: dict, timeout: int = 30) -> dict:
    url = f"{BASE_URL}{path}"
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _get_json(path: str, timeout: int = 30) -> dict:
    url = f"{BASE_URL}{path}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _clean_highlight(text: str) -> str:
    """Remove <em> highlight tags from title"""
    return re.sub(r"</?em[^>]*>", "", text) if text else text


def search(keyword: str, exact: bool = False, law_type: str = None,
           sxx_filter: list = None, gbrq_from: str = None, gbrq_to: str = None,
           content_search: bool = False,
           page: int = 1, size: int = 20) -> dict:
    """搜索法规列表
    
    Args:
        content_search: True=按正文内容搜索, False=按标题搜索(默认)
    """
    data = {
        "searchRange": 2 if content_search else 1,
        "searchContent": keyword,
        "searchType": 1 if exact else 2,
        "sxrq": [],
        "gbrq": [gbrq_from, gbrq_to] if gbrq_from and gbrq_to else [],
        "gbrqYear": [],
        "flfgCodeId": FLFG_CODES.get(law_type, []) if law_type else [],
        "zdjgCodeId": [],
        "sxx": sxx_filter or [],
        "xgzlSearch": False,
        "orderByParam": {"order": "-1", "sort": ""},
        "pageNum": page,
        "pageSize": size,
    }
    result = _post_json("/search/list", data)
    # Clean highlights in titles
    for row in result.get("rows", []):
        row["title"] = _clean_highlight(row.get("title", ""))
        row["sxx_text"] = SXX_MAP.get(row.get("sxx"), "未知")
    return result


def detail(bbbs_id: str) -> dict:
    """获取法规详情"""
    result = _get_json(f"/search/flfgDetails?bbbs={bbbs_id}")
    return result.get("data", result)


def search_all_pages(keyword: str, exact: bool = False, law_type: str = None,
                     content_search: bool = False, sxx_filter: list = None,
                     max_pages: int = 50, page_size: int = 20, delay: float = 0.5) -> dict:
    """搜索所有页的结果，自动翻页"""
    all_rows = []
    page = 1
    while page <= max_pages:
        resp = search(keyword, exact=exact, law_type=law_type,
                      content_search=content_search, sxx_filter=sxx_filter,
                      page=page, size=page_size)
        rows = resp.get("rows", [])
        total = resp.get("total", 0)
        all_rows.extend(rows)
        if len(all_rows) >= total or not rows:
            break
        page += 1
        if page <= max_pages and len(all_rows) < total:
            time.sleep(delay)
            print(f"  翻页: {page}/{(total + page_size - 1) // page_size} (已获取 {len(all_rows)}/{total})", file=sys.stderr)
    return {"total": total, "rows": all_rows}


def new_since(date_str: str, law_type: str = None, all_types: bool = False,
              page: int = 1, size: int = 100) -> dict:
    """查找指定日期后新发布的法规（单页）"""
    today = datetime.now().strftime("%Y-%m-%d")
    codes = []
    if not all_types:
        if law_type:
            codes = FLFG_CODES.get(law_type, [])
        else:
            codes = FLFG_CODES["法律"] + FLFG_CODES["行政法规"]
    data = {
        "searchRange": 1,
        "searchContent": "",
        "searchType": 2,
        "sxrq": [],
        "gbrq": [date_str, today],
        "gbrqYear": [],
        "flfgCodeId": codes,
        "zdjgCodeId": [],
        "sxx": [],
        "xgzlSearch": False,
        "orderByParam": {"order": "-1", "sort": ""},
        "pageNum": page,
        "pageSize": size,
    }
    result = _post_json("/search/list", data)
    for row in result.get("rows", []):
        row["title"] = _clean_highlight(row.get("title", ""))
        row["sxx_text"] = SXX_MAP.get(row.get("sxx"), "未知")
    return result


def new_since_all(date_str: str, law_type: str = None, all_types: bool = False,
                  page_size: int = 50, delay: float = 0.5) -> dict:
    """查找指定日期后新发布的所有法规（自动翻页）"""
    all_rows = []
    page = 1
    while True:
        resp = new_since(date_str, law_type=law_type, all_types=all_types,
                         page=page, size=page_size)
        rows = resp.get("rows", [])
        total = resp.get("total", 0)
        all_rows.extend(rows)
        if len(all_rows) >= total or not rows:
            break
        page += 1
        time.sleep(delay)
        print(f"  翻页: {page} (已获取 {len(all_rows)}/{total})", file=sys.stderr)
    return {"total": total, "rows": all_rows}


def batch_check(names: list, delay: float = 0.5) -> list:
    """批量检查法规状态"""
    results = []
    total = len(names)
    for i, name in enumerate(names):
        name = name.strip()
        if not name:
            continue
        # Clean name: remove book title marks and angle brackets
        clean = re.sub(r'[《》<>〈〉]', '', name).strip()
        try:
            resp = search(clean, exact=True, size=10)
            rows = resp.get("rows", [])
            if rows:
                # Find best match: prefer exact title match, then containment
                best = None
                for row in rows:
                    title = row.get("title", "")
                    title_clean = re.sub(r'[《》<>〈〉]', '', title).strip()
                    if title_clean == clean or title == name:
                        # For laws with multiple versions, prefer sxx=3 (有效)
                        if best is None or row.get("sxx") == 3:
                            best = row
                if best is None:
                    # Fallback: containment match, but only if clean name
                    # is a substantial part of the title (avoid false positives)
                    for row in rows:
                        title = row.get("title", "")
                        title_clean = re.sub(r'[《》<>〈〉]', '', title).strip()
                        # Only match if the clean name IS the core of the title
                        # (not just a substring of a much longer title)
                        if clean in title_clean and len(clean) > len(title_clean) * 0.5:
                            if best is None or row.get("sxx") == 3:
                                best = row
                        elif title_clean in clean and len(title_clean) > len(clean) * 0.5:
                            if best is None or row.get("sxx") == 3:
                                best = row
                if best:
                    bbbs = best.get("bbbs", "")
                    results.append({
                        "name": name,
                        "found": True,
                        "title": best["title"],
                        "status": best.get("sxx_text", "未知"),
                        "sxx": best.get("sxx"),
                        "gbrq": best.get("gbrq", ""),
                        "sxrq": best.get("sxrq", ""),
                        "bbbs": bbbs,
                        "flxz": best.get("flxz", ""),
                        "url": f"https://flk.npc.gov.cn/detail?id={bbbs}" if bbbs else "",
                        "total_versions": resp.get("total", 0),
                    })
                else:
                    results.append({"name": name, "found": False, "status": "未找到（搜索结果不匹配）"})
            else:
                results.append({"name": name, "found": False, "status": "未找到"})
        except Exception as e:
            results.append({"name": name, "found": False, "status": f"查询失败: {e}"})
        if i < total - 1:
            time.sleep(delay)
        # Progress
        if (i + 1) % 10 == 0 or i == total - 1:
            print(f"  进度: {i+1}/{total}", file=sys.stderr)
    return results


def download_info(bbbs_id: str, fmt: str = "docx") -> dict:
    """获取法规文件的签名下载链接"""
    result = _get_json(f"/download/pc?format={fmt}&bbbs={bbbs_id}")
    data = result.get("data", {})
    return {
        "url": data.get("url", ""),
        "url_internal": data.get("urlIn", ""),
    }


def make_download_filename(title: str, gbrq: str, flxz: str, sxx: int,
                           fmt: str = "docx") -> str:
    """生成下载文件名: 法规类型_法规名称_公布日期_状态.ext"""
    status = SXX_MAP.get(sxx, "未知")
    flxz = flxz or "未分类"
    parts = [flxz, title]
    if gbrq:
        parts.append(gbrq)
    parts.append(status)
    name = "_".join(parts) + f".{fmt}"
    return re.sub(r'[\\/:*?"<>|]', '_', name)


def download_file(bbbs_id: str, output_dir: str = ".", fmt: str = "docx",
                   filename: str = None) -> str:
    """下载法规文件到本地
    
    文件命名规则: 法规类型_法规名称_公布日期_状态.ext
    例如: 法律_中华人民共和国安全生产法_2021-06-10_有效.docx
    """
    import os
    info = download_info(bbbs_id, fmt)
    url = info.get("url")
    if not url:
        raise ValueError(f"No download URL for {bbbs_id} format={fmt}")

    if not filename:
        d = detail(bbbs_id)
        filename = make_download_filename(
            title=d.get("title", bbbs_id),
            gbrq=d.get("gbrq", ""),
            flxz=d.get("flxz", ""),
            sxx=d.get("sxx"),
            fmt=fmt,
        )

    filepath = os.path.join(output_dir, filename)
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=60) as resp:
        with open(filepath, "wb") as f:
            while True:
                chunk = resp.read(8192)
                if not chunk:
                    break
                f.write(chunk)
    return filepath


# ── CLI ──────────────────────────────────────────────────────────────

def _fmt_row(r: dict) -> str:
    return (f"  [{r.get('flxz','')}] {r['title']} | "
            f"状态:{r.get('sxx_text','?')} | "
            f"公布:{r.get('gbrq','')} | 施行:{r.get('sxrq','')} | "
            f"id:{r.get('bbbs','')}")


def cmd_search(args):
    if args.all_pages:
        resp = search_all_pages(args.keyword, exact=args.exact, law_type=args.type,
                                content_search=args.content, page_size=args.size)
    else:
        resp = search(args.keyword, exact=args.exact, law_type=args.type,
                      content_search=args.content, size=args.size)
    print(f"共 {resp.get('total', 0)} 条结果（返回 {len(resp.get('rows',[]))} 条）：")
    for r in resp.get("rows", []):
        print(_fmt_row(r))


def cmd_detail(args):
    d = detail(args.id)
    print(json.dumps(d, ensure_ascii=False, indent=2))


def cmd_new_since(args):
    if args.all_pages:
        resp = new_since_all(args.date, law_type=args.type, all_types=args.all_types)
    else:
        resp = new_since(args.date, law_type=args.type, all_types=args.all_types,
                         size=args.size)
    print(f"自 {args.date} 以来新发布法规，共 {resp.get('total', 0)} 条（返回 {len(resp.get('rows',[]))} 条）：")
    for r in resp.get("rows", []):
        print(_fmt_row(r))


def cmd_batch_check(args):
    with open(args.file, "r", encoding="utf-8") as f:
        names = [line.strip() for line in f if line.strip()]
    print(f"共 {len(names)} 条法规待检查", file=sys.stderr)
    results = batch_check(names, delay=args.delay)
    # Output as JSON
    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for r in results:
            if r["found"]:
                print(f"  ✓ {r['name']} → {r['status']} (公布:{r['gbrq']})")
            else:
                print(f"  ✗ {r['name']} → {r['status']}")
        # Summary
        found = [r for r in results if r["found"]]
        valid = [r for r in found if r.get("sxx") == 3]
        modified = [r for r in found if r.get("sxx") == 2]
        abolished = [r for r in found if r.get("sxx") == 1]
        not_found = [r for r in results if not r["found"]]
        print(f"\n汇总: 有效 {len(valid)}, 已修改 {len(modified)}, "
              f"已废止 {len(abolished)}, 未找到 {len(not_found)}")


def cmd_batch_check_excel(args):
    try:
        import openpyxl
    except ImportError:
        print("需要安装 openpyxl: pip install openpyxl", file=sys.stderr)
        sys.exit(1)
    wb = openpyxl.load_workbook(args.file, read_only=True, data_only=True)
    ws = wb.active
    col_letter = args.col.upper()
    col_idx = ord(col_letter) - ord('A') + 1
    names = []
    for row in ws.iter_rows(min_row=args.start_row, values_only=False):
        cell = row[col_idx - 1]
        if cell.value and isinstance(cell.value, str):
            # Clean: remove book title marks and extra whitespace
            name = cell.value.strip().strip("《》")
            if name and len(name) > 2:
                names.append(name)
    wb.close()
    names = list(dict.fromkeys(names))  # deduplicate preserving order
    print(f"从 Excel 提取 {len(names)} 条法规名称", file=sys.stderr)
    results = batch_check(names, delay=args.delay)
    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for r in results:
            mark = "✓" if r["found"] else "✗"
            print(f"  {mark} {r['name']} → {r.get('status', '?')}")


def cmd_download(args):
    import os
    os.makedirs(args.output, exist_ok=True)
    filepath = download_file(args.id, output_dir=args.output, fmt=args.format)
    print(f"已下载: {filepath}")


def cmd_batch_download(args):
    """从 batch-check 的 JSON 结果文件中批量下载法规原文"""
    import os
    os.makedirs(args.output, exist_ok=True)
    with open(args.file, "r", encoding="utf-8") as f:
        results = json.load(f)
    to_download = [r for r in results if r.get("found") and r.get("bbbs")]
    print(f"共 {len(to_download)} 条法规待下载", file=sys.stderr)
    success = 0
    fail = 0
    for i, r in enumerate(to_download):
        try:
            filepath = download_file(r["bbbs"], output_dir=args.output, fmt=args.format)
            print(f"  ✓ {r['name']} → {os.path.basename(filepath)}")
            success += 1
        except Exception as e:
            print(f"  ✗ {r['name']} → 下载失败: {e}")
            fail += 1
        if i < len(to_download) - 1:
            time.sleep(args.delay)
        if (i + 1) % 10 == 0:
            print(f"  进度: {i+1}/{len(to_download)}", file=sys.stderr)
    print(f"\n完成: 成功 {success}, 失败 {fail}")


def main():
    parser = argparse.ArgumentParser(description="国家法律法规数据库查询工具")
    sub = parser.add_subparsers(dest="command", required=True)

    # search
    p = sub.add_parser("search", help="搜索法规")
    p.add_argument("keyword", help="搜索关键词")
    p.add_argument("--exact", action="store_true", help="精确匹配")
    p.add_argument("--content", action="store_true", help="按正文内容搜索（默认按标题）")
    p.add_argument("--all-pages", action="store_true", help="获取所有页结果")
    p.add_argument("--type", help="法规类型: 法律/行政法规/地方性法规/司法解释")
    p.add_argument("--size", type=int, default=20, help="每页条数")
    p.set_defaults(func=cmd_search)

    # detail
    p = sub.add_parser("detail", help="查看法规详情")
    p.add_argument("id", help="法规 bbbs ID")
    p.set_defaults(func=cmd_detail)

    # new-since
    p = sub.add_parser("new-since", help="查找指定日期后新发布的法规")
    p.add_argument("date", help="起始日期 (yyyy-MM-dd)")
    p.add_argument("--type", help="法规类型")
    p.add_argument("--all-types", action="store_true", help="查所有类型")
    p.add_argument("--all-pages", action="store_true", help="获取所有页结果")
    p.add_argument("--size", type=int, default=100, help="每页条数")
    p.set_defaults(func=cmd_new_since)

    # batch-check
    p = sub.add_parser("batch-check", help="批量检查法规状态")
    p.add_argument("file", help="法规名称列表文件（每行一个）")
    p.add_argument("--delay", type=float, default=0.5, help="请求间隔秒数")
    p.add_argument("--json", action="store_true", help="输出 JSON 格式")
    p.set_defaults(func=cmd_batch_check)

    # batch-check-excel
    p = sub.add_parser("batch-check-excel", help="从 Excel 批量检查法规状态")
    p.add_argument("file", help="Excel 文件路径")
    p.add_argument("--col", default="D", help="法规名称所在列 (默认 D)")
    p.add_argument("--start-row", type=int, default=4, help="数据起始行 (默认 4)")
    p.add_argument("--delay", type=float, default=0.5, help="请求间隔秒数")
    p.add_argument("--json", action="store_true", help="输出 JSON 格式")
    p.set_defaults(func=cmd_batch_check_excel)

    # download
    p = sub.add_parser("download", help="下载法规文件")
    p.add_argument("id", help="法规 bbbs ID")
    p.add_argument("--format", default="docx", choices=["docx", "pdf"], help="文件格式")
    p.add_argument("--output", default=".", help="输出目录")
    p.set_defaults(func=cmd_download)

    # batch-download
    p = sub.add_parser("batch-download", help="从 batch-check JSON 结果批量下载法规原文")
    p.add_argument("file", help="batch-check --json 输出的 JSON 文件")
    p.add_argument("--format", default="docx", choices=["docx", "pdf"], help="文件格式")
    p.add_argument("--output", default=".", help="输出目录")
    p.add_argument("--delay", type=float, default=1.0, help="请求间隔秒数")
    p.set_defaults(func=cmd_batch_download)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
