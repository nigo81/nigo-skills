#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股权穿透：对被审计单位做 N 层股权穿透，发现所有子公司/参控股企业，输出核查名单。

为什么需要它：关联方核查的准确性依赖于"把被审计单位的子公司也纳入核查范围"。
造假常发生在子公司层面（如卓朗科技案，造假客户供应商挂在子公司"卓朗发展"名下，
只查母公司看不出关联）。本脚本先把子公司查出来，供后续下载工商数据 + 核查。

用法:
    python3 discover_scope.py --target "被审计单位全称" --depth 5 -o 核查名单.txt
    # 然后把客户/供应商追加进名单，调 cicpa --browser-export 下载全部公司

依赖: cicpa-company-query skill 的 search_company + ensure_session（复用其 cookie）
"""
import argparse
import os
import sys

CICPA_SCRIPTS = os.path.expanduser("~/.claude/skills/cicpa-company-query/scripts")
sys.path.insert(0, CICPA_SCRIPTS)

from cicpa_query import search_company, ensure_session  # noqa: E402

EQUITY_URL = "https://zsk-cmis.cicpa.org.cn/open/enterprise_info_api/v3/atlas/enterprise_equity"


def get_investments(session, headers, org_id):
    """调 equity API 拿一家公司的对外投资列表。"""
    if not org_id:
        return []
    try:
        resp = session.get(f"{EQUITY_URL}?orgid={org_id}", headers=headers, timeout=30)
        result = resp.json()
    except Exception as e:
        print(f"  ⚠️ equity API 失败 org_id={org_id}: {e}", file=sys.stderr)
        return []
    if result.get("status_code") != 0:
        return []
    data = result.get("data", {}) or {}
    # 兼容多种返回结构（invests 可能是 dict / None / list）
    invests_raw = data.get("invests")
    if isinstance(invests_raw, dict):
        invests = invests_raw.get("children", []) or invests_raw.get("list", [])
    elif isinstance(invests_raw, list):
        invests = invests_raw
    else:
        invests = data.get("investList", []) or []
    return invests or []


def penetrate(target, threshold=50.0, max_depth=5):
    """递归股权穿透，返回子公司列表 [{name, ratio, depth, parent, org_id}]。"""
    # 1. search 拿 org_id（用验证过的 search_company，绕过 discover_subsidiaries 的字段名 bug）
    results = search_company(target)
    if not results:
        print(f"❌ 未找到企业: {target}", file=sys.stderr)
        return []
    # 精确匹配优先
    matched = [r for r in results if r.get("name") == target] or results
    org_id = matched[0].get("org_id") or matched[0].get("id")
    if not org_id:
        print(f"❌ 未获取 org_id: {target}", file=sys.stderr)
        return []
    print(f"✓ {target} (org_id={org_id})")

    session = ensure_session()
    if session is None:
        print("❌ 无法获取 session", file=sys.stderr)
        return []
    xsrf = session.cookies.get("XSRF-TOKEN", "")
    headers = {"X-Xsrf-Token": xsrf} if xsrf else {}

    all_subs = []
    visited = {target}

    def _recurse(parent_name, parent_oid, depth):
        if depth > max_depth:
            return
        invests = get_investments(session, headers, parent_oid)
        for inv in invests:
            name = (inv.get("entName") or inv.get("name") or "").strip()
            ratio_str = inv.get("investRatio") or inv.get("ratio") or inv.get("czbl") or "0"
            sub_oid = inv.get("orgId") or inv.get("orgid") or inv.get("entId") or ""
            if not name or name in visited:
                continue
            try:
                ratio = float(str(ratio_str).replace("%", "").strip())
            except (ValueError, AttributeError):
                ratio = 0
            if ratio >= threshold:
                visited.add(name)
                all_subs.append({
                    "name": name, "ratio": ratio, "depth": depth,
                    "parent": parent_name, "org_id": sub_oid,
                })
                indent = "  " * depth
                print(f"{indent}└─ [第{depth}层] {name} (持股 {ratio}%, parent={parent_name})")
                if sub_oid:
                    _recurse(name, sub_oid, depth + 1)

    _recurse(target, org_id, 1)
    return all_subs


def main():
    ap = argparse.ArgumentParser(description="股权穿透：发现被审计单位的子公司/参控股企业")
    ap.add_argument("--target", required=True, help="被审计单位全称")
    ap.add_argument("--depth", type=int, default=5, help="穿透深度（默认5层）")
    ap.add_argument("--threshold", type=float, default=50.0, help="持股比例阈值（默认50%%）")
    ap.add_argument("-o", "--output", default="核查名单.txt", help="输出核查名单文件")
    args = ap.parse_args()

    print(f"🔍 对 [{args.target}] 做 {args.depth} 层股权穿透（持股 >= {args.threshold}%）...\n")
    subs = penetrate(args.target, args.threshold, args.depth)

    # 汇总：target + 所有子公司，去重保序
    names = [args.target] + [s["name"] for s in subs]
    seen, unique = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            unique.append(n)

    print(f"\n📊 穿透完成：发现 {len(subs)} 家子公司/参控股企业，核查名单共 {len(unique)} 家")
    # 按深度统计
    by_depth = {}
    for s in subs:
        by_depth.setdefault(s["depth"], []).append(s)
    for d in sorted(by_depth):
        print(f"   第{d}层: {len(by_depth[d])} 家")

    with open(args.output, "w", encoding="utf-8") as f:
        for n in unique:
            f.write(n + "\n")
    print(f"\n✅ 核查名单已输出: {args.output}")
    print(f"   下一步：把客户/供应商追加到此名单，然后调 cicpa 一次性下载全部公司工商数据：")
    print(f'   python3 <cicpa>/cicpa_query.py -n $(cat {args.output} | tr "\\n" " ") --browser-export')


if __name__ == "__main__":
    main()
