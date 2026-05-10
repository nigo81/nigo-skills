#!/usr/bin/env python3
"""
准则本地缓存下载工具
用法：
  python3 fetch_standards.py              # 下载常用准则
  python3 fetch_standards.py cas 14       # 下载 CAS 14
  python3 fetch_standards.py casi 16      # 下载 准则解释第16号
  python3 fetch_standards.py --all        # 下载全部准则（CAS 1-42 + 解释1-19）
  python3 fetch_standards.py --update     # 重新下载所有已缓存的准则
"""

import sys, os, json, time
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import markdownify

SCRIPT_DIR = Path(__file__).parent
STANDARDS_DIR = SCRIPT_DIR.parent / "references" / "standards"
MANIFEST_PATH = STANDARDS_DIR / "manifest.json"
BASE_URL = "https://docs.maoyanqing.com"

# 类别 → URL前缀
CATEGORIES = {
    "cas":   "accounting/ent/cas",
    "casi":  "accounting/ent/casi",
    "casg":  "accounting/ent/casg",
    "casq":  "accounting/ent/casq",
    "casc":  "accounting/ent/casc",
    "rlc":   "securities/rlc",
    "casca": "securities/casca",
    "rwas":  "securities/rwas",
    "asr":   "securities/asr",
    "csa":   "auditing/csa",
    "csag":  "auditing/csag",
}

# 常用准则（陈奕蔚答疑高频引用）
# CAS 1-9 的URL用两位数(01-09)，10以上用原始数字
COMMON = {
    "cas": [f"{n:02d}" for n in range(1,10)] + ["10","11","12","13","14","16","18","20","21","22","23","24","25","28","33","36","37","38","39","40","41","42"],
    "casi": [f"{n:02d}" for n in range(1,10)] + ["10","11","12","13","14","15","16","17","18","19"],
    "rlc": ["01","02","03","04","05","06","07","08"],
}

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"})


def fetch_and_save(category: str, num: str) -> dict | None:
    url_prefix = CATEGORIES.get(category)
    if not url_prefix:
        print(f"  ❌ 未知类别: {category}")
        return None

    url = f"{BASE_URL}/{url_prefix}/{num}.html"
    target_dir = STANDARDS_DIR / category
    target_file = target_dir / f"{num}.md"
    target_dir.mkdir(parents=True, exist_ok=True)

    try:
        resp = session.get(url, timeout=15)
        resp.raise_for_status()
        resp.encoding = "utf-8"
    except Exception as e:
        print(f"  ❌ {category}/{num} 下载失败: {e}")
        return None

    # 提取正文
    soup = BeautifulSoup(resp.text, "html.parser")

    # VuePress theme-hope: 主内容在 .vp-doc 或 article 或 main
    content_el = (
        soup.select_one(".vp-doc")
        or soup.select_one("article")
        or soup.select_one("main")
        or soup.body
    )

    if not content_el:
        print(f"  ❌ {category}/{num} 无法提取正文")
        return None

    # 转 markdown
    md_content = markdownify.markdownify(
        str(content_el),
        heading_style="ATX",
        bullets="-",
        strip=["script", "style", "nav", "footer", "header"],
    )

    # 清理多余空行
    lines = [line.rstrip() for line in md_content.split("\n")]
    cleaned = []
    prev_empty = False
    for line in lines:
        if line == "":
            if not prev_empty:
                cleaned.append("")
            prev_empty = True
        else:
            cleaned.append(line)
            prev_empty = False
    md_content = "\n".join(cleaned).strip()

    # 获取标题
    title_el = soup.select_one("h1")
    title = title_el.get_text(strip=True) if title_el else f"{category}/{num}"

    # 写入文件
    today = date.today().isoformat()
    frontmatter = f"""---
source: {url}
downloaded: {today}
title: {title}
---

"""
    target_file.write_text(frontmatter + md_content + "\n", encoding="utf-8")

    size = target_file.stat().st_size
    print(f"  ✅ {category}/{num} — {title} ({size:,} bytes)")
    return {"title": title, "downloaded": today}


def load_manifest() -> dict:
    if MANIFEST_PATH.exists():
        return json.loads(MANIFEST_PATH.read_text())
    return {"lastUpdated": "", "source": BASE_URL, "standards": {}}


def save_manifest(manifest: dict):
    manifest["lastUpdated"] = date.today().isoformat()
    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def download_batch(items: list[tuple[str, str]], label: str = ""):
    manifest = load_manifest()
    ok, fail = 0, 0

    for i, (cat, num) in enumerate(items, 1):
        print(f"[{i}/{len(items)}] {cat}/{num}")
        result = fetch_and_save(cat, num)
        if result:
            ok += 1
            cat_data = manifest["standards"].setdefault(cat, {})
            cat_data[num] = result
        else:
            fail += 1
        # 礼貌延迟
        if i < len(items):
            time.sleep(0.3)

    save_manifest(manifest)
    print(f"\n{'='*50}")
    print(f"完成: {ok} 成功, {fail} 失败")
    if label:
        print(f"类型: {label}")


def update_cached():
    manifest = load_manifest()
    items = []
    for cat, entries in manifest.get("standards", {}).items():
        for num in entries:
            items.append((cat, num))
    if not items:
        print("没有已缓存的准则需要更新")
        return
    print(f"更新 {len(items)} 个已缓存的准则...\n")
    download_batch(items, label="更新已缓存")


def main():
    if not STANDARDS_DIR.exists():
        STANDARDS_DIR.mkdir(parents=True)

    args = sys.argv[1:]

    if not args:
        # 默认：下载常用准则
        items = [(cat, num) for cat, nums in COMMON.items() for num in nums]
        print(f"下载常用准则（{len(items)} 个）...\n")
        download_batch(items, label="常用准则")
        return

    if args[0] == "--all":
        # 全部 CAS 1-42 + 解释 1-19
        items = [(cat, str(n)) for cat, nums in COMMON.items() for n in range(int(nums[0] if nums[0].isdigit() else 1), int(nums[-1]) + 1)]
        # 只取 cas 和 casi 的全部范围
        items = [("cas", str(n)) for n in range(1, 43)]
        items += [("casi", str(n)) for n in range(1, 20)]
        print(f"下载全部准则（{len(items)} 个）...\n")
        download_batch(items, label="全部准则")
        return

    if args[0] == "--update":
        update_cached()
        return

    # 单个下载: python3 fetch_standards.py cas 14
    if len(args) >= 2:
        cat, num = args[0], args[1]
        download_batch([(cat, num)], label="单个下载")
        return

    print(__doc__)


if __name__ == "__main__":
    main()
