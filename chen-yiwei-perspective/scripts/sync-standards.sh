#!/bin/bash
# 准则本地缓存下载脚本
# 用法：
#   ./sync-standards.sh cas 14          # 下载 CAS 14 收入
#   ./sync-standards.sh casi 16         # 下载 准则解释第16号
#   ./sync-standards.sh rlc 03          # 下载 监管规则适用指引第3号
#   ./sync-standards.sh all             # 下载常用准则（批量）
#
# 下载的文件保存到 references/standards/{类别}/{编号}.md

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
STANDARDS_DIR="$(dirname "$SCRIPT_DIR")/references/standards"
BASE_URL="https://docs.maoyanqing.com"

# URL 映射
get_url() {
  local category="$1"
  local num="$2"
  case "$category" in
    cas)   echo "${BASE_URL}/accounting/ent/cas/${num}.html" ;;
    casi)  echo "${BASE_URL}/accounting/ent/casi/${num}.html" ;;
    casg)  echo "${BASE_URL}/accounting/ent/casg/${num}.html" ;;
    rlc)   echo "${BASE_URL}/securities/rlc/${num}.html" ;;
    casca) echo "${BASE_URL}/securities/casca/${num}.html" ;;
    rwas)  echo "${BASE_URL}/securities/rwas/${num}.html" ;;
    asr)   echo "${BASE_URL}/securities/asr/${num}.html" ;;
    csa)   echo "${BASE_URL}/auditing/csa/${num}.html" ;;
    csag)  echo "${BASE_URL}/auditing/csag/${num}.html" ;;
    *)     echo "UNKNOWN"; exit 1 ;;
  esac
}

download_one() {
  local category="$1"
  local num="$2"
  local url
  url=$(get_url "$category" "$num")

  if [ "$url" = "UNKNOWN" ]; then
    echo "❌ 未知类别: $category"
    return 1
  fi

  local target_dir="${STANDARDS_DIR}/${category}"
  local target_file="${target_dir}/${num}.md"
  mkdir -p "$target_dir"

  local today
  today=$(date +%Y-%m-%d)

  echo "⬇️  下载 ${category}/${num} ← $url"

  # 用 curl 抓取 HTML，提取正文并转存
  # 注意：此脚本只保存原始 HTML。推荐在 Claude Code 中使用 webReader 工具下载，
  # 可获得干净的 Markdown 格式。此脚本作为离线备选方案。
  local html_content
  html_content=$(curl -sL "$url")
  if [ -z "$html_content" ]; then
    echo "❌ 下载失败: $url"
    return 1
  fi

  # 写入 frontmatter + 原始 HTML（后续可由 Claude 转为 Markdown）
  cat > "$target_file" << HEREDOC
---
source: $url
downloaded: $today
format: html
---

$html_content
HEREDOC

  echo "✅ 已保存: $target_file"
}

# 常用准则列表
download_common() {
  echo "📦 批量下载常用准则..."

  # 企业会计准则（高频引用）
  local cas_numbers="1 2 3 4 5 6 7 8 9 10 11 12 13 14 16 18 20 21 22 23 24 25 28 33 36 37 38 39 40 41 42"
  for num in $cas_numbers; do
    download_one cas "$num" 2>/dev/null || true
  done

  # 准则解释
  local casi_numbers="1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19"
  for num in $casi_numbers; do
    download_one casi "$num" 2>/dev/null || true
  done

  # 监管规则适用指引（会计类）
  local rlc_numbers="01 02 03 04 05 06 07 08"
  for num in $rlc_numbers; do
    download_one rlc "$num" 2>/dev/null || true
  done

  echo ""
  echo "⚠️  注意：此脚本保存的是 HTML 格式。建议在 Claude Code 中运行以下命令获取干净的 Markdown："
  echo "   /chen-yiwei-perspective 更新准则缓存"
}

# 主入口
case "${1:-}" in
  all)
    download_common
    ;;
  cas|casi|casg|casq|casc|rlc|casca|rwas|asr|csa|csag)
    if [ -z "${2:-}" ]; then
      echo "用法: $0 <类别> <编号>"
      echo "示例: $0 cas 14"
      exit 1
    fi
    download_one "$1" "$2"
    ;;
  *)
    echo "准则本地缓存下载工具"
    echo ""
    echo "用法:"
    echo "  $0 <类别> <编号>    下载单个准则"
    echo "  $0 all              批量下载常用准则"
    echo ""
    echo "类别: cas casi casg casq casc rlc casca rwas asr csa csag"
    echo ""
    echo "示例:"
    echo "  $0 cas 14          # CAS 14 收入"
    echo "  $0 casi 16         # 准则解释第16号"
    echo "  $0 rlc 03          # 监管规则适用指引第3号"
    echo ""
    echo "💡 提示：推荐在 Claude Code 中使用 webReader 工具下载，"
    echo "   可直接获得 Markdown 格式。此脚本作为离线备选。"
    ;;
esac
