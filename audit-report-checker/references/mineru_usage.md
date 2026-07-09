---
文件: mineru_usage.md
层级: 参考文档
用途: mineru 云端 API 调用说明（安装、认证、用法、隐私提示）
---

# mineru 云端 API 调用说明

本文件说明如何使用 `mineru-open-api`（npm 包）调用 mineru.net 云端 API，把扫描型 PDF 和图片报表精准提取成 markdown。这是审计报告四路分流提取策略中的一环（详见 `extraction.md`），仅在无文本层时使用。

**Claude 何时读这个文件**：
- 首次遇到扫描型 PDF 或图片报表，需要引导用户安装配置 mineru 时
- `parse_report.py` 调用 mineru 失败，需要排查问题时
- 用户问"mineru 是什么"/"为什么要装这个"/"数据安全吗"时
- 需要向用户解释 flash-extract vs extract 的选择时

---

## 一、mineru-open-api 是什么

`mineru-open-api` 是一个 npm 包，封装了 mineru.net **云端 API** 的调用（不是本地部署，不需要 GPU/模型）。

- **云端 API**：文档上传到 mineru.net 服务器，云端用视觉模型 OCR 识别后返回 markdown
- **跨平台**：mac/Windows/Linux 均可，只需 Node.js 环境
- **免费额度**：注册即送免费 token，足够日常审计核查使用
- **高精度模型**：`--model vlm` 用视觉语言模型，精度优于传统 OCR

> **与本地 mineru 的区别**：本机可能装了 `mineru-document-extractor` skill（本地封装），但底层也是调云端 API。`mineru-open-api` 是更底层的命令行工具，`parse_report.py` 直接 subprocess 调用它，不依赖其他 skill。

---

## 二、安装

需要 Node.js 环境（mac 用 brew install node，Windows 用官方安装包）。

```bash
# mac / Linux
npm install -g mineru-open-api

# Windows (PowerShell 或 CMD)
npm install -g mineru-open-api
```

安装后验证：

```bash
mineru-open-api --help
```

---

## 三、认证（配置免费 token）

1. **注册 token**：浏览器打开 https://mineru.net/apiManage/token ，注册账号并获取免费 API token
2. **配置 token**：在终端运行

```bash
mineru-open-api auth
```

按提示粘贴 token，配置信息保存在 `~/.mineru-open-api/config.json`（mac/Linux）或 `%USERPROFILE%\.mineru-open-api\config.json`（Windows）。

3. **验证认证**：

```bash
mineru-open-api auth --check
```

---

## 四、两种调用模式

### 4.1 flash-extract（免 token，快速场景）

- **免 token**：无需注册认证即可使用
- **限制**：单文件 ≤ 10MB，≤ 20 页
- **精度**：标准 OCR 精度（低于 vlm 模型）
- **适用**：快速预览、小文件、非关键场景

```bash
mineru-open-api flash-extract "<文件路径>" -o <输出目录> -f md
```

### 4.2 extract（需 token，高精度）

- **需 token**：首次需 `mineru-open-api auth` 配置
- **支持大文件**：无 10MB/20 页限制
- **两种模型**：
  - `--model vlm`：视觉语言模型，**高精度**（推荐，审计核查首选）
  - `--model pipeline`：流水线模型，**无幻觉**（适合纯文本场景，不编造内容）
- **语言**：`--language ch` 指定中文优化

```bash
# 高精度模式（审计核查推荐）
mineru-open-api extract "<文件路径>" -o <输出目录> -f md --model vlm --language ch

# 无幻觉模式（纯文本场景）
mineru-open-api extract "<文件路径>" -o <输出目录> -f md --model pipeline --language ch
```

**参数说明**：
- `<文件路径>`：PDF/图片的完整路径（含空格需用引号）
- `-o <输出目录>`：输出目录（自动创建）
- `-f md`：输出 markdown 格式（也支持 json/html）
- `--model vlm`：视觉语言模型
- `--language ch`：中文优化

### 4.3 模型选择建议

| 场景 | 推荐命令 | 理由 |
|---|---|---|
| 扫描型审计报告（含复杂表格） | `extract --model vlm` | vlm 对表格结构识别最准 |
| 图片报表（Word 插入的报表截图） | `extract --model vlm` | 同上 |
| 纯文本扫描件（无表格） | `extract --model pipeline` | 无幻觉，避免编造内容 |
| 快速预览小文件 | `flash-extract` | 免 token，速度快 |
| 大文件（>10MB 或 >20 页） | `extract`（不能用 flash-extract） | flash-extract 有大小限制 |

---

## 五、数据隐私提示（重要）

**审计报告含敏感财务数据，上传云端 API 前必须告知用户并获得知情同意。**

### 5.1 隐私风险

- 调用 mineru 云端 API 时，文档会上传到 mineru.net 服务器处理
- 审计报告可能含客户未公开的财务数据、商业秘密
- 上市公司年报属公开信息，无隐私问题；但审计底稿、内部报告、未公开财报需谨慎

### 5.2 隐私提示机制

`parse_report.py` 在调用 mineru 前，**自动在终端打印隐私提示**，例如：

```
⚠️  隐私提示：本文件需使用 mineru 云端 API 提取（扫描型 PDF / 图片报表）。
    文档将上传至 mineru.net 服务器处理，可能包含敏感财务数据。
    - 上市公司公开报告：无隐私风险，可继续
    - 未公开/保密报告：建议优先使用文本路径（pdfplumber/python-docx），或征得客户同意后再用
    继续调用 mineru 请按回车，取消请按 Ctrl+C...
```

Claude 在引导用户首次配置 mineru 时，也应主动提示数据隐私考量。

### 5.3 隐私保护建议

- **优先文本路径**：只要 PDF 有文本层或 Word 有文本表格，就用 pdfplumber/python-docx，数据不出本机
- **公开报告才用 mineru**：上市公司年报、公开披露文件用 mineru 无隐私问题
- **保密报告慎用**：客户内部报告、审计底稿，确认客户合规要求后再决定是否用 mineru
- **处理完即删**：mineru 云端处理完的文件，如非必要不留存在云端（mineru 默认不长期存储，但建议确认其隐私政策）

---

## 六、在 parse_report.py 中的调用方式

`parse_report.py` 通过 `subprocess` 调用 `mineru-open-api` 命令行工具。示例调用逻辑（伪代码）：

```python
import subprocess
import pathlib

def extract_with_mineru(file_path: pathlib.Path, output_dir: pathlib.Path) -> str:
    """调用 mineru 云端 API 提取扫描型 PDF / 图片报表，返回 markdown 内容。"""
    
    # 1. 打印隐私提示，等待用户确认
    print("⚠️  隐私提示：本文件需使用 mineru 云端 API 提取...")
    input("继续调用 mineru 请按回车，取消请按 Ctrl+C...")
    
    # 2. 构造命令
    cmd = [
        "mineru-open-api", "extract",
        str(file_path),
        "-o", str(output_dir),
        "-f", "md",
        "--model", "vlm",
        "--language", "ch",
    ]
    
    # 3. 调用（跨平台，subprocess 自动处理路径）
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    
    # 4. 读取输出的 markdown 文件
    md_file = output_dir / f"{file_path.stem}.md"
    return md_file.read_text(encoding="utf-8")
```

### 6.1 调用前的检查

`parse_report.py` 调用 mineru 前应检查：
- `mineru-open-api` 是否已安装（`which mineru-open-api` / `where mineru-open-api`）
- 是否已认证（`mineru-open-api auth --check`），未认证则提示用户运行 `mineru-open-api auth`
- 文件大小是否超限（vlm 模式无 10MB 限制，但过大文件可能超时）

### 6.2 调用失败的处理

- **网络错误**：提示用户检查网络，mineru 云端 API 需要外网访问
- **token 过期/额度用完**：提示用户去 https://mineru.net/apiManage/token 查看额度或更换 token
- **文件格式不支持**：检查文件是否为 PDF/图片，损坏的文件需用户重新提供
- **超时**：大文件可能超时，建议拆分或重试

### 6.3 跨平台兼容

- 命令用列表形式传参（`subprocess.run(cmd, ...)`），避免 shell 注入和路径空格问题
- 文件路径用 `pathlib.Path`，自动处理 mac/Linux 正斜杠和 Windows 反斜杠
- 不依赖平台特定 shell 命令（不用 `open`/`start`/`pbcopy` 等）

---

## 七、常见问题排查

| 问题 | 原因 | 解决 |
|---|---|---|
| `command not found: mineru-open-api` | 未安装或 PATH 未配置 | 重新 `npm install -g mineru-open-api`；检查 Node.js/bin 目录是否在 PATH |
| `Auth failed` / `401 Unauthorized` | token 未配置或失效 | 运行 `mineru-open-api auth` 重新配置 |
| `File too large` | flash-extract 超 10MB | 改用 `extract`（需 token） |
| `Page limit exceeded` | flash-extract 超 20 页 | 改用 `extract` |
| 提取结果数字错乱 | OCR 精度问题 | 确认用 `--model vlm`；关键勾稽数字人工复核原件 |
| 表格结构错乱 | 复杂表格 OCR 困难 | 人工核对；或改用文本路径（若有文本层） |
| 调用超时 | 网络慢或文件大 | 重试；或拆分大 PDF 后分批处理 |

---

## 八、参考链接

- mineru-open-api npm 包：https://www.npmjs.com/package/mineru-open-api
- mineru.net 官网：https://mineru.net
- token 管理：https://mineru.net/apiManage/token
- 提取策略原理：参见本 skill 的 `references/extraction.md`
