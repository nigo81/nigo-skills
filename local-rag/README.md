# Local RAG — 本地向量知识库

轻量级本地向量知识库，支持按项目管理文档，语义检索。面向审计/财务从业者，默认零模型安装。

## 特性

- **多项目隔离** — 一个审计项目建一个知识库，互不干扰
- **零模型默认安装** — 默认用硅基流动 API（免费额度），不需要 GPU
- **中文制度文档切片** — 三级递归（章→段落→条款），不会切到"第X章"中间
- **多格式支持** — Word (.docx/.doc)、PDF、Markdown、纯文本
- **扫描件 PDF 集成** — 自动检测扫描件，调用 MinerU 处理
- **MCP Server** — 可作为 Cursor/VS Code 的 MCP 工具使用

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置

```bash
# 交互式配置向导（推荐）
~/.claude/skills/local-rag/bin/local-rag setup

# 或手动设置环境变量：
export SILICONFLOW_API_KEY="sk-xxxxxxxx"
```

### 3. 使用

```bash
# 设置快捷方式（一次性）
export PATH="$HOME/.claude/skills/local-rag/bin:$PATH"

# 创建项目
local-rag create my-project

# 入库文档（文件或文件夹）
local-rag ingest my-project /path/to/docs

# 语义检索
local-rag search my-project "消防安全管理"

# 检索 + Rerank（更精准）
local-rag search my-project "消防安全管理" --rerank

# 列出所有项目
local-rag list

# 查看配置
local-rag info
```

## Embedding Provider 选择

| Provider | 适用场景 | 需要什么 |
|----------|---------|---------|
| **硅基流动**（默认） | 大多数场景，免费 | API Key（注册送免费额度） |
| **Ollama** | 离线/内网环境 | 安装 Ollama + 拉取模型 |
| **OpenAI** | 已有 OpenAI 账号 | API Key |

## 文件格式支持

| 格式 | 解析方式 |
|------|---------|
| `.md` / `.txt` | 直接读文本 |
| `.docx` | python-docx |
| `.doc` | macOS textutil |
| `.pdf`（文本型） | PyMuPDF |
| `.pdf`（扫描件） | MinerU flash-extract |

## 配置

配置文件：`~/.local/share/local-rag/config.yaml`

```yaml
embedding:
  provider: siliconflow
  model: BAAI/bge-m3
  api_key: ${SILICONFLOW_API_KEY}

reranker:
  provider: siliconflow
  model: BAAI/bge-reranker-v2-m3
  api_key: ${SILICONFLOW_API_KEY}

chunking:
  strategy: chinese_regulation    # chinese_regulation | generic
  chunk_size: 800
  chunk_overlap: 0.15
  min_chunk_size: 50
```

### 环境变量

| 变量 | 说明 |
|------|------|
| `SILICONFLOW_API_KEY` | 硅基流动 API Key |
| `OPENAI_API_KEY` | OpenAI API Key |
| `RAG_DATA_DIR` | 数据存储路径（默认 `~/.local/share/local-rag/`） |
| `RAG_EMBEDDING_MODEL` | 覆盖默认 embedding 模型 |

## MCP Server

```bash
# 启动 MCP Server
pip install fastmcp
python3 mcp_server.py
```

在 Cursor/VS Code 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "local-rag": {
      "command": "python3",
      "args": ["/path/to/local-rag/mcp_server.py"]
    }
  }
}
```

## 项目结构

```
local-rag/
├── src/                  # 模块化源码（v3）
│   ├── config.py         # 配置管理
│   ├── embedding.py      # Embedding 抽象层（SiliconFlow/Ollama/OpenAI）
│   ├── chunker.py        # 切片策略
│   ├── store.py          # ChromaDB 存储
│   ├── parser.py         # 文档解析 + MinerU
│   ├── reranker.py       # Reranker（SiliconFlow/Ollama/Local/Noop）
│   ├── pipeline.py       # 编排层
│   ├── cli.py            # CLI 入口
│   └── exceptions.py     # 异常体系
├── scripts/              # 旧版代码（保留向后兼容）
│   ├── rag_v2.py         # v2（Chonkie + Ollama）
│   └── rag_server.py     # v1（自写切片）
└── mcp_server.py         # MCP Server
```

## License

MIT
