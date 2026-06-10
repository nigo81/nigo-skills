# nigo-skills

> 审计师专属 AI 技能包 —— 让 AI 拥有资深审计专家的思维方式和专业判断力。

本仓库收集了一套面向审计、财务、会计从业者的 AI Skills（技能）。安装后，AI 助手将具备资深专家的分析视角，在处理会计准则判断、审计实务问题、财报分析时给出更专业的回答。

## 包含技能

| 技能 | 一句话说明 |
|------|-----------|
| **[chen-yiwei-perspective](./chen-yiwei-perspective)** | 陈奕蔚（会计视野论坛版主）的会计准则实务判断框架，20+ 年答疑精华 |
| **[tianchuan-perspective](./tianchuan-perspective)** | 田川（信永中和高级经理）的审计方法论与专业判断逻辑 |
| **[a-stock-financial](./a-stock-financial)** | A股上市公司三大报表一键查询与 Excel 导出（同花顺数据源，免费） |
| **[chenyiwei-bbs](./chenyiwei-bbs)** | 陈版主（会计视野论坛）实务问答实时检索，自动查询最新准则解答 |
| **[local-rag](./local-rag)** | 本地向量知识库，支持多项目隔离、语义检索、中英文文档解析（SiliconFlow 免费额度） |

## 安装方式

### 方式一：告诉 AI 直接安装（推荐）

如果你使用 **Claude Code**，在对话中直接说：

```
安装这个 skill：https://github.com/nigo81/nigo-skills/tree/main/chen-yiwei-perspective
```

AI 会自动从 GitHub 下载并放到正确的位置。每个 skill 文件夹都可以单独安装。

同样适用于支持 GitHub raw 文件读取的 AI 工具，如 **Cursor**、**Windsurf**、**Trae**、**WorkBuddy** 等——把上面的链接发给 AI，让它帮你下载 SKILL.md 和相关文件即可。

### 方式二：手动复制

```bash
# 克隆仓库
git clone https://github.com/nigo81/nigo-skills.git

# 安装到 Claude Code
cp -r nigo-skills/chen-yiwei-perspective ~/.claude/skills/

# 安装到 Cursor（放入项目 .cursor/rules/ 目录）
cp nigo-skills/chen-yiwei-perspective/SKILL.md your-project/.cursor/rules/chen-yiwei-perspective.md
```

### 各工具对应的安装位置

| 工具 | 安装方式 |
|------|---------|
| **Claude Code** | 复制到 `~/.claude/skills/<skill-name>/`，SKILL.md 作为入口 |
| **Cursor** | 放入项目的 `.cursor/rules/` 目录 |
| **Windsurf** | 放入项目的 `.windsurfrules` 或通过 Windsurf Rules 配置 |
| **Trae** | 放入项目的 `.trae/rules/` 目录 |
| **WorkBuddy** | 在对话中直接发送 SKILL.md 的内容作为上下文 |

> 不确定你的工具怎么装？直接把 SKILL.md 的内容粘贴到对话里，效果一样——只是每次新建对话需要重新粘贴。

## 使用示例

安装后，在 AI 对话中直接提问即可触发对应技能：

**会计准则判断：**
> "客户有一笔政府补助，与资产相关，应该怎么进行会计处理？"
> → AI 将以陈奕蔚的准则解读风格回答，引用具体准则条款和实务案例

**审计实务分析：**
> "存货监盘过程中发现盘点日与资产负债表日不一致，应该怎么做？"
> → AI 将以田川的审计方法论分析，给出实操建议

**A股财报查询：**
> "帮我查一下贵州茅台近3年的利润表，导出 Excel"
> → 自动调用 AkShare 查询并生成 Excel 文件

**陈版主问答检索：**
> "政府补助与收益相关和与资产相关如何区分？陈版主怎么说的？"
> → 自动查询会计视野论坛陈奕蔚的最新实务解答

**制度条款检索：**
> "把目标公司的制度文件建库，搜索消防安全管理相关条款"
> → 自动解析 Word/PDF/Markdown，切片入库，语义检索 + rerank 精排

## 谁适合用

- 审计助理 ~ 项目经理 ~ 合伙人：快速获取准则依据和实务判断
- 财务人员：查询上市公司财报，做同行对比分析
- 会计专业学生：以专家视角理解准则条文背后的实务逻辑
- 审计项目组：建立客户制度文档知识库，快速检索匹配条款

## 许可

MIT
