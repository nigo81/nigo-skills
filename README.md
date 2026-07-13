# nigo-skills

> 审计师专属 AI 技能包 —— 让 AI 拥有资深审计专家的思维方式和专业判断力。

本仓库收集了一套面向审计、财务、会计从业者的 AI Skills（技能）。安装后，AI 助手将具备资深专家的分析视角与一整套实战工具，覆盖会计准则判断、审计实务、财报分析、关联方核查、报告勾稽、法规检索等高频场景。

## 包含技能

### 思维视角类（专家视角）

| 技能 | 一句话说明 |
|------|-----------|
| **[chen-yiwei-perspective](./chen-yiwei-perspective)** | 陈奕蔚（会计视野论坛版主、中审众环合伙人）的会计准则实务判断框架，24 年答疑精华 |
| **[chenyiwei-bbs](./chenyiwei-bbs)** | 陈版主实务问答实时检索（curl 公开 API，无需 API Key/MCP） |
| **[tianchuan-perspective](./tianchuan-perspective)** | 田川（信永中和高级经理）的审计方法论与"业务优先"专业判断逻辑 |
| **[shaoniannu-perspective](./shaoniannu-perspective)** | 少年怒（审计小哥）的实务派思维，241 篇文章 + 千条评论提炼 |

### 数据查询类（一键取数）

| 技能 | 一句话说明 |
|------|-----------|
| **[a-stock-financial](./a-stock-financial)** | A股上市公司三大报表一键查询与 Excel 导出（同花顺数据源，免费） |
| **[china-law-search](./china-law-search)** | 统一查询国家法律法规库（flk.npc.gov.cn）+ 国家规章库（gov.cn） |
| **[cicpa-company-query](./cicpa-company-query)** | 中注协行业知识库查工商信息，支持单个查询与全量导出（52 维度 53 个 Excel） |

### 审计工具类（实战核查）

| 技能 | 一句话说明 |
|------|-----------|
| **[audit-report-checker](./audit-report-checker)** | 检查审计报告勾稽/加总/格式错误，输出 7 sheet Excel + Markdown 复核报告（算术全走代码） |
| **[related-party-identification](./related-party-identification)** | 基于工商数据自动识别"隐形关联方"，八大维度比对，对齐证监会处罚案例规则 |
| **[flowchart-generator](./flowchart-generator)** | AI 驱动的内控流程图绘制工具，自然语言描述自动生成标准泳道流程图（.drawio/PNG/VSDX/SVG） |
| **[local-rag](./local-rag)** | 本地向量知识库，按项目管理制度文档，语义检索条款（硅基流动免费 API，零安装） |

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
> → AI 以陈奕蔚的准则解读风格回答，引用具体准则条款和实务案例

**审计实务分析：**
> "存货监盘过程中发现盘点日与资产负债表日不一致，应该怎么做？"
> → AI 以田川的审计方法论分析，给出实操建议

**A股财报查询：**
> "帮我查一下贵州茅台近 3 年的利润表，导出 Excel"
> → 自动调用数据源查询并生成 Excel 文件

**审计报告核查：**
> "帮我检查这份审计报告的勾稽关系"
> → 自动解析报表+附注，校验加总和表注勾稽，输出复核报告

**关联方识别：**
> "核查一下被审计单位的客户供应商里有没有隐形关联方"
> → 拉取工商全维度数据，八大维度比对，揪出未披露关联方

## 谁适合用

- 审计助理 ~ 项目经理 ~ 合伙人：快速获取准则依据、实务判断与核查工具
- 财务人员：查询上市公司财报，做同行对比分析
- 会计专业学生：以专家视角理解准则条文背后的实务逻辑

## 许可

MIT
