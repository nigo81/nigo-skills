---
name: flowchart-generator
description: AI驱动的内控流程图绘制工具，专为审计和财务领域设计。把自然语言描述的流程生成标准泳道流程图（.drawio/PNG/VSDX/SVG）。触发词：流程图、内控流程、画流程图、泳道图、drawio、生成流程图、创建流程图、采购流程图、报销流程图、审批流程图、"画一个XXX流程图"。适用于审计工作流、财务审批流程、内控流程文档。
---

# Flowchart Generator Skill

> **制作人：nigo** | **微信公众号：逆行的狗**

把自然语言描述的内控/审计流程，生成为标准**泳道流程图**。输出 `.drawio`（可编辑）、`.png`（高清图）、`.vsdx`（Visio）、`.svg`。

内置"内控搭建"主题，对齐主流内控手册绘制规范：浅蓝标题栏 + 白底黑框内容节点 + 红色开始/结束。支持样式微调、主题管理（可新增自定义主题）。

## 触发时机

用户出现下列信号时激活：
- 直接指令：「画流程图」「生成XX流程」「画一个采购付款泳道图」
- 审计/内控语境：「把这段内控流程可视化」「这个控制活动画成图」
- 提到 drawio、泳道、审批流程、报销流程等

## 工作原理（三步）

```
用户描述 → [Step 1] AI 解析为节点数据 → [Step 2] 引擎生成 .drawio → [Step 3] 渲染导出
```

引擎路径：`~/.claude/skills/flowchart-generator`。调用前把该目录加入 `sys.path`：
```python
import sys
sys.path.insert(0, "/Users/<user>/.claude/skills/flowchart-generator")
from engine import StyleManager, FlowchartBuilder, Renderer
```

## Step 1 — 把流程描述解析为节点数据

引擎不直接接受自然语言。**你的核心职责**是把用户描述的内控流程解析成下面的结构化节点列表，再交给引擎。

### 节点数据格式

```python
nodes_data = [
    {
        "id": "1",                       # 必填，唯一字符串标识
        "step": "提交付款申请\n（附单据）", # 必填，节点文字（\n 换行）
        "type": "process",               # 必填，"process"=矩形 | "decision"=菱形
        "swimlane_role": "采购部",        # 必填，决定归属哪个泳道（部门名）
        "node_role": "采购经办",          # 必填，节点内角色标签
        "output_docs": "请款单",          # 可选，输出文档名（无则空字符串）
        "next_steps": [                  # 必填，下游连接列表
            {"id": "2", "condition": ""} # condition 空=普通流转
        ]
    },
    {
        "id": "2",
        "step": "金额>2000？",
        "type": "decision",
        "swimlane_role": "财务部",
        "node_role": "财务经理",
        "output_docs": "",
        "next_steps": [
            {"id": "3", "condition": "是"},
            {"id": "4", "condition": "否"}
        ]
    },
    {
        "id": "5",
        "step": "登记账簿",
        "type": "process",
        "swimlane_role": "财务部",
        "node_role": "会计",
        "output_docs": "",
        "next_steps": []                 # 空列表=流程终点（连到自动生成的"结束"节点）
    }
]
```

### 解析规则

- **泳道识别**：从描述中提取部门名（财务部、采购部、总经理办公室、出纳等）作为 `swimlane_role`。泳道按节点出现顺序自动生成。
- **节点类型**：含"是否/判断/大于/匹配"等条件的设为 `decision`；动作步骤设为 `process`。
- **开始/结束节点**：**不要手动添加**，引擎会自动在最前生成"开始"、在所有 `next_steps=[]` 的节点后生成"结束"。
- **分支条件**：判断节点的每个 `next_steps` 必须带 `condition`（"是"/"否"等），普通流转留空字符串。
- **输出文档**：节点产出单据/凭证/报告时填 `output_docs`（如"入库单""付款凭证"），会渲染为节点下方的小文档形状。

## Step 2 — 生成 .drawio

```python
from engine import StyleManager, FlowchartBuilder
from pathlib import Path

# 加载主题（默认 governance 内控搭建）
sm = StyleManager(theme_name="governance")

# 可选：样式微调（见下文"样式调整"）
sm.override({
    "edges": {"dashed_for_negative": True}
})

# 构建流程图
builder = FlowchartBuilder(
    style_manager=sm,
    sheet_name="采购付款循环-付款流程",  # 标题栏文字
    nodes_data=nodes_data
)

xml = builder.generate_xml_content()
Path("付款流程.drawio").write_text(xml, encoding="utf-8")

# builder.lanes 可查看实际生成的泳道顺序
print(builder.lanes)  # 例如 ['采购部', '财务部', '总经理办公室', '出纳']
```

## Step 3 — 导出为图片/Visio

导出需要本机安装 [draw.io Desktop](https://github.com/jgraph/drawio-desktop/releases)（macOS 放到 `/Applications/`）。仅生成 `.drawio` 不需要它。

```python
from engine import Renderer
from pathlib import Path

r = Renderer()
r.export_to_png(Path("付款流程.drawio"), Path("付款流程.png"), scale=2.0)  # 2倍高清
r.export_to_vsd(Path("付款流程.drawio"), Path("付款流程.vsdx"))            # Visio
r.export_to_svg(Path("付款流程.drawio"), Path("付款流程.svg"))             # 矢量
```

`.drawio` 也可在 https://app.diagrams.net/ 在线打开编辑。

## 样式调整

用 `StyleManager.override()` 临时调整样式（不修改主题文件）。**配置键映射表**：

| 用户表达 | override 键 | 取值示例 |
|---------|------------|---------|
| 泳道宽度改成300 | `layout.swimlane_width` | `300` |
| 标题栏颜色 | `colors.layout_bar.fill` | `parse_color("浅蓝")` |
| 泳道头颜色 | `colors.swimlane_header.fill` | `"#b4bff5"` |
| 开始/结束节点颜色 | `colors.start_end.fill` | `"#d24a4a"` |
| 角色节点背景 | `colors.node_role.fill` | `"#FFFFFF"` |
| 步骤节点背景 | `colors.node_step.fill` | `"#dae8fc"` |
| 文档节点背景 | `colors.node_doc.fill` | `"#f5f5f5"` |
| 判断节点背景 | `colors.decision.fill` | `"#fff2cc"` |
| 连线颜色 | `colors.edge.stroke` | `"#000000"` |
| 开始节点形状 | `shapes.start_end` | `"rounded"` |
| 否分支用虚线 | `edges.dashed_for_negative` | `True` |
| 自定义虚线关键词 | `edges.negative_keywords` | `["否","驳回","退回"]` |

**override 示例**：
```python
sm.override({
    "colors": {"decision": {"fill": "#FFA500"}},   # 判断节点橙色
    "edges":  {"dashed_for_negative": True}        # 否分支虚线
})
```

override 后需重新调用 `builder.generate_xml_content()` 生成。`override` 是叠加式（可多次调用），`reset_overrides()` 清空回到主题原样。

## 主题管理

内置 `governance`（内控搭建）一个主题。用户可随时新增自定义主题。

| 操作 | 调用 |
|------|------|
| 列出全部主题（内置+自定义） | `sm.list_themes()` |
| 查看某主题详情 | `sm.get_theme_info("governance")` |
| 加载主题 | `sm.load_theme("governance")` 或初始化时 `StyleManager(theme_name="...")` |
| 把当前样式保存为新主题 | `sm.save_as_theme(name="我的方案", description="...")` |
| 删除自定义主题 | `sm.delete_theme("我的方案")`（内置主题受保护，不可删） |
| 清空 override 回到主题原样 | `sm.reset_overrides()` |

`save_as_theme` 会自动创建 `themes/custom/` 目录并写入 `<name>.json`，之后 `StyleManager(theme_name="<name>")` 即可加载。主题完整 JSON 字段见 `themes/README.md`。

## 颜色系统

颜色值通过 `engine.colors.parse_color()` 解析，支持多种写法：

| 写法 | 示例 |
|------|------|
| 十六进制 | `#FF0000` |
| 中文 | `红色`、`浅蓝色`、`橙色` |
| 英文 | `red`、`blue` |

```python
from engine.colors import parse_color
parse_color("橙色")   # → "#FFA500"
parse_color("#b4bff5") # → "#B4BFF5"
```

override 时颜色值可填以上任意写法，引擎内部统一经 `parse_color()` 归一。

## 端到端示例

```python
import sys
sys.path.insert(0, "/Users/<user>/.claude/skills/flowchart-generator")
from engine import StyleManager, FlowchartBuilder, Renderer
from pathlib import Path

nodes_data = [
    {"id": "1", "step": "提交报销", "type": "process",
     "swimlane_role": "各部门", "node_role": "员工",
     "output_docs": "报销单", "next_steps": [{"id": "2", "condition": ""}]},
    {"id": "2", "step": "金额>5000？", "type": "decision",
     "swimlane_role": "各部门", "node_role": "部门经理",
     "next_steps": [{"id": "3", "condition": "是"},
                    {"id": "4", "condition": "否"}]},
    {"id": "3", "step": "总经理审批", "type": "process",
     "swimlane_role": "总经理办公室", "node_role": "总经理",
     "output_docs": "", "next_steps": [{"id": "5", "condition": ""}]},
    {"id": "4", "step": "财务审核", "type": "process",
     "swimlane_role": "财务部", "node_role": "财务",
     "output_docs": "", "next_steps": [{"id": "5", "condition": ""}]},
    {"id": "5", "step": "出纳付款", "type": "process",
     "swimlane_role": "财务部", "node_role": "出纳",
     "output_docs": "付款凭证", "next_steps": []},
]

sm = StyleManager(theme_name="governance")
builder = FlowchartBuilder(style_manager=sm, sheet_name="费用报销流程", nodes_data=nodes_data)
Path("费用报销.drawio").write_text(builder.generate_xml_content(), encoding="utf-8")

Renderer().export_to_png(Path("费用报销.drawio"), Path("费用报销.png"), scale=2.0)
```

更多示例见 `example_usage.py`。

## 故障排除

### AI 解析的流程不符合预期
- 用更明确的描述：写清部门名、步骤名。
- 用箭头符号（→）标注流转。
- 判断节点用"是否/判断/大于"等关键词。

### 样式调整不生效
- 确认 override 键拼写（见上方映射表）。
- override 后必须重新 `generate_xml_content()`。
- 手动编辑主题：`themes/governance.json` 或 `themes/custom/<name>.json`。

### PNG 导出失败
- 安装 draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases
- 放到 `/Applications/`（macOS）。
- 或用 https://app.diagrams.net/ 在线打开 `.drawio` 手动导出。

### Visio 打开
- `.vsdx` 可被 Visio 2013+ 直接打开。
- 较旧版本需转换。`.drawio` 也可在 Visio 中打开。

## 技术架构

引擎内部模块结构、布局/碰撞检测算法、容错降级、样式系统机制等进阶内容，面向二次开发者，见 **[references/architecture.md](references/architecture.md)**。普通使用无需阅读。
