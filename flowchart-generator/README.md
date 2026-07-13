# Flowchart Generator Skill

> **制作人：nigo** | **微信公众号：逆行的狗**

AI驱动的内控流程图绘制工具，专为审计和财务领域设计。

## 功能特点

- 把自然语言描述的流程生成标准泳道流程图
- 支持多种输出格式：.drawio、.png、.vsdx、.svg
- 内置"内控搭建"主题，对齐主流内控手册绘制规范
- 支持样式微调、主题管理（可新增自定义主题）

## 快速开始

```python
import sys
sys.path.insert(0, "/path/to/flowchart-generator")
from engine import StyleManager, FlowchartBuilder

# 定义流程节点
nodes_data = [
    {
        "id": "1",
        "step": "提交申请",
        "type": "process",
        "swimlane_role": "申请部门",
        "node_role": "申请人",
        "output_docs": "申请单",
        "next_steps": [{"id": "2", "condition": ""}]
    },
    # ... 更多节点
]

# 生成流程图
sm = StyleManager(theme_name="governance")
builder = FlowchartBuilder(style_manager=sm, sheet_name="示例流程", nodes_data=nodes_data)
xml = builder.generate_xml_content()
```

## 文件说明

- **SKILL.md** - 完整使用文档（推荐从这里开始）
- **example_usage.py** - 示例代码
- **references/architecture.md** - 技术架构文档（进阶阅读）
- **engine/** - 核心引擎代码

## 运行示例

```bash
cd /path/to/flowchart-generator
python example_usage.py
```

## 注意事项

- 本 skill 是为 OpenCode/Claude Code 设计，可配合自然语言流程描述使用
- 导出 PNG/VSDX 需要安装 draw.io Desktop
- `.drawio` 文件可在 https://app.diagrams.net/ 在线打开

## 许可

本 skill 可自由使用、分享、修改。
