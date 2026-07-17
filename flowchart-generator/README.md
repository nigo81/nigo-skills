# Flowchart Generator Skill

> **制作人：nigo** | **微信公众号：逆行的狗**

AI驱动的内控流程图绘制工具，专为审计和财务领域设计。

## 功能特点

- 把自然语言描述的流程生成标准泳道流程图
- 支持多种输出格式：.drawio、.png、.svg、.vsdx（Visio）
- 内置"内控搭建"主题，对齐主流内控手册绘制规范
- 支持样式微调、主题管理（可新增自定义主题）
- **跨平台支持**：Windows、macOS、Linux（自动识别）
- **智能版本管理**：Visio 导出自动检测并安装兼容版本

## 快速开始

```python
import sys
sys.path.insert(0, "/path/to/flowchart-generator")
from engine import StyleManager, FlowchartBuilder, Renderer
from pathlib import Path

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

# 导出
renderer = Renderer()
renderer.export_to_png(Path("example.drawio"), Path("example.png"), scale=2.0)
renderer.export_to_svg(Path("example.drawio"), Path("example.svg"))
renderer.export_to_vsdx(Path("example.drawio"), Path("example.vsdx"), auto_install=True)
```

## Visio 导出说明

**重要**：VSDX 导出需要 draw.io v26.0.16 或更早版本。

**智能自动安装**：
```python
# 自动检测版本，如不兼容则自动安装 v26.0.16
renderer.export_to_vsdx(
    Path("example.drawio"),
    Path("example.vsdx"),
    auto_install=True  # 推荐
)
```

**支持平台**：
- ✅ macOS (Intel + Apple Silicon)
- ✅ Windows (x64)
- ✅ Linux (x64 + ARM64)

## 文件说明

- **SKILL.md** - 完整使用文档
- **scripts/example_usage.py** - 示例代码
- **references/architecture.md** - 技术架构文档（进阶阅读）
- **engine/** - 核心引擎代码

## 运行示例

```bash
cd /path/to/flowchart-generator
python scripts/example_usage.py
```

## 注意事项

- 本 skill 是为 OpenCode/Claude Code 设计
- 导出 PNG/SVG 不需要特殊版本的 draw.io
- VSDX 导出会自动下载并安装兼容版本
- 自动安装会禁用 draw.io 的自动更新
- 如遇问题，请查看 SKILL.md 中的故障排除部分

## 许可

本 skill 可自由使用、分享、修改。