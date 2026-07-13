# 主题配置

本目录存放流程图样式主题（JSON 格式）。

## 内置主题

- `governance.json` — 内控搭建（默认）。浅蓝标题栏/泳道头 + 白底黑框内容节点 + 红色开始结束，朴素专业风格，对齐主流内控手册绘制规范。

## 自定义主题

skill 支持用户随时新增主题，有两种方式：

1. **交互式**：用自然语言描述样式调整，满意后说"保存为我的方案"，引擎会写入 `custom/` 目录（首次保存自动创建）。
2. **手写 JSON**：按下方格式在 `custom/` 下新建 `你的方案名.json`，加载时 `StyleManager(theme_name="你的方案名")`。

加载优先级：`custom/` 同名主题 > 内置主题。

## 主题配置格式

完整字段如下（以 governance 为例）：

```json
{
  "meta": {
    "name": "内控搭建",
    "description": "主题描述"
  },
  "layout": {
    "swimlane_width": 280,
    "swimlane_header_height": 40,
    "frame_top_height": 40,
    "frame_left_width": 40,
    "node_width": 120,
    "role_height": 30,
    "step_height": 40,
    "doc_height": 40,
    "decision_height": 80,
    "y_start_offset": 140,
    "y_gap": 100
  },
  "colors": {
    "layout_bar": { "fill": "#b4bff5", "stroke": "#000000" },
    "swimlane_header": { "fill": "#b4bff5", "stroke": "#000000" },
    "swimlane_body": { "fill": "#FFFFFF", "stroke": "#000000" },
    "start_end": { "fill": "#d24a4a", "stroke": "#000000" },
    "node_role": { "fill": "#FFFFFF", "stroke": "#000000" },
    "node_step": { "fill": "#FFFFFF", "stroke": "#000000" },
    "node_doc": { "fill": "#FFFFFF", "stroke": "#000000" },
    "decision": { "fill": "#FFFFFF", "stroke": "#000000" },
    "edge": { "stroke": "#000000" }
  },
  "fonts": {
    "layout_bar": { "size": 16, "family": "Arial", "color": "#000000", "style": "bold" },
    "swimlane_header": { "size": 14, "family": "Arial", "color": "#000000", "style": "normal" },
    "start_end": { "size": 14, "family": "Arial", "color": "#FFFFFF", "style": "bold" },
    "node_role": { "size": 12, "family": "Arial", "color": "#000000", "style": "normal" },
    "node_step": { "size": 12, "family": "Arial", "color": "#000000", "style": "normal" },
    "node_doc": { "size": 10, "family": "Arial", "color": "#000000", "style": "normal" },
    "decision": { "size": 12, "family": "Arial", "color": "#000000", "style": "normal" }
  },
  "shapes": {
    "start_end": "rounded"
  },
  "edges": {
    "dashed_for_negative": true,
    "negative_keywords": ["否", "不符合", "不", "未", "fail", "驳回"]
  }
}
```

### 字段说明

- `layout`：尺寸控制。泳道宽度、节点宽高、行间距等。
- `colors`：各元素的填充色 `fill` 与描边色 `stroke`。
- `fonts`：字体大小、字体族、颜色、是否粗体（`bold`/`normal`）。
- `shapes.start_end`：开始/结束节点形状，默认 `rounded`（圆角）。
- `edges.dashed_for_negative`：是否对"否/驳回"等分支自动使用虚线。
- `edges.negative_keywords`：触发虚线的关键词列表。
