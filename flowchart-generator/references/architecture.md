# 技术架构（进阶）

> **制作人：nigo** | **微信公众号：逆行的狗**

本文件面向想二次开发、调试布局算法或扩展节点类型的开发者。普通使用无需阅读——SKILL.md 已涵盖全部使用场景。

## 目录

- [引擎模块结构](#引擎模块结构)
- [样式配置目录](#样式配置目录)
- [布局算法](#布局算法)
- [连线位置计算](#连线位置计算)
- [容错降级](#容错降级)
- [样式系统机制](#样式系统机制)
- [错误处理](#错误处理)
- [无需外部依赖](#无需外部依赖)

## 引擎模块结构

```
engine/
├── __init__.py             # 公开导出：StyleManager / FlowchartBuilder / Renderer
├── flowchart_builder.py    # 流程图构建器（节点类型注册表模式）
│   ├── _parse_lanes()          # 解析泳道
│   ├── _calculate_layout()     # 核心布局算法（见下文）
│   └── generate_xml_content()  # 生成 .drawio XML
├── style_manager.py        # 样式管理器
│   ├── load_theme()            # 加载主题（可选保留覆盖）
│   ├── override()              # 应用用户覆盖
│   ├── reset_overrides()       # 重置覆盖
│   ├── save_as_theme()         # 保存为自定义主题（防误覆盖）
│   ├── list_themes()           # 列出内置+自定义
│   ├── get_theme_info()        # 主题详情
│   └── delete_theme()          # 删除自定义主题
├── renderer.py             # 渲染器
│   ├── export_to_png()         # PNG（支持 scale 高清）
│   ├── export_to_vsdx()        # Visio .vsdx
│   └── export_to_svg()         # SVG
├── validator.py            # 结构校验器
│   └── FlowchartValidator      # 悬空连线/重复ID/空节点/孤立节点/循环引用/孤儿下一步
└── colors.py               # 颜色解析
    └── parse_color()           # 中文/英文/十六进制 → 颜色值
```

## 样式配置目录

```
themes/
├── governance.json         # 内控搭建（默认，内置）
├── README.md               # 主题格式说明
└── custom/                 # 用户自定义主题（save_as_theme 自动创建）
    └── *.json
```

加载优先级：`custom/` 同名 > 内置。主题完整 JSON 格式见 `themes/README.md`。

## 布局算法

`_calculate_layout` 是经长期调优的核心算法，处理以下复杂场景：

1. **Spot Collision 检测**：防止同层同一泳道多个节点重叠。
2. **Path Collision 检测**：防止跨泳道连线穿过其他节点。
3. **行高度计算**：根据节点类型（判断/普通）和文档输出动态计算。
4. **基线 offset 计算**：确保同行连线的水平通道畅通。
5. **Grid 对齐**：所有坐标 snap 到 10 的倍数，保证视觉整齐。

## 连线位置计算

- **同行横向流转**：
  - 普通节点：根据 baseline 和节点高度计算 ratio。
  - 判断节点：强制使用 `ratio=0.5`（中心点出入）。
- **异行纵向流转**：统一使用顶部进入 / 底部出口。
- **虚线处理**：识别 `edges.negative_keywords`（默认含"否""不符合""驳回"等）自动应用虚线。

## 容错降级

系统在各环节都有降级，确保始终有输出：

- **数据校验降级**：validator 发现错误 → 警告但不阻止生成；给出质量评分（0-100），低分提示检查。
- **导出降级**：draw.io CLI 不可用 → 只输出 .drawio 不报错；PNG 失败 → 降级到 .drawio。
- **颜色降级**：颜色解析失败 → 默认黑色；角色名不识别 → 回退普通解析。
- **布局降级**：复杂场景碰撞检测耗时 → 降低精度但保证出图；Grid 对齐失败 → 用原始坐标。

## 样式系统机制

- **优先级**：用户覆盖（override） > 主题 > 默认配置（`_get_default_config`）。
- **深拷贝**：合并配置时深拷贝，避免污染原配置。
- **持久化**：自定义主题经 `save_as_theme` 写入 `custom/` 目录自动保存。

## 错误处理

- 字段缺失 → 使用默认值。
- 空数据 → 友好提示。
- 主题不存在 → 警告 + 回退默认。

## 无需外部依赖

本 skill 直接使用当前 AI 会话的推理能力完成流程解析与样式调整，**不需要**：
- 配置任何 LLM API Key
- 调用外部模型服务
- 联网

唯一外部依赖是导出 PNG/VSDX 时需要 [draw.io Desktop](https://github.com/jgraph/drawio-desktop/releases)（仅导出环节需要，生成 .drawio 不需要）。
