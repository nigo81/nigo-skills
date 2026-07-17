"""
流程图生成示例 - 测试skill核心功能

制作人：nigo
微信公众号：逆行的狗
"""
import sys
from pathlib import Path

# 添加 skill 根目录到 Python 路径
skill_root = Path(__file__).parent.parent
sys.path.insert(0, str(skill_root))

from engine import StyleManager, FlowchartBuilder, Renderer

# 示例节点数据（模拟AI解析结果）
nodes_data_example = [
    {
        "id": "1",
        "step": "初审",
        "type": "process",
        "swimlane_role": "财务部",
        "node_role": "财务专员",
        "output_docs": "",
        "next_steps": [{"id": "2", "condition": ""}]
    },
    {
        "id": "2",
        "step": "金额>10万？",
        "type": "decision",
        "swimlane_role": "财务部",
        "node_role": "财务经理",
        "next_steps": [
            {"id": "3", "condition": "是"},
            {"id": "4", "condition": "否"}
        ]
    },
    {
        "id": "3",
        "step": "总经理审批",
        "type": "process",
        "swimlane_role": "总经理办公室",
        "node_role": "总经理",
        "output_docs": "",
        "next_steps": [{"id": "5", "condition": ""}]
    },
    {
        "id": "4",
        "step": "财务部审批",
        "type": "process",
        "swimlane_role": "财务部",
        "node_role": "财务总监",
        "output_docs": "",
        "next_steps": [{"id": "5", "condition": ""}]
    },
    {
        "id": "5",
        "step": "付款",
        "type": "process",
        "swimlane_role": "财务部",
        "node_role": "出纳",
        "output_docs": "付款凭证",
        "next_steps": []
    }
]

print("="*60)
print("流程图生成示例 - Skill核心功能测试")
print("="*60)

# 步骤1：加载样式配置
print("\n[步骤1] 加载样式配置...")
style_manager = StyleManager(theme_name="governance")
print(f"✓ 主题: {style_manager.current_config.get('meta', {}).get('name', 'Unknown')}")

# 步骤2：应用样式调整
print("\n[步骤2] 应用样式调整...")
style_manager.override({
    'colors': {
        'decision': {
            'fill': '#FFA500'  # 橙色
        }
    },
    'edges': {
        'dashed_for_negative': True
    }
})
print("✓ 判断节点颜色: 橙色")
print("✓ 否分支连线: 虚线")

# 步骤3：生成流程图
print("\n[步骤3] 生成流程图...")
builder = FlowchartBuilder(
    style_manager=style_manager,
    sheet_name="采购付款流程",
    nodes_data=nodes_data_example
)

xml_content = builder.generate_xml_content()

# 输出到文件
output_path = "/tmp/example_flowchart.drawio"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(xml_content)

print(f"✓ 流程图已生成: {output_path}")
print(f"  泳道数: {len(builder.lanes)}")
print(f"  节点数: {len(nodes_data_example)}")

# 步骤4：保存自定义主题
print("\n[步骤4] 保存自定义方案...")
save_path = style_manager.save_as_theme(
    name="审计专属方案",
    description="橙色判断节点，虚线表示否分支"
)
if save_path:
    print(f"✓ 主题已保存: {save_path}")

# 步骤5：导出为图片
print("\n[步骤5] 导出为图片...")
renderer = Renderer()

# PNG 导出（高清）
renderer.export_to_png(Path(output_path), Path("example_flowchart.png"), scale=2.0)
print("✓ PNG 已导出")

# SVG 导出（矢量）
renderer.export_to_svg(Path(output_path), Path("example_flowchart.svg"))
print("✓ SVG 已导出")

# VSDX 导出（仅当版本兼容时）
if renderer._is_version_compatible_with_vsdx():
    try:
        renderer.export_to_vsdx(
            Path(output_path),
            Path("example_flowchart.vsdx"),
            auto_install=False  # 不自动安装，避免交互式询问
        )
        print("✓ VSDX 已导出")
    except Exception as e:
        print(f"⚠ VSDX 导出失败: {e}")
else:
    print("⚠ VSDX 导出跳过（当前版本不支持，设置 auto_install=True 可自动安装）")