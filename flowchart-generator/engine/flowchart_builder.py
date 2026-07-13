"""
Flowchart Builder - 流程图构建器
复用布局算法，样式从 StyleManager 读取
"""
import html
from typing import List, Dict, Callable
from .style_manager import StyleManager
from .validator import FlowchartValidator
from .colors import parse_color


# ==========================================
# 样式构建器函数（模块级函数）
# ==========================================

def build_layout_bar_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建布局栏样式"""
    fill = parse_color(builder.style_manager.get_color('layout_bar', 'fill', '#b4bff5'))
    stroke = parse_color(builder.style_manager.get_color('layout_bar', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('layout_bar')
    return (
        f"rounded=0;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};"
        f"fontColor={font['color']};align=center;verticalAlign=middle;"
        f"fontSize={font['size']};fontStyle=1;spacingLeft=10;"
    )


def build_swimlane_header_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建泳道标题样式"""
    fill = parse_color(builder.style_manager.get_color('swimlane_header', 'fill', '#b4bff5'))
    stroke = parse_color(builder.style_manager.get_color('swimlane_header', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('swimlane_header')
    return (
        f"verticalAlign=middle;align=center;overflow=width;"
        f"fillColor={fill};gradientColor=none;"
        f"shape=swimlane;startSize={builder.SWIMLANE_HEADER_HEIGHT};"
        f"strokeColor={stroke};"
        f"labelBackgroundColor=none;rounded=0;html=1;whiteSpace=wrap;"
        f"fontSize={font['size']};"
    )


def build_swimlane_body_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建泳道主体样式"""
    fill = parse_color(builder.style_manager.get_color('swimlane_body', 'fill', '#FFFFFF'))
    stroke = parse_color(builder.style_manager.get_color('swimlane_body', 'stroke', '#000000'))
    return (
        f"swimlane;startSize=0;html=1;childLayout=none;horizontal=1;"
        f"swimlaneFillColor={fill};fillColor=none;"
        f"collapsible=0;strokeColor={stroke};"
    )


def build_start_end_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建开始/结束节点样式"""
    fill = parse_color(builder.style_manager.get_color('start_end', 'fill', '#d24a4a'))
    stroke = parse_color(builder.style_manager.get_color('start_end', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('start_end')
    shape = overrides.get('shape', 'rounded')
    return (
        f"verticalAlign=middle;align=center;overflow=width;"
        f"fillColor={fill};gradientColor=none;shape={shape};"
        f"strokeColor={stroke};spacingTop=-1;spacingBottom=-1;"
        f"spacingLeft=-1;spacingRight=-1;"
        f"points=[[0.5,1,0],[0.5,0,0],[0,0.5,0],[1,0.5,0]];"
        f"labelBackgroundColor=none;rounded=0;html=1;whiteSpace=wrap;"
        f"fontColor={font['color']};fontStyle=1;fontSize={font['size']};"
    )


def build_node_role_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建节点角色样式"""
    fill = parse_color(builder.style_manager.get_color('node_role', 'fill', '#FFFFFF'))
    stroke = parse_color(builder.style_manager.get_color('node_role', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('node_role')
    return (
        f"rounded=0;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};"
        f"align=center;verticalAlign=middle;fontSize={font['size']};"
    )


def build_node_step_process_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建流程节点步骤样式"""
    fill = parse_color(builder.style_manager.get_color('node_step', 'fill', '#FFFFFF'))
    stroke = parse_color(builder.style_manager.get_color('node_step', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('node_step')
    return (
        f"fillColor={fill};gradientColor=none;strokeColor={stroke};"
        f"spacingTop=-1;spacingBottom=-1;spacingLeft=-1;spacingRight=-1;"
        f"points=[[0.5,1,0],[0,0.5,0],[1,0.5,0],[0.5,0,0],[0.5,0.5,0]];"
        f"labelBackgroundColor=none;rounded=0;html=1;whiteSpace=wrap;"
        f"fontSize={font['size']};"
    )


def build_node_step_decision_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建判断节点步骤样式"""
    fill = parse_color(builder.style_manager.get_color('decision', 'fill', '#ffffff'))
    stroke = parse_color(builder.style_manager.get_color('decision', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('decision')
    return (
        f"fillColor={fill};gradientColor=none;"
        f"shape=mxgraph.flowchart.decision;strokeColor={stroke};spacingTop=-1;spacingBottom=-1;"
        f"spacingLeft=-1;spacingRight=-1;"
        f"points=[[0,0.5,0],[1,0.5,0],[0.5,0,0],[0.5,1,0],[0.5,0.5,0]];"
        f"labelBackgroundColor=none;rounded=0;html=1;whiteSpace=wrap;"
        f"fontSize={font['size']};fontStyle=1;"
    )


def build_node_doc_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建文档节点样式"""
    fill = parse_color(builder.style_manager.get_color('node_doc', 'fill', '#FFFFFF'))
    stroke = parse_color(builder.style_manager.get_color('node_doc', 'stroke', '#000000'))
    font = builder.style_manager.get_font_config('node_doc')
    return (
        f"verticalAlign=middle;align=center;overflow=width;"
        f"fillColor={fill};gradientColor=none;shape=document;"
        f"strokeColor={stroke};spacingTop=-1;spacingBottom=-1;"
        f"spacingLeft=-1;spacingRight=-1;"
        f"points=[[0,0.5,0],[1,0.5,0],[0.5,0,0],[0.5,0.88,0]];"
        f"labelBackgroundColor=none;rounded=0;html=1;whiteSpace=wrap;"
        f"fontSize={font['size']};"
    )


def build_invisible_box_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建隐形框样式"""
    return "rounded=0;whiteSpace=wrap;html=1;fillColor=none;strokeColor=none;opacity=0;"


def build_edge_style(builder: 'FlowchartBuilder', **overrides) -> str:
    """构建连线样式"""
    stroke = parse_color(builder.style_manager.get_color('edge', 'stroke', '#000000'))
    dashed = overrides.get('dashed', False)
    style = "edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;"
    style += f"strokeColor={stroke};"
    if dashed:
        style += "dashed=1;"
    return style


# 样式构建器注册表
STYLE_BUILDERS: Dict[str, Callable] = {
    'layout_bar': build_layout_bar_style,
    'swimlane_header': build_swimlane_header_style,
    'swimlane_body': build_swimlane_body_style,
    'start_end': build_start_end_style,
    'node_role': build_node_role_style,
    'node_step_process': build_node_step_process_style,
    'node_step_decision': build_node_step_decision_style,
    'node_doc': build_node_doc_style,
    'invisible_box': build_invisible_box_style,
    'edge': build_edge_style,
}


# ==========================================
# 节点渲染器函数（模块级函数）
# ==========================================

def render_decision_node(builder: 'FlowchartBuilder', index: int, node: Dict, base_x: int, base_y: int):
    """渲染判断节点（菱形）"""
    ai_id = str(node.get('id', index))
    step_text = node.get('step', '步骤')

    total_node_height = builder.DECISION_HEIGHT

    # 隐形框 (用于定位)
    group_id = builder._get_new_id()
    xml_group = f"""
    <mxCell id="{group_id}" value="" style="{builder._build_style('invisible_box')}" vertex="1" parent="1">
        <mxGeometry x="{base_x}" y="{base_y}" width="{builder.NODE_WIDTH}" height="{total_node_height}" as="geometry"/>
    </mxCell>
    """
    builder.generated_cells_xml.append(xml_group)

    # 菱形本体
    step_id = builder._get_new_id()
    xml_step = f"""
    <mxCell id="{step_id}" value="{builder._clean_xml_text(step_text)}" style="{builder._build_style('node_step_decision')}" vertex="1" parent="1">
        <mxGeometry x="{base_x}" y="{base_y}" width="{builder.NODE_WIDTH}" height="{builder.DECISION_HEIGHT}" as="geometry"/>
    </mxCell>
    """
    builder.generated_cells_xml.append(xml_step)

    builder.node_id_map[ai_id] = {"entry": step_id, "exit": step_id, "step": group_id}


def render_process_node(builder: 'FlowchartBuilder', index: int, node: Dict, base_x: int, base_y: int):
    """渲染普通流程节点（三段式：角色 + 步骤 + 文档）"""
    ai_id = str(node.get('id', index))
    role_text = node.get('node_role', '')
    step_text = node.get('step', '步骤')
    doc_text = node.get('output_docs', '')

    has_doc = doc_text and str(doc_text).strip() and str(doc_text).lower() != "nan"
    total_node_height = builder.ROLE_HEIGHT + builder.STEP_HEIGHT + (builder.DOC_HEIGHT if has_doc else 0)

    # 隐形框
    group_id = builder._get_new_id()
    xml_group = f"""
    <mxCell id="{group_id}" value="" style="{builder._build_style('invisible_box')}" vertex="1" parent="1">
        <mxGeometry x="{base_x}" y="{base_y}" width="{builder.NODE_WIDTH}" height="{total_node_height}" as="geometry"/>
    </mxCell>
    """
    builder.generated_cells_xml.append(xml_group)

    # Role
    role_id = builder._get_new_id()
    xml_role = f"""
    <mxCell id="{role_id}" value="{builder._clean_xml_text(role_text)}" style="{builder._build_style('node_role')}" vertex="1" parent="1">
        <mxGeometry x="{base_x}" y="{base_y}" width="{builder.NODE_WIDTH}" height="{builder.ROLE_HEIGHT}" as="geometry"/>
    </mxCell>
    """
    builder.generated_cells_xml.append(xml_role)

    # Step
    step_id = builder._get_new_id()
    step_y = base_y + builder.ROLE_HEIGHT
    xml_step = f"""
    <mxCell id="{step_id}" value="{builder._clean_xml_text(step_text)}" style="{builder._build_style('node_step_process')}" vertex="1" parent="1">
        <mxGeometry x="{base_x}" y="{step_y}" width="{builder.NODE_WIDTH}" height="{builder.STEP_HEIGHT}" as="geometry"/>
    </mxCell>
    """
    builder.generated_cells_xml.append(xml_step)

    entry_id = role_id
    exit_id = step_id

    if has_doc:
        doc_id = builder._get_new_id()
        doc_y = step_y + builder.STEP_HEIGHT
        xml_doc = f"""
        <mxCell id="{doc_id}" value="{builder._clean_xml_text(doc_text)}" style="{builder._build_style('node_doc')}" vertex="1" parent="1">
            <mxGeometry x="{base_x}" y="{doc_y}" width="{builder.NODE_WIDTH}" height="{builder.DOC_HEIGHT}" as="geometry"/>
        </mxCell>
        """
        builder.generated_cells_xml.append(xml_doc)
        exit_id = doc_id

    builder.node_id_map[ai_id] = {"entry": entry_id, "exit": exit_id, "step": group_id}


# 节点渲染器注册表
NODE_RENDERERS: Dict[str, Callable] = {
    'decision': render_decision_node,
    'process': render_process_node,
}

class FlowchartBuilder:
    def __init__(self, style_manager: StyleManager, sheet_name: str, nodes_data: List[Dict]):
        self.style_manager = style_manager
        self.sheet_name = sheet_name
        self.nodes_data = nodes_data

        # 预处理数据：确保每个节点都有 next_step_ids
        for node in self.nodes_data:
            if 'next_steps' in node and isinstance(node['next_steps'], list):
                ids = [str(step.get('id')) for step in node['next_steps'] if step.get('id') is not None]
                node['next_step_ids'] = ids
            elif 'next_step_ids' not in node:
                node['next_step_ids'] = []

        # 从样式管理器获取布局参数
        layout = style_manager.get_layout_params()
        self.SWIMLANE_WIDTH = layout.get('swimlane_width', 280)
        self.SWIMLANE_HEADER_HEIGHT = layout.get('swimlane_header_height', 40)
        self.FRAME_TOP_HEIGHT = layout.get('frame_top_height', 40)
        self.FRAME_LEFT_WIDTH = layout.get('frame_left_width', 40)
        self.NODE_WIDTH = layout.get('node_width', 120)
        self.ROLE_HEIGHT = layout.get('role_height', 30)
        self.STEP_HEIGHT = layout.get('step_height', 40)
        self.DOC_HEIGHT = layout.get('doc_height', 40)
        self.DECISION_HEIGHT = layout.get('decision_height', 60)

        self.Y_START_OFFSET = layout.get('y_start_offset', self.FRAME_TOP_HEIGHT + 60)
        self.Y_GAP = layout.get('y_gap', 100)

        # 状态
        self.lanes = []
        self.lane_x_map = {}
        self.cell_id_counter = 2
        self.generated_cells_xml = []
        self.node_id_map = {}
        self.node_positions = {}
        self.node_row_indices = {}
        self.row_baseline_offsets = {}
        self.total_content_height = 0

    def _get_new_id(self) -> str:
        cid = str(self.cell_id_counter)
        self.cell_id_counter += 1
        return cid

    def _parse_lanes(self):
        """扫描所有节点，提取唯一的责任部门作为泳道"""
        seen = set()
        ordered_roles = []
        for node in self.nodes_data:
            r = node.get('swimlane_role', '未定义部门')
            r = r.strip()
            if r and r not in seen:
                ordered_roles.append(r)
                seen.add(r)
        self.lanes = ordered_roles

        current_x = self.FRAME_LEFT_WIDTH
        for lane in self.lanes:
            self.lane_x_map[lane] = current_x
            current_x += self.SWIMLANE_WIDTH

    def _snap_to_grid(self, value: int, grid: int = 10) -> int:
        """坐标对齐到网格（确保整洁）"""
        return int(round(value / grid) * grid)

    def _calculate_layout(self):
        """核心布局算法：网格化分层布局 + 路径碰撞检测"""
        predecessors = {}
        id_to_index = {str(n.get('id')): i for i, n in enumerate(self.nodes_data)}

        for idx, node in enumerate(self.nodes_data):
            next_ids = node.get('next_step_ids', [])
            # 只有当 next_steps 字段不存在或为 None 时，才自动连接下一个节点
            # 显式 next_steps=[] 表示节点终止，不应自动连接
            should_auto_connect = 'next_steps' not in node or node.get('next_steps') is None

            if not next_ids and should_auto_connect:
                if idx < len(self.nodes_data) - 1:
                    next_node_id = str(self.nodes_data[idx+1].get('id'))
                    if next_node_id not in predecessors:
                        predecessors[next_node_id] = []
                    predecessors[next_node_id].append(idx)
            else:
                for nid in next_ids:
                    nid_str = str(nid)
                    if nid_str not in predecessors:
                        predecessors[nid_str] = []
                    predecessors[nid_str].append(idx)

        lane_to_idx = {lane: i for i, lane in enumerate(self.lanes)}

        current_row_idx = 0
        current_row_lanes = set()
        self.node_row_indices = {}

        for index, node in enumerate(self.nodes_data):
            node_id = str(node.get('id'))
            lane = node.get('swimlane_role', '未定义部门').strip()
            lane_idx = lane_to_idx.get(lane, 0)

            collision = False

            # Check A: Spot Collision
            if lane in current_row_lanes:
                collision = True

            # Check B: Path Collision
            if not collision:
                my_preds_indices = predecessors.get(node_id, [])
                for pred_idx in my_preds_indices:
                    if self.node_row_indices.get(pred_idx) == current_row_idx:
                        pred_lane = self.nodes_data[pred_idx].get('swimlane_role', '未定义部门').strip()
                        pred_lane_idx = lane_to_idx.get(pred_lane, 0)
                        start_l = min(pred_lane_idx, lane_idx)
                        end_l = max(pred_lane_idx, lane_idx)
                        for check_l_idx in range(start_l + 1, end_l):
                            check_lane_name = self.lanes[check_l_idx]
                            if check_lane_name in current_row_lanes:
                                collision = True
                                break
                    if collision: break

            if collision:
                current_row_idx += 1
                current_row_lanes = set()

            self.node_row_indices[index] = current_row_idx
            current_row_lanes.add(lane)

        # Step 2: Calculate Row Heights
        row_heights = {}
        for index, node in enumerate(self.nodes_data):
            node_type = node.get('type', 'process')

            if node_type == 'decision':
                node_h = self.DECISION_HEIGHT
            else:
                doc_text = node.get('output_docs', '')
                node_h = self.ROLE_HEIGHT + self.STEP_HEIGHT
                if doc_text and str(doc_text).strip() and str(doc_text).lower() != "nan":
                    node_h += self.DOC_HEIGHT

            row_idx = self.node_row_indices[index]
            if row_idx not in row_heights:
                row_heights[row_idx] = 0
            row_heights[row_idx] = max(row_heights[row_idx], node_h)

        # Step 3: Determine Y Positions & Row Baselines (Horizontal Channels)
        current_y = self.Y_START_OFFSET + 40 + 80
        self.row_baseline_offsets = {}

        sorted_rows = sorted(row_heights.keys())
        for row_idx in sorted_rows:
            row_start_y = current_y

            # Find the "First Node" in this row to determine the horizontal baseline
            for index, node in enumerate(self.nodes_data):
                if self.node_row_indices[index] == row_idx:
                    # Found first node of the row (leftmost in process order)
                    if row_idx not in self.row_baseline_offsets:
                        node = self.nodes_data[index]
                        node_type = node.get('type', 'process')
                        if node_type == 'decision':
                            h = self.DECISION_HEIGHT
                        else:
                            doc_text = node.get('output_docs', '')
                            has_doc = doc_text and str(doc_text).strip() and str(doc_text).lower() != "nan"
                            h = self.ROLE_HEIGHT + self.STEP_HEIGHT + (self.DOC_HEIGHT if has_doc else 0)

                        self.row_baseline_offsets[row_idx] = h / 2.0

                    self.node_positions[index] = row_start_y

            current_y += (row_heights[row_idx] + self.Y_GAP)

        self.total_content_height = max(current_y + 150, 1169)
        self.node_positions['end'] = current_y

    def _build_style(self, style_type: str, **overrides) -> str:
        """构建样式字符串（使用注册表模式）"""
        builder = STYLE_BUILDERS.get(style_type)
        if builder is None:
            return ""
        return builder(self, **overrides)

    def generate_xml_content(self) -> str:
        # 前置校验
        validator = FlowchartValidator()
        result = validator.validate(self.nodes_data)

        if not result.is_valid:
            # 收集错误信息，打印警告但仍尝试生成（降级处理）
            for err in result.errors:
                print(f"⚠️ 数据校验警告: {err}")

        self._parse_lanes()
        self._calculate_layout()

        total_height = self.total_content_height
        total_width = self.FRAME_LEFT_WIDTH + (len(self.lanes) * self.SWIMLANE_WIDTH)

        # 0. 生成布局框架
        top_bar_id = self._get_new_id()
        xml_top_bar = f"""
        <mxCell id="{top_bar_id}" value="{self._clean_xml_text(self.sheet_name)}" style="{self._build_style('layout_bar')}" vertex="1" parent="1">
            <mxGeometry x="0" y="0" width="{total_width}" height="{self.FRAME_TOP_HEIGHT}" as="geometry"/>
        </mxCell>
        """
        self.generated_cells_xml.append(xml_top_bar)

        left_bar_id = self._get_new_id()
        left_bar_height = total_height - self.FRAME_TOP_HEIGHT
        xml_left_bar = f"""
        <mxCell id="{left_bar_id}" value="" style="{self._build_style('layout_bar')}" vertex="1" parent="1">
            <mxGeometry x="0" y="{self.FRAME_TOP_HEIGHT}" width="{self.FRAME_LEFT_WIDTH}" height="{left_bar_height}" as="geometry"/>
        </mxCell>
        """
        self.generated_cells_xml.append(xml_left_bar)

        # 1. 生成泳道
        for lane in self.lanes:
            lane_body_id = self._get_new_id()
            lane_header_id = self._get_new_id()
            x_pos = self.lane_x_map[lane]
            body_y = self.FRAME_TOP_HEIGHT
            body_height = total_height - self.FRAME_TOP_HEIGHT

            xml_body = f"""
            <mxCell id="{lane_body_id}" value="" style="{self._build_style('swimlane_body')}" vertex="1" parent="1">
                <mxGeometry x="{x_pos}" y="{body_y}" width="{self.SWIMLANE_WIDTH}" height="{body_height}" as="geometry"/>
            </mxCell>
            """
            self.generated_cells_xml.append(xml_body)

            xml_header = f"""
            <mxCell id="{lane_header_id}" value="{self._clean_xml_text(lane)}" style="{self._build_style('swimlane_header')}" vertex="1" parent="{lane_body_id}">
                <mxGeometry x="0" y="0" width="{self.SWIMLANE_WIDTH}" height="{self.SWIMLANE_HEADER_HEIGHT}" as="geometry"/>
            </mxCell>
            """
            self.generated_cells_xml.append(xml_header)

        # 2. 生成开始节点
        start_node_id = self._get_new_id()
        first_lane = self.nodes_data[0].get('swimlane_role', self.lanes[0] if self.lanes else '未定义部门').strip() if self.nodes_data else (self.lanes[0] if self.lanes else '未定义部门')
        start_x = self._snap_to_grid(self.lane_x_map.get(first_lane, 0) + (self.SWIMLANE_WIDTH - self.NODE_WIDTH) / 2)
        start_y = self._snap_to_grid(self.Y_START_OFFSET + 40)

        start_shape = self.style_manager.get('shapes.start_end', 'rounded')

        xml_start = f"""
        <mxCell id="{start_node_id}" value="开始" style="{self._build_style('start_end', shape=start_shape)}" vertex="1" parent="1">
            <mxGeometry x="{start_x}" y="{start_y}" width="{self.NODE_WIDTH}" height="40" as="geometry"/>
        </mxCell>
        """
        self.generated_cells_xml.append(xml_start)

        # 3. 生成流程节点
        for index, node in enumerate(self.nodes_data):
            lane = node.get('swimlane_role', '未定义部门').strip()
            node_type = node.get('type', 'process')

            lane_x = self.lane_x_map.get(lane, 0)
            base_x = self._snap_to_grid(lane_x + (self.SWIMLANE_WIDTH - self.NODE_WIDTH) / 2)
            base_y = self._snap_to_grid(self.node_positions.get(index, 0))

            # 使用注册表渲染节点
            renderer = NODE_RENDERERS.get(node_type, NODE_RENDERERS['process'])
            renderer(self, index, node, base_x, base_y)

            if index == 0:
                self._add_edge(start_node_id, self.node_id_map[str(node.get('id', index))]["entry"], "exitX=0.5;exitY=1;entryX=0.5;entryY=0;")

        # 4. 生成结束节点
        end_node_id = self._get_new_id()
        if self.nodes_data:
            last_node = self.nodes_data[-1]
            last_lane = last_node.get('swimlane_role', self.lanes[-1] if self.lanes else '未定义部门').strip()
        else:
            last_lane = self.lanes[-1] if self.lanes else '未定义部门'
        end_x = self._snap_to_grid(self.lane_x_map.get(last_lane, 0) + (self.SWIMLANE_WIDTH - self.NODE_WIDTH) / 2)
        end_y = self._snap_to_grid(self.node_positions.get('end', 0))

        xml_end = f"""
        <mxCell id="{end_node_id}" value="结束" style="{self._build_style('start_end', shape=start_shape)}" vertex="1" parent="1">
            <mxGeometry x="{end_x}" y="{end_y}" width="{self.NODE_WIDTH}" height="40" as="geometry"/>
        </mxCell>
        """
        self.generated_cells_xml.append(xml_end)

        if self.nodes_data:
             last_ai_id = str(self.nodes_data[-1].get('id'))
             if last_ai_id in self.node_id_map:
                 self._add_edge(self.node_id_map[last_ai_id]["exit"], end_node_id, "exitX=0.5;exitY=1;entryX=0.5;entryY=0;")

        # 5. 生成连线
        id_to_index = {str(n.get('id')): i for i, n in enumerate(self.nodes_data)}

        for node in self.nodes_data:
            ai_id = str(node.get('id'))
            if ai_id not in self.node_id_map: continue

            source_idx = id_to_index.get(ai_id)
            if source_idx is None:
                continue

            transitions = node.get('next_steps', [])
            if not transitions and node.get('next_step_ids'):
                transitions = [{'id': nid, 'condition': ''} for nid in node.get('next_step_ids')]

            if not transitions:
                if source_idx is not None and source_idx < len(self.nodes_data) - 1:
                    next_node = self.nodes_data[source_idx + 1]
                    transitions = [{'id': next_node.get('id'), 'condition': ''}]

            for trans in transitions:
                target_ai_id = str(trans.get('id'))
                label_text = trans.get('condition', '')

                if target_ai_id not in self.node_id_map: continue
                target_idx = id_to_index.get(target_ai_id)
                if target_idx is None: continue

                source_row = self.node_row_indices.get(source_idx)
                target_row = self.node_row_indices.get(target_idx)

                source_node = self.nodes_data[source_idx]
                target_node = self.nodes_data[target_idx]
                source_lane = source_node.get('swimlane_role', '未定义部门').strip()
                target_lane = target_node.get('swimlane_role', '未定义部门').strip()
                source_x = self.lane_x_map.get(source_lane, 0)
                target_x = self.lane_x_map.get(target_lane, 0)

                if source_row == target_row:
                    # 同行横向
                    baseline = self.row_baseline_offsets.get(source_row, 35)

                    src_type = source_node.get('type', 'process')
                    if src_type == 'decision':
                        # 菱形：强制使用 0.5（中心点）
                        src_y_ratio = 0.5
                        src_h = self.DECISION_HEIGHT
                    else:
                        src_doc = source_node.get('output_docs', '')
                        src_h = self.ROLE_HEIGHT + self.STEP_HEIGHT + (self.DOC_HEIGHT if src_doc and str(src_doc).strip() and str(src_doc).lower() != "nan" else 0)
                        src_y_ratio = round(baseline / src_h, 3)

                    tgt_type = target_node.get('type', 'process')
                    if tgt_type == 'decision':
                        # 菱形：强制使用 0.5（中心点）
                        tgt_y_ratio = 0.5
                        tgt_h = self.DECISION_HEIGHT
                    else:
                        tgt_doc = target_node.get('output_docs', '')
                        tgt_h = self.ROLE_HEIGHT + self.STEP_HEIGHT + (self.DOC_HEIGHT if tgt_doc and str(tgt_doc).strip() and str(tgt_doc).lower() != "nan" else 0)
                        tgt_y_ratio = round(baseline / tgt_h, 3)

                    source_xml = self.node_id_map[ai_id]["step"]
                    target_xml = self.node_id_map[target_ai_id]["step"]

                    if source_x < target_x:
                        style_points = f"exitX=1;exitY={src_y_ratio};entryX=0;entryY={tgt_y_ratio};"
                    else:
                        style_points = f"exitX=0;exitY={src_y_ratio};entryX=1;entryY={tgt_y_ratio};"

                else:
                    # 异行纵向
                    source_xml = self.node_id_map[ai_id]["exit"]
                    target_xml = self.node_id_map[target_ai_id]["entry"]
                    style_points = "exitX=0.5;exitY=1;entryX=0.5;entryY=0;"

                # 判断是否为否定分支，使用虚线
                dashed = False

                # 优先读取配置中的关键词列表
                negative_keywords = self.style_manager.get('edges.negative_keywords', ['否', '不符合', '不', '未', 'fail'])
                dashed_enabled = self.style_manager.get('edges.dashed_for_negative', True)

                if dashed_enabled and label_text and any(neg in label_text for neg in negative_keywords):
                    dashed = True

                self._add_edge(source_xml, target_xml, style_points, label_text, dashed=dashed)

        return self._build_full_xml()

    def _add_edge(self, source: str, target: str, style_points: str, label: str = "", dashed: bool = False):
        edge_id = self._get_new_id()
        full_style = self._build_style('edge', dashed=dashed) + style_points

        if label:
            val = self._clean_xml_text(label)
            full_style += "labelBackgroundColor=#ffffff;"
        else:
            val = ""

        xml_edge = f"""
        <mxCell id="{edge_id}" value="{val}" style="{full_style}" edge="1" parent="1" source="{source}" target="{target}">
            <mxGeometry relative="1" as="geometry"/>
        </mxCell>
        """
        self.generated_cells_xml.append(xml_edge)

    def _clean_xml_text(self, text: str) -> str:
        """清理XML文本，转义特殊字符"""
        if not text: return ""
        return html.escape(str(text))

    def _build_full_xml(self) -> str:
        header = """<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="Electron" modified="2024-05-23T00:00:00.000Z" agent="5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) draw.io/24.4.0 Chrome/120.0.6099.109 Electron/28.1.0 Safari/537.36" etag="YourEtag" version="24.4.0" type="device">
  <diagram id="ProcessDiagram" name="Page-1">
    <mxGraphModel dx="1422" dy="868" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169" math="0" shadow="0">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
"""
        footer = """
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
"""
        body = "\n".join(self.generated_cells_xml)
        return header + body + footer