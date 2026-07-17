"""
FlowchartValidator - 流程图数据结构校验器

校验节点数据的结构完整性，确保生成的流程图正确无误。
"""

from typing import List, Dict
from dataclasses import dataclass


@dataclass
class ValidationResult:
    """校验结果"""
    errors: List[str]      # 严重错误，必须修复
    warnings: List[str]    # 警告，建议修复
    info: List[str]        # 信息提示
    is_valid: bool         # 是否通过（无errors）
    score: int             # 质量评分 0-100


class FlowchartValidator:
    """流程图数据校验器"""

    # 允许的节点类型
    VALID_TYPES = {'process', 'decision'}

    # 必需的字段
    REQUIRED_FIELDS = {'id', 'step', 'type', 'swimlane_role', 'node_role'}

    def __init__(self):
        """初始化校验器"""
        pass

    def validate(self, nodes_data: List[Dict]) -> ValidationResult:
        """
        校验节点数据

        Args:
            nodes_data: 节点数据列表，每个节点是字典格式

        Returns:
            ValidationResult: 校验结果对象
        """
        errors = []
        warnings = []
        info = []

        if not nodes_data:
            errors.append("节点数据为空")
            return ValidationResult(errors, warnings, info, False, 0)

        # 1. 检查重复ID
        duplicate_ids = self._check_duplicate_ids(nodes_data)
        if duplicate_ids:
            for id_val, count in duplicate_ids.items():
                errors.append(f"重复ID: '{id_val}' 重复 {count} 次")

        # 2. 检查缺失字段
        missing_fields_errors = self._check_missing_fields(nodes_data)
        errors.extend(missing_fields_errors)

        # 3. 检查空节点
        empty_node_errors = self._check_empty_nodes(nodes_data)
        errors.extend(empty_node_errors)

        # 4. 检查孤儿下一步（None 或空字符串）
        orphan_next_steps = self._check_orphan_next_steps(nodes_data)
        errors.extend(orphan_next_steps)

        # 5. 检查悬空连线（next_steps 指向不存在的节点）
        dangling_edges = self._check_dangling_edges(nodes_data)
        errors.extend(dangling_edges)

        # 6. 检查孤立节点
        isolated_nodes = self._check_isolated_nodes(nodes_data)
        warnings.extend(isolated_nodes)

        # 7. 检查循环引用（警告）
        circular_refs = self._check_circular_references(nodes_data)
        warnings.extend(circular_refs)

        # 统计信息
        info.append(f"节点数量: {len(nodes_data)}")

        # 计算质量评分
        score = max(0, 100 - len(errors) * 20 - len(warnings) * 5)

        is_valid = len(errors) == 0

        return ValidationResult(errors, warnings, info, is_valid, score)

    def _check_duplicate_ids(self, nodes_data: List[Dict]) -> Dict[str, int]:
        """
        检查重复ID

        Returns:
            Dict: 重复ID及其出现次数
        """
        id_counts = {}
        duplicates = {}

        for node in nodes_data:
            node_id = node.get('id')
            if node_id:
                id_counts[node_id] = id_counts.get(node_id, 0) + 1

        for id_val, count in id_counts.items():
            if count > 1:
                duplicates[id_val] = count

        return duplicates

    def _check_missing_fields(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查缺失字段

        Returns:
            List[str]: 错误消息列表
        """
        errors = []

        for i, node in enumerate(nodes_data):
            node_id = node.get('id', f'节点{i}')

            for field in self.REQUIRED_FIELDS:
                if field not in node:
                    errors.append(f"节点 '{node_id}' 缺失字段: {field}")
                elif node[field] is None:
                    errors.append(f"节点 '{node_id}' 字段 '{field}' 为 None")
                elif field == 'type' and node[field] not in self.VALID_TYPES:
                    errors.append(f"节点 '{node_id}' 的 type 必须是 'process' 或 'decision'，实际值: {node[field]}")

        return errors

    def _check_empty_nodes(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查空节点（step 或 swimlane_role 为空）

        Returns:
            List[str]: 错误消息列表
        """
        errors = []

        for i, node in enumerate(nodes_data):
            node_id = node.get('id', f'节点{i}')

            step = node.get('step', '')
            if step == '' or (isinstance(step, str) and step.strip() == ''):
                errors.append(f"节点 '{node_id}' 的 step 为空")

            role = node.get('swimlane_role', '')
            if role == '' or (isinstance(role, str) and role.strip() == ''):
                errors.append(f"节点 '{node_id}' 的 swimlane_role 为空")

        return errors

    def _check_orphan_next_steps(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查孤儿下一步（id 为 None 或空字符串）

        Returns:
            List[str]: 错误消息列表
        """
        errors = []

        for node in nodes_data:
            node_id = node.get('id', '未知节点')
            next_steps = node.get('next_steps', [])

            if not isinstance(next_steps, list):
                continue

            for next_step in next_steps:
                if not isinstance(next_step, dict):
                    continue
                target_id = next_step.get('id')
                if target_id is None:
                    errors.append(f"节点 '{node_id}' 的 next_steps 包含 id 为 None 的项")
                elif isinstance(target_id, str) and target_id.strip() == '':
                    errors.append(f"节点 '{node_id}' 的 next_steps 包含空 id")

        return errors

    def _check_dangling_edges(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查悬空连线（next_steps 指向不存在的节点）

        Returns:
            List[str]: 错误消息列表
        """
        errors = []

        # 构建所有节点ID的集合
        valid_ids = {node.get('id') for node in nodes_data if node.get('id')}

        for node in nodes_data:
            node_id = node.get('id', '未知节点')
            next_steps = node.get('next_steps', [])

            if not isinstance(next_steps, list):
                continue

            for next_step in next_steps:
                # 处理 dict 类型（包含 id 和 condition）
                if isinstance(next_step, dict):
                    target_id = next_step.get('id')
                else:
                    target_id = next_step

                if target_id and target_id not in valid_ids:
                    errors.append(f"节点 '{node_id}' 的 next_steps 指向不存在的节点 ID: '{target_id}'")

        return errors

    def _check_isolated_nodes(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查孤立节点（既没有 next_steps，也没有任何节点指向它）

        例外：第一个节点和最后一个节点不算孤立

        Returns:
            List[str]: 警告消息列表
        """
        warnings = []

        if len(nodes_data) < 3:
            # 节点太少，不检查孤立节点
            return warnings

        # 构建所有节点ID的集合
        all_ids = {node.get('id') for node in nodes_data if node.get('id')}

        # 构建被引用的节点ID集合
        referenced_ids = set()

        for node in nodes_data:
            next_steps = node.get('next_steps', [])
            if isinstance(next_steps, list):
                for next_step in next_steps:
                    # 处理 dict 类型（包含 id 和 condition）
                    if isinstance(next_step, dict):
                        target_id = next_step.get('id')
                    else:
                        target_id = next_step

                    if target_id:
                        referenced_ids.add(target_id)

        # 找出孤立节点
        # 规则：没有 next_steps 且没有被任何节点引用的节点
        # 例外：第一个节点（索引0）和最后一个节点（索引-1）不算孤立

        first_node_id = nodes_data[0].get('id')
        last_node_id = nodes_data[-1].get('id')

        for i, node in enumerate(nodes_data):
            node_id = node.get('id')
            if not node_id:
                continue

            # 跳过第一个和最后一个节点
            if node_id == first_node_id or node_id == last_node_id:
                continue

            next_steps = node.get('next_steps', [])
            has_next = next_steps and isinstance(next_steps, list) and len(next_steps) > 0
            is_referenced = node_id in referenced_ids

            if not has_next and not is_referenced:
                warnings.append(f"孤立节点: '{node_id}'（既没有 next_steps，也没有任何节点指向它）")

        return warnings

    def _check_circular_references(self, nodes_data: List[Dict]) -> List[str]:
        """
        检查循环引用（A→B→A 这种直接循环）

        注意：这不是错误，只是警告，因为某些流程确实需要循环

        Returns:
            List[str]: 警告消息列表
        """
        warnings = []

        # 构建邻接表
        adj = {}
        for node in nodes_data:
            node_id = node.get('id')
            if node_id:
                next_steps = node.get('next_steps', [])
                if isinstance(next_steps, list):
                    # 处理 dict 类型（包含 id 和 condition）
                    adj[node_id] = []
                    for next_step in next_steps:
                        if isinstance(next_step, dict):
                            target_id = next_step.get('id')
                        else:
                            target_id = next_step
                        if target_id:
                            adj[node_id].append(target_id)

        # 检查直接循环（A→B 且 B→A）
        for node_id, next_steps in adj.items():
            for next_id in next_steps:
                if next_id in adj and node_id in adj[next_id]:
                    # 避免重复报告（只报告 A→B，不报告 B→A）
                    if node_id < next_id:  # 按字典序避免重复
                        warnings.append(f"检测到循环引用: '{node_id}' ↔ '{next_id}'")

        return warnings
