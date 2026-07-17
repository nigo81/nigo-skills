"""
Style Manager - 管理流程图样式配置
支持 YAML 和 JSON 两种格式
"""
from pathlib import Path
from typing import Dict, Any, Optional, List
import json
import copy
from datetime import datetime

# 尝试导入yaml，失败则使用json
try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

class StyleManager:
    def __init__(self, theme_name: str = "governance", themes_dir: Optional[Path] = None):
        # 自动检测themes目录
        if themes_dir is None:
            current_file = Path(__file__).resolve()
            skill_dir = current_file.parent.parent
            themes_dir = skill_dir / "themes"

        self.themes_dir = Path(themes_dir)
        self.custom_dir = self.themes_dir / "custom"
        self.current_theme = None
        self.current_config = {}
        self.user_overrides = {}  # 用户通过语言指令调整的样式

        # 加载默认主题
        self.load_theme(theme_name)

    def _load_config_file(self, config_path: Path) -> Dict[str, Any]:
        """加载配置文件（支持YAML和JSON）"""
        if not config_path.exists():
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        with open(config_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 尝试按JSON解析
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass

        # 尝试按YAML解析
        if HAS_YAML:
            try:
                return yaml.safe_load(content)
            except yaml.YAMLError:
                pass

        raise ValueError(f"无法解析配置文件: {config_path}")

    def _save_config_file(self, config_path: Path, config: Dict[str, Any]):
        """保存配置文件（按扩展名决定格式：.json 用JSON，.yaml/.yml 用YAML）"""
        config_path.parent.mkdir(parents=True, exist_ok=True)

        suffix = config_path.suffix.lower()
        if suffix in ('.yaml', '.yml') and HAS_YAML:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
            return config_path

        # 默认/回退：JSON（与内置主题一致）
        json_path = config_path if suffix == '.json' else config_path.with_suffix('.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return json_path

    def load_theme(self, theme_name: str = "governance", keep_overrides: bool = False) -> Dict[str, Any]:
        """加载指定主题配置

        Args:
            theme_name: 主题名称
            keep_overrides: 是否保留用户覆盖项（默认False，加载主题时清空overrides）

        Returns:
            主题配置字典
        """
        # 优先使用JSON（无依赖）
        extensions = ['.json', '.yaml']
        theme_path = None

        # 先检查自定义主题
        for ext in extensions:
            custom_path = self.custom_dir / f"{theme_name}{ext}"
            if custom_path.exists():
                theme_path = custom_path
                break

        # 如果自定义不存在，检查内置主题
        if theme_path is None:
            for ext in extensions:
                builtin_path = self.themes_dir / f"{theme_name}{ext}"
                if builtin_path.exists():
                    theme_path = builtin_path
                    break

        # 如果都不存在，使用默认主题并发出警告
        if theme_path is None:
            import warnings
            warnings.warn(f"主题 '{theme_name}' 不存在，使用默认主题", UserWarning)
            if not self.current_config:
                self.current_config = self._get_default_config()
            return self.current_config

        # 加载配置
        config = self._load_config_file(theme_path)

        # 清空用户覆盖（如果不保留）
        if not keep_overrides:
            self.user_overrides = {}
            self.current_theme = theme_name
            self.current_config = config
            return config

        # 保留用户覆盖，合并到新主题
        if self.user_overrides:
            config = self._merge_config(config, self.user_overrides)

        self.current_theme = theme_name
        self.current_config = config
        return config

    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            'meta': {'name': '默认', 'description': '默认配置'},
            'layout': {
                'swimlane_width': 280,
                'swimlane_header_height': 40,
                'frame_top_height': 40,
                'frame_left_width': 40,
                'node_width': 120,
                'role_height': 30,
                'step_height': 40,
                'doc_height': 40,
                'decision_height': 60,
                'y_start_offset': 140,
                'y_gap': 50
            },
            'colors': {
                'layout_bar': {'fill': '#b4bff5', 'stroke': '#000000'},
                'swimlane_header': {'fill': '#b4bff5', 'stroke': '#000000'},
                'swimlane_body': {'fill': '#FFFFFF', 'stroke': '#000000'},
                'start_end': {'fill': '#d24a4a', 'stroke': '#000000'},
                'node_role': {'fill': '#FFFFFF', 'stroke': '#000000'},
                'node_step': {'fill': '#FFFFFF', 'stroke': '#000000'},
                'node_doc': {'fill': '#FFFFFF', 'stroke': '#000000'},
                'decision': {'fill': '#ffffff', 'stroke': '#000000'},
                'edge': {'stroke': '#000000'}
            },
            'fonts': {
                'layout_bar': {'size': 16, 'family': 'Arial', 'color': '#000000', 'style': 'bold'},
                'swimlane_header': {'size': 14, 'family': 'Arial', 'color': '#000000', 'style': 'normal'},
                'start_end': {'size': 14, 'family': 'Arial', 'color': '#FFFFFF', 'style': 'bold'},
                'node_role': {'size': 12, 'family': 'Arial', 'color': '#000000', 'style': 'normal'},
                'node_step': {'size': 12, 'family': 'Arial', 'color': '#000000', 'style': 'normal'},
                'node_doc': {'size': 10, 'family': 'Arial', 'color': '#000000', 'style': 'normal'},
                'decision': {'size': 12, 'family': 'Arial', 'color': '#000000', 'style': 'normal'}
            },
            'shapes': {'start_end': 'rounded'},
            'edges': {'dashed_for_negative': False}
        }

    def _merge_config(self, base: Dict, override: Dict) -> Dict:
        """递归合并配置（深拷贝避免污染原配置）"""
        result = copy.deepcopy(base)
        for key, value in override.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_config(result[key], value)
            else:
                result[key] = copy.deepcopy(value)
        return result

    def save_as_theme(self, name: str, description: str = "", overwrite: bool = False) -> Dict[str, Any]:
        """保存当前配置为自定义主题

        Args:
            name: 主题名称
            description: 主题描述
            overwrite: 是否覆盖同名主题（默认False）

        Returns:
            包含保存结果和信息的字典：
            {
                'success': bool,
                'message': str,
                'path': Optional[Path],
                'theme_name': str,
                'is_overwrite': bool
            }
        """
        # 检查同名主题是否存在
        existing_path = self._find_theme_path(name)
        if existing_path and not overwrite:
            existing_info = self.get_theme_info(name)
            return {
                'success': False,
                'message': f"主题 '{name}' 已存在（路径: {existing_path}）。使用 overwrite=True 强制覆盖。",
                'path': existing_path,
                'theme_name': name,
                'is_overwrite': False,
                'existing_theme': existing_info
            }

        self.custom_dir.mkdir(parents=True, exist_ok=True)

        # 合并当前配置和用户覆盖
        config = self._merge_config(self.current_config, self.user_overrides)

        # 设置元数据
        config.setdefault('meta', {})
        config['meta']['name'] = name
        if description:
            config['meta']['description'] = description
        config['meta'].setdefault('version', '1.0')
        config['meta']['created_at'] = datetime.now().isoformat()
        config['meta']['saved_from'] = f"主题: {self.current_theme}"

        # 确定保存路径（与内置主题统一为 .json）
        save_path = self.custom_dir / f"{name}.json"

        try:
            saved_path = self._save_config_file(save_path, config)
            is_overwrite = existing_path is not None
            print(f"✓ 主题 '{name}' 保存成功！")
            print(f"  位置: {saved_path}")
            if is_overwrite:
                print(f"  说明: 已覆盖同名主题")
            return {
                'success': True,
                'message': f"主题 '{name}' 保存成功到 {saved_path}",
                'path': saved_path,
                'theme_name': name,
                'is_overwrite': is_overwrite
            }
        except Exception as e:
            error_msg = f"保存主题失败: {e}"
            print(f"⚠️ {error_msg}")
            return {
                'success': False,
                'message': error_msg,
                'path': None,
                'theme_name': name,
                'is_overwrite': False
            }

    def get_layout_params(self) -> Dict[str, float]:
        """获取布局参数"""
        return self.current_config.get('layout', {
            'swimlane_width': 280,
            'swimlane_header_height': 40,
            'frame_top_height': 40,
            'frame_left_width': 40,
            'node_width': 120,
            'role_height': 30,
            'step_height': 40,
            'doc_height': 40,
            'decision_height': 60,
            'y_start_offset': 140,
            'y_gap': 100
        })

    def get_color(self, element: str, key: str, default: str = "#000000") -> str:
        """获取颜色配置"""
        colors = self.current_config.get('colors', {})
        element_colors = colors.get(element, {})
        return element_colors.get(key, default)

    def get_font_config(self, element: str) -> Dict[str, Any]:
        """获取字体配置"""
        fonts = self.current_config.get('fonts', {})
        return fonts.get(element, {
            'size': 12,
            'family': 'Arial',
            'color': '#000000',
            'style': 'normal'
        })

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值（支持点分路径）"""
        parts = key.split('.')
        value = self.current_config

        for part in parts:
            if isinstance(value, dict) and part in value:
                value = value[part]
            else:
                return default

        return value

    def override(self, updates: Dict[str, Any]):
        """应用样式更新"""
        self.user_overrides = self._merge_config(self.user_overrides, updates)
        # 重新应用到当前配置
        if self.current_config:
            self.current_config = self._merge_config(self.current_config, updates)

    def reset_overrides(self) -> Dict[str, Any]:
        """清空用户覆盖项，恢复到主题原始配置

        Returns:
            清除前的 user_overrides 内容
        """
        old_overrides = copy.deepcopy(self.user_overrides)
        self.user_overrides = {}
        # 重新加载当前主题（不带覆盖）
        if self.current_theme:
            self.load_theme(self.current_theme, keep_overrides=False)
        return old_overrides

    def _find_theme_path(self, name: str) -> Optional[Path]:
        """查找主题路径（优先自定义，其次内置）

        Args:
            name: 主题名称

        Returns:
            主题文件路径（如果存在）
        """
        # 先检查自定义主题
        for ext in ['.yaml', '.json']:
            custom_path = self.custom_dir / f"{name}{ext}"
            if custom_path.exists():
                return custom_path

        # 检查内置主题
        for ext in ['.yaml', '.json']:
            builtin_path = self.themes_dir / f"{name}{ext}"
            if builtin_path.exists():
                return builtin_path

        return None

    def list_themes(self, detailed: bool = False) -> List[Dict[str, Any]]:
        """列出所有可用主题（内置+自定义）

        Args:
            detailed: 是否返回详细信息（默认False）

        Returns:
            主题列表，每个主题包含：
            - name: 主题名称
            - is_custom: 是否为自定义主题
            - file_path: 文件路径
            - description: 主题描述（如果detailed=True）
            - created_time: 创建时间（如果detailed=True）
        """
        themes = {}
        theme_names = set()

        # 收集内置主题
        for ext in ['.yaml', '.json']:
            for theme_file in self.themes_dir.glob(f"*{ext}"):
                if theme_file.name not in ['README.md', 'README.json']:
                    name = theme_file.stem
                    theme_names.add(name)
                    themes[name] = {
                        'name': name,
                        'is_custom': False,
                        'file_path': str(theme_file)
                    }

        # 收集自定义主题（覆盖内置同名主题）
        if self.custom_dir.exists():
            for ext in ['.yaml', '.json']:
                for theme_file in self.custom_dir.glob(f"*{ext}"):
                    name = theme_file.stem
                    theme_names.add(name)
                    themes[name] = {
                        'name': name,
                        'is_custom': True,
                        'file_path': str(theme_file)
                    }

        # 如果需要详细信息
        if detailed:
            for name in theme_names:
                try:
                    info = self.get_theme_info(name)
                    if info:
                        themes[name].update({
                            'description': info.get('description', ''),
                            'created_time': info.get('created_time', ''),
                            'version': info.get('version', '')
                        })
                except Exception:
                    # 读取失败时保持基本信息
                    pass

        # 按名称排序，自定义主题优先
        sorted_themes = sorted(
            themes.values(),
            key=lambda x: (not x['is_custom'], x['name'])
        )

        return sorted_themes


    def delete_theme(self, name: str) -> Dict[str, Any]:
        """删除自定义主题

        Args:
            name: 主题名称

        Returns:
            包含删除结果的字典：
            {
                'success': bool,
                'message': str,
                'deleted_path': Optional[Path],
                'theme_name': str
            }
        """
        theme_path = self._find_theme_path(name)

        if not theme_path:
            return {
                'success': False,
                'message': f"主题 '{name}' 不存在",
                'deleted_path': None,
                'theme_name': name
            }

        # 检查是否为内置主题
        if not str(theme_path).startswith(str(self.custom_dir)):
            return {
                'success': False,
                'message': f"主题 '{name}' 是内置主题，不能删除",
                'deleted_path': None,
                'theme_name': name
            }

        try:
            theme_path.unlink()
            print(f"✓ 自定义主题 '{name}' 已删除")
            print(f"  删除文件: {theme_path}")

            # 如果删除的是当前主题，重新加载默认主题
            if name == self.current_theme:
                self.load_theme("governance", keep_overrides=False)

            return {
                'success': True,
                'message': f"自定义主题 '{name}' 已删除",
                'deleted_path': theme_path,
                'theme_name': name
            }
        except Exception as e:
            error_msg = f"删除主题失败: {e}"
            print(f"⚠️ {error_msg}")
            return {
                'success': False,
                'message': error_msg,
                'deleted_path': None,
                'theme_name': name
            }

    def get_theme_info(self, name: str) -> Optional[Dict[str, Any]]:
        """获取主题详细信息

        Args:
            name: 主题名称

        Returns:
            主题信息字典，包含：
            - name: 主题名称
            - description: 主题描述
            - version: 版本号
            - created_at: 创建时间
            - saved_from: 来源主题
            - is_custom: 是否为自定义主题
            - file_path: 文件路径
            - key_colors: 关键颜色配置（主色调摘要）
            - key_shapes: 关键形状配置
        """
        theme_path = self._find_theme_path(name)

        if not theme_path:
            return None

        try:
            config = self._load_config_file(theme_path)
            meta = config.get('meta', {})

            # 提取关键颜色
            colors = config.get('colors', {})
            key_colors = {
                'start_end': colors.get('start_end', {}).get('fill', ''),
                'node_step': colors.get('node_step', {}).get('fill', ''),
                'decision': colors.get('decision', {}).get('fill', ''),
                'edge': colors.get('edge', {}).get('stroke', '')
            }

            # 提取关键形状
            shapes = config.get('shapes', {})
            key_shapes = {
                'start_end_shape': shapes.get('start_end', '')
            }

            return {
                'name': name,
                'description': meta.get('description', ''),
                'version': meta.get('version', '1.0'),
                'created_at': meta.get('created_at', ''),
                'saved_from': meta.get('saved_from', ''),
                'is_custom': str(theme_path).startswith(str(self.custom_dir)),
                'file_path': str(theme_path),
                'key_colors': key_colors,
                'key_shapes': key_shapes
            }
        except Exception:
            return None


