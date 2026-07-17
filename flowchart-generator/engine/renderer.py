"""
Renderer - 渲染器，支持PNG、SVG和Visio导出

制作人：nigo
微信公众号：逆行的狗
"""
import subprocess
import sys
import platform
import zipfile
from pathlib import Path
from typing import Optional
import urllib.request
import tempfile
import shutil
import os
import time

class Renderer:
    # VSDX 支持的最后一个版本
    VSDX_COMPATIBLE_VERSION = "26.0.16"

    def __init__(self, force_vsdx_compatible: bool = False):
        """
        初始化渲染器

        Args:
            force_vsdx_compatible: 强制使用支持 VSDX 的版本（v26.0.16）
                                 如果设置为 True 且当前版本不支持 VSDX，
                                 会自动下载并安装兼容版本
        """
        self.force_vsdx_compatible = force_vsdx_compatible
        self.drawio_path = self._find_drawio()
        self._version_checked = False
        self._current_version = None

        if force_vsdx_compatible:
            self._ensure_vsdx_compatible_version()

    def _get_platform_name(self) -> str:
        """获取平台名称"""
        system = platform.system().lower()
        machine = platform.machine().lower()

        if system == 'darwin':
            if machine in ('arm64', 'aarch64'):
                return 'macos-arm64'
            return 'macos-intel'
        elif system == 'windows':
            return 'windows'
        elif system == 'linux':
            if machine in ('x86_64', 'amd64'):
                return 'linux-x64'
            return 'linux-arm64'
        return 'unknown'

    def _get_download_urls(self) -> dict:
        """获取 draw.io v26.0.16 的下载链接"""
        return {
            'macos-arm64': 'https://github.com/jgraph/drawio-desktop/releases/download/v26.0.16/draw.io-arm64-26.0.16.dmg',
            'macos-intel': 'https://github.com/jgraph/drawio-desktop/releases/download/v26.0.16/draw.io-x86_64-26.0.16.dmg',
            'windows': 'https://github.com/jgraph/drawio-desktop/releases/download/v26.0.16/draw.io-26.0.16-windows-installer.exe',
            'linux-x64': 'https://github.com/jgraph/drawio-desktop/releases/download/v26.0.16/draw.io-amd64-26.0.16.deb',
            'linux-arm64': 'https://github.com/jgraph/drawio-desktop/releases/download/v26.0.16/drawio-arm64-26.0.16.deb',
        }

    def _get_common_paths(self) -> list:
        """获取跨平台常见安装路径"""
        paths = []

        if platform.system() == 'Darwin':  # macOS
            paths = [
                Path('/Applications/draw.io.app/Contents/MacOS/draw.io'),
                Path('/Applications/Draw.io.app/Contents/MacOS/draw.io'),
                Path('/Applications/drawio.app/Contents/MacOS/draw.io'),
            ]
        elif platform.system() == 'Windows':  # Windows
            # 常见 Windows 安装路径
            program_files = Path(os.environ.get('ProgramFiles', 'C:\\Program Files'))
            program_files_x86 = Path(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'))
            local_app_data = Path(os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local')))

            paths = [
                program_files / 'draw.io' / 'draw.io.exe',
                program_files_x86 / 'draw.io' / 'draw.io.exe',
                local_app_data / 'draw.io' / 'draw.io.exe',
                Path('C:\\draw.io\\draw.io.exe'),
            ]
        else:  # Linux
            paths = [
                Path('/usr/local/bin/draw.io'),
                Path('/opt/homebrew/bin/drawio'),
                Path('/usr/bin/drawio'),
            ]

        return paths

    def _find_drawio(self) -> Optional[Path]:
        """查找draw.io可执行文件"""
        # 检查常见路径
        for path in self._get_common_paths():
            if path.exists():
                return path

        # 尝试通过 which/where 查找
        try:
            if platform.system() == 'Windows':
                result = subprocess.run(['where', 'draw.io'], capture_output=True, text=True)
            else:
                result = subprocess.run(['which', 'draw.io'], capture_output=True, text=True)

            if result.returncode == 0 and result.stdout.strip():
                return Path(result.stdout.strip().split('\n')[0].strip())
        except:
            pass

        return None

    def _get_drawio_version(self) -> Optional[str]:
        """获取 draw.io 版本号"""
        if not self.drawio_path:
            return None

        try:
            result = subprocess.run(
                [str(self.drawio_path), '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                version = result.stdout.strip()
                self._current_version = version
                return version
        except:
            pass

        return None

    def _is_version_compatible_with_vsdx(self) -> bool:
        """检查当前版本是否支持 VSDX 导出"""
        version = self._get_drawio_version()
        if not version:
            return False

        # 版本号格式化处理
        version_parts = version.split('.')
        if len(version_parts) < 2:
            return False

        major, minor = int(version_parts[0]), int(version_parts[1])

        # v26.0.16 及更早版本支持 VSDX
        # v26.2.2 之后移除 VSDX 支持
        if major < 26:
            return True
        if major == 26 and minor <= 0:
            return True

        return False

    def _disable_auto_update(self):
        """禁用 draw.io 自动更新"""
        config_path = None

        if platform.system() == 'Darwin':  # macOS
            config_path = Path.home() / 'Library' / 'Application Support' / 'draw.io' / '.preferences'
        elif platform.system() == 'Windows':  # Windows
            config_path = Path(os.environ.get('APPDATA', os.path.expanduser('~\\AppData\\Roaming'))) / 'draw.io' / '.preferences'
        elif platform.system() == 'Linux':
            config_path = Path.home() / '.config' / 'draw.io' / '.preferences'

        if config_path:
            config_path.parent.mkdir(parents=True, exist_ok=True)
            config = {
                "checkForUpdates": False,
                "disableUpdate": True,
                "disableUpdates": True,
                "noUpdate": True
            }
            import json
            with open(config_path, 'w') as f:
                json.dump(config, f)

    def _download_file(self, url: str, dest: Path) -> Path:
        """下载文件"""
        dest.parent.mkdir(parents=True, exist_ok=True)

        print(f"正在下载: {url}")
        urllib.request.urlretrieve(url, dest)
        print(f"✓ 下载完成: {dest}")

        return dest

    def _install_on_windows(self, installer_path: Path):
        """在 Windows 上安装 draw.io"""
        import subprocess

        print(f"正在安装 draw.io v26.0.16 (Windows)...")
        # 静默安装
        subprocess.run([str(installer_path), '/S', '/D=C:\\draw.io-26.0.16'], check=True)

        # 更新路径
        self.drawio_path = Path('C:\\draw.io-26.0.16\\draw.io.exe')
        self._disable_auto_update()

    def _install_on_macos(self, dmg_path: Path):
        """在 macOS 上安装 draw.io"""
        import subprocess

        print(f"正在安装 draw.io v26.0.16 (macOS)...")

        # 挂载 DMG
        mount_result = subprocess.run(
            ['hdiutil', 'attach', str(dmg_path)],
            capture_output=True,
            text=True
        )

        if mount_result.returncode != 0:
            raise RuntimeError(f"DMG 挂载失败: {mount_result.stderr}")

        # 查找挂载点
        mount_point = None
        for line in mount_result.stdout.split('\n'):
            if '/Volumes/' in line:
                mount_point = line.split()[-1]
                break

        if not mount_point:
            raise RuntimeError("无法找到 DMG 挂载点")

        try:
            dmg_app = Path(mount_point) / 'draw.io.app'

            # 备份当前版本
            current_app = Path('/Applications/draw.io.app')
            if current_app.exists():
                backup_path = Path(f'/Applications/draw.io.app.backup_{int(time.time())}')
                shutil.move(current_app, backup_path)
                print(f"✓ 已备份当前版本: {backup_path}")

            # 复制新版本
            shutil.copytree(dmg_app, '/Applications/draw.io.app')
            print("✓ 安装完成")

            # 更新路径
            self.drawio_path = Path('/Applications/draw.io.app/Contents/MacOS/draw.io')

        finally:
            # 卸载 DMG
            subprocess.run(['hdiutil', 'detach', mount_point], capture_output=True)

        self._disable_auto_update()

    def _ensure_vsdx_compatible_version(self, skip_prompt: bool = False):
        """确保使用支持 VSDX 的版本
        
        Args:
            skip_prompt: 是否跳过用户确认（当从 export_to_vsdx 调用时为 True）
        """
        if not self.drawio_path:
            raise RuntimeError(
                "未找到 draw.io。\n"
                "请先安装 Draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases"
            )

        # 检查版本
        if self._is_version_compatible_with_vsdx():
            print(f"✓ 当前版本 {self._current_version} 支持 VSDX 导出")
            return

        print(f"⚠ 当前版本 {self._current_version} 不支持 VSDX 导出")
        print(f"VSDX 导出需要 draw.io v26.0.16 或更早版本")

        # 询问用户是否安装（如果未跳过）
        if not skip_prompt:
            try:
                response = input("\n是否自动安装 draw.io v26.0.16？(y/n): ").strip().lower()
                if response != 'y':
                    raise RuntimeError(
                        "需要支持 VSDX 的版本才能导出 Visio 格式。\n"
                        "手动下载地址：https://github.com/jgraph/drawio-desktop/releases/tag/v26.0.16"
                    )
            except EOFError:
                # 非交互式环境，直接安装
                pass

        # 自动下载和安装
        platform_name = self._get_platform_name()
        urls = self._get_download_urls()

        if platform_name not in urls:
            raise RuntimeError(f"不支持的平台: {platform_name}")

        download_url = urls[platform_name]

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir) / f"drawio_v26.0.16.{platform_name}"
            downloaded_file = self._download_file(download_url, temp_path)

            if platform_name.startswith('macos'):
                self._install_on_macos(downloaded_file)
            elif platform_name.startswith('windows'):
                self._install_on_windows(downloaded_file)
            else:
                raise RuntimeError(f"暂不支持 {platform_name} 的自动安装，请手动安装")

        # 验证安装
        new_version = self._get_drawio_version()
        if not self._is_version_compatible_with_vsdx():
            raise RuntimeError("安装后版本仍不支持 VSDX 导出")

        print(f"✓ 成功安装 draw.io {new_version}")

    def _run_export(self, cmd: list, output_file: Path):
        """执行导出命令的通用方法"""
        if not self.drawio_path:
            raise RuntimeError(
                "未找到draw.io可执行文件。\n"
                "请安装Draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases\n"
                "或使用在线版：https://app.diagrams.net/"
            )

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            if result.returncode != 0:
                raise RuntimeError(f"draw.io导出失败: {result.stderr}")
        except subprocess.TimeoutExpired:
            raise RuntimeError("draw.io导出超时（60秒）")
        except FileNotFoundError:
            raise RuntimeError(f"draw.io可执行文件不存在: {self.drawio_path}")

        if not output_file.exists():
            raise RuntimeError(f"文件未生成: {output_file}")

    def export_to_png(
        self,
        drawio_file: Path,
        output_file: Optional[Path] = None,
        scale: float = 2.0,
        page_index: int = 0
    ) -> Path:
        """
        导出为PNG图片

        Args:
            drawio_file: .drawio文件路径
            output_file: 输出PNG路径（默认同目录，同名.png）
            scale: 缩放比例（1.0-4.0，默认2.0高清）
            page_index: 页面索引（默认0，第一页）

        Returns:
            输出PNG文件路径
        """
        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.png')

        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 构建draw.io命令
        # 注意：page-index 在 draw.io 命令行中是 1-based
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--scale', str(scale),
            '--format', 'png',
            '--page-index', str(page_index + 1)
        ]

        self._run_export(cmd, output_file)
        return output_file

    def export_to_svg(
        self,
        drawio_file: Path,
        output_file: Optional[Path] = None,
        page_index: int = 0
    ) -> Path:
        """
        导出为SVG矢量图

        Args:
            drawio_file: .drawio文件路径
            output_file: 输出SVG路径（默认同目录，同名.svg）
            page_index: 页面索引（默认0，第一页）

        Returns:
            输出SVG文件路径
        """
        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.svg')

        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 构建draw.io命令
        # 注意：page-index 在 draw.io 命令行中是 1-based
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--format', 'svg',
            '--page-index', str(page_index + 1)
        ]

        self._run_export(cmd, output_file)
        return output_file

    def export_to_vsdx(
        self,
        drawio_file: Path,
        output_file: Optional[Path] = None,
        page_index: int = 0,
        auto_install: bool = True
    ) -> Path:
        """
        导出为Visio格式（.vsdx）

        **注意**：
        - VSDX 导出需要 draw.io v26.0.16 或更早版本
        - v26.2.2 之后已移除此功能
        - 设置 auto_install=True 可自动下载并安装兼容版本

        Args:
            drawio_file: .drawio文件路径
            output_file: 输出vsdx路径（默认同目录，同名.vsdx）
            page_index: 页面索引（默认0，第一页）
            auto_install: 如果版本不支持，是否自动安装兼容版本（默认True）

        Returns:
            输出vsdx文件路径

        Raises:
            RuntimeError: draw.io版本不支持VSDX导出且未启用自动安装
        """
        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.vsdx')

        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 检查并确保版本兼容
        if not self._is_version_compatible_with_vsdx():
            if auto_install:
                print(f"\n{'='*60}")
                print(f"⚠️  VSDX 导出需要兼容的 draw.io 版本")
                print(f"当前版本: {self._current_version}")
                print(f"需要版本: {self.VSDX_COMPATIBLE_VERSION} (支持 VSDX 的最后一个版本)")
                print(f"{'='*60}")
                
                try:
                    response = input("\n是否自动安装 draw.io v26.0.16？(y/n): ").strip().lower()
                    if response == 'y':
                        print("\n✓ 开始自动安装...")
                        # 传入 skip_prompt=True 避免重复询问
                        self._ensure_vsdx_compatible_version(skip_prompt=True)
                    else:
                        raise RuntimeError(
                            f"需要支持 VSDX 的版本才能导出 Visio 格式。\n"
                            f"手动下载地址：https://github.com/jgraph/drawio-desktop/releases/tag/v26.0.16\n"
                            f"安装后请禁用自动更新功能"
                        )
                except (EOFError, KeyboardInterrupt):
                    # 非交互式环境，默认为否
                    print("\n⚠ 非交互式环境，跳过自动安装")
                    raise RuntimeError(
                        f"需要支持 VSDX 的版本才能导出 Visio 格式。\n"
                        f"手动下载地址：https://github.com/jgraph/drawio-desktop/releases/tag/v26.0.16\n"
                        f"或设置 auto_install=False 且在交互式环境中运行"
                    )
            else:
                raise RuntimeError(
                    f"当前版本 {self._current_version} 不支持VSDX导出。\n"
                    f"VSDX导出需要 draw.io v26.0.16 或更早版本。\n"
                    f"下载地址：https://github.com/jgraph/drawio-desktop/releases/tag/v26.0.16\n"
                    f"或设置 auto_install=True 自动安装兼容版本"
                )

        # 构建draw.io命令
        # 注意：page-index 在 draw.io 命令行中是 1-based
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--format', 'vsdx',
            '--page-index', str(page_index + 1)
        ]

        self._run_export(cmd, output_file)

        # 验证是否是真正的 VSDX（应该是 ZIP 格式）
        try:
            with zipfile.ZipFile(output_file, 'r') as zf:
                # VSDX 应该包含 visio/ 目录
                if 'visio/' not in [name for name in zf.namelist() if name.endswith('/')]:
                    raise RuntimeError(
                        "导出的文件不是有效的VSDX格式。\n"
                        "可能使用了不支持的draw.io版本。"
                    )
        except zipfile.BadZipFile:
            raise RuntimeError(
                "导出的文件不是有效的VSDX格式。\n"
                "可能使用了不支持的draw.io版本。"
            )

        return output_file