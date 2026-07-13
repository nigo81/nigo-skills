"""
Renderer - 渲染器，支持PNG和Visio导出
"""
import subprocess
from pathlib import Path
from typing import Optional

class Renderer:
    def __init__(self):
        self.drawio_path = self._find_drawio()

    def _find_drawio(self) -> Optional[Path]:
        """查找draw.io可执行文件"""
        # 尝试常见路径
        possible_paths = [
            Path('/Applications/draw.io.app/Contents/MacOS/draw.io'),
            Path('/Applications/Draw.io.app/Contents/MacOS/draw.io'),
            Path('/Applications/drawio.app/Contents/MacOS/drawio'),
            Path('/usr/local/bin/drawio'),
            Path('/opt/homebrew/bin/drawio'),
        ]

        for path in possible_paths:
            if path.exists():
                return path

        # 尝试通过which查找
        try:
            result = subprocess.run(['which', 'draw.io'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return Path(result.stdout.strip())
        except:
            pass

        return None

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
        if not self.drawio_path:
            raise RuntimeError(
                "未找到draw.io可执行文件。\n"
                "请安装Draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases\n"
                "或使用在线版：https://app.diagrams.net/"
            )

        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.png')

        # 确保输出目录存在
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 构建draw.io命令
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--scale', str(scale),
            '--format', 'png',
            '--page-index', str(page_index)
        ]

        # 执行命令
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode != 0:
                raise RuntimeError(f"draw.io导出失败: {result.stderr}")
        except subprocess.TimeoutExpired:
            raise RuntimeError("draw.io导出超时（30秒）")
        except FileNotFoundError:
            raise RuntimeError(f"draw.io可执行文件不存在: {self.drawio_path}")

        if not output_file.exists():
            raise RuntimeError(f"PNG文件未生成: {output_file}")

        return output_file

    def export_to_vsd(
        self,
        drawio_file: Path,
        output_file: Optional[Path] = None,
        page_index: int = 0
    ) -> Path:
        """
        导出为Visio格式（.vsdx）

        Args:
            drawio_file: .drawio文件路径
            output_file: 输出vsdx路径（默认同目录，同名.vsdx）
            page_index: 页面索引（默认0，第一页）

        Returns:
            输出vsdx文件路径
        """
        if not self.drawio_path:
            raise RuntimeError(
                "未找到draw.io可执行文件。\n"
                "请安装Draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases\n"
                "或使用在线版：https://app.diagrams.net/"
            )

        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.vsdx')

        # 确保输出目录存在
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 构建draw.io命令
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--format', 'vsdx',
            '--page-index', str(page_index)
        ]

        # 执行命令
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode != 0:
                raise RuntimeError(f"draw.io导出失败: {result.stderr}")
        except subprocess.TimeoutExpired:
            raise RuntimeError("draw.io导出超时（30秒）")
        except FileNotFoundError:
            raise RuntimeError(f"draw.io可执行文件不存在: {self.drawio_path}")

        if not output_file.exists():
            raise RuntimeError(f"vsdx文件未生成: {output_file}")

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
        if not self.drawio_path:
            raise RuntimeError(
                "未找到draw.io可执行文件。\n"
                "请安装Draw.io Desktop：https://github.com/jgraph/drawio-desktop/releases\n"
                "或使用在线版：https://app.diagrams.net/"
            )

        if not drawio_file.exists():
            raise FileNotFoundError(f"文件不存在: {drawio_file}")

        if output_file is None:
            output_file = drawio_file.with_suffix('.svg')

        # 确保输出目录存在
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # 构建draw.io命令
        cmd = [
            str(self.drawio_path),
            '--export',
            str(drawio_file),
            '--output', str(output_file),
            '--format', 'svg',
            '--page-index', str(page_index)
        ]

        # 执行命令
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode != 0:
                raise RuntimeError(f"draw.io导出失败: {result.stderr}")
        except subprocess.TimeoutExpired:
            raise RuntimeError("draw.io导出超时（30秒）")
        except FileNotFoundError:
            raise RuntimeError(f"draw.io可执行文件不存在: {self.drawio_path}")

        if not output_file.exists():
            raise RuntimeError(f"SVG文件未生成: {output_file}")

        return output_file
