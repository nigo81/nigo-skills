"""Pipeline 编排层 — 唯一的业务编排入口"""

import os
import sys
from .config import Config
from .embedding import create_embedding_provider
from .reranker import create_reranker_provider
from .store import ChromaStore
from .parser import parse_file
from .chunker import chunk_text
from .exceptions import ParseError, UnsupportedFormatError

SUPPORTED_EXTENSIONS = {".md", ".txt", ".docx", ".doc", ".pdf"}


class Pipeline:
    """Pipeline 编排层 — 唯一的业务编排入口"""

    def __init__(self, config_path=None):
        """初始化 Pipeline

        Args:
            config_path: 配置文件路径（可选，默认自动搜索）
        """
        self.config = Config.load(config_path)
        self.embedder = create_embedding_provider(self.config)
        self.reranker = create_reranker_provider(self.config)
        self.store = ChromaStore(self.config, self.embedder)

    def create_project(self, name):
        """创建项目

        Args:
            name: 项目名称

        Returns:
            {"name": name, "status": "created"}
        """
        return self.store.create_project(name)

    def delete_project(self, name):
        """删除项目

        Args:
            name: 项目名称

        Returns:
            {"name": name, "status": "deleted"}
        """
        return self.store.delete_project(name)

    def list_projects(self):
        """列出所有项目

        Returns:
            项目名称列表
        """
        return self.store.list_projects()

    def ingest(self, project, path, label=""):
        """ ingestion 文件或目录到项目

        Args:
            project: 项目名称
            path: 文件或目录路径
            label: 可选标签

        Returns:
            {"project": project, "total_files": int, "total_chunks": int, "errors": int}
        """
        if os.path.isfile(path):
            # 单个文件：用文件所在目录作为 base_path，source = 文件名
            files_added, chunks_added, has_error = self._ingest_file(
                project, path, label, os.path.dirname(path)
            )
            return {
                "project": project,
                "total_files": files_added,
                "total_chunks": chunks_added,
                "errors": 1 if has_error else 0,
            }
        else:
            # 目录遍历
            return self._ingest_directory(project, path, label)

    def _ingest_file(self, project, filepath, label, base_path):
        """处理单个文件

        Returns:
            (total_files, total_chunks, has_error)
        """
        ext = os.path.splitext(filepath)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            return 0, 0, False

        # 提前计算 source，避免 except 中 UnboundLocalError
        source = os.path.relpath(filepath, base_path)

        try:
            # 解析
            text = parse_file(filepath)
            # 切片
            chunks = chunk_text(text, self.config)
            # 入库
            self.store.add_documents(project, chunks, source=source, label=label)
            return 1, len(chunks), False
        except Exception as e:
            print(f"  ⚠️  文件处理失败: {source} - {e}", file=sys.stderr)
            return 0, 0, True

    def _ingest_directory(self, project, dir_path, label):
        """遍历目录处理所有支持格式的文件

        Returns:
            {"project": project, "total_files": int, "total_chunks": int, "errors": int}
        """
        # 先遍历一次获取总文件数
        total_files = 0
        files_to_process = []
        for root, dirs, files in os.walk(dir_path):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in SUPPORTED_EXTENSIONS:
                    total_files += 1
                    files_to_process.append(os.path.join(root, f))

        # 处理文件
        total_chunks = 0
        error_count = 0
        processed_count = 0

        for fpath in files_to_process:
            processed_count += 1
            if processed_count % 10 == 0:
                print(f"  已处理 {processed_count}/{total_files} 文件...", file=sys.stderr)

            files_added, chunks_added, has_error = self._ingest_file(
                project, fpath, label, dir_path
            )
            total_chunks += chunks_added
            if has_error:
                error_count += 1

        return {
            "project": project,
            "total_files": total_files,
            "total_chunks": total_chunks,
            "errors": error_count,
        }

    def search(self, project, query, top_k=15, label=None):
        """语义检索

        Args:
            project: 项目名称
            query: 查询文本
            top_k: 返回结果数量
            label: 可选标签过滤

        Returns:
            结果列表 [{"text": str, "source": str, "label": str, "distance": float}, ...]
        """
        return self.store.query(project, query, top_k, label=label)

    def search_with_rerank(self, project, query, final_k=10):
        """语义检索 + 重排序

        先检索 top_k * 3 个候选，再用 reranker 重排序到 final_k

        Args:
            project: 项目名称
            query: 查询文本
            final_k: 最终返回结果数量

        Returns:
            结果列表 [{"text": str, "source": str, "label": str, "distance": float, "rerank_score": float}, ...]
        """
        candidates = self.store.query(project, query, top_k=final_k * 3)
        return self.reranker.rerank(query, candidates, final_k)

    def chunk_test(self, filepath):
        """测试切片效果

        Args:
            filepath: 文件路径

        Returns:
            {"file": filepath, "text_length": int, "chunks": int, "preview": list[str]}
        """
        text = parse_file(filepath)
        chunks = chunk_text(text, self.config)
        return {
            "file": filepath,
            "text_length": len(text),
            "chunks": len(chunks),
            "preview": chunks[:5],
        }