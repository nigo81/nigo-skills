"""ChromaDB 存储模块"""
import os
import hashlib
import chromadb
from .exceptions import ProjectNotFoundError, ProjectExistsError


class ChromaStore:
    """ChromaDB 存储封装"""

    def __init__(self, config, embedding_provider):
        """接收 Config 和 EmbeddingProvider"""
        self._config = config
        self._provider = embedding_provider
        self._client = None  # 延迟初始化
        self._ef = None  # ChromaDBEmbeddingFunction

    @property
    def client(self):
        """延迟初始化 ChromaDB PersistentClient"""
        if self._client is None:
            data_dir = os.path.expanduser(
                self._config.get("storage", {}).get("data_dir", "~/.local/share/local-rag")
            )
            os.makedirs(data_dir, exist_ok=True)
            self._client = chromadb.PersistentClient(path=data_dir)
        return self._client

    @property
    def ef(self):
        """延迟初始化 ChromaDBEmbeddingFunction"""
        if self._ef is None:
            from .embedding import ChromaDBEmbeddingFunction
            self._ef = ChromaDBEmbeddingFunction(self._provider)
        return self._ef

    def create_project(self, name):
        """创建项目（collection）"""
        self.client.get_or_create_collection(name, embedding_function=self.ef)
        return {"name": name, "status": "created"}

    def delete_project(self, name):
        """删除项目"""
        try:
            self.client.delete_collection(name)
        except ValueError as e:
            raise ProjectNotFoundError(f"项目不存在: {name} (原始错误: {e})")
        return {"name": name, "status": "deleted"}

    def list_projects(self):
        """列出所有项目（带文档数）"""
        collections = self.client.list_collections()
        result = []
        for col in collections:
            name = col.name if hasattr(col, "name") else col
            count = col.count() if hasattr(col, "count") else 0
            result.append({"name": name, "count": count})
        return result

    def get_collection(self, name):
        """获取 collection 辅助方法"""
        try:
            return self.client.get_collection(name, embedding_function=self.ef)
        except ValueError as e:
            raise ProjectNotFoundError(f"项目不存在: {name} (原始错误: {e})")

    def add_documents(self, project, chunks, source="", label=""):
        """核心入库方法"""
        col = self.get_collection(project)

        # 生成唯一 ID（16 字符，碰撞概率极低）
        source_hash = hashlib.md5(source.encode()).hexdigest()[:16]
        ids = [f"{source_hash}-{i:04d}" for i in range(len(chunks))]

        # 构建元数据
        metadatas = [
            {"source": source, "label": label, "chunk_index": i}
            for i in range(len(chunks))
        ]

        # 分批入库，每批 100 条
        batch_size = 100
        for i in range(0, len(chunks), batch_size):
            batch_end = i + batch_size
            col.add(
                ids=ids[i:batch_end],
                documents=chunks[i:batch_end],
                metadatas=metadatas[i:batch_end],
            )

        return {"project": project, "chunks": len(chunks), "source": source}

    def query(self, project, query_text, top_k=15, label=None):
        """语义检索"""
        col = self.get_collection(project)

        # 构建 where 条件
        where = {"label": label} if label else None

        # 查询
        results = col.query(
            query_texts=[query_text],
            n_results=top_k,
            where=where,
        )

        # 格式化结果
        documents = results["documents"][0]
        metadatas = results["metadatas"][0]
        distances = results["distances"][0]

        return [
            {
                "text": doc,
                "source": meta.get("source", ""),
                "label": meta.get("label", ""),
                "distance": dist,
            }
            for doc, meta, dist in zip(documents, metadatas, distances)
        ]

    def get_file_text(self, project, source_filename):
        """按 source 获取完整文档"""
        col = self.get_collection(project)

        # 获取该 source 的所有 chunks
        result = col.get(where={"source": source_filename})

        if not result["documents"]:
            return ""

        # 按 chunk_index 排序后拼接
        indexed_chunks = sorted(
            zip(result["documents"], result["metadatas"]),
            key=lambda x: x[1]["chunk_index"],
        )

        return "\n".join([chunk[0] for chunk in indexed_chunks])