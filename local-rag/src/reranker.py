"""Reranker 模块

提供多种重排序实现：
- SiliconFlowReranker: 硅基流动 API（免费）
- OllamaReranker: 本地 Ollama（embedding-based rerank）
- LocalTransformersReranker: 本地 transformers 模型
- NoopReranker: 不做重排序（用于调试）
"""

import time
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING

import requests

from .exceptions import (
    APIKeyMissingError,
    EmbeddingConnectionError,
    EmbeddingError,
    RateLimitError,
)

if TYPE_CHECKING:
    from .config import Config


class RerankerProvider(ABC):
    """Reranker 抽象基类"""

    @abstractmethod
    def rerank(
        self, query: str, documents: list[dict], top_k: int = 10
    ) -> list[dict]:
        """对候选文档重排序。

        Args:
            query: 查询文本
            documents: [{"text": str, "source": str, "label": str, "distance": float}, ...]
            top_k: 返回前 k 个结果

        Returns:
            同格式但增加 "rerank_score" 字段，按 rerank_score 降序
        """
        pass


class SiliconFlowReranker(RerankerProvider):
    """硅基流动 Reranker（Cohere 风格 API）"""

    def __init__(
        self,
        api_key: str,
        model: str = "BAAI/bge-reranker-v2-m3",
        base_url: str = "https://api.siliconflow.cn/v1",
    ):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url.rstrip("/")

    def rerank(
        self, query: str, documents: list[dict], top_k: int = 10
    ) -> list[dict]:
        """调用 SiliconFlow rerank API

        API 格式（Cohere 风格）:
            POST {base_url}/rerank
            {"model": model, "query": query, "documents": [...], "top_n": top_n, "return_documents": true}

        响应格式:
            {"results": [{"index": int, "relevance_score": float, "document": {"text": str}}, ...]}
        """
        if not documents:
            return []

        url = f"{self.base_url}/rerank"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        payload = {
            "model": self.model,
            "query": query,
            "documents": [doc["text"] for doc in documents],
            "top_n": top_k,
            "return_documents": True,
        }

        # 3次重试，指数退避
        for attempt in range(3):
            try:
                resp = requests.post(url, json=payload, headers=headers, timeout=30)
            except requests.exceptions.RequestException as e:
                raise EmbeddingConnectionError(
                    f"Reranker 连接失败: {e}",
                    hint="请检查网络连接或服务地址",
                )

            if resp.status_code == 200:
                data = resp.json()
                results = []
                for item in data.get("results", []):
                    idx = item["index"]
                    score = item["relevance_score"]
                    doc = dict(documents[idx])
                    doc["rerank_score"] = float(score)
                    results.append(doc)
                return results

            elif resp.status_code == 401:
                raise APIKeyMissingError(
                    "SiliconFlow API Key 无效",
                    hint="请检查 config.yaml 中 reranker.api_key 配置\n"
                    "或设置环境变量 SILICONFLOW_API_KEY",
                )

            elif resp.status_code == 429:
                wait_time = 2 ** (attempt + 1)
                if attempt < 2:
                    time.sleep(wait_time)
                    continue
                else:
                    raise RateLimitError(
                        "超出 API 速率限制",
                        hint="请稍后重试，或升级为付费版本",
                    )

            else:
                error_msg = resp.text
                raise EmbeddingError(
                    f"Reranker 调用失败 (HTTP {resp.status_code})",
                    hint=f"错误信息: {error_msg}",
                )

        return []


class OllamaReranker(RerankerProvider):
    """Ollama Reranker（基于 embedding 的余弦相似度）

    由于 Ollama 没有原生的 rerank API，使用 embedding 端点计算余弦相似度。
    虽然不如 cross-encoder 精准，但可靠且无需额外服务。
    """

    def __init__(
        self, model: str = "linux6200/bge-reranker-v2-m3", url: str = "http://localhost:11434"
    ):
        self.model = model
        self.url = url.rstrip("/")

    def _cosine_similarity(self, a: list[float], b: list[float]) -> float:
        """纯 Python 实现余弦相似度（不依赖 numpy）"""
        dot = sum(x * y for x, y in zip(a, b))
        norm_a = sum(x * x for x in a) ** 0.5
        norm_b = sum(x * x for x in b) ** 0.5
        if norm_a == 0 or norm_b == 0:
            return 0.0
        return dot / (norm_a * norm_b)

    def rerank(
        self, query: str, documents: list[dict], top_k: int = 10
    ) -> list[dict]:
        """基于 embedding 余弦相似度排序"""
        if not documents:
            return []

        try:
            # 1. 获取 query embedding
            resp = requests.post(
                f"{self.url}/api/embed",
                json={"model": self.model, "input": query},
                timeout=30,
            )
            resp.raise_for_status()
            query_emb = resp.json()["embeddings"][0]

            # 2. 批量获取 document embeddings
            doc_texts = [doc["text"] for doc in documents]
            resp = requests.post(
                f"{self.url}/api/embed",
                json={"model": self.model, "input": doc_texts},
                timeout=60,
            )
            resp.raise_for_status()
            doc_embs = resp.json()["embeddings"]

            # 3. 计算余弦相似度
            scores = [
                self._cosine_similarity(query_emb, doc_emb) for doc_emb in doc_embs
            ]

            # 4. 排序取 top_k
            indexed = list(enumerate(scores))
            indexed.sort(key=lambda x: x[1], reverse=True)

            results = []
            for idx, score in indexed[:top_k]:
                doc = dict(documents[idx])
                doc["rerank_score"] = float(score)
                results.append(doc)

            return results

        except requests.exceptions.RequestException as e:
            raise EmbeddingConnectionError(
                f"Ollama 连接失败: {e}",
                hint="请确保 Ollama 服务正在运行: ollama serve",
            )


class LocalTransformersReranker(RerankerProvider):
    """本地 Transformers Reranker（懒加载模型）"""

    def __init__(self, model_name: str = "BAAI/bge-reranker-v2-m3"):
        self.model_name = model_name
        self._model = None
        self._tokenizer = None

    @property
    def model(self):
        """懒加载模型"""
        if self._model is None:
            try:
                from transformers import AutoModelForSequenceClassification, AutoTokenizer
                import torch

                self._tokenizer = AutoTokenizer.from_pretrained(self.model_name)
                self._model = AutoModelForSequenceClassification.from_pretrained(
                    self.model_name
                )
                self._model.eval()
            except ImportError as e:
                raise EmbeddingError(
                    "transformers 未安装",
                    hint="请安装: pip install transformers torch",
                )
            except Exception as e:
                raise EmbeddingError(
                    f"模型加载失败: {e}",
                    hint="请检查模型名称是否正确或网络连接",
                )
        return self._model

    @property
    def tokenizer(self):
        """懒加载 tokenizer"""
        if self._tokenizer is None:
            # 访问 model 属性会同时加载 tokenizer
            _ = self.model
        return self._tokenizer

    def rerank(
        self, query: str, documents: list[dict], top_k: int = 10
    ) -> list[dict]:
        """使用本地 transformers 模型重排序"""
        if not documents:
            return []

        try:
            import torch
            from transformers import AutoTokenizer

            model = self.model
            tokenizer = self.tokenizer

            scores = []
            batch_size = 8

            for i in range(0, len(documents), batch_size):
                batch = documents[i : i + batch_size]
                inputs = tokenizer(
                    [query] * len(batch),
                    [doc["text"] for doc in batch],
                    padding=True,
                    truncation=True,
                    return_tensors="pt",
                )

                with torch.no_grad():
                    outputs = model(**inputs)
                    batch_scores = (
                        torch.sigmoid(outputs.logits[:, 0]).cpu().numpy().tolist()
                    )
                    scores.extend(batch_scores)

            # 排序取 top_k
            indexed = list(enumerate(scores))
            indexed.sort(key=lambda x: x[1], reverse=True)

            results = []
            for idx, score in indexed[:top_k]:
                doc = dict(documents[idx])
                doc["rerank_score"] = float(score)
                results.append(doc)

            return results

        except Exception as e:
            raise EmbeddingError(
                f"Reranker 推理失败: {e}",
                hint="请检查模型是否正确加载或输入数据格式",
            )


class NoopReranker(RerankerProvider):
    """无操作 Reranker（直接返回前 top_k，不计算分数）"""

    def rerank(
        self, query: str, documents: list[dict], top_k: int = 10
    ) -> list[dict]:
        """直接返回原始列表"""
        return documents[:top_k]


def create_reranker_provider(config: "Config") -> RerankerProvider:
    """工厂函数：根据配置创建 RerankerProvider

    配置格式:
        reranker:
            provider: "siliconflow" | "ollama" | "local" | "none"
            model: str
            base_url: str
            api_key: str
    """
    provider = config.get("reranker", {}).get("provider", "siliconflow")

    if provider == "none":
        return NoopReranker()

    elif provider == "siliconflow":
        return SiliconFlowReranker(
            api_key=config.require_api_key("reranker"),
            model=config.get("reranker", {}).get("model", "BAAI/bge-reranker-v2-m3"),
            base_url=config.get("reranker", {}).get("base_url", "https://api.siliconflow.cn/v1"),
        )

    elif provider == "ollama":
        return OllamaReranker(
            model=config.get("reranker", {}).get("model", "linux6200/bge-reranker-v2-m3"),
            url=config.get("reranker", {}).get("base_url", "http://localhost:11434"),
        )

    elif provider == "local":
        return LocalTransformersReranker(
            model_name=config.get("reranker", {}).get("model", "BAAI/bge-reranker-v2-m3"),
        )

    else:
        raise ValueError(f"不支持的 reranker provider: {provider}")