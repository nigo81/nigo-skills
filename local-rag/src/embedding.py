"""Embedding 提供者抽象层"""
import time
import requests
from abc import ABC, abstractmethod

from .exceptions import EmbeddingError, EmbeddingConnectionError, APIKeyMissingError, RateLimitError


class EmbeddingProvider(ABC):
    """Embedding 提供者抽象基类"""

    def embed_query(self, text: str) -> list[float]:
        """生成单个文本的 embedding（默认用 embed 取第一个）"""
        return self.embed([text])[0]

    @property
    @abstractmethod
    def name(self) -> str:
        """提供者名称"""
        pass


class SiliconFlowEmbedding(EmbeddingProvider):
    """SiliconFlow Embedding 提供者"""

    def __init__(
        self,
        api_key: str,
        model: str = "BAAI/bge-m3",
        base_url: str = "https://api.siliconflow.cn/v1",
        batch_size: int = 20,
    ):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.batch_size = batch_size

    @property
    def name(self) -> str:
        return "siliconflow"

    def embed(self, texts: list[str]) -> list[list[float]]:
        """批量生成 embeddings"""
        if not texts:
            return []

        # 分批处理
        all_embeddings = []
        for i in range(0, len(texts), self.batch_size):
            batch = texts[i : i + self.batch_size]
            embeddings = self._embed_batch(batch)
            all_embeddings.extend(embeddings)

        return all_embeddings

    def _embed_batch(self, texts: list[str]) -> list[list[float]]:
        """处理一批 texts"""
        url = f"{self.base_url}/embeddings"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {"model": self.model, "input": texts}

        # 重试逻辑（3 次，指数退避）
        for attempt in range(3):
            try:
                response = requests.post(url, headers=headers, json=payload)
                response.raise_for_status()
                data = response.json()

                # 提取 embeddings
                embeddings = [item["embedding"] for item in data["data"]]
                return embeddings

            except requests.exceptions.ConnectionError as e:
                raise EmbeddingConnectionError(f"无法连接到 SiliconFlow: {e}")

            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code

                if status_code == 401:
                    raise APIKeyMissingError("SiliconFlow API Key 无效")

                if status_code == 429:
                    # 速率限制，重试
                    if attempt < 2:
                        sleep_time = 2**attempt  # 1s, 2s, 4s
                        time.sleep(sleep_time)
                        continue
                    else:
                        raise RateLimitError("超出 SiliconFlow 速率限制")

                # 其他 HTTP 错误
                raise EmbeddingError(f"SiliconFlow API 错误 ({status_code}): {e.response.text}")

            except Exception as e:
                raise EmbeddingError(f"SiliconFlow Embedding 调用失败: {e}")

        return []


class OllamaEmbedding(EmbeddingProvider):
    """Ollama Embedding 提供者"""

    def __init__(self, model: str = "qwen3-embedding:4b", url: str = "http://localhost:11434"):
        self.model = model
        self.url = url

    @property
    def name(self) -> str:
        return "ollama"

    def embed(self, texts: list[str]) -> list[list[float]]:
        """批量生成 embeddings（Ollama 不支持批量，逐个处理）"""
        if not texts:
            return []

        all_embeddings = []
        for text in texts:
            embedding = self._embed_single(text)
            all_embeddings.append(embedding)

        return all_embeddings

    def _embed_single(self, text: str) -> list[float]:
        """生成单个 embedding"""
        url = f"{self.url}/api/embed"
        payload = {"model": self.model, "input": text}

        try:
            response = requests.post(url, json=payload, timeout=120)
            response.raise_for_status()
            data = response.json()
            # Ollama /api/embed 返回 {"embeddings": [[...]]}
            return data["embeddings"][0]

        except requests.exceptions.ConnectionError as e:
            raise EmbeddingConnectionError(f"无法连接到 Ollama: {e}")

        except Exception as e:
            raise EmbeddingError(f"Ollama Embedding 调用失败: {e}")


class OpenAIEmbedding(EmbeddingProvider):
    """OpenAI Embedding 提供者"""

    def __init__(
        self,
        api_key: str,
        model: str = "text-embedding-3-small",
        base_url: str = "https://api.openai.com/v1",
        batch_size: int = 20,
    ):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.batch_size = batch_size

    @property
    def name(self) -> str:
        return "openai"

    def embed(self, texts: list[str]) -> list[list[float]]:
        """批量生成 embeddings"""
        if not texts:
            return []

        # 分批处理
        all_embeddings = []
        for i in range(0, len(texts), self.batch_size):
            batch = texts[i : i + self.batch_size]
            embeddings = self._embed_batch(batch)
            all_embeddings.extend(embeddings)

        return all_embeddings

    def _embed_batch(self, texts: list[str]) -> list[list[float]]:
        """处理一批 texts"""
        url = f"{self.base_url}/embeddings"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {"model": self.model, "input": texts}

        # 重试逻辑（3 次，指数退避）
        for attempt in range(3):
            try:
                response = requests.post(url, headers=headers, json=payload)
                response.raise_for_status()
                data = response.json()

                # 提取 embeddings
                embeddings = [item["embedding"] for item in data["data"]]
                return embeddings

            except requests.exceptions.ConnectionError as e:
                raise EmbeddingConnectionError(f"无法连接到 OpenAI: {e}")

            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code

                if status_code == 401:
                    raise APIKeyMissingError("OpenAI API Key 无效")

                if status_code == 429:
                    # 速率限制，重试
                    if attempt < 2:
                        sleep_time = 2**attempt  # 1s, 2s, 4s
                        time.sleep(sleep_time)
                        continue
                    else:
                        raise RateLimitError("超出 OpenAI 速率限制")

                # 其他 HTTP 错误
                raise EmbeddingError(f"OpenAI API 错误 ({status_code}): {e.response.text}")

            except Exception as e:
                raise EmbeddingError(f"OpenAI Embedding 调用失败: {e}")

        return []


class ChromaDBEmbeddingFunction:
    """ChromaDB 兼容的 embedding function 适配器"""

    def __init__(self, provider: EmbeddingProvider):
        self._provider = provider

    def name(self) -> str:
        return self._provider.name

    def __call__(self, input: list) -> list:
        return self._provider.embed(input)

    def embed_documents(self, input: list) -> list:
        return self._provider.embed(input)

    def embed_query(self, input: list) -> list:
        return self._provider.embed(input)


def create_embedding_provider(config) -> EmbeddingProvider:
    """根据配置创建 embedding provider"""
    provider = config.get("embedding", {}).get("provider", "siliconflow")

    if provider == "siliconflow":
        return SiliconFlowEmbedding(
            api_key=config.require_api_key("embedding"),
            model=config.get("embedding", {}).get("model", "BAAI/bge-m3"),
            base_url=config.get("embedding", {}).get("base_url", "https://api.siliconflow.cn/v1"),
        )
    elif provider == "ollama":
        return OllamaEmbedding(
            model=config.get("embedding", {}).get("model", "qwen3-embedding:4b"),
            url=config.get("embedding", {}).get("base_url", "http://localhost:11434"),
        )
    elif provider == "openai":
        return OpenAIEmbedding(
            api_key=config.require_api_key("embedding"),
            model=config.get("embedding", {}).get("model", "text-embedding-3-small"),
            base_url=config.get("embedding", {}).get("base_url", "https://api.openai.com/v1"),
        )
    else:
        raise ValueError(f"不支持的 embedding provider: {provider}")