"""配置管理模块

读取 config.yaml + 环境变量 + 默认值，合并为最终配置。
"""

import os
import re
import sys
from pathlib import Path
from typing import Any

import yaml

from .exceptions import APIKeyMissingError, ConfigError


def _default_data_dir() -> str:
    """跨平台默认数据目录"""
    if sys.platform == "win32":
        # Windows: C:\Users\xxx\AppData\Local\local-rag
        base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~/AppData/Local"))
        return os.path.join(base, "local-rag")
    elif sys.platform == "darwin":
        # macOS: ~/Library/Application Support/local-rag
        return os.path.expanduser("~/Library/Application Support/local-rag")
    else:
        # Linux/其他: ~/.local/share/local-rag
        return os.path.expanduser("~/.local/share/local-rag")


DEFAULTS: dict[str, Any] = {
    "embedding": {
        "provider": "siliconflow",
        "model": "BAAI/bge-m3",
        "api_key": "",
        "base_url": "https://api.siliconflow.cn/v1",
    },
    "reranker": {
        "provider": "siliconflow",
        "model": "BAAI/bge-reranker-v2-m3",
        "api_key": "",
        "base_url": "https://api.siliconflow.cn/v1",
    },
    "chunking": {
        "strategy": "chinese_regulation",
        "chunk_size": 800,
        "chunk_overlap": 0.15,
        "min_chunk_size": 50,
    },
    "storage": {
        "data_dir": _default_data_dir(),
    },
    "logging": {
        "level": "INFO",
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    """深度合并两个字典，override 覆盖 base"""
    result = base.copy()
    for key, value in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _resolve_env_vars(value: Any) -> Any:
    """解析 ${ENV_VAR} 语法"""
    if isinstance(value, str):
        # 匹配 ${VAR} 或 $VAR
        def _replacer(match):
            var_name = match.group(1) or match.group(2)
            resolved = os.environ.get(var_name)
            if resolved is None:
                import warnings
                warnings.warn(
                    f"配置中引用了环境变量 ${var_name}，但该变量未设置",
                    stacklevel=3,
                )
                return ""
            return resolved

        return re.sub(r"\$\{([^}]+)\}|\$([A-Za-z_][A-Za-z0-9_]*)", _replacer, value)
    elif isinstance(value, dict):
        return {k: _resolve_env_vars(v) for k, v in value.items()}
    elif isinstance(value, list):
        return [_resolve_env_vars(v) for v in value]
    return value


class Config(dict):
    """配置管理

    用法:
        config = Config.load()
        config.get("embedding", {}).get("provider")
        config.require_api_key("embedding")
    """

    def __init__(self, data: dict[str, Any] | None = None):
        super().__init__(data or {})

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值，支持多级 key"""
        return super().get(key, default)

    def require_api_key(self, purpose: str = "embedding") -> str:
        """获取 API Key，缺失时抛出异常

        Args:
            purpose: "embedding" 或 "reranker"

        Returns:
            API Key 字符串

        Raises:
            APIKeyMissingError: API Key 未配置
        """
        section = self.get(purpose, {})
        api_key = section.get("api_key", "")

        if not api_key:
            provider = section.get("provider", "unknown")
            if provider == "siliconflow":
                raise APIKeyMissingError(
                    "硅基流动 API Key 未设置",
                    hint="请设置环境变量 SILICONFLOW_API_KEY\n"
                    "或在 config.yaml 中配置 embedding.api_key\n"
                    "注册地址: https://cloud.siliconflow.cn",
                )
            elif provider == "openai":
                raise APIKeyMissingError(
                    "OpenAI API Key 未设置",
                    hint="请设置环境变量 OPENAI_API_KEY\n"
                    "或在 config.yaml 中配置 embedding.api_key",
                )
            else:
                raise APIKeyMissingError(
                    f"{purpose} 的 API Key 未设置",
                    hint="请在 config.yaml 中配置或设置对应的环境变量",
                )

        return api_key

    @classmethod
    def load(cls, config_path: str | None = None) -> "Config":
        """加载配置

        优先级:
        1. 函数参数指定的 config_path
        2. 当前目录的 config.yaml
        3. 数据目录的 config.yaml (~/.local/share/local-rag/config.yaml)
        4. 环境变量覆盖
        5. 内置默认值
        """
        # 从默认值开始
        merged = DEFAULTS.copy()

        # 搜索配置文件路径
        search_paths = []
        if config_path:
            search_paths.append(config_path)
        search_paths.append(os.path.join(os.getcwd(), "config.yaml"))

        # 数据目录路径（需要先从环境变量或默认值获取）
        data_dir = os.environ.get("RAG_DATA_DIR", _default_data_dir())
        search_paths.append(os.path.join(data_dir, "config.yaml"))

        # 加载第一个找到的配置文件
        for path in search_paths:
            if os.path.isfile(path):
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        file_config = yaml.safe_load(f) or {}
                    # 解析 ${ENV_VAR} 引用
                    file_config = _resolve_env_vars(file_config)
                    merged = _deep_merge(merged, file_config)
                except Exception as e:
                    raise ConfigError(
                        f"配置文件读取失败: {path}",
                        hint=f"请检查 YAML 语法是否正确\n错误: {e}",
                    )
                break  # 只加载第一个找到的

        # 环境变量覆盖（传入已加载的配置，用于判断 provider）
        env_overrides = _build_env_overrides(merged)
        if env_overrides:
            merged = _deep_merge(merged, env_overrides)

        return cls(merged)


def _build_env_overrides(merged_config: dict[str, Any] | None = None) -> dict[str, Any]:
    """从环境变量构建覆盖配置

    Args:
        merged_config: 已加载的配置（用于判断 provider），可选
    """
    overrides: dict[str, Any] = {}

    # SILICONFLOW_API_KEY → embedding.api_key + reranker.api_key
    sf_key = os.environ.get("SILICONFLOW_API_KEY")
    if sf_key:
        overrides.setdefault("embedding", {})["api_key"] = sf_key
        overrides.setdefault("reranker", {})["api_key"] = sf_key

    # OPENAI_API_KEY → embedding.api_key
    # 仅当 provider 是 openai（或尚未设置 provider）时才覆盖
    # 避免 OPENAI_API_KEY 意外覆盖 SILICONFLOW_API_KEY
    openai_key = os.environ.get("OPENAI_API_KEY")
    if openai_key:
        current_provider = (merged_config or {}).get("embedding", {}).get("provider", "siliconflow")
        if current_provider == "openai":
            overrides.setdefault("embedding", {})["api_key"] = openai_key

    # RAG_DATA_DIR → storage.data_dir
    rag_dir = os.environ.get("RAG_DATA_DIR")
    if rag_dir:
        overrides.setdefault("storage", {})["data_dir"] = rag_dir

    # RAG_EMBEDDING_MODEL → embedding.model
    rag_model = os.environ.get("RAG_EMBEDDING_MODEL")
    if rag_model:
        overrides.setdefault("embedding", {})["model"] = rag_model

    # OLLAMA_URL → embedding.base_url + reranker.base_url (仅当 provider 是 ollama)
    ollama_url = os.environ.get("OLLAMA_URL")
    if ollama_url:
        overrides.setdefault("embedding", {})["base_url"] = ollama_url
        overrides.setdefault("reranker", {})["base_url"] = ollama_url

    return overrides
