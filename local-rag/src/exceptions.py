"""Local RAG 异常定义"""


class LocalRAGError(Exception):
    """Local RAG 基础异常类"""

    def __init__(self, message: str, hint: str = ""):
        super().__init__(message)
        self.message = message
        self.hint = hint

    def __str__(self):
        if self.hint:
            return f"{self.message}\n💡 {self.hint}"
        return self.message


class ConfigError(LocalRAGError):
    """配置错误"""

    pass


class APIKeyMissingError(ConfigError):
    """API Key 缺失"""

    def __init__(self, message: str = "API Key 未设置", hint: str = "请设置环境变量或在 config.yaml 中配置"):
        super().__init__(message, hint)


class EmbeddingError(LocalRAGError):
    """Embedding 调用失败"""

    pass


class EmbeddingConnectionError(EmbeddingError):
    """Embedding 连接失败"""

    def __init__(
        self, message: str = "Embedding 服务连接失败", hint: str = "请检查网络连接或服务地址"
    ):
        super().__init__(message, hint)


class RateLimitError(EmbeddingError):
    """超出速率限制"""

    def __init__(
        self, message: str = "超出 API 速率限制", hint: str = "请稍后重试，或升级为付费版本"
    ):
        super().__init__(message, hint)


class ParseError(LocalRAGError):
    """文档解析失败"""

    pass


class UnsupportedFormatError(ParseError):
    """不支持的文件格式"""

    def __init__(
        self,
        message: str = "不支持的文件格式",
        hint: str = "支持的格式: .md, .txt, .docx, .doc, .pdf",
    ):
        super().__init__(message, hint)


class EmptyFileError(ParseError):
    """文件为空"""

    def __init__(self, message: str = "文件为空"):
        super().__init__(message)


class MinerUNotInstalledError(ParseError):
    """MinerU 未安装"""

    def __init__(
        self,
        message: str = "MinerU 未安装",
        hint: str = "安装命令: pip install mineru-open-api",
    ):
        super().__init__(message, hint)


class MinerUTokenRequiredError(ParseError):
    """MinerU 需要 Token"""

    def __init__(
        self,
        message: str = "MinerU 需要 API Token",
        hint: str = "请注册 https://mineru.net 获取 API Token，然后设置环境变量 MINERU_API_TOKEN",
    ):
        super().__init__(message, hint)


class ProjectError(LocalRAGError):
    """项目操作失败"""

    pass


class ProjectNotFoundError(ProjectError):
    """项目不存在"""

    def __init__(self, message: str = "项目不存在"):
        super().__init__(message)


class ProjectExistsError(ProjectError):
    """项目已存在"""

    def __init__(self, message: str = "项目已存在"):
        super().__init__(message)