"""文本切片模块

根据配置的切片策略将文本切分为 chunks。
"""

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Any, Dict


def _chinese_numbers(max_n: int = 99) -> list[str]:
    """生成中文数字列表（一到九十九）"""
    digits = "一二三四五六七八九"
    result = []
    for i in range(1, min(max_n + 1, 100)):
        if i <= 9:
            result.append(digits[i - 1])
        elif i == 10:
            result.append("十")
        elif i < 20:
            result.append(f"十{digits[i - 11]}")
        else:
            tens = i // 10
            ones = i % 10
            s = f"{digits[tens - 1]}十"
            if ones > 0:
                s += digits[ones - 1]
            result.append(s)
    return result


def chunk_text(text: str, config: "Dict[str, Any]") -> list[str]:
    """根据配置的切片策略将文本切分为 chunks

    Args:
        text: 待切片的文本
        config: 配置对象（dict-like），需要 chunking.strategy, chunking.chunk_size 等

    Returns:
        切片后的文本列表
    """
    # 导入 Chonkie（延迟导入以避免硬依赖）
    try:
        from chonkie import RecursiveChunker, OverlapRefinery
        from chonkie.types import RecursiveLevel, RecursiveRules
    except ImportError:
        raise ImportError(
            "Chonkie is required for chunking. Install with: pip install chonkie"
        )

    # 获取配置参数
    chunking_config = config.get("chunking", {})
    strategy = chunking_config.get("strategy", "chinese_regulation")
    chunk_size = chunking_config.get("chunk_size", 800)
    chunk_overlap = chunking_config.get("chunk_overlap", 0.15)
    min_chunk_size = chunking_config.get("min_chunk_size", 50)

    # 根据策略创建规则
    if strategy == "chinese_regulation":
        # 中文法规策略：按章节切分
        chinese_nums = _chinese_numbers(99)
        level1_delimiters = [f"\n第{num}章" for num in chinese_nums]

        rules = RecursiveRules(
            levels=[
                RecursiveLevel(delimiters=level1_delimiters),
                RecursiveLevel(delimiters=["\n\n"]),
                RecursiveLevel(delimiters=["\n"]),
            ]
        )
    else:  # generic
        # 通用策略：按 markdown 标题切分
        rules = RecursiveRules(
            levels=[
                RecursiveLevel(delimiters=["\n# "]),
                RecursiveLevel(delimiters=["\n## "]),
                RecursiveLevel(delimiters=["\n### "]),
                RecursiveLevel(delimiters=["\n\n"]),
                RecursiveLevel(delimiters=["\n"]),
            ]
        )

    # 创建切片器
    chunker = RecursiveChunker(
        tokenizer="character",
        chunk_size=chunk_size,
        rules=rules,
        min_characters_per_chunk=min_chunk_size,
    )

    # 创建重叠处理器
    # context_size 需要整数（字符数），chunk_overlap 默认 0.15 表示比例
    overlap_size = int(chunk_size * chunk_overlap) if isinstance(chunk_overlap, float) else chunk_overlap
    refinery = OverlapRefinery(context_size=overlap_size, method="suffix")

    # 执行切片
    chunks = chunker.chunk(text)

    # 应用重叠
    chunks = refinery.refine(chunks)

    # 提取文本并过滤短片段
    result = [chunk.text for chunk in chunks if len(chunk.text) >= min_chunk_size]

    return result