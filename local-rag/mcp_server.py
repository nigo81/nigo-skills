#!/usr/bin/env python3
"""Local RAG MCP Server — 在 skill 根目录下运行: python3 mcp_server.py"""

import sys
import os

# 确保 src/ 可导入（无论从哪个目录运行）
_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)

from src.pipeline import Pipeline

try:
    from fastmcp import FastMCP
except ImportError:
    print("请安装 FastMCP: pip install fastmcp")
    raise

mcp = FastMCP("local-rag")

_pipeline = None

def _get_pipeline():
    global _pipeline
    if _pipeline is None:
        _pipeline = Pipeline()
    return _pipeline

@mcp.tool()
def create_project(name: str) -> dict:
    """创建知识库项目"""
    return _get_pipeline().create_project(name)

@mcp.tool()
def delete_project(name: str) -> dict:
    """删除知识库项目"""
    return _get_pipeline().delete_project(name)

@mcp.tool()
def list_projects() -> list:
    """列出所有知识库项目"""
    return _get_pipeline().list_projects()

@mcp.tool()
def ingest(project: str, path: str, label: str = "") -> dict:
    """文档入库（支持文件或文件夹）"""
    return _get_pipeline().ingest(project, path, label)

@mcp.tool()
def search(project: str, query: str, top_k: int = 15) -> list:
    """语义检索"""
    return _get_pipeline().search(project, query, top_k)

@mcp.tool()
def rerank_search(project: str, query: str, top_k: int = 10) -> list:
    """检索 + Rerank 二次精排"""
    return _get_pipeline().search_with_rerank(project, query, top_k)

if __name__ == "__main__":
    mcp.run()