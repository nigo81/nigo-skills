"""Local RAG CLI 命令行工具"""
import argparse
import sys
import traceback
from .pipeline import Pipeline
from .exceptions import LocalRAGError
from .config import Config


def main():
    """CLI 主入口"""
    parser = argparse.ArgumentParser(
        prog="local-rag",
        description="Local RAG - 本地向量知识库工具"
    )
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # setup 命令
    setup_parser = subparsers.add_parser("setup", help="交互式配置向导")
    setup_parser.add_argument("--provider", choices=["siliconflow", "ollama", "openai"], help="非交互模式：指定 Embedding 服务提供商")
    setup_parser.add_argument("--api-key", help="非交互模式：API Key（siliconflow 或 openai 需要此参数）")
    setup_parser.add_argument("--disable-reranker", action="store_true", help="非交互模式：禁用重排序")

    # create 命令
    create_parser = subparsers.add_parser("create", help="创建项目")
    create_parser.add_argument("name", help="项目名称")

    # delete 命令
    delete_parser = subparsers.add_parser("delete", help="删除项目")
    delete_parser.add_argument("name", help="项目名称")

    # list 命令
    subparsers.add_parser("list", help="列出所有项目")

    # ingest 命令
    ingest_parser = subparsers.add_parser("ingest", help="入库文件/文件夹")
    ingest_parser.add_argument("project", help="项目名称")
    ingest_parser.add_argument("path", help="文件或文件夹路径")
    ingest_parser.add_argument("--label", default="", help="可选标签")

    # search 命令
    search_parser = subparsers.add_parser("search", help="语义检索")
    search_parser.add_argument("project", help="项目名称")
    search_parser.add_argument("query", help="查询文本")
    search_parser.add_argument("--top-k", type=int, default=15, help="返回结果数量")
    search_parser.add_argument("--label", default="", help="筛选标签")
    search_parser.add_argument("--rerank", action="store_true", help="启用重排序")

    # chunk-test 命令
    chunk_parser = subparsers.add_parser("chunk-test", help="测试切片效果")
    chunk_parser.add_argument("filepath", help="文件路径")

    # info 命令
    subparsers.add_parser("info", help="显示配置和环境信息")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    try:
        cmd_handlers = {
            "setup": cmd_setup,
            "create": cmd_create,
            "delete": cmd_delete,
            "list": cmd_list,
            "ingest": cmd_ingest,
            "search": cmd_search,
            "chunk-test": cmd_chunk_test,
            "info": cmd_info,
        }

        handler = cmd_handlers.get(args.command)
        if handler:
            handler(args)
    except LocalRAGError as e:
        print(f"❌ {e.message}", file=sys.stderr)
        if e.hint:
            print(f"💡 {e.hint}", file=sys.stderr)
        sys.exit(1)
    except Exception:
        print(f"❌ 发生未预期的错误:", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)


def _save_env_to_shell(var_name: str, value: str):
    """将 API Key 写入 shell 配置文件（跨平台：macOS/Linux ~/.zshrc 或 ~/.bashrc，Windows %USERPROFILE%\\.env.local-rag）"""
    import os
    import platform
    from pathlib import Path

    system = platform.system()

    if system == "Windows":
        # Windows: 写入 %USERPROFILE%\.env.local-rag，提示用户用 setx 或手动设置
        env_file = Path.home() / ".env.local-rag"
        export_line = f"{var_name}={value}"
        lines = []
        if env_file.exists():
            content = env_file.read_text(encoding="utf-8")
            import re
            pattern = rf"^{var_name}=.*$"
            if re.search(pattern, content, re.MULTILINE):
                content = re.sub(pattern, export_line, content, flags=re.MULTILINE)
                env_file.write_text(content, encoding="utf-8")
            else:
                with open(env_file, "a", encoding="utf-8") as f:
                    f.write(f"\n{export_line}\n")
        else:
            env_file.write_text(f"# Local RAG API Key\n{export_line}\n", encoding="utf-8")

        # 用 subprocess 安全调用 setx（避免 shell 注入）
        import subprocess
        subprocess.run(["setx", var_name, value], check=False)
        os.environ[var_name] = value
        print(f"✅ 已通过 setx 设置 {var_name}（新终端窗口生效）")
        print(f"   同时写入: {env_file}")

    else:
        # macOS / Linux: 写入 ~/.zshrc 或 ~/.bashrc
        shell = os.environ.get("SHELL", "")
        if "zsh" in shell:
            rc_file = Path.home() / ".zshrc"
        elif "bash" in shell:
            rc_file = Path.home() / ".bashrc"
        else:
            rc_file = Path.home() / ".zshrc"

        export_line = f'export {var_name}="{value}"'
        if rc_file.exists():
            content = rc_file.read_text(encoding="utf-8")
            import re
            pattern = rf"^export {var_name}=.*$"
            if re.search(pattern, content, re.MULTILINE):
                content = re.sub(pattern, export_line, content, flags=re.MULTILINE)
                rc_file.write_text(content, encoding="utf-8")
                print(f"✅ 已更新 {rc_file} 中的 {var_name}")
            else:
                with open(rc_file, "a", encoding="utf-8") as f:
                    f.write(f"\n# Local RAG API Key\n{export_line}\n")
                print(f"✅ 已将 {var_name} 写入 {rc_file}")
        else:
            rc_file.write_text(f"# Local RAG API Key\n{export_line}\n", encoding="utf-8")
            print(f"✅ 已创建 {rc_file} 并写入 {var_name}")

        os.environ[var_name] = value
        print(f"💡 新终端窗口将自动生效，或执行: source {rc_file}")


def cmd_setup(args):
    """交互式配置向导"""
    import os
    import requests
    import yaml
    from pathlib import Path
    from .embedding import SiliconFlowEmbedding, OpenAIEmbedding

    # 检测 Ollama
    ollama_url = "http://localhost:11434"
    ollama_available = False
    try:
        requests.get(f"{ollama_url}/api/tags", timeout=2)
        ollama_available = True
        print("✅ 检测到 Ollama 服务")
    except Exception:
        print("⚠️  未检测到 Ollama 服务")

    # 判断是否为非交互模式
    non_interactive = hasattr(args, 'provider') and args.provider is not None

    # 选择 embedding provider
    if non_interactive:
        # 非交互模式：使用 CLI 参数
        choice_map = {
            "siliconflow": "1",
            "ollama": "2",
            "openai": "3"
        }
        choice = choice_map.get(args.provider, "1")
        print(f"使用非交互模式: provider={args.provider}")
    else:
        # 交互模式：提示用户选择
        print("\n请选择 Embedding 服务提供商:")
        print("1. 硅基流动 (推荐，免费额度)")
        print("2. Ollama (本地运行)")
        print("3. OpenAI")

        try:
            choice = input("请输入选项 (1-3) [1]: ").strip() or "1"
        except EOFError:
            print("⚠️  检测到非交互环境（stdin 已关闭）", file=sys.stderr)
            print("💡 请使用 --provider 和 --api-key 参数进行非交互配置，例如:", file=sys.stderr)
            print("   local-rag setup --provider siliconflow --api-key YOUR_API_KEY", file=sys.stderr)
            sys.exit(1)

    config = {"embedding": {}, "reranker": {}}

    if choice == "1":
        # 硅基流动
        if non_interactive:
            if not args.api_key:
                print("❌ --api-key 参数是 siliconflow provider 必需的", file=sys.stderr)
                sys.exit(1)
            api_key = args.api_key
        else:
            try:
                api_key = input("请输入 API Key: ").strip()
            except EOFError:
                print("⚠️  检测到非交互环境（stdin 已关闭）", file=sys.stderr)
                print("💡 请使用 --provider 和 --api-key 参数进行非交互配置", file=sys.stderr)
                sys.exit(1)

        # 验证
        embedder = SiliconFlowEmbedding(api_key=api_key)
        try:
            embedder.embed(["test"])
            print("✅ API Key 验证成功")
        except Exception as e:
            print(f"❌ API Key 验证失败: {e}")
            sys.exit(1)

        config["embedding"]["provider"] = "siliconflow"
        config["embedding"]["api_key"] = "${SILICONFLOW_API_KEY}"
        config["embedding"]["model"] = "BAAI/bge-m3"
        config["reranker"]["provider"] = "siliconflow"
        config["reranker"]["api_key"] = "${SILICONFLOW_API_KEY}"
        config["reranker"]["model"] = "BAAI/bge-reranker-v2-m3"

        # 写入 shell 配置文件
        _save_env_to_shell("SILICONFLOW_API_KEY", api_key)

    elif choice == "2":
        if not ollama_available:
            print("❌ Ollama 服务不可用")
            sys.exit(1)
        config["embedding"]["provider"] = "ollama"
        config["embedding"]["model"] = "qwen3-embedding:4b"
        config["reranker"]["provider"] = "ollama"
        config["reranker"]["model"] = "linux6200/bge-reranker-v2-m3"

    elif choice == "3":
        # OpenAI
        if non_interactive:
            if not args.api_key:
                print("❌ --api-key 参数是 openai provider 必需的", file=sys.stderr)
                sys.exit(1)
            api_key = args.api_key
        else:
            try:
                api_key = input("请输入 API Key: ").strip()
            except EOFError:
                print("⚠️  检测到非交互环境（stdin 已关闭）", file=sys.stderr)
                print("💡 请使用 --provider 和 --api-key 参数进行非交互配置", file=sys.stderr)
                sys.exit(1)

        embedder = OpenAIEmbedding(api_key=api_key)
        try:
            embedder.embed(["test"])
            print("✅ API Key 验证成功")
        except Exception as e:
            print(f"❌ API Key 验证失败: {e}")
            sys.exit(1)

        config["embedding"]["provider"] = "openai"
        config["embedding"]["api_key"] = "${OPENAI_API_KEY}"
        config["embedding"]["model"] = "text-embedding-3-small"
        config["reranker"]["provider"] = "none"

        # 写入 shell 配置文件
        _save_env_to_shell("OPENAI_API_KEY", api_key)

    else:
        print("❌ 无效选项")
        sys.exit(1)

    # 询问是否启用 reranker
    if non_interactive:
        # 非交互模式：根据 --disable-reranker 标志决定
        if hasattr(args, 'disable_reranker') and args.disable_reranker:
            config["reranker"]["provider"] = "none"
            print("重排序已禁用（--disable-reranker）")
    else:
        try:
            enable_rerank = input("\n是否启用重排序 (Y/n) [Y]: ").strip().lower()
            if enable_rerank == "n":
                config["reranker"]["provider"] = "none"
        except EOFError:
            print("⚠️  检测到非交互环境（stdin 已关闭）", file=sys.stderr)
            print("💡 请使用 --provider 和 --api-key 参数进行非交互配置", file=sys.stderr)
            sys.exit(1)

    # 保存配置（使用跨平台数据目录）
    from .config import _default_data_dir
    data_dir = os.environ.get("RAG_DATA_DIR", _default_data_dir())
    config_path = Path(os.path.expanduser(data_dir)) / "config.yaml"
    config_path.parent.mkdir(parents=True, exist_ok=True)
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)

    print(f"\n✅ 配置已保存到: {config_path}")

    print("\n快速开始:")
    print("  python3 -m src create my-project")
    print("  python3 -m src ingest my-project /path/to/docs")
    print("  python3 -m src search my-project \"查询内容\"")


def cmd_create(args):
    """创建项目"""
    pipeline = Pipeline()
    result = pipeline.create_project(args.name)
    print(f"✅ 项目 '{result['name']}' 创建成功")


def cmd_delete(args):
    """删除项目"""
    pipeline = Pipeline()
    result = pipeline.delete_project(args.name)
    print(f"✅ 项目 '{result['name']}' 已删除")


def cmd_list(args):
    """列出所有项目"""
    pipeline = Pipeline()
    projects = pipeline.list_projects()

    if not projects:
        print("📂 暂无项目")
        return

    print("\n项目列表:")
    print("-" * 50)
    print(f"{'项目名':<20} | {'文档数':<10}")
    print("-" * 50)
    for p in projects:
        print(f"{p['name']:<20} | {p['count']:<10}")
    print("-" * 50)


def cmd_ingest(args):
    """入库文件/文件夹"""
    pipeline = Pipeline()
    result = pipeline.ingest(args.project, args.path, args.label)
    print(f"✅ 入库完成: {result['total_files']} 个文件, {result['total_chunks']} 个 chunks")
    if result.get("errors", 0) > 0:
        print(f"⚠️  {result['errors']} 个文件处理失败")


def cmd_search(args):
    """语义检索"""
    pipeline = Pipeline()

    if args.rerank:
        results = pipeline.search_with_rerank(args.project, args.query, final_k=args.top_k)
    else:
        results = pipeline.search(args.project, args.query, top_k=args.top_k, label=args.label or None)

    if not results:
        print("🔍 未找到相关内容")
        return

    print(f"\n找到 {len(results)} 条相关结果:\n")
    for i, r in enumerate(results, 1):
        distance = r.get("distance", 0)
        rerank_score = r.get("rerank_score")
        score_str = f"{rerank_score:.3f}" if rerank_score else f"{distance:.3f}"
        score_label = "重排序" if rerank_score else "距离"

        print(f"{i}. [{r['source']}]")
        print(f"   {r['text'][:200]}{'...' if len(r['text']) > 200 else ''}")
        print(f"   ({score_label}: {score_str})")
        if r.get("label"):
            print(f"   标签: {r['label']}")
        print()


def cmd_chunk_test(args):
    """测试切片效果"""
    pipeline = Pipeline()
    result = pipeline.chunk_test(args.filepath)

    print(f"文件: {result['file']}")
    print(f"文本长度: {result['text_length']} 字符")
    print(f"切片数: {result['chunks']} 个")
    print(f"\n前 5 个切片预览:\n")
    for i, chunk in enumerate(result["preview"], 1):
        print(f"--- 切片 {i} (长度: {len(chunk)}) ---")
        print(f"{chunk[:150]}{'...' if len(chunk) > 150 else ''}\n")


def cmd_info(args):
    """显示配置和环境信息"""
    config = Config.load()

    emb = config.get("embedding", {})
    rerank = config.get("reranker", {})
    storage = config.get("storage", {})

    print("\n📋 当前配置:\n")
    print(f"Embedding Provider: {emb.get('provider', 'siliconflow')}")
    print(f"Embedding Model: {emb.get('model', 'BAAI/bge-m3')}")
    print(f"Reranker: {'启用' if rerank.get('provider', 'siliconflow') != 'none' else '禁用'}")
    print(f"Reranker Model: {rerank.get('model', 'BAAI/bge-reranker-v2-m3')}")
    from .config import DEFAULTS
    default_dir = DEFAULTS["storage"]["data_dir"]
    print(f"Storage Path: {storage.get('data_dir', default_dir)}")
    print(f"Chunking Strategy: {config.get('chunking', {}).get('strategy', 'chinese_regulation')}")
    print()


if __name__ == "__main__":
    main()