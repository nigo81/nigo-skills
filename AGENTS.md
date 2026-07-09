# nigo-skills 协作规范

本文件是本仓库的长期记忆，AI 助手在此仓库工作时必须遵守。

## 作者署名规范（强制）

**每个 skill 的 SKILL.md 必须包含作者署名**，格式统一为：

```markdown
---

> 作者：nigo
> 微信公众号：逆行的狗
```

要求：
- 署名块放在 SKILL.md 的**最末尾**（`---` 分隔 + 引用块）
- 新增 skill 时必须追加；若 skill 由其他工具（如女娲 Skill 造人术）生成，原工具署名保留在上，nigo 署名追加在最下方
- 文件名不区分大小写：`SKILL.md` 与 `skill.md` 都要遵守
- 已有不同措辞署名（如「作者：nigo（公众号：逆行的狗）」）视为满足要求，无需强行改写

## 隐私与安全红线（强制）

skill 文件**严禁包含**以下内容：

- ❌ API Key、Token、Cookie、密码、Bearer 凭证等任何凭证（示例占位符如 `sk-xxx` 除外）
- ❌ 运行时缓存文件（如 `.athena_key_cache.json`、`.security-key`、`*.cache`）
- ❌ 个人手机号、身份证号、家庭住址等个人隐私信息（公开监管处罚决定书中的号码属公开记录，可保留）
- ❌ 具体审计项目的客户名称、项目编号、底稿数据等未公开业务信息
- ❌ 内网 IP、服务器凭证、内部系统地址

API Key 等凭证**一律走环境变量**（如 `SILICONFLOW_API_KEY`），config 文件中用 `${ENV_VAR}` 引用，禁止硬编码。

## 文件卫生

- `.gitignore` 已忽略：`__pycache__/`、`*.pyc`、`.DS_Store`、`logs/`、`*.cache`、`config.yaml`、`.env` 等
- Python 临时产物（`__pycache__`、`.pyc`、`*.egg-info`、`.pytest_cache`）不得提交
- 同一逻辑只保留一份代码，删除重复文件（如 `lib/` 与 `scripts/` 中的同名副本只保留实际被引用的那份）
- skill 之间若有依赖关系，在 SKILL.md 中明确标注（如 related-party-identification 依赖 cicpa-company-query 取数）

## Git 提交

- commit message 用中文简要说明改动内容
- 提交前自查：无隐私泄露、无临时文件、每个 SKILL.md 都带署名
- 不要提交 `.env`、凭证文件或本地 config

## 仓库结构

```
nigo-skills/
├── AGENTS.md                      # 本文件（协作规范）
├── README.md                      # 仓库说明与 skill 列表
├── .gitignore
└── <skill-name>/                  # 每个 skill 一个目录
    ├── SKILL.md                   # 入口（必须含作者署名）
    ├── references/                # 参考资料（可选）
    └── scripts/                   # 可执行脚本（可选）
```
