# SNS Market Article Generator

这是从项目中整理出来的 SNS 内容生成和市场文章/报告生成模块，包含一套 FastAPI 后端和一个单页前端。

![none]([https://raw.githubusercontent.com/username/repo/main/image.png](https://github.com/jjjadand/SNS-Generate/blob/main/desktop.png))


## 包含内容

- `main.py`: 后端 API，负责多平台 SNS 文案生成、素材上传、网页素材抓取、实时数据抓取、KOL 管理、市场新闻搜索和市场报告生成。
- `sns-marketing-hub.html`: 前端页面，启动后直接访问 `/`。
- `sns_data.db`: 本地 SQLite 数据库，首次启动时会自动创建；该文件默认不提交到 GitHub。
- `requirements.txt`: 独立运行所需 Python 依赖。
- `start.sh`: 默认以 `0.0.0.0:8005` 启动服务。

## 启动命令

```bash
cd /other_data/Agent-SharedReport/sns-market-article-generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
./start.sh
```

然后访问：

```text
http://127.0.0.1:8005/
```

数据库文件会生成在当前目录的 `sns_data.db`。如需迁移已有数据，可以手动复制原项目中的数据库文件到这个目录。

## API Key

后端会优先读取本机 Codex 配置：

- `~/.codex/config.toml`
- `~/.codex/auth.json`

也可以复制 `.env.example` 为 `.env`，填写 `OPENAI_API_KEY` 和可选的 `OPENAI_BASE_URL`。

## 常用环境变量

```bash
HOST=0.0.0.0 PORT=8005 RELOAD=true ./start.sh
```

如果 8005 已被占用，可以换端口：

```bash
PORT=8010 ./start.sh
```
