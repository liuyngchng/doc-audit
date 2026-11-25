# 文档审核（doc-audit）—— Docker 服务器版

基于自包含 Node 服务端（`server/server.bundle.js`）与前端静态资源构建的文档审核服务，可部署为 Docker 容器，核心功能完整保留。

## 一、功能范围

| 功能 | Docker 版 |
| --- | --- |
| 文档排版（docx / doc / pdf / md / txt / 在线文档快照 json） | ✅ 可用 |
| 模板集、模板设计器、从 Word 导入模板 | ✅ 可用 |
| 内容校对（错别字/标点/格式，Word 修订模式输出） | ✅ 可用 |
| 人员姓名校对（同音/形近/缺字多字） | ✅ 可用 |
| 任务并行、历史记录、重新转换 | ✅ 可用 |
| 接自定义 OpenAI 兼容模型服务 | ✅ 可用（通过 `cfg.yml`） |

> 文档排版与内容校对依赖 LLM。Docker 版请自备一个 **OpenAI 兼容服务**（baseUrl + apiKey，走 `/chat/completions`）。

## 二、目录结构

```
.
├── Dockerfile
├── docker-compose.yml
├── docker-entrypoint.sh          # 容器入口：读 cfg.yml → 生成 llms.json → 启动服务
├── cfg.example.yml               # LLM 配置模板
├── .dockerignore
├── server/server.bundle.js       # 服务端（已 patch：监听 0.0.0.0）
├── public/                       # Web 界面
├── config/                       # 出厂模板 + apply-cfg.js
├── docs/                         # 产品说明
└── buildinfo.json
```

## 三、快速开始

### 1. 准备 LLM 配置

复制 `cfg.example.yml` 为 `cfg.yml`，填入你的 OpenAI 兼容服务：

```yaml
llm:
  name: "我的模型服务"
  baseUrl: "https://api.example.com/v1"   # 网关地址
  apiKey: "sk-xxxx"                       # 密钥
  timeoutMs: 60
  models:
    - name: "gpt-4o"
      default: true
      maxTokens: 8192
```

> 支持任意 OpenAI 兼容网关（如 OneAPI / New API 等中转），以及自建 vLLM、Ollama 等服务的 OpenAI 兼容端点。

### 2. 构建并启动

```bash
docker compose up -d --build
```

启动后访问 **http://服务器IP:30001**。

### 3. 纯 docker 命令（不用 compose）

```bash
docker build -t doc-audit:1.0.0 .
docker run -d --name doc-audit \
  -p 8765:8765 \
  -v "$PWD/data:/data" \
  -v "$PWD/cfg.yml:/cfg/cfg.yml:ro" \
  --restart unless-stopped \
  doc-audit:1.0.0
```

## 四、配置说明

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `DOC_AUDIT_PORT` | `30001` | 服务端口（同时用于容器内监听与 EXPOSE） |
| `DOC_AUDIT_BIND_HOST` | `0.0.0.0` | 监听地址，一般无需改 |
| `DOC_AUDIT_DATA_DIR` | `/data` | 数据目录（模板、任务、名单、历史） |
| `DOC_AUDIT_CFG` | `/cfg/cfg.yml` | LLM 配置文件路径 |
| `NO_BROWSER` | `1` | 已固定关闭自动开浏览器 |

**持久化目录**（挂载到 `/data`）：

- `/data/templates/` — 用户模板集
- `/data/Tasks/` — 转换产物（输出 docx）
- `/data/config/llms.json` — LLM 配置（由 cfg.yml 自动生成；若不用 cfg.yml，也可在 Web 设置页里配置后持久化在此）
- `/data/rosters.json` — 人员名单

> ⚠️ 容器默认以 root 运行，`/data` 内文件属主为 root。如需非 root 运行，可在 Dockerfile 追加 `USER node` 并保证 `/data` 权限，生产环境建议自行加固。

## 五、改动说明

1. **监听地址 patch**：`server/server.bundle.js` 中 `zl="127.0.0.1"` 改为读取 `DOC_AUDIT_BIND_HOST`，默认 `0.0.0.0`。
2. **补 Linux 原生库**：安装 `@firecrawl/anydoc@0.2.4`（自动拉取 `@firecrawl/anydoc-linux-x64-gnu`），用于 `.doc`/`.pdf` 等文档解析。
3. **新增 cfg.yml 注入**：`config/apply-cfg.js`（依赖 `yaml`）在容器启动时把 `cfg.yml` 转成服务端读取的 `llms.json`。

## 六、常见问题

- **「文档排版需要启用模型服务」**：说明 `cfg.yml` 未挂载或解析失败。检查 `docker logs doc-audit` 中 `[entrypoint]`/`[apply-cfg]` 输出，以及 `/api/llms` 是否返回你的服务。
- **模型调用超时/连不上**：容器内 Node fetch 默认不读系统代理。若网关在公网且需要代理，可通过 `NODE_USE_ENV_PROXY=1` + `http_proxy/https_proxy` 环境变量注入。
- **端口冲突**：改 `DOC_AUDIT_PORT` 并同步 `docker-compose.yml` 的端口映射。
- **如何重置 LLM 配置**：删除 `/data/config/llms.json` 后重启容器（cfg.yml 会重新生成）；或直接在 Web 设置页修改。
