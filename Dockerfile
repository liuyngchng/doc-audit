# 文档审核（doc-audit）—— 服务器 / Docker 版
# 基于自包含 Node 服务端 server/server.bundle.js 与前端静态资源改造而来。
# 保留：文档排版、模板设计/导入、内容校对、人员姓名校对、修订模式输出、任务历史。
# 移除（依赖 Windows 本机能力）：内置登录、.doc->.docx 批量转换（需本机 Word）、自动更新。
FROM node:20-bookworm-slim

# 依赖：anydoc 原生库在 Linux 上需要 libstdc++（gcc 运行库）
RUN apt-get update \
    && apt-get install -y --no-install-recommends libstdc++6 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 先安装 @firecrawl/anydoc（含 Linux x64 原生绑定），npm 会按平台拉取
# @firecrawl/anydoc-linux-x64-gnu 这个 optionalDependency。
# 锁 0.2.4 与原包一致，避免跨版本行为差异。
RUN npm init -y >/dev/null \
    && npm install @firecrawl/anydoc@0.2.4 yaml --no-audit --no-fund

# 复制服务端、前端、配置与内置模板
COPY server ./server
COPY public ./public
COPY config ./config
COPY buildinfo.json ./

# 锚点唯一：Bs=JD(),zl="127.0.0.1"  ->  zl 改成读取环境变量 DOC_AUDIT_BIND_HOST，默认 0.0.0.0
RUN sed -i 's/zl="127\.0\.0\.1"/zl=process.env.DOC_AUDIT_BIND_HOST||"0.0.0.0"/' server/server.bundle.js \
    && grep -c 'DOC_AUDIT_BIND_HOST' server/server.bundle.js

# 让容器可重载 config（cfg.yml 生成 llms.json 时会读这里）
RUN mkdir -p /app/config

# 持久化数据目录：模板、LLM 配置、任务产物、名单等
VOLUME ["/data"]

ENV DOC_AUDIT_DATA_DIR=/data \
    NO_BROWSER=1 \
    DOC_AUDIT_BIND_HOST=0.0.0.0 \
    DOC_AUDIT_PORT=8765

EXPOSE 8765

COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh
RUN chmod +x /usr/local/bin/docker-entrypoint.sh

ENTRYPOINT ["docker-entrypoint.sh"]
