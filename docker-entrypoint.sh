#!/bin/sh
set -e

PORT="${DOC_AUDIT_PORT:-30001}"
export PORT

# 从 cfg.yml 生成 LLM 配置（OpenAI 兼容服务）。
# 优先用 $DOC_AUDIT_CFG 指向的文件；默认挂载 /cfg/cfg.yml。
CFG_FILE="${DOC_AUDIT_CFG:-/cfg/cfg.yml}"
if [ -f "$CFG_FILE" ]; then
    echo "[entrypoint] 检测到 $CFG_FILE，正在生成 llms.json ..."
    node /app/config/apply-cfg.js "$CFG_FILE" "$DOC_AUDIT_DATA_DIR/config/llms.json"
else
    echo "[entrypoint] 未检测到 $CFG_FILE，使用已存在的 /data/config/llms.json（如有）或 Web 设置页配置。"
fi

echo "[entrypoint] 启动文档审核服务，监听 ${DOC_AUDIT_BIND_HOST:-0.0.0.0}:${PORT}，数据目录 ${DOC_AUDIT_DATA_DIR}"
exec node /app/server/server.bundle.js "$PORT"
