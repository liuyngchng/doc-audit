#!/usr/bin/env bash
# 部署（启动/更新）doc-audit 服务。
#
# 用法：
#   ./deploy.sh              # 构建并启动（docker compose，默认）
#   ./deploy.sh --no-build   # 跳过构建，直接用已有镜像启动
#   ./deploy.sh --docker-run # 不用 compose，改用 docker run
#   VERSION=1.2.3 ./deploy.sh

set -euo pipefail

cd "$(dirname "$0")"

IMAGE_NAME="${IMAGE_NAME:-doc-audit}"
COMPOSE="${COMPOSE:-docker-compose}"
PORT="${DOC_AUDIT_PORT:-30001}"
# cfg.yml 只在运行时从宿主机挂载，绝不打包进镜像（Dockerfile 无 COPY cfg.yml，.dockerignore 已排除）。
CFG_FILE="$PWD/cfg.yml"

NO_BUILD=0
USE_DOCKER_RUN=0

while [ $# -gt 0 ]; do
  case "$1" in
    --no-build)   NO_BUILD=1 ;;
    --docker-run) USE_DOCKER_RUN=1 ;;
    *) echo "未知参数：$1" >&2; exit 2 ;;
  esac
  shift
done

# 1. 检查宿主机 cfg.yml：compose/run 都会在运行时挂载它到 /cfg/cfg.yml。
#    缺失时 docker 会把同名路径当目录创建，这里先兜底（从示例复制，记得填真实密钥）。
if [ ! -f "$CFG_FILE" ]; then
  echo "!! 未找到 $CFG_FILE，已从 cfg.example.yml 复制一份。请填入真实的 baseUrl/apiKey 后再重新部署。"
  cp cfg.example.yml "$CFG_FILE"
fi

# 2. 持久化目录（容器以 root 挂载，缺目录时 docker 会自动创建，这里提前建好避免属主混乱）。
mkdir -p data

# 3. 启动
if [ "$USE_DOCKER_RUN" -eq 1 ]; then
  # docker run 模式：镜像需提前构建，这里默认构建。
  if [ "$NO_BUILD" -eq 0 ]; then
    ./build.sh
  fi

  echo "==> 使用 docker run 启动 ${IMAGE_NAME}，端口 ${PORT}"
  echo "==> cfg.yml 映射：${CFG_FILE} -> /cfg/cfg.yml（只读，运行时挂载）"
  docker rm -f doc-audit >/dev/null 2>&1 || true
  docker run -d --name doc-audit \
    -p "${PORT}:30001" \
    -e DOC_AUDIT_PORT=30001 \
    -v "$PWD/data:/data" \
    -v "$CFG_FILE:/cfg/cfg.yml:ro" \
    --restart unless-stopped \
    "${IMAGE_NAME}:latest"
else
  echo "==> 使用 docker compose 启动，端口 ${PORT}"
  if [ "$NO_BUILD" -eq 0 ]; then
    $COMPOSE up -d --build
  else
    $COMPOSE up -d
  fi
fi

echo "==> 当前状态："
$COMPOSE ps 2>/dev/null || docker ps --filter "name=doc-audit"
echo "==> 访问：http://服务器IP:${PORT}"
