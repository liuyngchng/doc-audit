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
# 版本号：命令行参数 > 环境变量 VERSION > buildinfo.json 里的 version
VERSION="${VERSION:-}"
if [ -z "$VERSION" ]; then
  VERSION="$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' buildinfo.json | head -n1)"
fi
if [ -z "$VERSION" ]; then
  echo "错误：无法确定版本号。请用 ./deploy.sh <版本> 或 VERSION=<版本> ./deploy.sh 指定，或检查 buildinfo.json" >&2
  exit 1
fi
# 把版本号透传给 docker-compose（其 image 标签引用 ${VERSION}）
export VERSION

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

  echo "==> 使用 docker run 启动 ${IMAGE_NAME}:${VERSION}，端口 ${PORT}"
  echo "==> cfg.yml 映射：${CFG_FILE} -> /cfg/cfg.yml（只读，运行时挂载）"
  docker rm -f doc-audit >/dev/null 2>&1 || true
  docker run -d --name doc-audit \
    -p "${PORT}:30001" \
    -e DOC_AUDIT_PORT=30001 \
    -v "$PWD/data:/data" \
    -v "$CFG_FILE:/cfg/cfg.yml:ro" \
    --restart unless-stopped \
    "${IMAGE_NAME}:${VERSION}"
else
  echo "==> 使用 docker compose 启动，端口 ${PORT}"
  # --force-recreate：每次部署都重建容器，让 entrypoint 重新读取 cfg.yml 并重新生成 llms.json，
  # 避免「改了 cfg.yml 但容器未重启、旧配置仍生效」的问题。
  if [ "$NO_BUILD" -eq 0 ]; then
    $COMPOSE up -d --build --force-recreate
  else
    $COMPOSE up -d --force-recreate
  fi
fi

echo "==> 当前状态："
$COMPOSE ps 2>/dev/null || docker ps --filter "name=doc-audit"

# 自动探测本机 IP（IPv4），全部列出；取不到时回退 localhost。
echo "==> 访问地址："
IPS="$( (hostname -I 2>/dev/null || ip -4 -o addr show scope global 2>/dev/null || true) \
  | tr ' ' '\n' | grep -oE '([0-9]{1,3}\.){3}[0-9]{1,3}' | grep -v '^127\.' | sort -u )"
if [ -n "$IPS" ]; then
  echo "$IPS" | while read -r ip; do
    echo "    http://${ip}:${PORT}"
  done
else
  echo "    http://localhost:${PORT}"
fi
