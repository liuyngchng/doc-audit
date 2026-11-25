#!/usr/bin/env bash
# 构建 doc-audit Docker 镜像。
#
# 用法：
#   ./build.sh             # 使用 buildinfo.json 里的 version 作为 tag
#   ./build.sh 1.2.3       # 手动指定 tag
#   VERSION=1.2.3 ./build.sh

set -euo pipefail

cd "$(dirname "$0")"

IMAGE_NAME="${IMAGE_NAME:-doc-audit}"

# tag 优先级：命令行参数 > 环境变量 VERSION > buildinfo.json 里的 version > latest
VERSION="${1:-${VERSION:-}}"
if [ -z "$VERSION" ]; then
  VERSION="$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' buildinfo.json | head -n1)"
fi
if [ -z "$VERSION" ]; then
  VERSION="latest"
fi

echo "==> 构建镜像 ${IMAGE_NAME}:${VERSION}"
docker build -t "${IMAGE_NAME}:${VERSION}" -t "${IMAGE_NAME}:latest" .

echo "==> 构建完成"
docker images "${IMAGE_NAME}" | head -n 5
