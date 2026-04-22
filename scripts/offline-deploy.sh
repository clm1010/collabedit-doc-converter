#!/bin/bash
# 离线部署脚本（Puppeteer 版）
#
# 使用场景：
#   生产机没有外网，在有网的"打包机"上 build 好镜像 + 保存 tar，
#   然后把 tar 包传到生产机 docker load 并 docker compose up。
#
# 流程：
#   【打包机】bash scripts/offline-deploy.sh save
#   【复制到生产机】doc-converter-images.tar + docker-compose.yml + .env + fonts/（如不在镜像内）
#   【生产机】bash scripts/offline-deploy.sh load

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="${SCRIPT_DIR}/.."
IMAGES_FILE="doc-converter-images.tar"

cd "$PROJECT_DIR"

case "${1:-}" in
  save)
    echo "=== [1/4] 下载预置 Chrome for Testing 到 vendor/ ==="
    bash "$SCRIPT_DIR/fetch-vendor-chrome.sh"

    echo "=== [2/4] 本地 pnpm build 生成 dist/ ==="
    if ! command -v pnpm >/dev/null 2>&1; then
      echo "需要 pnpm。请先 corepack enable && corepack prepare pnpm@latest --activate" >&2
      exit 1
    fi
    pnpm install --frozen-lockfile
    pnpm build

    echo "=== [3/4] 拉基础镜像 + 构建 converter 镜像 ==="
    docker pull node:20-bookworm-slim
    docker compose build converter

    echo "=== [4/4] 导出镜像到 ${IMAGES_FILE} ==="
    docker save \
      node:20-bookworm-slim \
      collabedit-doc-converter-converter:latest \
      -o "${IMAGES_FILE}"

    echo ""
    echo "完成。请将以下文件复制到离线生产机："
    echo "  - ${IMAGES_FILE} （镜像 tar 包）"
    echo "  - docker-compose.yml"
    echo "  - .env  （由 .env.example 复制并按需调整）"
    echo "  - fonts/  （如生产机通过 volume 挂载字体；若已内置镜像可省）"
    ;;

  load)
    echo "=== [1/2] 导入镜像 ==="
    if [ ! -f "${IMAGES_FILE}" ]; then
      echo "找不到 ${IMAGES_FILE}，请确认已从打包机复制过来。" >&2
      exit 1
    fi
    docker load -i "${IMAGES_FILE}"

    echo "=== [2/2] 启动服务 ==="
    docker compose up -d

    echo ""
    echo "完成。查看状态：docker compose ps"
    echo "查看日志：     docker compose logs -f converter"
    echo "健康检查：     curl http://localhost:\${PORT:-3002}/health"
    ;;

  *)
    echo "使用方式: $0 {save|load}"
    echo "  save  - 在有网的打包机上：下载 Chrome + 构建镜像 + 保存 tar"
    echo "  load  - 在离线生产机上：导入 tar + 启动服务"
    exit 1
    ;;
esac
