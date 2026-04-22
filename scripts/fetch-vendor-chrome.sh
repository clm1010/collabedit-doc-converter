#!/bin/bash
# 下载 Chrome for Testing（linux64）到 vendor/ 目录，供 Dockerfile build 使用。
#
# 为什么需要预下载？
#   Docker build 过程中，容器访问 CDN（storage.googleapis.com / cdn.npmmirror.com）
#   下载 ~170MB 大包极不稳定，常在中途被 aborted / EOF。改为宿主机预下载 + COPY 进镜像，
#   build 完全离线可重复，速度快且稳定。
#
# 使用：
#   bash scripts/fetch-vendor-chrome.sh                  # 用默认版本
#   CHROME_VERSION=147.0.7727.57 bash scripts/fetch-vendor-chrome.sh
#
# 版本约束：必须与 Puppeteer 自带的 Chrome for Testing 版本一致，否则运行时会报
# "Could not find Chrome (ver. X.Y.Z.W)"。Puppeteer 升级时更新下面的 DEFAULT_VERSION。

set -euo pipefail

# 与 package.json 的 puppeteer 版本一一对应：
#   puppeteer ^24.42.0  →  Chrome for Testing 147.0.7727.57
DEFAULT_VERSION=147.0.7727.57
CHROME_VERSION="${CHROME_VERSION:-$DEFAULT_VERSION}"

# 下载源优先级：国内镜像 → 官方源
MIRRORS=(
  "https://cdn.npmmirror.com/binaries/chrome-for-testing/${CHROME_VERSION}/linux64/chrome-linux64.zip"
  "https://storage.googleapis.com/chrome-for-testing-public/${CHROME_VERSION}/linux64/chrome-linux64.zip"
)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENDOR_DIR="${SCRIPT_DIR}/../vendor"
OUTPUT="${VENDOR_DIR}/chrome-linux64.zip"

mkdir -p "$VENDOR_DIR"

# 已存在且大小正常则跳过（171MB 左右）
if [ -f "$OUTPUT" ]; then
  SIZE=$(wc -c < "$OUTPUT")
  if [ "$SIZE" -gt 150000000 ]; then
    echo "[fetch-vendor-chrome] $OUTPUT 已存在 ($(($SIZE/1024/1024)) MB)，跳过下载。"
    echo "[fetch-vendor-chrome] 如需重新下载请先删除。"
    exit 0
  else
    echo "[fetch-vendor-chrome] $OUTPUT 存在但大小异常 (${SIZE}B)，删除后重新下载..."
    rm -f "$OUTPUT"
  fi
fi

for url in "${MIRRORS[@]}"; do
  echo "[fetch-vendor-chrome] 尝试下载: $url"
  if curl -fL --retry 5 --retry-all-errors --retry-delay 3 -C - -o "$OUTPUT" "$url"; then
    SIZE=$(wc -c < "$OUTPUT")
    if [ "$SIZE" -gt 150000000 ]; then
      echo "[fetch-vendor-chrome] 下载成功: $OUTPUT ($(($SIZE/1024/1024)) MB)"
      exit 0
    fi
    echo "[fetch-vendor-chrome] 文件大小异常 (${SIZE}B)，换下一个源..."
    rm -f "$OUTPUT"
  fi
done

echo "[fetch-vendor-chrome] 所有下载源均失败。请检查网络或手动下载：" >&2
for url in "${MIRRORS[@]}"; do echo "  - $url" >&2; done
echo "  并保存为: $OUTPUT" >&2
exit 1
