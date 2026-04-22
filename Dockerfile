# 基础镜像换到 bookworm-slim：Puppeteer 对 glibc 环境依赖（Alpine musl 兼容性差）
FROM node:20-bookworm-slim

# apt 源保持 Debian 官方 deb.debian.org（含 CDN）：
#   曾尝试华为云/阿里云/USTC 等国内 mirror，对 chromium/fonts-noto-cjk 等大包频繁 EOF，
#   甚至小包也不稳定（502/500）。现已不依赖大包，官方源在国内虽略慢但对小包更稳定。
#
# apt 网络策略：重试、禁用流水线（缓解偶发 EOF）
RUN set -eux; \
    { \
      echo 'Acquire::Retries "5";'; \
      echo 'Acquire::http::Timeout "120";'; \
      echo 'Acquire::Queue-Mode "access";'; \
      echo 'Acquire::http::Pipeline-Depth "0";'; \
      echo 'Acquire::http::No-Cache "true";'; \
    } > /etc/apt/apt.conf.d/80retries

# 层 1：Chrome 运行时依赖（不安装 Debian 的 chromium 包，也不安装 fonts-noto-cjk 大包）
#
# 方案演进：最初尝试 `apt install chromium + fonts-noto-cjk`，但国内 mirror 对 chromium
# ~120MB 的包频繁 500 EOF / 502 Bad Gateway，build 体验极差。现改为：
#   - Chrome 由 puppeteer 在 pnpm install 时从 npmmirror 下载（走国内 CDN，稳定快速）；
#   - 中文字体由本仓库 fonts/ 目录覆盖（已包含 simsun/simhei/wqy 等）；
#   - 此处只装 Chrome 运行需要的小系统库（都是 <1MB 的小包，单次抖动 apt 重试即可）。
#
# 套壳的 retry 循环兜底 mirror 偶发 502。
RUN set -ux; \
    PKGS="unzip \
          libnss3 libnspr4 \
          libatk1.0-0 libatk-bridge2.0-0 libatspi2.0-0 \
          libcups2 libdbus-1-3 libdrm2 libexpat1 \
          libxkbcommon0 libxcomposite1 libxdamage1 libxrandr2 \
          libxfixes3 libxext6 libxcb1 libx11-6 libx11-xcb1 \
          libgbm1 libxss1 libasound2 libglib2.0-0 \
          libpangocairo-1.0-0 libpango-1.0-0 libcairo2 \
          ca-certificates fontconfig \
          fonts-liberation fonts-wqy-zenhei fonts-wqy-microhei"; \
    ok=0; \
    for i in 1 2 3 4 5; do \
      echo "[apt] attempt $i"; \
      apt-get update; \
      if apt-get install -y --no-install-recommends $PKGS; then \
        ok=1; break; \
      fi; \
      echo "[apt] attempt $i failed, sleeping 10s..."; \
      sleep 10; \
    done; \
    if [ "$ok" != "1" ]; then echo "apt-get install failed after 5 attempts"; exit 1; fi; \
    fc-cache -f; \
    rm -rf /var/lib/apt/lists/*

# 层 2：自定义字体（公文专用 simsun/方正/等，~650MB）
# 放到 /usr/share/fonts/custom/ 由 fontconfig 自动发现
COPY fonts/ /usr/share/fonts/custom/
RUN fc-cache -f

# Puppeteer 配置：
#   - 不让 puppeteer 在 pnpm install 时自动下载 Chrome（容器网络访问 CDN 大包极不稳定，
#     多次尝试 cdn.npmmirror.com 均在 10MB 左右被 aborted）；
#   - 改由宿主机预下载 chrome-linux64.zip 并 COPY 进镜像，稳定且可离线 build；
#   - PUPPETEER_CACHE_DIR 指定 puppeteer-core 运行时查找 Chrome 的位置。
ENV PUPPETEER_SKIP_DOWNLOAD=true \
    PUPPETEER_CACHE_DIR=/home/node/.cache/puppeteer

# 层 3a：预置 Chrome for Testing（由宿主机下载，见 vendor/README）
# 目录结构必须与 @puppeteer/browsers 约定一致：
#   $PUPPETEER_CACHE_DIR/chrome/linux-<version>/chrome-linux64/chrome
ARG CHROME_VERSION=147.0.7727.57
COPY vendor/chrome-linux64.zip /tmp/chrome.zip
RUN set -eux; \
    mkdir -p "$PUPPETEER_CACHE_DIR/chrome/linux-${CHROME_VERSION}"; \
    unzip -q /tmp/chrome.zip -d "$PUPPETEER_CACHE_DIR/chrome/linux-${CHROME_VERSION}/"; \
    rm /tmp/chrome.zip; \
    chmod +x "$PUPPETEER_CACHE_DIR/chrome/linux-${CHROME_VERSION}/chrome-linux64/chrome"; \
    "$PUPPETEER_CACHE_DIR/chrome/linux-${CHROME_VERSION}/chrome-linux64/chrome" --version

# 层 3b：Node 依赖（PUPPETEER_SKIP_DOWNLOAD=true，不会触发 Chrome 下载）
# corepack 走国内 npm 镜像（默认 npmjs.com 在容器内不稳定；npmmirror 偶发超时需 retry）
ENV COREPACK_NPM_REGISTRY=https://registry.npmmirror.com \
    npm_config_registry=https://registry.npmmirror.com
RUN set -ux; \
    for i in 1 2 3 4 5; do \
      echo "[corepack] attempt $i"; \
      if corepack enable && corepack prepare pnpm@latest --activate; then \
        exit 0; \
      fi; \
      echo "[corepack] attempt $i failed, sleeping 5s..."; \
      sleep 5; \
    done; \
    echo "corepack failed after 5 attempts"; exit 1
WORKDIR /app
COPY package.json pnpm-lock.yaml ./
RUN pnpm install --frozen-lockfile --prod

# 层 4：应用代码（高频变动）
COPY dist/ ./dist/

EXPOSE 3002

CMD ["node", "dist/main.js"]
