# 下载 Chrome for Testing（linux64）到 vendor/ 目录，供 Dockerfile build 使用。
# 详见 scripts/fetch-vendor-chrome.sh 的说明。
#
# 使用：
#   powershell -ExecutionPolicy Bypass -File scripts\fetch-vendor-chrome.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\fetch-vendor-chrome.ps1 -ChromeVersion 147.0.7727.57

param(
    [string]$ChromeVersion = "147.0.7727.57"
)

$ErrorActionPreference = "Stop"

$mirrors = @(
    "https://cdn.npmmirror.com/binaries/chrome-for-testing/$ChromeVersion/linux64/chrome-linux64.zip",
    "https://storage.googleapis.com/chrome-for-testing-public/$ChromeVersion/linux64/chrome-linux64.zip"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vendorDir = Join-Path $scriptDir "..\vendor"
$output = Join-Path $vendorDir "chrome-linux64.zip"

New-Item -ItemType Directory -Force -Path $vendorDir | Out-Null

if (Test-Path $output) {
    $size = (Get-Item $output).Length
    if ($size -gt 150MB) {
        Write-Host "[fetch-vendor-chrome] $output 已存在 ($([int]($size/1MB)) MB)，跳过下载。"
        Write-Host "[fetch-vendor-chrome] 如需重新下载请先删除。"
        exit 0
    } else {
        Write-Host "[fetch-vendor-chrome] $output 存在但大小异常 ($size B)，删除后重新下载..."
        Remove-Item $output
    }
}

foreach ($url in $mirrors) {
    Write-Host "[fetch-vendor-chrome] 尝试下载: $url"
    try {
        # curl.exe 自带于 Win10+，比 Invoke-WebRequest 快得多（后者对大文件极慢）
        & curl.exe -fL --retry 5 --retry-all-errors --retry-delay 3 -C - -o $output $url
        if ($LASTEXITCODE -eq 0) {
            $size = (Get-Item $output).Length
            if ($size -gt 150MB) {
                Write-Host "[fetch-vendor-chrome] 下载成功: $output ($([int]($size/1MB)) MB)"
                exit 0
            }
            Write-Host "[fetch-vendor-chrome] 文件大小异常 ($size B)，换下一个源..."
            Remove-Item $output -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "[fetch-vendor-chrome] 下载失败: $_"
    }
}

Write-Host "[fetch-vendor-chrome] 所有下载源均失败。请检查网络或手动下载：" -ForegroundColor Red
foreach ($url in $mirrors) { Write-Host "  - $url" -ForegroundColor Red }
Write-Host "  并保存为: $output" -ForegroundColor Red
exit 1
