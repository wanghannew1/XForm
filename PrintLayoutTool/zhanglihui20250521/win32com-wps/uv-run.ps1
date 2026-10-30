# uv-run.ps1 - 一键：安装 uv / 创建虚拟环境 / 安装 pip 依赖 / 运行 Python 脚本
# 用法：在资源管理器右键 → “使用 PowerShell 运行” 本脚本，或直接执行 .\uv-run.ps1
$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Push-Location $scriptDir | Out-Null

$pyScript   = "XFormz_zlh_v7.py"
$reqFile    = "requirements.txt"
$venvDir    = Join-Path $scriptDir ".venv"
$pypiIndex  = "https://pypi.tuna.tsinghua.edu.cn/simple"
$uvInstall  = "https://astral.sh/uv/install.ps1"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  XForm uv 一键运行脚本" -ForegroundColor Cyan
Write-Host "  工作目录: $scriptDir" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 1. 确保 uv 已安装
$uvInstalled = $false
$uvCmd = Get-Command uv 2>$null
if ($uvCmd) {
    $uvInstalled = $true
    Write-Host "[1/4] uv 已安装。" -ForegroundColor Green
}

if (-not $uvInstalled) {
    Write-Host "`n[1/4] 未检测到 uv，正在安装..." -ForegroundColor Yellow
    try {
        iex (iwr $uvInstall -UseBasicParsing)
    } catch {
        Write-Host "uv 安装失败：$_" -ForegroundColor Red
        exit 1
    }
    # 刷新当前会话 PATH，使本次运行可直接使用 uv
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    $machinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $env:PATH = "$userPath;$machinePath"
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        Write-Host "uv 安装后仍未生效，请重新打开 PowerShell 后再运行本脚本。" -ForegroundColor Red
        exit 1
    }
    Write-Host "uv 安装完成。" -ForegroundColor Green
}

# 2. 创建虚拟环境（如不存在）
if (-not (Test-Path $venvDir)) {
    Write-Host "`n[2/4] 正在创建虚拟环境 .venv ..." -ForegroundColor Yellow
    uv venv
    Write-Host "虚拟环境创建完成。" -ForegroundColor Green
} else {
    Write-Host "[2/4] 虚拟环境 .venv 已存在，跳过创建。" -ForegroundColor Green
}

# 3. 安装 pip 依赖（如有 requirements.txt，走清华镜像）
if (Test-Path $reqFile) {
    Write-Host "`n[3/4] 正在安装 pip 依赖（清华镜像加速）..." -ForegroundColor Yellow
    uv pip install -r $reqFile -i $pypiIndex
    Write-Host "依赖安装完成。" -ForegroundColor Green
} else {
    Write-Host "[3/4] 未找到 requirements.txt，跳过依赖安装。" -ForegroundColor Yellow
}

# 4. 运行 Python 脚本
Write-Host "`n[4/4] 正在运行 $pyScript ..." -ForegroundColor Cyan
uv run .\$pyScript

Pop-Location | Out-Null
