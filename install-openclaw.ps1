<#
.SYNOPSIS
  OpenClaw 一键安装脚本 (Windows)
.DESCRIPTION
  自动检测并安装 Node.js v22+、Git，然后安装并配置 OpenClaw。
.NOTES
  用法:
    powershell -ExecutionPolicy Bypass -File install-openclaw.ps1
  在线一键安装（推荐，显式 UTF-8 编码，兼容 PowerShell 5.x）:
    powershell -c "$w=New-Object Net.WebClient;$w.Encoding=[Text.Encoding]::UTF8;iex $w.DownloadString('https://www.codefather.cn/openclaw_install/install-openclaw.ps1')"
  在线一键安装（简短，需服务端返回 charset=utf-8 或文件含 BOM）:
    irm https://www.clawfather.cn/install-openclaw.ps1 | iex
#>

# ── 强制 UTF-8 编码（解决中文乱码）──
try { $null = & cmd /c "chcp 65001 >nul 2>&1" } catch {}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Stop"

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "  [FAIL] 需要 PowerShell 5.0 或更高版本，当前版本: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    exit 1
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ── 颜色输出 ──

function Write-Info    { param($Msg) Write-Host "  [INFO] " -ForegroundColor Blue -NoNewline; Write-Host $Msg }
function Write-Ok      { param($Msg) Write-Host "  [OK]   " -ForegroundColor Green -NoNewline; Write-Host $Msg }
function Write-Warn    { param($Msg) Write-Host "  [WARN] " -ForegroundColor Yellow -NoNewline; Write-Host $Msg }
function Write-Err     { param($Msg) Write-Host "  [FAIL] " -ForegroundColor Red -NoNewline; Write-Host $Msg }
function Write-Step    { param($Msg) Write-Host "`n━━━ $Msg ━━━`n" -ForegroundColor Cyan }

# ── 全局变量 ──

$script:NodeBinDir = $null
$script:RequiredNodeMajor = 22
$script:Arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "x64" }

function Get-LocalAppData {
    if ($env:LOCALAPPDATA) { return $env:LOCALAPPDATA }
    return (Join-Path $HOME "AppData\Local")
}

# ── 工具函数 ──

function Refresh-PathEnv {
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH = "$machinePath;$userPath"
}

function Add-ToUserPath {
    param([string]$Dir)
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    if ($currentPath -notlike "*$Dir*") {
        [Environment]::SetEnvironmentVariable("PATH", "$Dir;$currentPath", "User")
        $env:PATH = "$Dir;$env:PATH"
        Write-Info "已将 $Dir 添加到用户 PATH"
    }
}

function Get-NodeVersion {
    param([string]$NodeExe = "node")
    try {
        $output = & $NodeExe -v 2>$null
        if ($output -match "v(\d+)") {
            $major = [int]$Matches[1]
            if ($major -ge $script:RequiredNodeMajor) {
                return $output.Trim()
            }
        }
    } catch {}
    return $null
}

function Pin-NodePath {
    foreach ($dir in $env:PATH.Split(";")) {
        if (-not $dir) { continue }
        $nodeExe = Join-Path $dir "node.exe"
        if (Test-Path $nodeExe) {
            try {
                $output = & $nodeExe -v 2>$null
                if ($output -match "v(\d+)" -and [int]$Matches[1] -ge $script:RequiredNodeMajor) {
                    $script:NodeBinDir = $dir
                    $rest = ($env:PATH.Split(";") | Where-Object { $_ -ne $dir }) -join ";"
                    $env:PATH = "$dir;$rest"
                    Write-Info "锁定 Node.js v22 路径: $dir"
                    return
                }
            } catch {}
        }
    }
}

function Get-NpmCmd {
    if ($script:NodeBinDir) {
        $cmd = Join-Path $script:NodeBinDir "npm.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    return "npm"
}

function Get-PnpmCmd {
    if ($script:NodeBinDir) {
        $cmd = Join-Path $script:NodeBinDir "pnpm.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    $defaultPnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    $cmd = Join-Path $defaultPnpmHome "pnpm.cmd"
    if (Test-Path $cmd) { return $cmd }
    try {
        $resolved = (Get-Command pnpm.cmd -ErrorAction Stop).Source
        if (Test-Path $resolved) { return $resolved }
    } catch {}
    return "pnpm.cmd"
}

function Get-OpenclawCmd {
    # 优先查找 PNPM_HOME
    if ($env:PNPM_HOME) {
        $cmd = Join-Path $env:PNPM_HOME "openclaw.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    # 查找 NodeBinDir
    if ($script:NodeBinDir) {
        $cmd = Join-Path $script:NodeBinDir "openclaw.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    # 查找默认 pnpm 全局路径
    $defaultPnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    $cmd = Join-Path $defaultPnpmHome "openclaw.cmd"
    if (Test-Path $cmd) { return $cmd }
    return "openclaw"
}

function Ensure-PnpmHome {
    $pnpmHome = $env:PNPM_HOME
    if (-not $pnpmHome) {
        $pnpmHome = [Environment]::GetEnvironmentVariable("PNPM_HOME", "User")
    }
    if (-not $pnpmHome) {
        $pnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    }

    $env:PNPM_HOME = $pnpmHome
    if ($env:PATH -notlike "*$pnpmHome*") { $env:PATH = "$pnpmHome;$env:PATH" }

    $savedHome = [Environment]::GetEnvironmentVariable("PNPM_HOME", "User")
    if ($savedHome -ne $pnpmHome) {
        [Environment]::SetEnvironmentVariable("PNPM_HOME", $pnpmHome, "User")
        Write-Info "已持久化 PNPM_HOME=$pnpmHome"
    }

    Add-ToUserPath $pnpmHome
}

function Download-File {
    param([string]$Dest, [string[]]$Urls)
    foreach ($url in $Urls) {
        $hostName = ([Uri]$url).Host
        Write-Info "正在从 $hostName 下载..."
        try {
            Invoke-WebRequest -Uri $url -OutFile $Dest -UseBasicParsing -TimeoutSec 300
            Write-Ok "下载完成"
            return $true
        } catch {
            Write-Warn "从 $hostName 下载失败，尝试备用源..."
        }
    }
    return $false
}

function Get-LatestNodeVersion {
    param([int]$Major)
    $urls = @(
        "https://npmmirror.com/mirrors/node/latest-v${Major}.x/SHASUMS256.txt",
        "https://nodejs.org/dist/latest-v${Major}.x/SHASUMS256.txt"
    )
    foreach ($url in $urls) {
        try {
            $content = (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15).Content
            if ($content -match "node-(v\d+\.\d+\.\d+)") {
                return $Matches[1]
            }
        } catch {}
    }
    return $null
}

# ── 安装 Node.js ──

function Install-NodeViaNvm {
    try {
        $nvmOut = & cmd /c "nvm version" 2>$null
        if (-not $nvmOut) { return $false }
    } catch { return $false }

    Write-Info "检测到 nvm，正在使用 nvm 安装 Node.js v22..."
    try { & cmd /c "nvm node_mirror https://npmmirror.com/mirrors/node/" 2>$null } catch {}

    try {
        & cmd /c "nvm install 22" 2>$null
        & cmd /c "nvm use 22" 2>$null
        Refresh-PathEnv

        $ver = Get-NodeVersion
        if ($ver) {
            Write-Ok "Node.js $ver 已通过 nvm 安装"
            return $true
        }
    } catch {
        Write-Warn "nvm 安装 Node.js 失败: $_"
    }
    return $false
}

function Install-NodeDirect {
    Write-Info "正在直接下载安装 Node.js v22..."

    $version = Get-LatestNodeVersion -Major 22
    if (-not $version) {
        Write-Err "无法获取 Node.js 版本信息，请检查网络连接"
        return $false
    }
    Write-Info "最新 LTS 版本: $version"

    $filename = "node-$version-win-$($script:Arch).zip"
    $tmpPath = Join-Path $env:TEMP "openclaw-install"
    $tmpFile = Join-Path $tmpPath $filename
    $extractedName = "node-$version-win-$($script:Arch)"
    $installDir = Join-Path (Get-LocalAppData) "nodejs"

    New-Item -ItemType Directory -Force -Path $tmpPath | Out-Null

    $downloaded = Download-File -Dest $tmpFile -Urls @(
        "https://npmmirror.com/mirrors/node/$version/$filename",
        "https://nodejs.org/dist/$version/$filename"
    )

    if (-not $downloaded) {
        Write-Err "Node.js 下载失败，请检查网络连接"
        return $false
    }

    try {
        Write-Info "正在解压安装..."
        Expand-Archive -Path $tmpFile -DestinationPath $tmpPath -Force
        if (Test-Path $installDir) { Remove-Item $installDir -Recurse -Force }
        Move-Item (Join-Path $tmpPath $extractedName) $installDir

        $env:PATH = "$installDir;$env:PATH"
        Add-ToUserPath $installDir
    } catch {
        Write-Err "安装失败: $_"
        Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue

    $ver = Get-NodeVersion
    if ($ver) {
        Write-Ok "Node.js $ver 安装成功"
        return $true
    }
    Write-Warn "Node.js 安装完成但验证失败"
    return $false
}

# ── 安装 Git ──

function Get-GitVersion {
    try {
        $output = & git --version 2>$null
        return $output.Trim()
    } catch {}

    $gitPaths = @(
        (Join-Path $env:ProgramFiles "Git\cmd"),
        (Join-Path ${env:ProgramFiles(x86)} "Git\cmd")
    )
    foreach ($gp in $gitPaths) {
        $gitExe = Join-Path $gp "git.exe"
        if (Test-Path $gitExe) {
            try {
                $output = & $gitExe --version 2>$null
                if (-not ($env:PATH -like "*$gp*")) { $env:PATH = "$gp;$env:PATH" }
                return $output.Trim()
            } catch {}
        }
    }
    return $null
}

function Get-LatestGitRelease {
    $url = "https://registry.npmmirror.com/-/binary/git-for-windows/"
    try {
        $content = (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15).Content
        $regexMatches = [regex]::Matches($content, "v(\d+)\.(\d+)\.(\d+)\.windows\.(\d+)/")
        if ($regexMatches.Count -eq 0) { return $null }

        $best = $regexMatches | Sort-Object {
            [int]$_.Groups[1].Value * 1000000 + [int]$_.Groups[2].Value * 10000 +
            [int]$_.Groups[3].Value * 100 + [int]$_.Groups[4].Value
        } -Descending | Select-Object -First 1

        $version = "$($best.Groups[1].Value).$($best.Groups[2].Value).$($best.Groups[3].Value)"
        $winBuild = $best.Groups[4].Value
        $tag = "v$version.windows.$winBuild"
        $fileVersion = if ($winBuild -eq "1") { $version } else { "$version.$winBuild" }
        return @{ Version = $version; Tag = $tag; FileVersion = $fileVersion }
    } catch {}

    try {
        $ghUrl = "https://api.github.com/repos/git-for-windows/git/releases/latest"
        $content = (Invoke-WebRequest -Uri $ghUrl -UseBasicParsing -TimeoutSec 15).Content
        if ($content -match '\x22tag_name\x22\s*:\s*\x22(v(\d+\.\d+\.\d+)\.windows\.(\d+))\x22') {
            $version = $Matches[2]; $winBuild = $Matches[3]; $tag = $Matches[1]
            $fileVersion = if ($winBuild -eq "1") { $version } else { "$version.$winBuild" }
            return @{ Version = $version; Tag = $tag; FileVersion = $fileVersion }
        }
    } catch {}
    return $null
}

function Install-GitViaWinget {
    try {
        Get-Command winget -ErrorAction Stop | Out-Null
    } catch { return $false }

    Write-Info "检测到 winget，正在安装 Git..."
    try {
        & winget install --id Git.Git -e --source winget --silent --accept-package-agreements --accept-source-agreements 2>$null
    } catch {
        Write-Warn "winget 命令返回了非零退出码，检查 Git 是否已可用..."
    }

    Refresh-PathEnv
    $ver = Get-GitVersion
    if ($ver) {
        Write-Ok "$ver 已可用"
        return $true
    }
    Write-Warn "winget 安装后仍未检测到 Git"
    return $false
}

function Install-GitDirect {
    Write-Info "正在下载 Git for Windows..."

    $release = Get-LatestGitRelease
    if (-not $release) {
        Write-Err "无法获取 Git 版本信息，请检查网络连接"
        return $false
    }
    Write-Info "最新版本: Git $($release.FileVersion)"

    $archStr = if ($script:Arch -eq "arm64") { "arm64" } else { "64-bit" }
    $filename = "Git-$($release.FileVersion)-$archStr.exe"
    $tmpPath = Join-Path $env:TEMP "openclaw-install"
    $tmpFile = Join-Path $tmpPath $filename

    New-Item -ItemType Directory -Force -Path $tmpPath | Out-Null

    $downloaded = Download-File -Dest $tmpFile -Urls @(
        "https://registry.npmmirror.com/-/binary/git-for-windows/$($release.Tag)/$filename",
        "https://github.com/git-for-windows/git/releases/download/$($release.Tag)/$filename"
    )

    if (-not $downloaded) {
        Write-Err "Git 下载失败，请检查网络连接"
        Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    Write-Info "正在静默安装 Git..."
    try {
        Start-Process -FilePath $tmpFile -ArgumentList "/VERYSILENT","/NORESTART","/NOCANCEL","/SP-","/CLOSEAPPLICATIONS","/RESTARTAPPLICATIONS" -Wait
        Refresh-PathEnv

        $ver = Get-GitVersion
        if ($ver) {
            Write-Ok "$ver 安装成功"
            Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
            return $true
        }
    } catch {
        Write-Err "Git 安装失败: $_"
    }

    Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
    return $false
}

# ── 主流程步骤 ──

function Step-CheckNode {
    Write-Step "步骤 1/7: 准备 Node.js 环境"

    # 先检查脚本之前安装过的路径是否已有合格版本
    $scriptInstallDir = Join-Path (Get-LocalAppData) "nodejs"
    $scriptNodeExe = Join-Path $scriptInstallDir "node.exe"
    if (Test-Path $scriptNodeExe) {
        $ver = Get-NodeVersion -NodeExe $scriptNodeExe
        if ($ver) {
            Write-Ok "Node.js $ver 已安装，版本满足要求 (>= 22)"
            $script:NodeBinDir = $scriptInstallDir
            $rest = ($env:PATH.Split(";") | Where-Object { $_ -ne $scriptInstallDir }) -join ";"
            $env:PATH = "$scriptInstallDir;$rest"
            Add-ToUserPath $scriptInstallDir
            return $true
        }
    }

    # 再从 PATH 中查找合格版本
    $ver = Get-NodeVersion
    if ($ver) {
        Write-Ok "Node.js $ver 已安装，版本满足要求 (>= 22)"
        Pin-NodePath
        return $true
    }

    $existingVer = try { & node -v 2>$null } catch { $null }
    if ($existingVer) {
        Write-Warn "检测到 Node.js $existingVer，版本过低，需要 v22 以上"
    } else {
        Write-Warn "未检测到 Node.js"
    }

    Write-Info "正在自动安装 Node.js v22..."

    if (Install-NodeViaNvm) { Pin-NodePath; return $true }
    if (Install-NodeDirect) { Pin-NodePath; return $true }

    Write-Err "所有安装方式均失败，请检查网络连接后重试"
    return $false
}

function Step-CheckGit {
    Write-Step "步骤 2/7: 准备 Git 环境"

    $ver = Get-GitVersion
    if ($ver) {
        Write-Ok "$ver 已安装"
        return $true
    }

    Write-Warn "未检测到 Git，正在自动安装..."

    if (Install-GitViaWinget) { return $true }
    if (Install-GitDirect) { return $true }

    Write-Err "Git 自动安装失败，请手动安装 Git 后重试"
    Write-Host "  下载地址: https://git-scm.com/downloads"
    return $false
}

function Step-SetMirror {
    Write-Step "步骤 3/7: 设置国内 npm 镜像"

    $npmCmd = Get-NpmCmd
    try {
        & $npmCmd config set registry https://registry.npmmirror.com 2>$null
        Write-Ok "npm 镜像已设置为 https://registry.npmmirror.com"
        return $true
    } catch {
        Write-Err "设置 npm 镜像失败: $_"
        return $false
    }
}

function Step-InstallPnpm {
    Write-Step "步骤 4/7: 安装 pnpm"

    $pnpmCmd = Get-PnpmCmd
    try {
        $pnpmVer = (& $pnpmCmd -v 2>$null).Trim()
        if ($pnpmVer) {
            Write-Ok "pnpm $pnpmVer 已安装，跳过安装步骤"
            Ensure-PnpmHome
            $currentRegistry = try { (& $pnpmCmd config get registry 2>$null).Trim() } catch { "" }
            if ($currentRegistry -notlike "*npmmirror*") {
                try {
                    & $pnpmCmd config set registry https://registry.npmmirror.com 2>$null
                    Write-Ok "pnpm 镜像已设置为 https://registry.npmmirror.com"
                } catch {
                    Write-Warn "设置 pnpm 镜像失败，将使用默认源"
                }
            }
            return $true
        }
    } catch {}

    $npmCmd = Get-NpmCmd
    Write-Info "正在安装 pnpm..."
    try {
        & $npmCmd install -g pnpm
        $pnpmCmd = Get-PnpmCmd
        Write-Info "正在验证 pnpm 安装..."
        $pnpmVer = (& $pnpmCmd -v 2>$null).Trim()
        Write-Ok "pnpm $pnpmVer 安装成功"

        Write-Info "正在配置 pnpm 全局路径 (pnpm setup)..."
        try { & $pnpmCmd setup 2>$null } catch { Write-Warn "pnpm setup 执行未成功，不影响后续安装" }

        try {
            & $pnpmCmd config set registry https://registry.npmmirror.com 2>$null
            Write-Ok "pnpm 镜像已设置为 https://registry.npmmirror.com"
        } catch {
            Write-Warn "设置 pnpm 镜像失败，将使用默认源"
        }

        Ensure-PnpmHome
        return $true
    } catch {
        Write-Err "pnpm 安装失败: $_"
        return $false
    }
}

function Run-PnpmInstall {
    param([string]$PnpmCmd, [string]$Label = "安装")

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$PnpmCmd`" add -g openclaw@latest"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
    } catch {
        Write-Err "启动${Label}进程失败: $_"
        return @{ Success = $false; Stderr = ""; Stdout = "" }
    }

    $progress = 0
    $width = 30
    while (-not $proc.HasExited) {
        if ($progress -lt 30) { $progress += 3 }
        elseif ($progress -lt 60) { $progress += 2 }
        elseif ($progress -lt 90) { $progress += 1 }
        if ($progress -gt 90) { $progress = 90 }
        $filled = [math]::Floor($progress * $width / 100)
        $empty = $width - $filled
        $bar = ([string]::new([char]0x2588, $filled)) + ([string]::new([char]0x2591, $empty))
        Write-Host "`r  ${Label}进度 [$bar] $($progress.ToString().PadLeft(3))%" -NoNewline
        Start-Sleep -Seconds 1
    }

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    try { $null = & cmd /c "chcp 65001 >nul 2>&1" } catch {}
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    $fullBar = [string]::new([char]0x2588, $width)
    if ($proc.ExitCode -eq 0) {
        Write-Host "`r  ${Label}进度 [$fullBar] 100%"
        return @{ Success = $true; Stderr = $stderr; Stdout = $stdout }
    }

    Write-Host "`r  ${Label}进度 [$fullBar] 失败"
    return @{ Success = $false; Stderr = $stderr; Stdout = $stdout; ExitCode = $proc.ExitCode }
}

function Step-InstallOpenClaw {
    Write-Step "步骤 5/7: 安装 OpenClaw"
    Write-Info "正在安装 OpenClaw，请耐心等待..."

    $pnpmCmd = Get-PnpmCmd
    if (-not (Test-Path $pnpmCmd -ErrorAction SilentlyContinue)) {
        try { Get-Command $pnpmCmd -ErrorAction Stop | Out-Null } catch {
            Write-Err "找不到 pnpm 命令"
            return $false
        }
    }

    try {
        & git config --global --add "url.https://github.com/.insteadOf" "git+ssh://git@github.com/" 2>$null
        & git config --global --add "url.https://github.com/.insteadOf" "ssh://git@github.com/" 2>$null
    } catch {}

    function Clear-GitMirror([string]$Mirror) {
        if ($Mirror) {
            try { & git config --global --unset-all "url.${Mirror}.insteadOf" 2>$null } catch {}
        }
    }

    function Set-GitMirror([string]$Mirror) {
        & git config --global "url.${Mirror}.insteadOf" "https://github.com/" 2>$null
    }

    function Try-InstallWithCleanup([string]$PnpmCmd, [ref]$Result) {
        $combinedOutput = "$($Result.Value.Stderr)`n$($Result.Value.Stdout)"
        $isPnpmStoreError = $combinedOutput -match "VIRTUAL_STORE_DIR" -or $combinedOutput -match "broken lockfile" -or $combinedOutput -match "not compatible with current pnpm"
        if ($isPnpmStoreError) {
            Write-Warn "检测到 pnpm 全局 store 状态不兼容，正在清理后重试..."
            $pnpmGlobalDir = Join-Path (Get-LocalAppData) "pnpm\global"
            if (Test-Path $pnpmGlobalDir) {
                Remove-Item $pnpmGlobalDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Info "已清理 $pnpmGlobalDir"
            }
            try { & $PnpmCmd store prune 2>$null } catch {}
            $retryResult = Run-PnpmInstall -PnpmCmd $PnpmCmd -Label "重试安装"
            $Result.Value = $retryResult
            return $retryResult.Success
        }
        return $false
    }

    # 第 1 轮：使用官方 GitHub 地址（不走镜像）
    $result = Run-PnpmInstall -PnpmCmd $pnpmCmd -Label "安装"
    if ($result.Success) {
        Write-Ok "OpenClaw 安装完成"
        return $true
    }
    if (Try-InstallWithCleanup $pnpmCmd ([ref]$result)) {
        Write-Ok "OpenClaw 安装完成"
        return $true
    }

    # 第 2 轮：官方失败，依次尝试镜像
    Write-Warn "官方源安装失败，正在尝试镜像加速..."
    $fallbackMirrors = @(
        "https://bgithub.xyz/",
        "https://kkgithub.com/",
        "https://github.ur1.fun/",
        "https://ghproxy.net/https://github.com/",
        "https://gitclone.com/github.com/"
    )
    foreach ($mirror in $fallbackMirrors) {
        try {
            Set-GitMirror $mirror
            Write-Info "正在使用镜像 $mirror 重试..."
            $result = Run-PnpmInstall -PnpmCmd $pnpmCmd -Label "重试安装"
            Clear-GitMirror $mirror
            if ($result.Success) {
                Write-Ok "OpenClaw 安装完成 (通过镜像加速)"
                return $true
            }
        } catch {
            Clear-GitMirror $mirror
        }
    }

    Write-Err "OpenClaw 安装失败 (exit code: $($result.ExitCode))"
    if ($result.Stderr) {
        Write-Err "错误信息:"
        $result.Stderr.Trim().Split("`n") | ForEach-Object { Write-Host "         $_" -ForegroundColor Red }
    }
    if ($result.Stdout) {
        Write-Info "安装输出:"
        $result.Stdout.Trim().Split("`n") | Select-Object -Last 15 | ForEach-Object { Write-Host "         $_" }
    }
    return $false
}

function Step-Verify {
    Write-Step "步骤 6/7: 验证安装结果"

    Refresh-PathEnv
    Ensure-PnpmHome

    try {
        $pnpmCmd = Get-PnpmCmd
        $pnpmBin = (& $pnpmCmd bin -g 2>$null).Trim()
        if ($pnpmBin -and (Test-Path $pnpmBin)) {
            if ($env:PATH -notlike "*$pnpmBin*") {
                $env:PATH = "$pnpmBin;$env:PATH"
                Add-ToUserPath $pnpmBin
                Write-Info "已发现 pnpm 全局目录: $pnpmBin"
            }
        }
    } catch {}

    $openclawCmd = Get-OpenclawCmd
    $ver = $null
    try { $ver = & $openclawCmd -v 2>$null } catch {}
    if (-not $ver) { try { $ver = & openclaw -v 2>$null } catch {} }

    if ($ver) {
        Write-Ok "OpenClaw $ver 安装成功！"
        Write-Host "`n  🦞 恭喜！你的龙虾已就位！`n" -ForegroundColor Green
        return $true
    }

    Write-Warn "未能验证 OpenClaw 安装，请尝试重新打开终端后执行 openclaw -v"
    return $true
}

function Step-Onboard {
    Write-Step "步骤 7/7: 配置 OpenClaw"

    Write-Host "  请选择 AI 厂商:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   1) openai         - OpenAI (GPT-5.1 Codex, o3, o4-mini 等)"
    Write-Host "   2) anthropic      - Anthropic (Claude Sonnet 4.5, Opus 4.6)"
    Write-Host "   3) gemini         - Google Gemini (2.5 Pro, 2.5 Flash)"
    Write-Host "   4) mistral        - Mistral (Large, Codestral)"
    Write-Host "   5) zai            - 智谱 AI (GLM-5, GLM-4.7)"
    Write-Host "   6) moonshot       - Moonshot (Kimi K2.5)"
    Write-Host "   7) kimi-coding    - Kimi Coding (K2.5)"
    Write-Host "   8) qianfan        - 百度千帆 (ERNIE 4.5, DeepSeek R1)"
    Write-Host "   9) xiaomi         - 小米 (MiMo V2 Flash)"
    Write-Host "  10) custom         - 自定义 (OpenAI/Anthropic 兼容接口)"
    Write-Host "   0) 跳过配置"
    Write-Host ""

    $choice = (Read-Host "  请输入编号 [0-10]").Trim()

    if ($choice -eq "0") {
        Write-Info "已跳过配置"
        Write-Info "重新执行本安装脚本或运行 openclaw onboard 即可进入配置"
        return $true
    }

    $providerMap = @{
        "1"  = @{ Name="openai";      AuthChoice="openai-api-key";     KeyFlag="--openai-api-key" }
        "2"  = @{ Name="anthropic";   AuthChoice="apiKey";              KeyFlag="--anthropic-api-key" }
        "3"  = @{ Name="gemini";      AuthChoice="gemini-api-key";      KeyFlag="--gemini-api-key" }
        "4"  = @{ Name="mistral";     AuthChoice="mistral-api-key";     KeyFlag="--mistral-api-key" }
        "5"  = @{ Name="zai";         AuthChoice="zai-api-key";         KeyFlag="--zai-api-key" }
        "6"  = @{ Name="moonshot";    AuthChoice="moonshot-api-key";    KeyFlag="--moonshot-api-key" }
        "7"  = @{ Name="kimi-coding"; AuthChoice="kimi-code-api-key";   KeyFlag="--kimi-code-api-key" }
        "8"  = @{ Name="qianfan";     AuthChoice="qianfan-api-key";     KeyFlag="--qianfan-api-key" }
        "9"  = @{ Name="xiaomi";      AuthChoice="xiaomi-api-key";      KeyFlag="--xiaomi-api-key" }
        "10" = @{ Name="custom";      AuthChoice="custom-api-key";      KeyFlag="--custom-api-key" }
    }

    if (-not $providerMap.ContainsKey($choice)) {
        Write-Warn "无效选择，跳过配置"
        return $true
    }

    $provider = $providerMap[$choice]
    Write-Host ""
    $apiKey = (Read-Host "  请输入 API Key").Trim()
    if (-not $apiKey) {
        Write-Err "API Key 不能为空"
        return $false
    }

    $openclawCmd = Get-OpenclawCmd

    $onboardArgs = @(
        "onboard", "--non-interactive",
        "--accept-risk",
        "--mode", "local",
        "--auth-choice", $provider.AuthChoice,
        $provider.KeyFlag, $apiKey,
        "--secret-input-mode", "plaintext",
        "--gateway-port", "18789",
        "--gateway-bind", "loopback",
        "--skip-skills"
    )

    $customBaseUrl = ""
    $customModelId = ""
    if ($provider.Name -eq "custom") {
        Write-Host ""
        $customBaseUrl = (Read-Host "  请输入自定义 Base URL").Trim()
        $customModelId = (Read-Host "  请输入自定义 Model ID").Trim()

        if ($customBaseUrl) { $onboardArgs += @("--custom-base-url", $customBaseUrl) }
        if ($customModelId) { $onboardArgs += @("--custom-model-id", $customModelId) }
        $onboardArgs += @("--custom-compatibility", "openai")
    }

    Write-Info "正在配置 OpenClaw..."
    $savedEAP = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    & $openclawCmd @onboardArgs *>$null
    $onboardExit = $LASTEXITCODE
    $ErrorActionPreference = $savedEAP

    $configFile = Join-Path $env:USERPROFILE ".openclaw\openclaw.json"
    if (Test-Path $configFile) {
        Write-Ok "OpenClaw 配置完成！"
    } else {
        Write-Err "配置失败，请检查 API Key 是否正确"
        return $false
    }

    # 选择默认模型
    if ($provider.Name -ne "custom") {
        $modelMap = @{
            "openai"      = @(
                @{ Id="openai/gpt-5.1-codex"; Label="GPT-5.1 Codex" },
                @{ Id="openai/o3"; Label="o3" },
                @{ Id="openai/o4-mini"; Label="o4-mini" },
                @{ Id="openai/gpt-4.1"; Label="GPT-4.1" }
            )
            "anthropic"   = @(
                @{ Id="anthropic/claude-sonnet-4-5"; Label="Claude Sonnet 4.5" },
                @{ Id="anthropic/claude-opus-4-6"; Label="Claude Opus 4.6" }
            )
            "gemini"      = @(
                @{ Id="gemini/gemini-2.5-pro"; Label="Gemini 2.5 Pro" },
                @{ Id="gemini/gemini-2.5-flash"; Label="Gemini 2.5 Flash" }
            )
            "mistral"     = @(
                @{ Id="mistral/mistral-large-latest"; Label="Mistral Large" },
                @{ Id="mistral/codestral-latest"; Label="Codestral" }
            )
            "zai"         = @(
                @{ Id="zai/glm-5"; Label="GLM-5" },
                @{ Id="zai/glm-4.7"; Label="GLM-4.7" }
            )
            "moonshot"    = @(
                @{ Id="moonshot/kimi-k2.5"; Label="Kimi K2.5" },
                @{ Id="moonshot/kimi-k2-thinking"; Label="Kimi K2 Thinking" },
                @{ Id="moonshot/kimi-k2-thinking-turbo"; Label="Kimi K2 Thinking Turbo" }
            )
            "kimi-coding" = @(
                @{ Id="kimi-coding/k2p5"; Label="Kimi K2.5 (Coding)" }
            )
            "qianfan"     = @(
                @{ Id="qianfan/ernie-4.5-turbo-vl-32k"; Label="ERNIE 4.5 Turbo" },
                @{ Id="qianfan/deepseek-r1"; Label="DeepSeek R1" }
            )
            "xiaomi"      = @(
                @{ Id="xiaomi/mimo-v2-flash"; Label="MiMo V2 Flash" }
            )
        }

        $models = $modelMap[$provider.Name]
        if ($models -and $models.Count -gt 0) {
            Write-Host ""
            Write-Host "  请选择默认模型:" -ForegroundColor Cyan
            Write-Host ""
            for ($i = 0; $i -lt $models.Count; $i++) {
                Write-Host "   $($i+1)) $($models[$i].Label)"
            }
            Write-Host "   0) 跳过"
            Write-Host ""
            $modelChoice = (Read-Host "  请选择 [0-$($models.Count)]").Trim()

            if ($modelChoice -ne "0" -and $modelChoice) {
                $idx = [int]$modelChoice - 1
                if ($idx -ge 0 -and $idx -lt $models.Count) {
                    $selectedModel = $models[$idx]
                    Write-Info "正在设置默认模型: $($selectedModel.Id)"
                    try {
                        & $openclawCmd models set $selectedModel.Id 2>$null
                        Write-Ok "默认模型已设置为 $($selectedModel.Label)"
                    } catch {
                        Write-Warn "模型设置未成功，可稍后通过 openclaw models set 手动设置"
                    }
                }
            }
        }
    }

    # 启动 gateway 服务（Windows 无 systemd，不能用 daemon，直接启动 gateway）
    Write-Host ""
    Write-Info "正在启动 OpenClaw 服务..."
    try {
        Start-Process -FilePath "cmd.exe" `
            -ArgumentList "/c `"$openclawCmd`" gateway" `
            -WindowStyle Hidden
        Start-Sleep -Seconds 3
        Write-Ok "OpenClaw 已在后台启动，正在打开浏览器..."
        try { & $openclawCmd dashboard 2>$null } catch { Start-Process "http://127.0.0.1:18789" }
    } catch {
        Write-Warn "无法自动启动服务，请手动执行: openclaw gateway"
    }

    return $true
}

# ── 主函数 ──

function Main {
    Write-Host ""
    Write-Host "  🦞 OpenClaw 一键安装脚本" -ForegroundColor Green
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Blue
    Write-Host ""

    Refresh-PathEnv

    # 检测是否已安装
    $existingVer = $null
    try { $existingVer = & openclaw -v 2>$null } catch {}
    if ($existingVer) {
        Write-Ok "OpenClaw $existingVer 已安装，无需重复安装"
        Write-Host "`n  🦞 你的龙虾已就位！`n" -ForegroundColor Green
        $reconfig = (Read-Host "  是否要重新配置 OpenClaw? [y/N]").Trim()
        if ($reconfig -match "^[Yy]") {
            Step-Onboard | Out-Null
        }
        return
    }

    if (-not (Step-CheckNode))       { Write-Host "`n按任意键退出..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-CheckGit))        { Write-Host "`n按任意键退出..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-SetMirror))       { Write-Host "`n按任意键退出..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-InstallPnpm))     { Write-Host "`n按任意键退出..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-InstallOpenClaw)) { Write-Host "`n按任意键退出..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    Step-Verify | Out-Null
    Step-Onboard | Out-Null

    Write-Host ""
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
    Write-Host "  🦞 安装完成！请打开新终端窗口开始使用 OpenClaw" -ForegroundColor Green
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
    Write-Host ""
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Main
