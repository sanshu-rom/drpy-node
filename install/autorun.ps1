param(
    [switch]$UseProxy,
    [string]$ProxyHost = "127.0.0.1:7890",
    [switch]$SkipConfirm
)
$ErrorActionPreference = "Stop"

# -------------------------------------------------
# 0. 工具函数
# -------------------------------------------------
function Use-ProxyIfNeeded {
    param([scriptblock]$Script)
    if ($UseProxy) {
        $oldHttp  = [Environment]::GetEnvironmentVariable("HTTP_PROXY")
        $oldHttps = [Environment]::GetEnvironmentVariable("HTTPS_PROXY")
        [Environment]::SetEnvironmentVariable("HTTP_PROXY",  "http://$ProxyHost", "Process")
        [Environment]::SetEnvironmentVariable("HTTPS_PROXY", "http://$ProxyHost", "Process")
        try { & $Script } finally {
            [Environment]::SetEnvironmentVariable("HTTP_PROXY",  $oldHttp,  "Process")
            [Environment]::SetEnvironmentVariable("HTTPS_PROXY", $oldHttps, "Process")
        }
    } else { & $Script }
}
function Test-Cmd { param($cmd); $null -ne (Get-Command $cmd -ErrorAction SilentlyContinue) }
function Invoke-WebRequestWithProxy([string]$Url, [string]$OutFile) {
    if ($UseProxy) { Invoke-WebRequest $Url -OutFile $OutFile -Proxy "http://$ProxyHost" }
    else           { Invoke-WebRequest $Url -OutFile $OutFile }
}

# -------------------------------------------------
# 1. 自动提权 & 回到脚本目录
# -------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "当前非管理员权限，正在尝试以管理员身份重新启动，如果闪退请走代理..." -ForegroundColor Yellow
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    if ($UseProxy)    { $arguments += " -UseProxy -ProxyHost `"$ProxyHost`"" }
    if ($SkipConfirm) { $arguments += " -SkipConfirm" }
    Start-Process powershell -ArgumentList $arguments -Verb RunAs
    exit
}
Set-Location -LiteralPath $PSScriptRoot

# -------------------------------------------------
# 2. 按需安装 winget
# -------------------------------------------------
function Install-Winget {
    if (Test-Cmd winget) { return }
    Write-Host "未检测到 winget，正在安装 App Installer..." -ForegroundColor Yellow
    $url = "https://aka.ms/getwinget"
    $out = "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    Invoke-WebRequestWithProxy $url $out
    Add-AppxPackage -Path $out -ErrorAction SilentlyContinue
    Remove-Item $out
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# -------------------------------------------------
# 3. 用户确认
# -------------------------------------------------
if (-not $SkipConfirm) {
    Write-Host "警告：此脚本仅适用于 Windows 10/11 64 位" -ForegroundColor Green
    Write-Host "建议使用 Windows Terminal 终端管理员方式运行" -ForegroundColor Green
    Write-Host "如果 nvm、git、python 安装失败，建议手动安装" -ForegroundColor Green
    Write-Host "下载失败可指定旁路由代理：.\drpys-auto.ps1 -UseProxy -ProxyHost `"192.168.1.21:7890`"" -ForegroundColor Green
    Write-Host "如果旁路由也下载失败那就换成道长那个白嫖地址" -ForegroundColor Green
    $confirm = Read-Host "您是否理解并同意继续？(y/n) 默认(y)"
    if ($confirm -eq "n") { exit 1 }
}

# -------------------------------------------------
# 4. 安装 nvm
# -------------------------------------------------
if (-not (Test-Cmd "nvm")) {
    Write-Host "正在安装 nvm-windows..." -ForegroundColor Green
    $nvmSetup = "$env:TEMP\nvm-setup.exe"
    Invoke-WebRequestWithProxy "https://github.com/coreybutler/nvm-windows/releases/latest/download/nvm-setup.exe" $nvmSetup
    Start-Process -Wait -FilePath $nvmSetup -ArgumentList "/silent"
    Remove-Item $nvmSetup
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Yellow
    Write-Host "nvm 已安装完毕，但当前 PowerShell 会话尚未识别到它。" -ForegroundColor Yellow
    Write-Host "请执行以下任意一步后再继续："                         -ForegroundColor Cyan
    Write-Host "  1) 关闭本窗口，重新打开一个『管理员』PowerShell后，再次执行脚本；"      -ForegroundColor Cyan
    Write-Host "  2) 或者再次右键选PS运行本脚本。"                         -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Yellow
    Read-Host "按 Enter 键退出本窗口"
    exit
} else {
    Write-Host "已检测到 nvm，跳过安装" -ForegroundColor Green
}

# -------------------------------------------------
# 5. 安装/切换 Node
# -------------------------------------------------
$needNode = $false
if (Test-Cmd "node") {
    $nodeVer = (node -v) -replace '^v','' -split '\.' | ForEach-Object { [int]$_ }
    $current = $nodeVer[0]*10000 + $nodeVer[1]*100 + $nodeVer[2]
    $require = 20*10000 + 18*100 + 3          # 20.18.3
    if ($current -ge $require) {
        Write-Host "已检测到 Node v$($nodeVer -join '.') ≥20.18.3，跳过安装" -ForegroundColor Green
    } else {
        Write-Host "Node 版本低于 20.18.3，将使用 nvm 安装/切换到 20.18.3" -ForegroundColor Yellow
        $needNode = $true
    }
} else {
    Write-Host "未检测到 Node，准备安装" -ForegroundColor Yellow
    $needNode = $true
}
if ($needNode) {
    nvm install 20.18.3
    nvm use 20.18.3
}

# -------------------------------------------------
# 6. 安装 Python 3.11（优先 winget）
# -------------------------------------------------
$pyNeed = $false
try {
    $ver = (python -V 2>$null) -replace 'Python ',''
    if ($ver -match '^3\.11') {
        Write-Host "已检测到 Python 3.11 ($ver)，跳过安装" -ForegroundColor Green
    } else {
        Write-Host "检测到非 3.11 版本，准备覆盖安装 3.11" -ForegroundColor Yellow
        $pyNeed = $true
    }
} catch {
    Write-Host "未检测到 Python，准备安装 3.11" -ForegroundColor Yellow
    $pyNeed = $true
}
if ($pyNeed) {
    Install-Winget
    Write-Host "正在通过 winget 安装 Python 3.11..." -ForegroundColor Green
    winget install --id Python.Python.3.11 -e --accept-source-agreements --accept-package-agreements
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
}

# -------------------------------------------------
# 7. 安装 Git：winget 优先，失败自动离线
# -------------------------------------------------
if (-not (Test-Cmd "git")) {
    # 1) winget 交互式安装
    Install-Winget
    if (Test-Cmd winget) {
        Write-Host "正在通过 winget 安装 Git（交互模式）..." -ForegroundColor Green
        try {
            winget install --id Git.Git -e --source winget
            if (Test-Cmd git) {
                Write-Host "Git 安装成功（winget）" -ForegroundColor Green
                $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
                continue
            }
        } catch {
            Write-Host "winget 安装失败，将使用离线包..." -ForegroundColor Yellow
        }
    }

    # 2) winget 失败 → 离线安装
    Write-Host "正在解析 Git 最新版本..." -ForegroundColor Green
    try {
        $latestUri = (Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/latest" -MaximumRedirection 0 -ErrorAction SilentlyContinue).Headers.Location
        $ver = if ($latestUri) { $latestUri -replace '.*/tag/v([0-9.]+).*$','$1' } else { "2.51.0" }
    } catch {
        $ver = "2.51.0"
    }

    Write-Host "正在下载 Git $ver ..." -ForegroundColor Green
    $gitSetup = "$env:TEMP\Git-$ver-64-bit.exe"
    $gitUrl   = "https://github.com/git-for-windows/git/releases/download/v$ver.windows.1/Git-$ver-64-bit.exe"
    Invoke-WebRequestWithProxy $gitUrl $gitSetup
    Start-Process -Wait -FilePath $gitSetup -ArgumentList "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS"
    Remove-Item $gitSetup -Force
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
} else {
    Write-Host "已检测到 Git，跳过安装" -ForegroundColor Green
}

# -------------------------------------------------
# 8. 安装全局 npm 工具
# -------------------------------------------------
$tools = @{
    yarn = { npm install -g yarn }
    pm2  = { npm install -g pm2 }
}
foreach ($kv in $tools.GetEnumerator()) {
    if (-not (Test-Cmd $kv.Key)) {
        Write-Host "正在安装 $($kv.Key) ..." -ForegroundColor Yellow
        & $kv.Value
    } else {
        Write-Host "已检测到 $($kv.Key)，跳过安装" -ForegroundColor Green
    }
}
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# -------------------------------------------------
# 9. 克隆仓库 / 配置 / 依赖 / PM2
# -------------------------------------------------
$repoDir = Read-Host "请输入项目存放目录（留空则使用当前目录）"
if ([string]::IsNullOrWhiteSpace($repoDir)) { $repoDir = (Get-Location).Path }
$projectPath = Join-Path $repoDir "drpy-node"
$remoteRepo  = "https://github.com/hjdhnx/drpy-node.git"

# 记录路径供后续计划任务使用
$projectPath | Out-File "$env:ProgramData\drpys-path.txt" -Encoding UTF8 -Force

Use-ProxyIfNeeded -Script {
    if (-not (Test-Path $projectPath)) {
        Write-Host "正在克隆仓库..." -ForegroundColor Yellow
        if ($UseProxy) {
            git -c http.proxy="http://$ProxyHost" clone $remoteRepo $projectPath
        } else {
            git clone $remoteRepo $projectPath
        }
    }
    Set-Location $projectPath

    # 生成 env.json
    $configDir  = Join-Path $projectPath "config"
    $configJson = Join-Path $configDir "env.json"
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
    if (-not (Test-Path $configJson)) {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $jsonText  = @{
            ali_token = ""; ali_refresh_token = ""; quark_cookie = ""
            uc_cookie = ""; bili_cookie = ""; thread = "10"
            enable_dr2 = "1"; enable_py = "2"
        } | ConvertTo-Json
        [System.IO.File]::WriteAllLines($configJson, $jsonText, $utf8NoBom)
    }

# 生成 .env（UTF-8 无 BOM，不乱码）
$envFile = Join-Path $projectPath ".env"
if (-not (Test-Path $envFile)) {
    # 如果仓库没带模板，就写一份最小模板（同样无 BOM）
    $template = Join-Path $projectPath ".env.development"
    if (-not (Test-Path $template)) {
        @"
NODE_ENV=development
COOKIE_AUTH_CODE=drpys
API_AUTH_NAME=admin
API_AUTH_CODE=drpys
API_PWD=dzyyds
"@ | Out-File $template -Encoding UTF8
    }

    # 复制模板
    Copy-Item $template $envFile

    # 依次输入
    $cookieAuth = (Read-Host "网盘入库密码（默认 drpys）").Trim()
    $apiUser    = (Read-Host "登录用户名（默认 admin）").Trim()
    $apiPass    = (Read-Host "登录密码（默认 drpys）").Trim()
    $apiPwd     = (Read-Host "订阅PWD值（默认 dzyyds）").Trim()

    # 空值兜底
    if ([string]::IsNullOrWhiteSpace($cookieAuth)) { $cookieAuth = 'drpys' }
    if ([string]::IsNullOrWhiteSpace($apiUser))    { $apiUser    = 'admin' }
    if ([string]::IsNullOrWhiteSpace($apiPass))    { $apiPass    = 'drpys' }
    if ([string]::IsNullOrWhiteSpace($apiPwd))     { $apiPwd     = 'dzyyds' }

    # 逐行替换，最后统一 UTF-8 无 BOM 写回
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $lines = [System.IO.File]::ReadAllLines($template, $utf8NoBom)

    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*COOKIE_AUTH_CODE\s*=') {
            $lines[$i] = "COOKIE_AUTH_CODE = $cookieAuth"
        }
        elseif ($lines[$i] -match '^\s*API_AUTH_NAME\s*=') {
            $lines[$i] = "API_AUTH_NAME = $apiUser"
        }
        elseif ($lines[$i] -match '^\s*API_AUTH_CODE\s*=') {
            $lines[$i] = "API_AUTH_CODE = $apiPass"
        }
        elseif ($lines[$i] -match '^\s*API_PWD\s*=') {
            $lines[$i] = "API_PWD = $apiPwd"
        }
    }

    [System.IO.File]::WriteAllLines($envFile, $lines, $utf8NoBom)
}

    # Node 依赖
    if (-not (Test-Path "node_modules")) {
        Write-Host "首次安装 Node 依赖..." -ForegroundColor Yellow
        yarn config set registry https://registry.npmmirror.com/
        yarn
    } elseif ((git diff HEAD^ HEAD --name-only 2>$null) -match [regex]::Escape("yarn.lock")) {
        Write-Host "检测到 yarn.lock 变动，更新 Node 依赖..." -ForegroundColor Yellow
        yarn install --registry https://registry.npmmirror.com/
    }

    # Python 虚拟环境 & 依赖
    $venvActivate = Join-Path $projectPath ".venv\Scripts\Activate.ps1"
    if (-not (Test-Path ".venv\pyvenv.cfg")) {
        Write-Host "首次创建 Python 虚拟环境..." -ForegroundColor Yellow
        python -m venv .venv
    }
    & $venvActivate
    python -m pip install --upgrade pip -q
    pip install -r spider\py\base\requirements.txt -i https://mirrors.cloud.tencent.com/pypi/simple -q

    if ((git diff HEAD^ HEAD --name-only 2>$null) -match [regex]::Escape("spider\py\base\requirements.txt")) {
        Write-Host "检测到 requirements.txt 变动，更新 Python 依赖..." -ForegroundColor Yellow
        pip install -r spider\py\base\requirements.txt -i https://mirrors.cloud.tencent.com/pypi/simple -q
    }

    # PM2
    if (-not (pm2 list | Select-String "drpyS.*online")) {
        Write-Host "首次启动 PM2 进程..." -ForegroundColor Yellow
        pm2 start index.js --name drpyS --update-env
        pm2 save
    } else {
        Write-Host "PM2 进程 drpyS 已在运行，跳过启动" -ForegroundColor Green
    }
}

# -------------------------------------------------
# 10. 计划任务（开机自启 + 6 小时更新）
# -------------------------------------------------
$pathFile = "$env:ProgramData\drpys-path.txt"
if (Test-Path $pathFile) {
    $projectPath = (Get-Content $pathFile -Raw).Trim()
} else {
    Write-Host "首次未检测到路径记录，请先完整运行一次主脚本！" -ForegroundColor Red
    exit 1
}

$taskStartup = "drpyS_PM2_Startup"
$taskUpdate  = "drpyS_Update"
$pm2     = (Get-Command pm2.cmd  -ErrorAction SilentlyContinue).Source
$nodeExe = (Get-Command node.exe -ErrorAction SilentlyContinue).Source

if ($pm2 -and $nodeExe) {
    $taskStartup,$taskUpdate | ForEach-Object {
        if (Get-ScheduledTask -TaskName $_ -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $_ -Confirm:$false
        }
    }

    $commonSettings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    # 开机自启（激活 venv）
    $action  = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -Command `"& { `$env:PM2_HOME='C:\$env:USERNAME\.pm2'; Set-Location '$projectPath'; & .\.venv\Scripts\Activate.ps1; & '$pm2' start '$projectPath\index.js' --name drpyS --update-env }`"" `
        -WorkingDirectory $projectPath
    $trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -Seconds 30)
    Register-ScheduledTask -TaskName $taskStartup -Action $action -Trigger $trigger -Settings $commonSettings -User "SYSTEM" -RunLevel Highest -Force | Out-Null
    Write-Host "已创建/更新开机自启任务：$taskStartup" -ForegroundColor Green

    # 每 6 小时更新（激活 venv）
    $action  = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"& { `$env:PM2_HOME='C:\$env:USERNAME\.pm2'; Set-Location '$projectPath'; & .\.venv\Scripts\Activate.ps1; git fetch origin; if (git status -uno | Select-String 'Your branch is behind') { git reset --hard origin/main; yarn install --registry https://registry.npmmirror.com/; pip install -r spider\py\base\requirements.txt -i https://mirrors.cloud.tencent.com/pypi/simple; & '$pm2' restart drpyS } }`"" `
        -WorkingDirectory $projectPath
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 6)
    Register-ScheduledTask -TaskName $taskUpdate -Action $action -Trigger $trigger -Settings $commonSettings -User "SYSTEM" -RunLevel Highest -Force | Out-Null
    Write-Host "已创建/更新每 6 小时更新任务：$taskUpdate" -ForegroundColor Green
}

if (Test-Path "$projectPath\.venv\Scripts\Activate.ps1") {
    & "$projectPath\.venv\Scripts\Activate.ps1" | Out-Null
}
# -------------------------------
# 12. 结束提示 & 清理
# -------------------------------
# 1. 退出虚拟环境（如有）
deactivate *>$null
# 2. 回到脚本启动目录
Set-Location -LiteralPath $PSScriptRoot
# 3. 输出地址 & 暂停
$ip = (ipconfig | Select-String "IPv4 地址" | Select-Object -First 1).ToString().Split(":")[-1].Trim()
$public = (Invoke-RestMethod "https://ipinfo.io/ip")
Write-Host "内网地址：http://${ip}:5757" -ForegroundColor Green
Write-Host "公网地址：http://${public}:5757" -ForegroundColor Green
Write-Host "脚本执行完成！重启后 drpyS 自动启动并每 6 小时检查更新。" -ForegroundColor Green
Write-Host "按任意键退出！！！" -ForegroundColor Green
Read-Host
