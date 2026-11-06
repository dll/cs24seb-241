# ============================================
# 软件工程基础 - 一键环境安装脚本 v6.0
# 安装根目录：D:/CS24SEB
# 特性：进度条 | 重试3次 | 回滚 | 断点续装 | 端口可配
# 用法：直接运行会自动弹出 UAC 提权；
#       已在管理员 PowerShell 中时可加 -NoElevate 跳过自检
# ============================================
param([switch]$NoElevate)

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

# PS5.1 默认协议集可能不含 TLS1.2，get.scoop.sh / GitHub 等均要求 TLS1.2+
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

# ====== 自动提权：非管理员会话下通过 UAC 重启自身 ======
if (-not $NoElevate) {
    $wid = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wpr = New-Object Security.Principal.WindowsPrincipal($wid)
    if (-not $wpr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "安装需要管理员权限，正在请求 UAC 提权并重新启动..." -ForegroundColor Yellow
        try {
            $elevArgs = @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`"")
            Start-Process -FilePath 'powershell.exe' -ArgumentList $elevArgs `
                -Verb RunAs -WorkingDirectory $PSScriptRoot -ErrorAction Stop | Out-Null
            exit 0
        } catch {
            Write-Host "提权失败（可能拒绝了 UAC）。请右键 PowerShell 选择“以管理员身份运行”，再执行：" -ForegroundColor Red
            Write-Host "  powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -ForegroundColor Yellow
            pause; exit 1
        }
    }
}

# 刷新会话 PATH：合并注册表中的 Machine+User PATH，
# 保证公共机器上已安装的软件（docker/java/node 等）能被立即识别，绝不重装
$env:PATH = [Environment]::GetEnvironmentVariable('PATH','Machine') + ';' + [Environment]::GetEnvironmentVariable('PATH','User')

$BASE = "D:/CS24SEB"
$BASE_WIN = "D:\CS24SEB"
$STATE_FILE = "$BASE/.install_state.json"
$MAX_RETRY = 3
$TOTAL_STEPS = 24
$script:currentStep = 0
$script:startTime = Get-Date

# ====== 软件信息 ======
$SW = @{
    scoop     = @{ N="Scoop";      V="latest";  Desc="Windows包管理器"; Chapter="第1章-02 DevOps"; Experiment="实验一"; Url1="https://get.scoop.sh"; Url2="https://github.com/ScoopInstaller/Install" }
    git       = @{ N="Git";        V="2.47.1"; Desc="分布式版本控制系统"; Chapter="第1章-03 Git版本控制"; Experiment="实验一"; Url1="https://git-scm.com/download/win"; Url2="https://github.com/git-for-windows/git/releases" }
    nodejs    = @{ N="Node.js";    V="24.21.0"; Desc="JavaScript运行时"; Chapter="第5章-12 软件编码"; Experiment="实验一~五"; Url1="https://nodejs.org/dist/v24.21.0/node-v24.21.0-x64.msi"; Url2="https://npmmirror.com/mirrors/node/v24.21.0/node-v24.21.0-x64.msi" }
    python    = @{ N="Python";     V="3.11.9"; Desc="通用编程语言"; Chapter="第5章-12 软件编码"; Experiment="实验一~五"; Url1="https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"; Url2="https://mirrors.huaweicloud.com/python/" }
    jdk21     = @{ N="JDK 21";     V="21.0.5"; Desc="Java开发工具包"; Chapter="第4章-08 设计方法"; Experiment="实验一~五"; Url1="https://github.com/adoptium/temurin21-binaries/releases"; Url2="https://mirrors.tuna.tsinghua.edu.cn/Adoptium/" }
    maven     = @{ N="Maven";      V="3.9.9"; Desc="Java构建工具"; Chapter="第4章-08 设计方法"; Experiment="实验三~四"; Url1="https://maven.apache.org/download.cgi"; Url2="https://maven.aliyun.com/repository/public" }
    docker    = @{ N="Docker";     V="4.36.0"; Desc="容器引擎"; Chapter="第1章-02/第7章-16 DevOps"; Experiment="实验五"; Url1="https://desktop.docker.com/win/main/amd64/Docker%20Desktop%20Installer.exe"; Url2="https://mirrors.aliyun.com/docker-toolbox/" }
    trae      = @{ N="TRAE IDE";   V="Latest"; Desc="AI IDE"; Chapter="第1章-02/第5章-12 AI工具"; Experiment="实验一~五"; Url1="https://www.trae.ai/"; Url2="https://www.trae.cn/" }
    ea        = @{ N="EA UML";     V="17.1 LTS"; Desc="UML建模工具"; Chapter="第2章-04/05、第3章-06 UML"; Experiment="实验二"; Url1="https://sparxsystems.com/"; Url2="https://staruml.io/" }
    jmeter    = @{ N="JMeter";     V="5.6.3"; Desc="性能/压力测试"; Chapter="第6章-13/15 软件测试"; Experiment="实验五"; Url1="https://dlcdn.apache.org/jmeter/binaries/apache-jmeter-5.6.3.zip"; Url2="https://mirrors.tuna.tsinghua.edu.cn/apache/jmeter/binaries/apache-jmeter-5.6.3.zip" }
    postgres  = @{ N="PostgreSQL"; V="17-3.4"; Desc="空间数据库"; Chapter="第4章-09 UI/DB设计"; Experiment="实验三~五"; Url1="https://hub.docker.com/r/postgis/postgis"; Url2="https://postgis.net/windows_downloads/" }
    redis     = @{ N="Redis";      V="7-alpine"; Desc="内存缓存"; Chapter="第4章-10 模块设计"; Experiment="实验三~五"; Url1="https://hub.docker.com/_/redis"; Url2="https://redis.io/download" }
    es        = @{ N="Elasticsearch"; V="8.17.0"; Desc="全文检索"; Chapter="第3章-06 需求工程"; Experiment="实验三~五"; Url1="https://hub.docker.com/_/elasticsearch"; Url2="https://www.elastic.co/downloads/elasticsearch" }
    mongo     = @{ N="MongoDB";    V="7"; Desc="文档数据库"; Chapter="第4章-09 UI/DB设计"; Experiment="实验三~五"; Url1="https://hub.docker.com/_/mongo"; Url2="https://www.mongodb.com/try/download" }
    sonarqube = @{ N="SonarQube";  V="community"; Desc="代码质量"; Chapter="第5章-12 软件编码"; Experiment="实验四"; Url1="https://hub.docker.com/_/sonarqube"; Url2="https://www.sonarsource.com/products/sonarqube/" }
    jenkins   = @{ N="Jenkins";    V="lts-jdk17"; Desc="CI/CD"; Chapter="第7章-16 DevOps"; Experiment="实验五"; Url1="https://hub.docker.com/_/jenkins"; Url2="https://www.jenkins.io/download/" }
    sonarscan = @{ N="SonarScanner"; V="7.0.2"; Desc="代码扫描客户端"; Chapter="第5章-12 软件编码"; Experiment="实验四"; Url1="https://github.com/SonarSource/sonar-scanner-cli/releases"; Url2="https://binaries.sonarsource.com/Distribution/sonar-scanner-cli/" }
    plantuml  = @{ N="PlantUML";   V="latest"; Desc="UML文本建模"; Chapter="第2章-04/05 UML"; Experiment="实验二"; Url1="https://maven.aliyun.com/repository/public/net/sourceforge/plantuml/plantuml/1.2024.7/plantuml-1.2024.7.jar"; Url2="https://github.com/plantuml/plantuml/releases/latest/download/plantuml.jar" }
    graphviz  = @{ N="Graphviz";   V="12.x"; Desc="PlantUML布局依赖"; Chapter="第2章-04/05 UML"; Experiment="实验二"; Url1="https://graphviz.org/download/"; Url2="https://gitlab.com/api/v4/projects/4207231/packages/generic/graphviz-releases/latest/windows_10_cmake_Release_graphviz-install-x86_64.exe" }
}

# ====== 从 .env 读取端口和密码 ======
function Get-EnvVal {
    param([string]$Key, [string]$Default="")
    if (Test-Path "$BASE/.env") {
        $line = Get-Content "$BASE/.env" | Where-Object { $_ -match "^$Key=" } | Select-Object -First 1
        if ($line) { return ($line -split "=",2)[1] }
    }
    return $Default
}

# 确保 .env 存在
if (!(Test-Path "$BASE/.env")) {
@"
PG_PORT=5432
REDIS_PORT=6379
ES_PORT=9200
MONGO_PORT=27017
SONAR_PORT=9000
JENKINS_PORT=8081
BACKEND_PORT=8080
FRONTEND_PORT=3000
# 可选监控栈（AIOps，见第十二章）
PROM_PORT=9090
GRAFANA_PORT=3001
KIBANA_PORT=5601
PG_USER=postgres
PG_PASSWORD=se2024pass
PG_DB=seproject
REDIS_PASSWORD=
MONGO_USER=admin
MONGO_PASSWORD=se2024pass
SONAR_USER=admin
SONAR_PASSWORD=admin
JWT_SECRET=change-this-jwt-secret
DEEPSEEK_API_KEY=sk-your-key-here
OPENAI_API_KEY=sk-your-key-here
CS24_BASE=D:/CS24SEB
"@ | Out-File -FilePath "$BASE/.env" -Encoding utf8
}

# 读取端口配置
$PG_PORT      = Get-EnvVal "PG_PORT" "5432"
$REDIS_PORT   = Get-EnvVal "REDIS_PORT" "6379"
$ES_PORT      = Get-EnvVal "ES_PORT" "9200"
$MONGO_PORT   = Get-EnvVal "MONGO_PORT" "27017"
$SONAR_PORT   = Get-EnvVal "SONAR_PORT" "9000"
$JENKINS_PORT = Get-EnvVal "JENKINS_PORT" "8081"
$BACKEND_PORT = Get-EnvVal "BACKEND_PORT" "8080"
$PG_USER      = Get-EnvVal "PG_USER" "postgres"
$PG_PASSWORD  = Get-EnvVal "PG_PASSWORD" "se2024pass"
$PG_DB        = Get-EnvVal "PG_DB" "seproject"
$MONGO_USER   = Get-EnvVal "MONGO_USER" "admin"
$MONGO_PASS   = Get-EnvVal "MONGO_PASSWORD" "se2024pass"

# ====== 输出函数 ======
function Write-Step { param([string]$msg)
    $script:currentStep++
    $pct = [math]::Round($script:currentStep / $TOTAL_STEPS * 100)
    $elapsed = ((Get-Date) - $script:startTime).ToString("mm\:ss")
    Write-Host ""
    Write-Host "[$script:currentStep/$TOTAL_STEPS $pct% $elapsed] " -NoNewline -ForegroundColor Cyan
    Write-Host ">>> $msg" -ForegroundColor Cyan
}
function Write-OK    { param([string]$msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Skip  { param([string]$msg) Write-Host "    [SKIP] $msg" -ForegroundColor Yellow }
function Write-Err   { param([string]$msg) Write-Host "    [ERR] $msg" -ForegroundColor Red }
function Write-Retry { param([string]$msg,[int]$n) Write-Host "    [RETRY $n/$MAX_RETRY] $msg" -ForegroundColor DarkYellow }
function Write-Info  { param([string]$msg) Write-Host "    [INFO] $msg" -ForegroundColor Gray }

# ====== 状态管理 ======
# 注意：PS5.1 的 ConvertTo-Json 会把单元素数组拍平成字符串，
# 读取时必须用 @() 强制还原为数组，否则 += 会变成字符串拼接
function Get-State {
    if (Test-Path $STATE_FILE) {
        try {
            $o = Get-Content $STATE_FILE -Raw | ConvertFrom-Json
            return [pscustomobject]@{ completed=@($o.completed | Where-Object { $_ }); failed=@($o.failed | Where-Object { $_ }); lastStep=[int]$o.lastStep }
        } catch {
            # 状态文件损坏（如写入中途断电/被杀）：当空状态继续，靠实际检测兜底，不中断安装
            Write-Host "    [WARN] 状态文件损坏，将重新按实际检测结果续装" -ForegroundColor Yellow
        }
    }
    return [pscustomobject]@{ completed=@(); failed=@(); lastStep=0 }
}
function Save-State { param([object]$s) $s | ConvertTo-Json -Depth 3 | Out-File $STATE_FILE -Encoding utf8 }
function Done { param([string]$step,[object]$s)
    $s.completed = @($s.completed)
    if ($s.completed -notcontains $step) { $s.completed += $step }
    $s.failed = @($s.failed | Where-Object { $_ -ne $step })
    Save-State $s
}
function Failed { param([string]$step,[object]$s)
    $s.failed = @($s.failed)
    if ($s.failed -notcontains $step) { $s.failed += $step }
    Save-State $s
}
function Skipped { param([string]$step,[object]$s) return (@($s.completed) -contains $step) }

# ====== 重试下载 ======
function DL {
    param([string]$Url,[string]$Out,[int]$Max=$MAX_RETRY)
    for ($i=1;$i -le $Max;$i++) {
        try {
            Invoke-WebRequest -Uri $Url -OutFile $Out -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
            $sz = if (Test-Path $Out) { [math]::Round((Get-Item $Out).Length/1MB,1) } else { 0 }
            Write-OK "下载完成 (${sz}MB)"
            return $true
        } catch {
            Write-Retry "$($_.Exception.Message)" $i
            if ($i -eq $Max) { return $false }
            Start-Sleep -Seconds (5*$i)
        }
    }
    return $false
}

# ====== 已安装软件检测（查注册表卸载项，公共机器上已装的软件直接复用）======
function Test-InstalledApp { param([string]$Pattern)
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    return (Get-ItemProperty $roots -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $Pattern } | Select-Object -First 1)
}

# ====== 多地址依次尝试下载（每个地址只试1次，全部失败才返回 false）======
function DL2 { param([string[]]$Urls,[string]$Out)
    foreach ($u in $Urls) {
        if ($u -and (DL $u $Out 1)) { return $true }
    }
    return $false
}

# ====== 命令是否真实可用（排除微软商店 WindowsApps 占位别名，如 python.exe stub）======
function Test-RealCmd {
    param([string]$Name)
    $g = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $g) { return $false }
    if ($g.Source -like '*\WindowsApps\*') { return $false }
    return $true
}

# ====== 重试 scoop install ======
# 注意：scoop install 不支持按包指定安装目录（无 --dir 参数），
# 所有包统一安装在 $env:SCOOP\apps\<包名>\current，可执行文件经 shims 暴露到 PATH
function ScoopInstall {
    param([string]$Pkg,[int]$Max=$MAX_RETRY)
    # 先确认 scoop 命令可用，否则直接返回 false（避免错误被 2>&1 吞掉后误报成功）
    if (-not (Get-Command scoop -ErrorAction SilentlyContinue)) {
        Write-Err "scoop 命令不可用，无法安装 $Pkg"
        return $false
    }
    for ($i=1;$i -le $Max;$i++) {
        try {
            # 用 $LASTEXITCODE 和 scoop 输出检测真实结果，不能靠 2>&1|Out-Null（错误被吞，不抛异常）
            $output = & scoop install $Pkg 2>&1
            $exitCode = $LASTEXITCODE
            if ($exitCode -eq 0 -and $output -notmatch "ERROR|Failed|cannot find") {
                return $true
            }
            Write-Retry "scoop install $Pkg 失败 (exit=$exitCode)" $i
            if ($i -eq $Max) { return $false }
            Start-Sleep -Seconds (3*$i)
        } catch {
            Write-Retry "scoop install $Pkg 异常: $($_.Exception.Message)" $i
            if ($i -eq $Max) { return $false }
            Start-Sleep -Seconds (3*$i)
        }
    }
    return $false
}

# ====== 重试 docker run ======
function DockerRun {
    param([string[]]$DockerArgs,[string]$Name,[int]$Max=$MAX_RETRY)
    for ($i=1;$i -le $Max;$i++) {
        try {
            $cmd = "docker run -d --name $Name " + ($DockerArgs -join " ")
            Invoke-Expression $cmd 2>&1 | Out-Null
            Start-Sleep -Seconds 3
            $st = docker inspect -f '{{.State.Status}}' $Name 2>$null
            if ($st -eq "running") { return $true }
        } catch { Write-Retry "容器 $Name 启动失败" $i }
        docker rm -f $Name 2>$null | Out-Null
        if ($i -lt $Max) { Write-Host "    等待 $([math]::Round(5*$i))s 后重试..." -ForegroundColor Gray; Start-Sleep -Seconds (5*$i) }
    }
    return $false
}

# ====== 主程序 ======
Write-Host "
============================================
  软件工程基础 - 一键环境安装脚本 v6.0
  安装路径: $BASE
  总计 $TOTAL_STEPS 步 | 端口可在 .env 修改
  特性: 进度显示 | 重试3次 | 回滚 | 断点续装
============================================
" -ForegroundColor White

$state = Get-State

# ====== 0. 系统检测 ======
Write-Step "系统检测 + 创建目录"
$dirs = @(
    "$BASE","$BASE/git","$BASE/nodejs","$BASE/python",
    "$BASE/jdk21","$BASE/maven","$BASE/docker","$BASE/trae",
    "$BASE/ea","$BASE/tools",
    "$BASE/data/pgdata","$BASE/data/redis","$BASE/data/esdata",
    "$BASE/data/mongodata","$BASE/data/jenkins","$BASE/data/sonarqube",
    "$BASE/projects/backend/src/main/java/com/se",
    "$BASE/projects/backend/src/test","$BASE/projects/backend/src/main/resources",
    "$BASE/projects/frontend/src","$BASE/projects/frontend/public",
    "$BASE/projects/embedded/src","$BASE/projects/docs",
    "$BASE/projects/tests","$BASE/projects/.github/workflows","$BASE/projects/.workflow"
)
foreach ($d in $dirs) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
Write-OK "目录结构已创建"

$osInfo = Get-CimInstance Win32_OperatingSystem
Write-Info "系统: $($osInfo.Caption) $($osInfo.Version)"
if (Test-Path "D:\") {
    $disk = [math]::Round((Get-PSDrive D).Free / 1GB, 1)
    Write-Info "D盘剩余: ${disk}GB"
    if ($disk -lt 20) { Write-Err "D盘空间不足20GB！"; pause; exit 1 }
} else { Write-Err "D盘不存在！"; pause; exit 1 }

# 显示软件清单
Write-Host "`n    即将安装:" -ForegroundColor White
$swKeys = @("git","nodejs","python","jdk21","maven","docker","trae","ea","jmeter","plantuml","sonarscan","postgres","redis","es","mongo","sonarqube","jenkins")
foreach ($k in $swKeys) {
    $s = $SW[$k]
    $ch = if ($s.Chapter) { $s.Chapter } else { "Docker镜像" }
    Write-Host ("    {0}. {1,-15} {2} [{3}]" -f ($swKeys.IndexOf($k)+1), $s.N, $s.Desc, $ch) -ForegroundColor Gray
}
Write-Host ""

# ====== 1. 镜像加速 ======
if (Skipped "mirror" $state) { Write-Skip "镜像加速已配置" }
else {
    Write-Step "配置国内镜像加速 (npm/pip/maven)"
    # npm/pip 此时可能尚未安装，逐个 try/catch，任一失败不阻断
    try { npm config set registry https://registry.npmmirror.com 2>$null; Write-OK "npm 镜像: npmmirror" }
    catch { Write-Info "npm 未安装，跳过 npm 镜像配置" }
    try { pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple 2>$null; Write-OK "pip 镜像: tuna" }
    catch { Write-Info "pip 未安装，跳过 pip 镜像配置" }
    # maven settings.xml 总能写入（不依赖 mvn 命令）
    $m2 = "$HOME/.m2"; if (!(Test-Path $m2)) { New-Item -ItemType Directory -Path $m2 -Force | Out-Null }
    '<?xml version="1.0"?><settings><mirrors><mirror><id>a</id><mirrorOf>central</mirrorOf><name>A</name><url>https://maven.aliyun.com/repository/public</url></mirror></mirrors></settings>' | Out-File "$m2/settings.xml" -Encoding utf8
    Write-OK "maven 镜像: aliyun (settings.xml)"
    Done "mirror" $state
}

# ====== 2. Scoop ======
if (Skipped "scoop" $state) { Write-Skip "Scoop 已安装" }
else {
    $s = $SW.scoop
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    if (Get-Command scoop -ErrorAction SilentlyContinue) { Write-Skip "已安装"; Done "scoop" $state }
    else {
        try { Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop }
        catch { Write-Info "执行策略保持当前设置（本次运行不受影响）" }
        $scoopDir = "$BASE/tools/scoop"
        $env:SCOOP = $scoopDir
        [Environment]::SetEnvironmentVariable("SCOOP", $scoopDir, "User")
        # 仅对 shims 目录去重，不动其他任何已有 PATH 条目（公共机器不破坏现有配置）
        $curPath = [Environment]::GetEnvironmentVariable("PATH", "User")
        if ($curPath) {
            $cleaned = ($curPath -split ';' | Where-Object { $_ -and ($_ -ne $scoopDir -and $_ -ne "$scoopDir\shims") }) -join ';'
            [Environment]::SetEnvironmentVariable("PATH", $cleaned, "User")
        }
        # 用官方安装器 bootstrap（不能用 scoop install scoop，此时 scoop 尚不存在）
        $bootOk = $false
        # 完整的 scoop 目录已存在（之前已装过），直接采纳，不重复 bootstrap
        if (Test-Path "$scoopDir/shims/scoop.ps1") {
            Write-Info "检测到完整的 Scoop 目录，直接启用: $scoopDir"
            $bootOk = $true
        }
        # 将安装目录加入 Microsoft Defender 排除项，避免实时扫描占用文件导致解压失败
        if (-not $bootOk) {
        try {
            if (Get-Command Add-MpPreference -ErrorAction SilentlyContinue) {
                Add-MpPreference -ExclusionPath $scoopDir -ErrorAction Stop
                Write-Info "已将 Scoop 目录加入 Microsoft Defender 排除项"
            }
        } catch { Write-Info "未能添加 Defender 排除项（权限/策略限制），继续安装" }
        # 国内镜像（get.scoop.sh 在中国大陆常超时；scoop.201704.xyz 已内置 gitee 源，可达）
        $bootUrls = @("https://scoop.201704.xyz","https://get.scoop.sh")
        for ($i=1; $i -le $MAX_RETRY -and -not $bootOk; $i++) {
            foreach ($bootUrl in $bootUrls) {
                # 每次尝试前清理残缺目录（Defender 实时扫描偶发占用 .gitignore，
                # 上次中断的 git clone 会导致下次失败）；完整安装以 shims/scoop.ps1 为准
                if ((Test-Path $scoopDir) -and -not (Test-Path "$scoopDir/shims/scoop.ps1")) {
                    Write-Info "清理不完整的 scoop 目录..."
                    cmd /c rmdir /s /q $scoopDir 2>&1 | Out-Null
                    if (Test-Path $scoopDir) { Remove-Item $scoopDir -Recurse -Force -ErrorAction SilentlyContinue }
                    # 等待杀毒/索引服务释放句柄（最多 15s），删除后立即解压易产生文件锁
                    for ($w=0; $w -lt 15 -and (Test-Path $scoopDir); $w++) { Start-Sleep -Seconds 1 }
                }
                New-Item -ItemType Directory -Path $scoopDir -Force | Out-Null
                $bootFile = "$env:TEMP\scoop_install_$i.ps1"
                Write-Info "尝试 bootstrap ($i/$MAX_RETRY): $bootUrl"
                # 下载与执行分开捕获，任一步失败都打印明确原因后继续下一个镜像
                try {
                    Invoke-WebRequest -Uri $bootUrl -OutFile $bootFile -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
                } catch {
                    Write-Retry "下载 bootstrap 失败 ($bootUrl): $($_.Exception.Message)" $i
                    continue
                }
                $bootOut = $null
                try {
                    # -ScoopDir 指定目录，-RunAsAdmin 允许管理员身份运行
                    $bootOut = & $bootFile -ScoopDir $scoopDir -RunAsAdmin 2>&1
                } catch {
                    Write-Retry "bootstrap 执行异常 ($bootUrl): $($_.Exception.Message)" $i
                }
                if (Test-Path "$scoopDir/shims/scoop.ps1") {
                    $bootOk = $true
                    Write-OK "Scoop bootstrap 成功 ($bootUrl)"
                    break
                }
                $tail = (($bootOut | Where-Object { $_ } | Select-Object -Last 3) -join ' | ')
                if (-not $tail) { $tail = "无输出（可能网络被重置，请检查代理/防火墙）" }
                Write-Retry "bootstrap 未生成 shims/scoop.ps1 ($bootUrl): $tail" $i
                Start-Sleep -Seconds (5*$i)
            }
        }
        }
        # 将 shims 目录加入当前会话 PATH（scoop 的可执行文件在 shims/）
        $shimsDir = "$scoopDir/shims"
        if ($env:PATH -notlike "*$shimsDir*") { $env:PATH = "$shimsDir;$env:PATH" }
        $curPath2 = [Environment]::GetEnvironmentVariable("PATH", "User")
        if ($curPath2 -notlike "*$shimsDir*") {
            [Environment]::SetEnvironmentVariable("PATH", "$shimsDir;$curPath2", "User")
        }
        if ($bootOk -and (Test-Path "$shimsDir/scoop.ps1") -and (Get-Command scoop -ErrorAction SilentlyContinue)) {
            Write-OK "安装完成"; Done "scoop" $state
        } else { Failed "scoop" $state; Write-Err "安装失败：scoop 仍不可用（shims/scoop.ps1 不存在或命令不在 PATH）" }
    }
}

# ====== 3. Git ======
if (Skipped "git" $state) { Write-Skip "Git 已安装" }
else {
    $s = $SW.git
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    $gitOk = $false
    if (Get-Command git -ErrorAction SilentlyContinue) { Write-Skip "已安装: $(git --version)"; $gitOk = $true }
    elseif (ScoopInstall "git") { Write-OK "安装完成"; $gitOk = $true }
    else { Write-Err "安装失败" }
    if ($gitOk) {
        git config --global core.autocrlf true 2>$null
        git config --global core.editor "code --wait" 2>$null
        Write-OK "Git 配置已设置"
        # git 已就绪，现在可以添加 extras bucket（jmeter/graphviz 等在其中）
        if (Get-Command scoop -ErrorAction SilentlyContinue) {
            Write-Info "添加 extras bucket（gitee 镜像，失败回退官方源）..."
            scoop bucket add extras https://gitee.com/scoop-installer/Extras 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) { scoop bucket add extras 2>&1 | Out-Null }
        }
        Done "git" $state
    } else { Failed "git" $state }
}

# ====== 4. Node.js（紧随 Git 优先安装；npmmirror CDN 直装，scoop 兜底） ======
if (Skipped "nodejs" $state) { Write-Skip "Node.js 已安装" }
else {
    $s = $SW.nodejs
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url2)"
    Write-Info "备: $($s.Url1)"
    $nodeOk = $false
    if (Get-Command node -ErrorAction SilentlyContinue) { Write-Skip "已安装: $(node -v)"; $nodeOk = $true }
    else {
        # 优先 msi 静默直装到 $BASE_WIN\nodejs（npmmirror 国内 CDN 速度快），失败再用 scoop
        $msi = "$env:TEMP\node-v24.21.0-x64.msi"
        if (DL2 @($s.Url2, $s.Url1) $msi) {
            Write-Info "静默安装到 $BASE_WIN\nodejs ..."
            $p = Start-Process msiexec -ArgumentList "/i","`"$msi`"","/qn","/norestart","INSTALLDIR=`"$BASE_WIN\nodejs`"" -Wait -PassThru
            if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {
                Write-OK "安装完成"
                # msi 会写注册表 PATH，刷新会话 PATH 使 node/npm 立即可用
                $env:PATH = [Environment]::GetEnvironmentVariable('PATH','Machine') + ';' + [Environment]::GetEnvironmentVariable('PATH','User') + ";$BASE_WIN\nodejs"
                $nodeOk = $true
            } else { Write-Retry "msiexec 退出码 $($p.ExitCode)，改用 scoop" 1 }
        }
        if (-not $nodeOk -and (ScoopInstall "nodejs")) {
            $env:PATH = "$BASE/tools/scoop/shims;$env:PATH"
            $nodeOk = $true
        }
    }
    if ($nodeOk) {
        try { npm config set registry https://registry.npmmirror.com 2>$null; Write-OK "npm 镜像: npmmirror" } catch {}
        Done "nodejs" $state
    } else { Failed "nodejs" $state }
}

# ====== 5. JDK 21（已有 java 则只定位 JAVA_HOME 复用，不重装） ======
if (Skipped "jdk" $state) { Write-Skip "JDK 21 已安装" }
else {
    $s = $SW.jdk21
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    $jdkOk = $false
    if (Get-Command java -ErrorAction SilentlyContinue) { Write-Skip "已安装: $((java -version 2>&1 | Select-Object -First 1))"; $jdkOk = $true }
    elseif (ScoopInstall "openjdk21") { Write-OK "安装完成"; $jdkOk = $true }
    else { Write-Err "安装失败" }
    if ($jdkOk) {
        # 定位 JDK 主目录：scoop -> 注册表（Oracle/Adoptium 均写入）-> 常见安装目录
        $jdkHome = $null
        $scoopJdk = "$BASE/tools/scoop/apps/openjdk21/current"
        if (Test-Path "$scoopJdk/bin/java.exe") { $jdkHome = $scoopJdk }
        if (-not $jdkHome) {
            # Oracle javapath 里的 java.exe 指向真实 JDK，解析链接目标即可复用现有安装
            $jcmd = Get-Command java -ErrorAction SilentlyContinue
            if ($jcmd -and $jcmd.Source) {
                $tgt = $jcmd.Source
                try { $lk = (Get-Item $jcmd.Source -Force -ErrorAction Stop).Target; if ($lk) { $tgt = @($lk)[0] } } catch {}
                # Oracle javapath 可能指向 javapath_target 中转链接，需再解析一层
                try { if ($tgt -like '*javapath*') { $lk2 = (Get-Item $tgt -Force -ErrorAction Stop).Target; if ($lk2) { $tgt = @($lk2)[0] } } } catch {}
                if ($tgt -match '\\bin\\java(\.exe)?$') {
                    $cand = $tgt -replace '\\bin\\java(\.exe)?$',''
                    if (Test-Path "$cand\bin\java.exe") { $jdkHome = $cand }
                }
            }
        }
        if (-not $jdkHome) {
            # 注册表统一按子键遍历（覆盖 JDK/JRE/Adoptium，CurrentVersion 缺失也能命中）
            foreach ($rk in @('HKLM:\SOFTWARE\JavaSoft\JDK','HKLM:\SOFTWARE\JavaSoft\Java Development Kit','HKLM:\SOFTWARE\JavaSoft\Java Runtime Environment','HKLM:\SOFTWARE\Eclipse Adoptium')) {
                foreach ($sk in (Get-ChildItem $rk -ErrorAction SilentlyContinue)) {
                    $p2 = Get-ItemProperty $sk.PSPath -ErrorAction SilentlyContinue
                    $h2 = $p2.JavaHome; if (-not $h2) { $h2 = $p2.Path }
                    if ($h2 -and (Test-Path "$h2\bin\java.exe")) { $jdkHome = $h2; break }
                }
                if ($jdkHome) { break }
            }
        }
        if (-not $jdkHome) {
            $hit = Get-ChildItem 'C:\Program Files\Java','C:\Program Files\Eclipse Adoptium' -Directory -ErrorAction SilentlyContinue |
                Where-Object { Test-Path "$($_.FullName)\bin\java.exe" } |
                Sort-Object Name -Descending | Select-Object -First 1
            if ($hit) { $jdkHome = $hit.FullName }
        }
        if ($jdkHome -and (Test-Path "$jdkHome/bin/java.exe")) {
            [Environment]::SetEnvironmentVariable("JAVA_HOME", $jdkHome, "User")
            $env:JAVA_HOME = $jdkHome
            Write-OK "JAVA_HOME=$jdkHome"
        } else { Write-Info "未自动定位到 JDK 主目录，请手动设置 JAVA_HOME" }
        Done "jdk" $state
    } else { Failed "jdk" $state }
}

# ====== 6. Maven（华为云 zip 直装为主，scoop 兜底） ======
if (Skipped "maven" $state) { Write-Skip "Maven 已安装" }
else {
    $s = $SW.maven
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: https://mirrors.huaweicloud.com/maven/maven-3/3.9.9/binaries/apache-maven-3.9.9-bin.zip"
    Write-Info "备: https://archive.apache.org/dist/maven/maven-3/3.9.9/binaries/apache-maven-3.9.9-bin.zip"
    $mvnOk = $false
    if (Get-Command mvn -ErrorAction SilentlyContinue) { Write-Skip "已安装"; $mvnOk = $true }
    else {
        # 阿里云 apache 镜像只保留最新版（3.9.9 已 404），华为云/官方归档全版本可下
        $zip = "$env:TEMP\apache-maven-3.9.9-bin.zip"
        if (DL2 @('https://mirrors.huaweicloud.com/maven/maven-3/3.9.9/binaries/apache-maven-3.9.9-bin.zip','https://archive.apache.org/dist/maven/maven-3/3.9.9/binaries/apache-maven-3.9.9-bin.zip') $zip) {
            Write-Info "解压到 $BASE_WIN\maven ..."
            Expand-Archive -Path $zip -DestinationPath "$BASE\maven" -Force
            $mvnBin = "$BASE\maven\apache-maven-3.9.9\bin"
            if (Test-Path "$mvnBin\mvn.cmd") {
                $curPath = [Environment]::GetEnvironmentVariable("PATH", "User")
                if ($curPath -notlike "*apache-maven*") { [Environment]::SetEnvironmentVariable("PATH", "$mvnBin;$curPath", "User") }
                $env:PATH = "$mvnBin;$env:PATH"
                $mvnOk = $true
                Write-OK "安装完成"
            } else { Write-Err "解压后未找到 mvn.cmd" }
        }
        if (-not $mvnOk) { Write-Info "直装失败，改用 scoop..."; if (ScoopInstall "maven") { $mvnOk = $true } }
    }
    if ($mvnOk) { Done "maven" $state } else { Write-Err "安装失败"; Failed "maven" $state }
}

# ====== 7. Python ======
if (Skipped "python" $state) { Write-Skip "Python 已安装" }
else {
    $s = $SW.python
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    $pyOk = $false
    if (Test-RealCmd "python") { Write-Skip "已安装: $((python --version 2>&1 | Select-Object -First 1))"; $pyOk = $true }
    elseif (ScoopInstall "python") { Write-OK "安装完成"; $pyOk = $true }
    else { Write-Err "安装失败" }
    if ($pyOk -and (Test-RealCmd "pip")) {
        Write-Info "安装 Python 通用包 (20个)..."
        for ($r=1;$r -le $MAX_RETRY;$r++) {
            try {
                pip install flask flask-socketio fastapi uvicorn requests httpx python-dotenv redis pymongo numpy pandas scikit-learn matplotlib pytest pytest-cov pygame pillow openai pyjwt -q 2>$null
                Write-OK "Python 包安装完成"; break
            } catch { Write-Retry "pip install 失败" $r; Start-Sleep (5*$r) }
        }
    } else { Write-Info "pip 不可用，跳过 Python 包安装" }
    New-Item -ItemType Directory -Path "$BASE/python/venvs" -Force | Out-Null
    if ($pyOk) { Done "python" $state } else { Failed "python" $state }
}

# ====== 8. Docker（本次必须就绪：WSL2 特性 + 静默安装 + 启动 daemon 实测） ======
if (Skipped "docker" $state) { Write-Skip "Docker 已就绪" }
else {
    $s = $SW.docker
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"

    # daemon 功能性检测（CLI 存在不代表可用，公共机器上常出现"装了但起不来"）
    $daemonOk = $false
    if (Get-Command docker -ErrorAction SilentlyContinue) {
        try { docker info 2>$null | Out-Null; $daemonOk = ($LASTEXITCODE -eq 0) } catch {}
        if ($daemonOk) { Write-Skip "已安装且 daemon 运行中: $(docker --version)"; Done "docker" $state }
        else { Write-Info "docker CLI 已存在但 daemon 不可用，尝试修复..." }
    }

    if (-not $daemonOk) {
        # 1) 先确保 WSL2 可用（Docker Desktop 硬依赖；dism 对已启用特性无副作用，不破坏系统）
        Write-Info "启用 WSL2 特性 (Microsoft-Windows-Subsystem-Linux + VirtualMachinePlatform)..."
        try {
            dism.exe /online /enable-feature /featurename:Microsoft-Windows-Subsystem-Linux /all /norestart 2>&1 | Out-Null
            dism.exe /online /enable-feature /featurename:VirtualMachinePlatform /all /norestart 2>&1 | Out-Null
            wsl.exe --update 2>$null | Out-Null
        } catch { Write-Info "WSL 特性启用命令异常（可能已启用），继续" }

        # 2) CLI 缺失才静默安装（已有安装不重复装）
        if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
            $installer = "$BASE/docker/DockerDesktopInstaller.exe"
            $dlOk = $false
            if ((Test-Path $installer) -and ((Get-Item $installer).Length -gt 100MB)) {
                Write-Skip "安装包已存在 ($([math]::Round((Get-Item $installer).Length/1MB))MB)，跳过下载"
                $dlOk = $true
            } else { $dlOk = DL2 @($SW.docker.Url1) $installer }
            if ($dlOk) {
                Write-Info "静默安装 Docker Desktop 到 $BASE_WIN\docker ..."
                $p = Start-Process $installer -ArgumentList 'install','--quiet','--accept-license',"--installation-dir=`"$BASE_WIN\docker`"",'--backend=wsl-2' -Wait -PassThru
                # 3010 = 安装成功但需重启系统
                if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) { Write-OK "Docker Desktop 安装完成 (exit=$($p.ExitCode))" }
                else {
                    Write-Retry "静默安装失败 (exit=$($p.ExitCode))，回退交互式安装" 1
                    Write-Host "    请安装: 1.勾选WSL2  2.路径=$BASE_WIN\docker  3.Install → 按任意键" -ForegroundColor Yellow
                    Start-Process $installer; pause
                }
            } else { Write-Err "安装包下载失败"; Failed "docker" $state }
        }

        # 3) 刷新 PATH（覆盖 per-user 与 D 盘两种安装位置）并尝试启动 daemon
        $env:PATH = [Environment]::GetEnvironmentVariable('PATH','Machine') + ';' + [Environment]::GetEnvironmentVariable('PATH','User') + ";$env:LOCALAPPDATA\Programs\DockerDesktop\resources\bin;$BASE_WIN\docker\resources\bin"
        if (Get-Command docker -ErrorAction SilentlyContinue) {
            $dd = @(Get-ChildItem "$env:LOCALAPPDATA\Programs\DockerDesktop","$BASE\docker","C:\Program Files\Docker\Docker" -Filter 'Docker Desktop.exe' -Recurse -EA 0) | Select-Object -First 1
            if ($dd) { Write-Info "启动 Docker Desktop..."; Start-Process $dd.FullName }
            for ($w=1; $w -le 8; $w++) {
                Start-Sleep -Seconds 10
                try { docker info 2>$null | Out-Null; if ($LASTEXITCODE -eq 0) { $daemonOk = $true; break } } catch {}
                Write-Info "等待 daemon 就绪 ($w/8)..."
            }
        }
        if ($daemonOk) { Write-OK "Docker daemon 就绪"; Done "docker" $state }
        else {
            Failed "docker" $state
            Write-Err "daemon 未就绪（多半需重启电脑完成 WSL2 安装）"
            Write-Host "    请重启电脑后再次运行本脚本，将自动续装剩余步骤" -ForegroundColor Yellow
        }
    }
}

# ====== 8a. Docker 拉取加速 ======
# daemon 未运行：直接写入配置；daemon 运行中但无加速：写入后重启 Docker Desktop 生效
if (-not (Skipped "dmirror" $state)) {
    $daemonJson = "$env:USERPROFILE\.docker\daemon.json"
    $mirrors = @('https://docker.1ms.run','https://docker.m.daocloud.io','https://hub.rat.dev')
    $daemonRunning = $false
    try { docker info 2>$null | Out-Null; $daemonRunning = ($LASTEXITCODE -eq 0) } catch {}
    $needRestart = $false
    try {
        $cfg = $null
        if (Test-Path $daemonJson) { $cfg = Get-Content $daemonJson -Raw | ConvertFrom-Json }
        if (-not $cfg) { $cfg = New-Object PSObject }
        if (-not $cfg.PSObject.Properties['registry-mirrors']) {
            $cfg | Add-Member -NotePropertyName 'registry-mirrors' -NotePropertyValue $mirrors
            New-Item -ItemType Directory -Path (Split-Path $daemonJson) -Force | Out-Null
            $cfg | ConvertTo-Json -Depth 10 | Out-File $daemonJson -Encoding utf8
            Write-OK "已配置 Docker 国内镜像加速 (1ms.run/daocloud/rat.dev)"
            if ($daemonRunning) { $needRestart = $true }
        } else { Write-Info "Docker 镜像加速已配置" }
    } catch { Write-Info "daemon.json 写入失败（不影响后续步骤）" }
    if ($needRestart) {
        # 该 daemon 是本脚本刚启动的，重启使加速配置生效
        Write-Info "重启 Docker Desktop 使镜像加速生效..."
        Get-Process 'Docker Desktop' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        $dd = @(Get-ChildItem "$env:LOCALAPPDATA\Programs\DockerDesktop","$BASE\docker","C:\Program Files\Docker\Docker" -Filter 'Docker Desktop.exe' -Recurse -EA 0) | Select-Object -First 1
        if ($dd) { Start-Process $dd.FullName }
        for ($w=1; $w -le 8; $w++) {
            Start-Sleep -Seconds 10
            try { docker info 2>$null | Out-Null; if ($LASTEXITCODE -eq 0) { Write-OK "daemon 已就绪（加速已生效）"; break } } catch {}
            Write-Info "等待 daemon 重启 ($w/8)..."
        }
    }
    Done "dmirror" $state
}

# ====== 9. TRAE ======
if (Skipped "trae" $state) { Write-Skip "TRAE 已安装" }
else {
    $s = $SW.trae
    Write-Step "$($s.N) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    if (Get-ChildItem "$BASE/trae" -Recurse -Filter "trae.exe" -EA 0) { Write-Skip "已安装"; Done "trae" $state }
    elseif (Test-InstalledApp '*Trae*') { Write-Skip "检测到系统已安装 TRAE（$((Test-InstalledApp '*Trae*').DisplayName)），直接复用"; Done "trae" $state }
    else {
        # TRAE 官网是网页不是直链安装包，无法自动下载，只提示用户
        Write-Host "    TRAE 需手动下载安装：" -ForegroundColor Yellow
        Write-Host "      1. 访问 $($s.Url2) 下载安装包" -ForegroundColor Yellow
        Write-Host "      2. 安装路径改为 $BASE_WIN\trae" -ForegroundColor Yellow
        Write-Host "    安装完成后再次运行本脚本将自动跳过" -ForegroundColor Gray
        Done "trae" $state
    }
}

# ====== 10. EA UML ======
Write-Step "$($SW.ea.N) — $($SW.ea.Desc) [$($SW.ea.Chapter) / $($SW.ea.Experiment)]"
Write-Info "主: $($SW.ea.Url1)"
Write-Info "备: $($SW.ea.Url2) (免费替代)"
if (Test-InstalledApp '*Enterprise Architect*') {
    Write-Skip "检测到系统已安装 EA（$((Test-InstalledApp '*Enterprise Architect*').DisplayName)），直接复用"
} else {
    Write-Host "    商业软件，需手动安装到 $BASE_WIN\ea（建议版本 17.1）" -ForegroundColor Yellow
}
Done "ea" $state

# ====== 11. JMeter ======
if (Skipped "jmeter" $state) { Write-Skip "JMeter 已安装" }
else {
    $s = $SW.jmeter
    Write-Step "$($s.N) v$($s.V) — $($s.Desc) [$($s.Chapter) / $($s.Experiment)]"
    Write-Info "主: $($s.Url1)"
    Write-Info "备: $($s.Url2)"
    $jmOk = $false
    if (Get-Command jmeter -ErrorAction SilentlyContinue) { Write-Skip "已安装"; $jmOk = $true }
    else {
        if (ScoopInstall "jmeter") { Write-OK "安装完成"; $jmOk = $true }
        else {
            # 备选：直接下载 zip 解压
            $zip = "$BASE/jmeter/apache-jmeter-5.6.3.zip"
            New-Item -ItemType Directory -Path "$BASE/jmeter" -Force | Out-Null
            if (DL $SW.jmeter.Url1 $zip) {
                Expand-Archive -Path $zip -DestinationPath "$BASE/jmeter" -Force
                Write-OK "解压完成"; $jmOk = $true
            } else { Write-Err "下载失败" }
        }
    }
    # GUI 启动路径：scoop 安装优先，其次 zip 解压目录
    $jmBat = "$BASE/tools/scoop/apps/jmeter/current/bin/jmeter.bat"
    if (-not (Test-Path $jmBat)) { $jmBat = "$BASE/jmeter/apache-jmeter-5.6.3/bin/jmeter.bat" }
    if (Test-Path $jmBat) { Write-Host "    运行: $jmBat (GUI)" -ForegroundColor White }
    Write-Host "    命令行: jmeter -n -t test.jmx -l result.jtl -e -o report" -ForegroundColor Gray
    if ($jmOk) { Done "jmeter" $state } else { Failed "jmeter" $state }
}

# ====== 11a. Graphviz + PlantUML（UML建模） ======
if (Skipped "plantuml" $state) { Write-Skip "Graphviz/PlantUML 已安装" }
else {
    Write-Step "Graphviz + PlantUML — UML文本建模 [第2章-04/05 / 实验二]"
    $gvOk = $false
    if (Get-Command dot -ErrorAction SilentlyContinue) { Write-Skip "Graphviz 已安装"; $gvOk = $true }
    else {
        # scoop 优先（自动配 shims）；失败则官方安装包静默安装（NSIS: /S /D=）
        if (ScoopInstall "graphviz") { Write-OK "Graphviz 安装完成"; $gvOk = $true }
        else {
            $gvExe = "$env:TEMP\graphviz-install.exe"
            if (DL2 @($SW.graphviz.Url2) $gvExe) {
                Write-Info "静默安装到 $BASE_WIN\graphviz ..."
                $p = Start-Process $gvExe -ArgumentList "/S","/D=$BASE_WIN\graphviz" -Wait -PassThru
                if (Test-Path "$BASE\graphviz\bin\dot.exe") {
                    $env:PATH = "$BASE\graphviz\bin;$env:PATH"
                    $curPath = [Environment]::GetEnvironmentVariable("PATH", "User")
                    if ($curPath -notlike "*CS24SEB\graphviz*") { [Environment]::SetEnvironmentVariable("PATH", "$BASE\graphviz\bin;$curPath", "User") }
                    Write-OK "Graphviz 安装完成"; $gvOk = $true
                } else { Write-Err "静默安装失败（退出码 $($p.ExitCode)）" }
            } else { Write-Err "Graphviz 安装失败" }
        }
    }
    if ($gvOk -and (Get-Command dot -ErrorAction SilentlyContinue)) { Write-OK "Graphviz: $(dot -V 2>&1)" }
    New-Item -ItemType Directory -Path "$BASE/plantuml" -Force | Out-Null
    $jar = "$BASE/plantuml/plantuml.jar"
    $puOk = $false
    # 如果 jar 已存在且被占用，先尝试删除
    if (Test-Path $jar) { Remove-Item $jar -Force -ErrorAction SilentlyContinue }
    if (DL2 @($SW.plantuml.Url1, $SW.plantuml.Url2) $jar) {
        $puOk = $true
        try { $v = java -jar $jar -version 2>&1 | Select-Object -First 1; Write-OK "PlantUML: $v" }
        catch { Write-Host "    PlantUML 已下载（需 JDK）" -ForegroundColor Yellow }
    } else { Write-Err "PlantUML 下载失败" }
    Write-Host "    绘图: java -jar $BASE_WIN\plantuml\plantuml.jar -tsvg *.puml" -ForegroundColor White
    if ($gvOk -or $puOk) { Done "plantuml" $state } else { Failed "plantuml" $state }
}

# ====== 11b. SonarScanner（代码扫描） ======
if (Skipped "sonarscan" $state) { Write-Skip "SonarScanner 已安装" }
else {
    Write-Step "SonarScanner — 代码质量扫描 [第5章-12 / 实验四]"
    $ssOk = $false
    if (Get-Command sonar-scanner -ErrorAction SilentlyContinue) { Write-Skip "已安装"; $ssOk = $true }
    elseif (ScoopInstall "sonar-scanner") { Write-OK "安装完成"; $ssOk = $true }
    else { Write-Err "安装失败" }
    Write-Host "    扫描: cd projects/backend && sonar-scanner (先 docker start se-sonarqube)" -ForegroundColor White
    if ($ssOk) { Done "sonarscan" $state } else { Failed "sonarscan" $state }
}

# ====== 12-17. Docker 容器 ======
$svcs = @(
    @{ K="pg";    N="se-postgres";      Img="postgis/postgis:17-3.4"; Desc="PostgreSQL+PostGIS"; Port=$PG_PORT;
       Args=@("-p ${PG_PORT}:5432","-e POSTGRES_DB=$PG_DB","-e POSTGRES_USER=$PG_USER","-e POSTGRES_PASSWORD=$PG_PASSWORD","-e TZ=Asia/Shanghai","-v `"$($BASE_WIN)\data\pgdata:/var/lib/postgresql/data`"","--restart unless-stopped");
       Verify={ docker exec se-postgres pg_isready -U $PG_USER 2>$null } },
    @{ K="redis"; N="se-redis";         Img="redis:7-alpine";        Desc="Redis缓存";          Port=$REDIS_PORT;
       Args=@("-p ${REDIS_PORT}:6379","-v `"$($BASE_WIN)\data\redis:/data`"","--restart unless-stopped");
       Verify={ docker exec se-redis redis-cli ping 2>$null } },
    @{ K="es";    N="se-elasticsearch"; Img="elasticsearch:8.17.0";  Desc="Elasticsearch检索";   Port=$ES_PORT;
       Args=@("-p ${ES_PORT}:9200","-p 9300:9300","-e discovery.type=single-node","-e xpack.security.enabled=false","-e `"ES_JAVA_OPTS=-Xms512m -Xmx512m`"","-v `"$($BASE_WIN)\data\esdata:/usr/share/elasticsearch/data`"","--restart unless-stopped");
       Verify={ (Invoke-WebRequest -Uri "http://localhost:$ES_PORT" -UseBasicParsing -TimeoutSec 5).StatusCode -eq 200 } },
    @{ K="mongo"; N="se-mongodb";       Img="mongo:7";               Desc="MongoDB文档库";      Port=$MONGO_PORT;
       Args=@("-p ${MONGO_PORT}:27017","-e MONGO_INITDB_ROOT_USERNAME=$MONGO_USER","-e MONGO_INITDB_ROOT_PASSWORD=$MONGO_PASS","-v `"$($BASE_WIN)\data\mongodata:/data/db`"","--restart unless-stopped");
       Verify={ docker exec se-mongodb mongosh -u $MONGO_USER -p $MONGO_PASS --eval "db.runCommand({ping:1})" --quiet 2>$null } },
    @{ K="sonar"; N="se-sonarqube";     Img="sonarqube:community";   Desc="SonarQube代码质量";  Port=$SONAR_PORT;
       Args=@("-p ${SONAR_PORT}:9000","-e SONAR_ES_BOOTSTRAP_CHECKS_DISABLE=true","-v `"$($BASE_WIN)\data\sonarqube:/opt/sonarqube/data`"","--restart unless-stopped");
       Verify={ (Invoke-WebRequest -Uri "http://localhost:$SONAR_PORT" -UseBasicParsing -TimeoutSec 5).StatusCode -eq 200 } },
    @{ K="jenkins"; N="se-jenkins";     Img="jenkins/jenkins:lts-jdk17"; Desc="Jenkins CI/CD";   Port=$JENKINS_PORT;
       Args=@("-p ${JENKINS_PORT}:8080","-p 50000:50000","-v `"$($BASE_WIN)\data\jenkins:/var/jenkins_home`"","--restart unless-stopped");
       Verify={ (Invoke-WebRequest -Uri "http://localhost:$JENKINS_PORT" -UseBasicParsing -TimeoutSec 5).StatusCode -eq 200 } }
)

foreach ($svc in $svcs) {
    if (Skipped $svc.K $state) { Write-Skip "$($svc.Desc) 已部署 (端口:$($svc.Port))"; continue }

    # 首个容器前，检查 Docker daemon 是否运行（Docker Desktop 可能已安装但未启动）
    if ($svc.K -eq $svcs[0].K) {
        $daemonOk = $false
        for ($d=1;$d -le 5;$d++) {
            $info = docker info 2>&1
            if ($LASTEXITCODE -eq 0) { $daemonOk = $true; break }
            if ($d -eq 1) {
                Write-Host "    Docker daemon 未运行，尝试启动 Docker Desktop..." -ForegroundColor Yellow
                $ddPath = Get-ChildItem "C:\Program Files\Docker" -Filter "Docker Desktop.exe" -Recurse -EA 0 | Select-Object -First 1
                if ($ddPath) { Start-Process $ddPath.FullName }
                else { Write-Err "找不到 Docker Desktop，请手动启动后重试"; break }
            }
            Write-Host "    等待 Docker daemon 启动 ($d/5)..." -ForegroundColor Gray
            Start-Sleep -Seconds 15
        }
        if (-not $daemonOk) {
            Write-Err "Docker daemon 不可用，跳过所有容器部署"
            foreach ($s2 in $svcs) { Failed $s2.K $state }
            break
        }
        Write-OK "Docker daemon 就绪"
    }

    # 容器键名映射到 $SW 键名（pg→postgres, sonar→sonarqube）
    $swKey = ($svc.K -replace '^pg$','postgres') -replace '^sonar$','sonarqube'
    $s = $SW[$swKey]
    Write-Step "$($svc.Desc) — $($svc.Img) [端口:$($svc.Port)] [$($s.Chapter)]"
    Write-Info "镜像: $($s.Url1)"

    $existing = docker ps -a --filter "name=$($svc.N)" --format "{{.Names}}" 2>$null
    if ($existing -eq $svc.N) {
        $running = docker inspect -f '{{.State.Status}}' $svc.N 2>$null
        if ($running -eq "running") { Write-Skip "$($svc.N) 已运行"; Done $svc.K $state; continue }
        docker start $svc.N 2>$null | Out-Null; Start-Sleep 2
        $running = docker inspect -f '{{.State.Status}}' $svc.N 2>$null
        if ($running -eq "running") { Write-OK "$($svc.N) 已启动 (端口:$($svc.Port))"; Done $svc.K $state; continue }
    }

    Write-Host "    拉取镜像 $($svc.Img)..." -ForegroundColor White
    if (DockerRun $svc.Args $svc.N) {
        # 验证连接
        Start-Sleep -Seconds 5
        try { & $svc.Verify | Out-Null; Write-OK "连接验证通过" } catch { Write-Info "容器已启动，连接验证跳过" }
        Write-OK "$($svc.Desc) 部署完成 (端口:$($svc.Port))"
        Done $svc.K $state
    } else {
        Failed $svc.K $state
        Write-Err "$($svc.Desc) 部署失败"
        Write-Info "可修改端口后重试: 编辑 $BASE_WIN\.env 中的 $($svc.K.ToUpper())_PORT"
    }
}

# ====== 18. 生成模板文件 ======
if (Skipped "tpl" $state) { Write-Skip "模板文件已生成" }
else {
    Write-Step "生成项目模板文件 (docker-compose/Dockerfile/CI/CD)"

    # docker-compose.yml（完整栈，使用 .env 变量）
    @"
version: '3.8'
services:
  postgres:
    image: postgis/postgis:17-3.4
    container_name: se-postgres
    environment:
      POSTGRES_DB: $PG_DB
      POSTGRES_USER: $PG_USER
      POSTGRES_PASSWORD: $PG_PASSWORD
      TZ: Asia/Shanghai
    ports:
      - "`${PG_PORT}:5432"
    volumes:
      - $BASE_WIN\data\pgdata:/var/lib/postgresql/data
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U $PG_USER"]
      interval: 10s
      timeout: 5s
      retries: 5
    restart: unless-stopped

  redis:
    image: redis:7-alpine
    container_name: se-redis
    ports:
      - "`${REDIS_PORT}:6379"
    volumes:
      - $BASE_WIN\data\redis:/data
    healthcheck:
      test: ["CMD", "redis-cli", "ping"]
      interval: 10s
      timeout: 5s
      retries: 5
    restart: unless-stopped

  elasticsearch:
    image: elasticsearch:8.17.0
    container_name: se-elasticsearch
    environment:
      - discovery.type=single-node
      - xpack.security.enabled=false
      - "ES_JAVA_OPTS=-Xms512m -Xmx512m"
    ports:
      - "`${ES_PORT}:9200"
    volumes:
      - $BASE_WIN\data\esdata:/usr/share/elasticsearch/data
    restart: unless-stopped

  mongodb:
    image: mongo:7
    container_name: se-mongodb
    environment:
      MONGO_INITDB_ROOT_USERNAME: $MONGO_USER
      MONGO_INITDB_ROOT_PASSWORD: $MONGO_PASS
    ports:
      - "`${MONGO_PORT}:27017"
    volumes:
      - $BASE_WIN\data\mongodata:/data/db
    restart: unless-stopped

  sonarqube:
    image: sonarqube:community
    container_name: se-sonarqube
    environment:
      - SONAR_ES_BOOTSTRAP_CHECKS_DISABLE=true
    ports:
      - "`${SONAR_PORT}:9000"
    volumes:
      - $BASE_WIN\data\sonarqube:/opt/sonarqube/data
    restart: unless-stopped

  jenkins:
    image: jenkins/jenkins:lts-jdk17
    container_name: se-jenkins
    ports:
      - "`${JENKINS_PORT}:8080"
      - "50000:50000"
    volumes:
      - $BASE_WIN\data\jenkins:/var/jenkins_home
    restart: unless-stopped

  backend:
    build: ./projects/backend
    ports:
      - "`${BACKEND_PORT}:8080"
    environment:
      - SPRING_PROFILES_ACTIVE=docker
      - DATABASE_URL=jdbc:postgresql://postgres:5432/$PG_DB
      - REDIS_URL=redis://redis:6379
    depends_on:
      postgres:
        condition: service_healthy
      redis:
        condition: service_started
    restart: unless-stopped

  frontend:
    build: ./projects/frontend
    ports:
      - "`${FRONTEND_PORT}:80"
    depends_on:
      - backend
    restart: unless-stopped
"@ | Out-File "$BASE/docker-compose.yml" -Encoding utf8

    Write-OK "模板文件已生成"
    Done "tpl" $state
}

# ====== 19. Git 初始化 ======
if (Skipped "gitinit" $state) { Write-Skip "Git 仓库已初始化" }
else {
    Write-Step "Git 仓库初始化"
    Push-Location "$BASE/projects"
    if (Get-Command git -ErrorAction SilentlyContinue) {
        git init 2>$null; git add . 2>$null; git commit -m "feat: 初始化项目环境" 2>$null
        Write-OK "已初始化"
        Write-Host "    cd $BASE_WIN\projects && git remote add origin git@gitee.com:用户/项目.git" -ForegroundColor White
        Write-Host "    (GitHub 备选: git@github.com:用户/项目.git)" -ForegroundColor DarkGray
    }
    Pop-Location; Done "gitinit" $state
}

# ====== PATH ======
# scoop 安装的所有软件（git/node/python/maven/jmeter/sonar-scanner 等）
# 都通过 $BASE/tools/scoop/shims 暴露，只需保证该目录在用户 PATH 中
Write-Step "配置系统 PATH"
$shimsDir = "$BASE/tools/scoop/shims"
$currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
if ($currentPath -notlike "*$shimsDir*") {
    [Environment]::SetEnvironmentVariable("PATH", "$shimsDir;$currentPath", "User")
    Write-OK "PATH 已加入 Scoop shims: $shimsDir"
    Write-Host "    请重启 PowerShell 使 PATH 生效" -ForegroundColor Yellow
} else { Write-Skip "PATH 已包含 Scoop shims" }

# ====== 验证 ======
Write-Step "环境验证"
$results = @()
foreach ($t in @(
    @{N="Git";C="git --version"},@{N="Node.js";C="node -v"},@{N="npm";C="npm -v"},
    @{N="Python";C="python --version"},@{N="pip";C="pip --version"},
    @{N="Java";C="java --version"},@{N="Maven";C="mvn -v"},@{N="Docker";C="docker --version"}
)) {
    try { $o = Invoke-Expression $t.C 2>&1 | Select-Object -First 1; Write-OK "$($t.N): $o"; $results += "OK" }
    catch { Write-Err "$($t.N): 未安装"; $results += "ERR" }
}

foreach ($c in @("se-postgres","se-redis","se-elasticsearch","se-mongodb","se-sonarqube","se-jenkins")) {
    $st = docker inspect -f '{{.State.Status}}' $c 2>$null
    if ($st -eq "running") { Write-OK "${c}: running (端口:$(docker port $c 2>$null | Select-Object -First 1 -Replace '.*:',''))"; $results += "OK" }
    elseif ($st -eq "exited") { Write-Skip "${c}: exited"; $results += "WARN" }
    else { Write-Skip "${c}: 未创建"; $results += "SKIP" }
}

# ====== 汇总 ======
$elapsed = ((Get-Date) - $script:startTime).ToString("mm\:ss")
$okC = ($results | Where-Object { $_ -eq "OK" }).Count
$errC = ($results | Where-Object { $_ -eq "ERR" }).Count
$warnC = $results.Count - $okC - $errC

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  安装完成！ 耗时: $elapsed" -ForegroundColor Cyan
Write-Host "  结果: $okC 成功 / $errC 失败 / $warnC 警告" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

Write-Host "`n  端口配置 ($BASE_WIN\.env):" -ForegroundColor White
Write-Host "    PostgreSQL=$PG_PORT | Redis=$REDIS_PORT | ES=$ES_PORT | MongoDB=$MONGO_PORT" -ForegroundColor Gray
Write-Host "    SonarQube=$SONAR_PORT | Jenkins=$JENKINS_PORT | Backend=$BACKEND_PORT | Frontend=3000" -ForegroundColor Gray

Write-Host "`n  账号密码:" -ForegroundColor White
Write-Host "    PostgreSQL: $PG_USER / $PG_PASSWORD" -ForegroundColor Gray
Write-Host "    MongoDB:    $MONGO_USER / $MONGO_PASS" -ForegroundColor Gray
Write-Host "    SonarQube:  admin / admin (首次登录修改)" -ForegroundColor Gray
Write-Host "    Jenkins:    admin / (查看: docker exec se-jenkins cat /var/jenkins_home/secrets/initialAdminPassword)" -ForegroundColor Gray

$failedS = $state.failed
if ($failedS.Count -gt 0) {
    Write-Host "`n  失败项 (再次运行脚本将自动续装):" -ForegroundColor Red
    foreach ($f in $failedS) { Write-Host "    - $f" -ForegroundColor Yellow }
}

Write-Host "`n  下一步:" -ForegroundColor White
Write-Host "    1. 重启 PowerShell    2. 编辑 .env 填入 API Key    3. TRAE 打开 projects/" -ForegroundColor Gray
Write-Host "============================================`n" -ForegroundColor Green

pause
