param(
    [int]$Port = 8086,
    [int]$StartupTimeoutSeconds = 45
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ("[wenxing-web] " + $Message) -ForegroundColor Cyan
}

function Resolve-PythonPath {
    param([string]$AppRoot)

    $candidates = @(
        (Join-Path $AppRoot "venv\Scripts\python.exe"),
        (Join-Path $AppRoot ".venv\Scripts\python.exe"),
        "C:\Users\f\AppData\Local\Programs\Python\Python313\python.exe"
    )

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if (-not (Test-Path $candidate)) {
            continue
        }

        $resolved = (Resolve-Path $candidate).Path
        & $resolved -c "import uvicorn" *> $null
        if ($LASTEXITCODE -eq 0) {
            return $resolved
        }
    }

    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        & $pythonCommand.Source -c "import uvicorn" *> $null
        if ($LASTEXITCODE -eq 0) {
            return $pythonCommand.Source
        }
    }

    throw "No usable Python with uvicorn was found. Please confirm venv or system Python has project dependencies installed."
}

function Get-PortListener {
    param([int]$TargetPort)

    if (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue) {
        return Get-NetTCPConnection -LocalPort $TargetPort -State Listen -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }

    $pattern = ":{0}\s+.*LISTENING\s+(\d+)$" -f $TargetPort
    $line = netstat -ano | Select-String -Pattern $pattern | Select-Object -First 1
    if (-not $line) {
        return $null
    }

    $parts = ($line.ToString() -replace "\s+", " ").Trim().Split(" ")
    if ($parts.Length -lt 5) {
        return $null
    }

    return [pscustomobject]@{
        LocalPort = $TargetPort
        OwningProcess = [int]$parts[-1]
    }
}

function Stop-ExistingProjectProcess {
    param([int]$TargetPort)

    $listener = Get-PortListener -TargetPort $TargetPort
    if (-not $listener) {
        return
    }

    $processInfo = Get-CimInstance Win32_Process -Filter ("ProcessId = {0}" -f $listener.OwningProcess) -ErrorAction SilentlyContinue
    if (-not $processInfo) {
        throw "Port $TargetPort is already in use by PID $($listener.OwningProcess), but process details are unavailable."
    }

    $commandLine = [string]$processInfo.CommandLine
    $looksLikeThisProject = ($commandLine -match "run_web\.py") -or ($commandLine -match "wenxing-web")

    if (-not $looksLikeThisProject) {
        throw "Port $TargetPort is occupied by another process: PID=$($processInfo.ProcessId), Name=$($processInfo.Name)."
    }

    Write-Step ("Stopping previous service instance, PID=" + $processInfo.ProcessId)
    Stop-Process -Id $processInfo.ProcessId -Force

    $deadline = (Get-Date).AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 500
        $stillListening = Get-PortListener -TargetPort $TargetPort
    } while ($stillListening -and (Get-Date) -lt $deadline)

    if ($stillListening) {
        throw "Previous service instance did not release port $TargetPort in time."
    }
}

function Test-AppEntry {
    param([string]$AppRoot)

    $mainPy = Join-Path $AppRoot "api\main.py"
    if (-not (Test-Path $mainPy)) {
        throw "api\\main.py not found."
    }

    $runtimeMain = Join-Path $AppRoot "api\main_runtime.py"
    if (-not (Test-Path $runtimeMain)) {
        throw "api\\main_runtime.py not found."
    }

    $templatePath = Join-Path $AppRoot "Template_File\template_file.xlsx"
    if (-not (Test-Path $templatePath)) {
        throw "Template file not found: $templatePath"
    }
}

$appRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $appRoot

$logDir = Join-Path $appRoot "logs"
$serviceLog = Join-Path $logDir "service.log"
$errorLog = Join-Path $logDir "error.log"
$pidFile = Join-Path $logDir "wenxing-web.pid"
$healthUrl = "http://127.0.0.1:$Port/api/health"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

Write-Step "Checking runtime environment"
$pythonExe = Resolve-PythonPath -AppRoot $appRoot
Test-AppEntry -AppRoot $appRoot

Write-Step ("Using Python: " + $pythonExe)
Stop-ExistingProjectProcess -TargetPort $Port

$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

Write-Step "Starting web service"
$process = Start-Process `
    -FilePath $pythonExe `
    -ArgumentList @("-X", "utf8", "run_web.py") `
    -WorkingDirectory $appRoot `
    -RedirectStandardOutput $serviceLog `
    -RedirectStandardError $errorLog `
    -WindowStyle Hidden `
    -PassThru

$process.Id | Set-Content -Path $pidFile -Encoding ascii
Write-Step ("Process started, PID=" + $process.Id)

Write-Step "Waiting for health check"
$deadline = (Get-Date).AddSeconds($StartupTimeoutSeconds)
$healthy = $false

while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 2
    $process.Refresh()

    if ($process.HasExited) {
        break
    }

    try {
        $health = Invoke-RestMethod -Uri $healthUrl -TimeoutSec 5
        if ($health.status -eq "ok") {
            $healthy = $true
            break
        }
    }
    catch {
    }
}

if (-not $healthy) {
    Write-Host ""
    Write-Host "Service failed to pass health check within the timeout." -ForegroundColor Red
    Write-Host ("Stdout log: " + $serviceLog)
    Write-Host ("Stderr log: " + $errorLog)

    if (Test-Path $errorLog) {
        Write-Host ""
        Write-Host "Recent error log:" -ForegroundColor Yellow
        Get-Content -Path $errorLog -Tail 40
    }

    exit 1
}

Write-Host ""
Write-Host "Service started successfully." -ForegroundColor Green
Write-Host ("Backend health URL: " + $healthUrl)
Write-Host "Continue using the current IIS site address for browser access, for example: http://server-ip:1026/"
Write-Host ("Stdout log: " + $serviceLog)
Write-Host ("Stderr log: " + $errorLog)
