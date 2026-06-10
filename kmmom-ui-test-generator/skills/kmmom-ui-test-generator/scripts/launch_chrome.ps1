param(
    [string]$Url = "http://192.168.30.69:40000",
    [int]$RemoteDebuggingPort = 9223,
    [string]$UserDataDir = "",
    [string]$ChromePath = ""
)

$ErrorActionPreference = "Stop"

function Resolve-ChromePath {
    param([string]$Candidate)

    if ($Candidate -and (Test-Path -LiteralPath $Candidate)) {
        return (Resolve-Path -LiteralPath $Candidate).Path
    }

    $paths = @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )

    foreach ($path in $paths) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    throw "Google Chrome executable was not found."
}

if (-not $UserDataDir) {
    $workspace = Get-Location
    $UserDataDir = Join-Path $workspace "chrome_kmmom_profile"
}

$resolvedChrome = Resolve-ChromePath -Candidate $ChromePath
$resolvedProfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($UserDataDir)
New-Item -ItemType Directory -Force -Path $resolvedProfile | Out-Null

$arguments = @(
    "--remote-debugging-port=$RemoteDebuggingPort",
    "--user-data-dir=$resolvedProfile",
    "--no-first-run",
    "--no-default-browser-check",
    $Url
)

Start-Process -FilePath $resolvedChrome -ArgumentList $arguments -WindowStyle Normal | Out-Null

Write-Output "Chrome launched"
Write-Output "ChromePath=$resolvedChrome"
Write-Output "RemoteDebuggingPort=$RemoteDebuggingPort"
Write-Output "UserDataDir=$resolvedProfile"
Write-Output "Url=$Url"
