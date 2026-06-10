param(
    [string]$CodexSkillsDir = "C:\Users\Administrator\.codex\skills",
    [string]$PasswordEnv = "KMMOM_PASSWORD",
    [switch]$SetPassword
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$skillSource = Join-Path $repoRoot "skills\kmmom-ui-test-generator"
$skillTarget = Join-Path $CodexSkillsDir "kmmom-ui-test-generator"

if (-not (Test-Path -LiteralPath $skillSource)) {
    throw "Skill source not found: $skillSource"
}

New-Item -ItemType Directory -Force -Path $CodexSkillsDir | Out-Null

if (Test-Path -LiteralPath $skillTarget) {
    Remove-Item -LiteralPath $skillTarget -Recurse -Force
}

Copy-Item -LiteralPath $skillSource -Destination $skillTarget -Recurse -Force

Write-Output "Skill installed:"
Write-Output $skillTarget

if ($SetPassword) {
    $secure = Read-Host "Input KMMOM password for environment variable $PasswordEnv" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        [Environment]::SetEnvironmentVariable($PasswordEnv, $plain, "User")
        Set-Item -Path "Env:$PasswordEnv" -Value $plain
        Write-Output "Password environment variable saved for current Windows user: $PasswordEnv"
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}
else {
    Write-Output "Password was not configured. To configure it, rerun with -SetPassword."
}

Write-Output "Restart Codex after installing or updating the skill."
