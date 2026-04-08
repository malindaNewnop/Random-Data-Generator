param(
    [string]$OutputFile,
    [string]$MessagePrefix = "ProgramExecution functional test"
)

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$line = "$MessagePrefix | heartbeat | $timestamp"

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $folder = Split-Path -Parent $OutputFile
    if (-not [string]::IsNullOrWhiteSpace($folder)) {
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }

    Add-Content -Path $OutputFile -Value $line
}
else {
    Write-Output $line
}
