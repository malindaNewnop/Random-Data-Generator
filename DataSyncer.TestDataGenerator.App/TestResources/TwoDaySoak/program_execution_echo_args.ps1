param(
    [string]$OutputFile,
    [string]$MessagePrefix = "Arg Test Value"
)

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$lines = @(
    "timestamp=$timestamp"
    "messagePrefix=$MessagePrefix"
    "outputFile=$OutputFile"
)

foreach ($line in $lines) {
    Write-Output $line
}

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $folder = Split-Path -Parent $OutputFile
    if (-not [string]::IsNullOrWhiteSpace($folder)) {
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }

    Add-Content -Path $OutputFile -Value ($lines -join [Environment]::NewLine)
}
