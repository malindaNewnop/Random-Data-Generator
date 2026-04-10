param(
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64",
    [string]$Version,
    [switch]$SkipPublish,
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Get-ProjectVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    [xml]$projectXml = Get-Content -Path $ProjectPath
    $propertyGroups = @($projectXml.Project.PropertyGroup)

    foreach ($group in $propertyGroups) {
        if ($group.Version -and -not [string]::IsNullOrWhiteSpace($group.Version)) {
            return $group.Version.Trim()
        }
    }

    return "1.0.0"
}

function Get-InnoCompiler {
    $candidates = @(
        $env:INNO_SETUP_COMPILER,
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return (Resolve-Path $candidate).Path
        }
    }

    return $null
}

$repoRoot = Get-RepoRoot
$projectPath = Join-Path $repoRoot "DataSyncer.TestDataGenerator.App\DataSyncer.TestDataGenerator.App.csproj"
$installerScript = Join-Path $repoRoot "installer\DataSyncer.TestDataGenerator.iss"
$publishDir = Join-Path $repoRoot ("artifacts\publish\" + $RuntimeIdentifier)
$installerOutputDir = Join-Path $repoRoot "artifacts\installer"

if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = Get-ProjectVersion -ProjectPath $projectPath
}

if (-not $SkipPublish) {
    if (Test-Path $publishDir) {
        Remove-Item -LiteralPath $publishDir -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $publishDir | Out-Null

    $publishArgs = @(
        "publish",
        $projectPath,
        "-c", $Configuration,
        "-r", $RuntimeIdentifier,
        "--self-contained", "true",
        "-p:PublishSingleFile=false",
        "-p:PublishTrimmed=false",
        "-p:GenerateAssemblyInfo=false",
        "-p:GenerateTargetFrameworkAttribute=false",
        "-o", $publishDir
    )

    & dotnet @publishArgs
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed with exit code $LASTEXITCODE."
    }
}

if (-not (Test-Path $publishDir)) {
    throw "Publish output was not found at '$publishDir'."
}

if (-not $SkipInstaller) {
    $compiler = Get-InnoCompiler
    if (-not $compiler) {
        throw "Inno Setup Compiler (ISCC.exe) was not found. Install Inno Setup 6 or set the INNO_SETUP_COMPILER environment variable."
    }

    New-Item -ItemType Directory -Force -Path $installerOutputDir | Out-Null

    $compilerArgs = @(
        "/DAppVersion=$Version",
        "/DPublishDir=$publishDir",
        "/DOutputDir=$installerOutputDir",
        $installerScript
    )

    & $compiler @compilerArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup compilation failed with exit code $LASTEXITCODE."
    }
}

Write-Host ""
Write-Host "Publish output:  $publishDir"
if (-not $SkipInstaller) {
    Write-Host "Installer output: $installerOutputDir"
}
