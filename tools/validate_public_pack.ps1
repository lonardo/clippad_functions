[CmdletBinding()]
param(
    [string]$Root = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Join-Path $PSScriptRoot ".."
}

$rootPath = (Resolve-Path -LiteralPath $Root).ProviderPath

function Fail([string]$Message) {
    throw $Message
}

function Read-Json([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Fail ("Missing JSON file: " + $Path)
    }
    return Get-Content -Raw -Encoding UTF8 $Path | ConvertFrom-Json
}

$requiredFiles = @(
    "README.md",
    "LICENSE",
    "NOTICE.md",
    "EXPORT_MANIFEST.json",
    "host/host-api-v1.json",
    "prompts/vbs-system-prompt.md"
)

foreach ($relativePath in $requiredFiles) {
    $path = Join-Path $rootPath $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Fail ("Missing required public file: " + $relativePath)
    }
}

$exportManifest = Read-Json (Join-Path $rootPath "EXPORT_MANIFEST.json")
$hostManifest = Read-Json (Join-Path $rootPath "host/host-api-v1.json")

$expectedVbs = [int]$exportManifest.included.vbs.total
$expectedExamples = [int]$exportManifest.included.vbsExamples
$expectedWorkflows = [int]$exportManifest.included.workflowExamples
$expectedHostMethods = [int]$exportManifest.included.hostMethods

$vbsRoots = @(
    (Join-Path $rootPath "vbs"),
    (Join-Path $rootPath "examples/vbs")
)
$vbsFiles = @()
foreach ($vbsRoot in $vbsRoots) {
    if (Test-Path -LiteralPath $vbsRoot -PathType Container) {
        $vbsFiles += @(Get-ChildItem -LiteralPath $vbsRoot -Recurse -Filter "*.vbs" -File)
    }
}

if ($vbsFiles.Count -ne ($expectedVbs + $expectedExamples)) {
    Fail ("Unexpected public VBS count: expected " + ($expectedVbs + $expectedExamples) + ", got " + $vbsFiles.Count)
}

$unsafePattern = "CreateObject|GetObject|WScript\.Shell|Shell\.Application|ADODB\.Stream|FileSystemObject|InputBox|MsgBox"
foreach ($file in $vbsFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191) {
        Fail ("VBS file is not UTF-8 BOM: " + $file.FullName)
    }

    $text = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
    if ($text -notmatch "(?im)^\s*Function\s+Main\s*\(") {
        Fail ("VBS file has no Function Main(): " + $file.FullName)
    }
    if ($text -match $unsafePattern) {
        Fail ("Unsafe API found in public VBS: " + $file.FullName)
    }
}

$workflowRoot = Join-Path $rootPath "examples/workflows"
$workflowFiles = @(Get-ChildItem -LiteralPath $workflowRoot -Filter "*.json" -File)
if ($workflowFiles.Count -ne $expectedWorkflows) {
    Fail ("Unexpected workflow example count: expected " + $expectedWorkflows + ", got " + $workflowFiles.Count)
}
foreach ($file in $workflowFiles) {
    $workflowText = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
    if ($workflowText -match $unsafePattern) {
        Fail ("Unsafe API found in workflow example: " + $file.FullName)
    }
}

$jsonFiles = @(Get-ChildItem -LiteralPath $rootPath -Recurse -Filter "*.json" -File)
foreach ($file in $jsonFiles) {
    Get-Content -Raw -Encoding UTF8 $file.FullName | ConvertFrom-Json | Out-Null
}

if ($hostManifest.methods.Count -ne $expectedHostMethods) {
    Fail ("Unexpected Host method count: expected " + $expectedHostMethods + ", got " + $hostManifest.methods.Count)
}
if (-not $hostManifest.selectionSemantics.'Host.GetSelection()'.available) {
    Fail "Host.GetSelection() must be marked available in the public contract."
}

$excludedSegments = @(
    "copyright_submission",
    "icp_filing",
    "patent_submission",
    "app_market_submission",
    "backend",
    "logs",
    "dist"
)
$allFiles = @(Get-ChildItem -LiteralPath $rootPath -Recurse -File)
foreach ($file in $allFiles) {
    $relativePath = $file.FullName.Substring($rootPath.Length).TrimStart("\", "/")
    foreach ($segment in $excludedSegments) {
        if ($relativePath -match "(^|[\\/])" + [regex]::Escape($segment) + "([\\/]|$)") {
            Fail ("Excluded repository material found in public pack: " + $relativePath)
        }
    }
}

Write-Host ("[public-pack] validation passed: VBS={0}, workflows={1}, hostMethods={2}, JSON={3}" -f $vbsFiles.Count, $workflowFiles.Count, $hostManifest.methods.Count, $jsonFiles.Count)
