[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$ScriptPath = ""
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

function Get-VbsHeaderMetadata([string]$Text) {
    $metadata = @{}
    $lines = $Text -split "`r?`n"
    foreach ($line in $lines) {
        $trim = $line.Trim()
        while (-not [string]::IsNullOrEmpty($trim) -and [int][char]$trim[0] -eq 0xFEFF) {
            $trim = $trim.Substring(1).TrimStart()
        }
        if ([string]::IsNullOrWhiteSpace($trim)) {
            continue
        }
        if (-not $trim.StartsWith("'")) {
            break
        }

        $content = $trim.Substring(1).Trim()
        $colon = $content.IndexOf(":")
        if ($colon -lt 0) {
            $colon = $content.IndexOf("：")
        }
        if ($colon -gt 0) {
            $key = $content.Substring(0, $colon).Trim()
            $value = $content.Substring($colon + 1).Trim()
            $metadata[$key] = $value
        }
    }
    return $metadata
}

function Test-VbsScript([string]$Path, [bool]$RequireMetadata) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Fail ("VBS file not found: " + $Path)
    }
    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne ".vbs") {
        Fail ("Expected a .vbs file: " + $Path)
    }

    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191) {
        Fail ("VBS file is not UTF-8 BOM: " + $Path)
    }

    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    if ($text -notmatch "(?im)^\s*Function\s+Main\s*\(\s*\)") {
        Fail ("VBS file must contain a parameterless Function Main(): " + $Path)
    }

    $unsafePattern = "CreateObject|GetObject|WScript\.Shell|Shell\.Application|ADODB\.Stream|FileSystemObject|InputBox|MsgBox|(^|[^A-Za-z])Shell([^A-Za-z]|$)"
    if ($text -match $unsafePattern) {
        Fail ("Unsafe API found in public VBS: " + $Path)
    }

    if ($RequireMetadata) {
        $metadata = Get-VbsHeaderMetadata $text
        $hasFunctionName = $metadata.ContainsKey("函数名") -and -not [string]::IsNullOrWhiteSpace([string]$metadata["函数名"])
        $hasDescription = ($metadata.ContainsKey("描述") -and -not [string]::IsNullOrWhiteSpace([string]$metadata["描述"])) -or
            ($metadata.ContainsKey("Description") -and -not [string]::IsNullOrWhiteSpace([string]$metadata["Description"])) -or
            ($metadata.ContainsKey("说明") -and -not [string]::IsNullOrWhiteSpace([string]$metadata["说明"]))
        $hasApp = $metadata.ContainsKey("适用应用") -and -not [string]::IsNullOrWhiteSpace([string]$metadata["适用应用"])
        if (-not $hasFunctionName -or -not $hasDescription -or -not $hasApp) {
            Fail ("VBS header must contain 函数名, 描述/Description, and 适用应用 before code: " + $Path)
        }
    }
}

if (-not [string]::IsNullOrWhiteSpace($ScriptPath)) {
    $scriptFullPath = (Resolve-Path -LiteralPath $ScriptPath -ErrorAction Stop).ProviderPath
    Test-VbsScript $scriptFullPath $true
    Write-Host ("[public-pack] single VBS validation passed: " + $scriptFullPath)
    exit 0
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

foreach ($file in $vbsFiles) {
    Test-VbsScript $file.FullName $true
}

$unsafePattern = "CreateObject|GetObject|WScript\.Shell|Shell\.Application|ADODB\.Stream|FileSystemObject|InputBox|MsgBox"
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
