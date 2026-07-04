 [CmdletBinding(PositionalBinding = $false)]
param(
    [string]$CacheDir,
    [string]$DefaultCacheDir,
    [switch]$Integration,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PytestArguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$uvArgs = @("run", "--no-sync", "pytest", "-q")
if ($Integration) {
    $uvArgs += @("-m", "integration")
}
if ($PytestArguments) {
    $uvArgs += $PytestArguments
}

$uvRunScript = Join-Path $PSScriptRoot "uv-run.ps1"
$forwarded = @{}
if ($CacheDir) {
    $forwarded["CacheDir"] = $CacheDir
}
if ($DefaultCacheDir) {
    $forwarded["DefaultCacheDir"] = $DefaultCacheDir
}

& $uvRunScript @forwarded @uvArgs
exit $LASTEXITCODE
