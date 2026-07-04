 [CmdletBinding(PositionalBinding = $false)]
param(
    [string]$CacheDir,
    [string]$DefaultCacheDir,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-WritableDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue
    )

    try {
        $directory = New-Item -ItemType Directory -Path $PathValue -Force -ErrorAction Stop
        $probe = Join-Path $directory.FullName (".uv-probe-" + [guid]::NewGuid().ToString("N"))
        Set-Content -LiteralPath $probe -Value "ok" -Encoding ascii -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

if (-not $env:UV_CACHE_DIR) {
    if ($DefaultCacheDir) {
        $candidate = $DefaultCacheDir
    }
    elseif ($env:LOCALAPPDATA) {
        $candidate = Join-Path $env:LOCALAPPDATA "uv\cache"
    }
    else {
        $candidate = $null
    }

    $useFallback = $true
    if ($candidate) {
        $useFallback = -not (Test-WritableDirectory -PathValue $candidate)
    }

    if ($useFallback) {
        if (-not $CacheDir) {
            $CacheDir = Join-Path ([System.IO.Path]::GetTempPath()) "msword-cli-uv-cache"
        }
        New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null
        $env:UV_CACHE_DIR = $CacheDir
    }
}

& uv @Arguments
exit $LASTEXITCODE
