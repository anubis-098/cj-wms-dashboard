param(
    [string]$SourcePath = "\\10.84.194.51\CJWMSDashboard",
    [string]$DestinationPath = (Join-Path $PSScriptRoot "..\smd-cache"),
    [string]$SyncApiUrl = "http://localhost:8000/file-server/sync"
)

$ErrorActionPreference = "Stop"
$destination = [System.IO.Path]::GetFullPath($DestinationPath)
[System.IO.Directory]::CreateDirectory($destination) | Out-Null

$latest = Get-ChildItem -LiteralPath $SourcePath -File -Filter "*.xlsx" |
    Where-Object { -not $_.Name.StartsWith("~$") } |
    Sort-Object LastWriteTimeUtc, Name -Descending |
    Select-Object -First 1

if ($null -eq $latest) {
    throw "No .xlsx files found in $SourcePath"
}

$target = Join-Path $destination $latest.Name
$existing = Get-Item -LiteralPath $target -ErrorAction SilentlyContinue
if ($null -ne $existing -and $existing.Length -eq $latest.Length -and $existing.LastWriteTimeUtc -eq $latest.LastWriteTimeUtc) {
    Write-Output "Already current: $($latest.Name)"
} else {
    $temporary = Join-Path $destination (".{0}.{1}.tmp" -f $latest.Name, [guid]::NewGuid().ToString("N"))
    try {
        Copy-Item -LiteralPath $latest.FullName -Destination $temporary -Force
        Move-Item -LiteralPath $temporary -Destination $target -Force
        (Get-Item -LiteralPath $target).LastWriteTimeUtc = $latest.LastWriteTimeUtc
        Write-Output "Mirrored: $($latest.Name)"
    } finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force
        }
    }
}

$syncResult = Invoke-RestMethod -Method Post -Uri $SyncApiUrl -TimeoutSec 300
Write-Output "Backend sync: changed=$($syncResult.data.changed), file=$($syncResult.data.filename)"
