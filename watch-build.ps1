# Re-runs build.bat when main.py or requirements.txt change (poll every 2s).
# Usage: .\watch-build.ps1
# Stop with Ctrl+C

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
Set-Location $root

$watched = @(
    (Join-Path $root "main.py")
    (Join-Path $root "requirements.txt")
)

$last = @{}
foreach ($f in $watched) {
    if (Test-Path $f) {
        $last[$f] = (Get-Item -LiteralPath $f).LastWriteTimeUtc
    }
}

Write-Host "Watching for changes (Ctrl+C to stop). Project: $root"
Write-Host ""

while ($true) {
    Start-Sleep -Seconds 2
    $changed = $false
    foreach ($f in $watched) {
        if (-not (Test-Path -LiteralPath $f)) { continue }
        $t = (Get-Item -LiteralPath $f).LastWriteTimeUtc
        if (-not $last.ContainsKey($f)) {
            $last[$f] = $t
            continue
        }
        if ($last[$f] -ne $t) {
            $last[$f] = $t
            $changed = $true
        }
    }
    if ($changed) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Change detected — running build.bat ..."
        & cmd.exe /c "cd /d `"$root`" && build.bat"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Build failed (exit $LASTEXITCODE)." -ForegroundColor Red
        } else {
            Write-Host "Build finished." -ForegroundColor Green
        }
        Write-Host ""
    }
}
