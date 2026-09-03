$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $scriptDir "data.js"
$logFile = Join-Path $scriptDir "auto_publish.log"
$gitExe = "C:\Program Files\Git\cmd\git.exe"

function Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Output $msg
}

Log "== Day file data.js nhan duoc (khong doc lai Excel) =="

if (-not (Test-Path $dataPath)) {
    Log "LOI: Khong tim thay data.js trong folder."
    exit 1
}

$content = Get-Content $dataPath -Raw
if ($content -notmatch "DASHBOARD_DATA_SOURCES") {
    Log "LOI: File data.js hien tai khong dung dinh dang (thieu DASHBOARD_DATA_SOURCES). Kiem tra lai file nhan duoc."
    exit 1
}

Set-Location $scriptDir

$diff = & $gitExe diff --stat -- data.js
if (-not $diff) {
    Log "Khong co thay doi gi so voi ban dang chay tren web. Khong can day len."
    exit 0
}

& $gitExe add data.js
$commitMsg = "Cap nhat du lieu (file nhan tu dong nghiep) - " + (Get-Date -Format "yyyy-MM-dd HH:mm")
& $gitExe commit -m $commitMsg *> $null
if ($LASTEXITCODE -ne 0) {
    Log "LOI khi commit (exit code $LASTEXITCODE)"
    exit 1
}
Log "Da commit: $commitMsg"

& $gitExe push origin main *> $null
if ($LASTEXITCODE -eq 0) {
    Log "Da day len GitHub thanh cong. Web se tu cap nhat sau ~30-60s."
} else {
    Log "LOI khi push (exit code $LASTEXITCODE) - kiem tra ket noi mang / dang nhap GitHub."
    exit 1
}
