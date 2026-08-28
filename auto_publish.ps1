# Script tu dong: doc Excel -> cap nhat data.js -> commit -> push len GitHub
# Chi commit/push khi du lieu thuc su thay doi (tranh commit rong)
# Duoc goi dinh ky boi Windows Task Scheduler (moi 15 phut)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptDir "auto_publish.log"
$gitExe = "C:\Program Files\Git\cmd\git.exe"

function Log($msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

try {
    Set-Location $scriptDir
    Log "== Bat dau chu ky cap nhat =="

    # 1. Doc Excel, sinh lai data.js
    & (Join-Path $scriptDir "update_data.ps1") | Out-Null
    Log "Da doc lai Excel va sinh data.js"

    # 2. Kiem tra co thay doi thuc su khong
    $diff = & $gitExe diff --stat -- data.js
    if (-not $diff) {
        Log "Khong co thay doi du lieu, bo qua commit/push."
        exit 0
    }

    # 3. Commit
    & $gitExe add data.js
    $commitMsg = "Tu dong cap nhat du lieu - " + (Get-Date -Format "yyyy-MM-dd HH:mm")
    & $gitExe commit -m $commitMsg *> $null
    if ($LASTEXITCODE -ne 0) {
        Log "LOI khi commit (exit code $LASTEXITCODE)"
        exit 1
    }
    Log "Da commit: $commitMsg"

    # 4. Push (khong dung 2>&1 vi PowerShell 5.1 hieu nham stderr cua git la loi)
    & $gitExe push origin main *> $null
    if ($LASTEXITCODE -eq 0) {
        Log "Da day len GitHub thanh cong. Web se tu cap nhat sau ~30-60s."
    } else {
        Log "LOI khi push (exit code $LASTEXITCODE) - kiem tra ket noi mang / dang nhap GitHub."
    }
}
catch {
    Log "LOI: $($_.Exception.Message)"
}
