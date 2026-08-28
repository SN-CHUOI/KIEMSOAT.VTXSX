﻿param(
    [string]$XoaFiles = ""   # So thu tu file muon xoa, cach nhau bang dau phay (vd: "1,3"). De trong se hoi.
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-XlsxList {
    Get-ChildItem -Path $scriptDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike "~`$*" -and $_.Name -notlike "test_*" } |
        Sort-Object Name
}

$xlsxCandidates = @(Get-XlsxList)

Write-Output "============================================"
Write-Output "  DANH SACH FILE EXCEL DANG DUOC DUNG"
Write-Output "============================================"

if ($xlsxCandidates.Count -eq 0) {
    Write-Output "(Khong co file .xlsx nao trong thu muc nay)"
} else {
    for ($i = 0; $i -lt $xlsxCandidates.Count; $i++) {
        Write-Output ("  {0}. {1}" -f ($i + 1), $xlsxCandidates[$i].Name)
    }
}
Write-Output ""

if ($XoaFiles -eq "") {
    $XoaFiles = Read-Host "Nhap SO THU TU file muon XOA (vd: 1,3) - roi bo qua bang cach nhan Enter"
}

$deletedAny = $false

if ($XoaFiles.Trim() -ne "") {
    $indices = $XoaFiles -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } | Sort-Object -Unique -Descending

    $toDelete = @()
    foreach ($idx in $indices) {
        if ($idx -ge 1 -and $idx -le $xlsxCandidates.Count) {
            $toDelete += $xlsxCandidates[$idx - 1]
        } else {
            Write-Output "  (Bo qua so $idx - khong hop le)"
        }
    }

    if ($toDelete.Count -gt 0) {
        Write-Output ""
        Write-Output "Se XOA (chuyen vao Thung rac) cac file sau:"
        $toDelete | ForEach-Object { Write-Output "  - $($_.Name)" }
        $confirm = Read-Host "`nGo YES de xac nhan xoa, hoac Enter de HUY"
        if ($confirm -eq "YES") {
            Add-Type -AssemblyName Microsoft.VisualBasic
            foreach ($f in $toDelete) {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $f.FullName,
                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
                )
                Write-Output "  Da xoa (vao Thung rac): $($f.Name)"
            }
            $deletedAny = $true
        } else {
            Write-Output "Da HUY, khong xoa gi ca."
        }
    }
}

Write-Output ""
Write-Output "============================================"
Write-Output "  DANG CAP NHAT VA DAY LEN WEB..."
Write-Output "============================================"
& (Join-Path $scriptDir "auto_publish.ps1")
Get-Content (Join-Path $scriptDir "auto_publish.log") -Tail 3 | ForEach-Object { Write-Output $_ }

Write-Output ""
Write-Output "============================================"
Write-Output "  DANH SACH FILE SAU KHI CAP NHAT"
Write-Output "============================================"
$finalList = @(Get-XlsxList)
if ($finalList.Count -eq 0) {
    Write-Output "(Khong con file .xlsx nao)"
} else {
    $finalList | ForEach-Object { Write-Output "  - $($_.Name)" }
}
Write-Output ""
Write-Output "Web: https://sn-chuoi.github.io/KIEMSOATVATU.XSX/dashboard.html (cap nhat sau ~30-60s)"
