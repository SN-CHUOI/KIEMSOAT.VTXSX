# Script cap nhat du lieu cho Dashboard Kiem soat vat tu
# Cach dung: chuot phai -> "Run with PowerShell", hoac chay trong PowerShell:
#   powershell -ExecutionPolicy Bypass -File update_data.ps1
# Script tu dong tim file Excel (.xlsx) dau tien trong thu muc nay (khong phu
# thuoc ten file cu the, doi ten van chay duoc) co sheet "Tổng hợp theo nhóm",
# va xuat ra file data.js de dashboard.html su dung. Sau khi chay xong, mo lai
# (hoac F5) file dashboard.html de thay du lieu moi.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sheetName = "Tổng hợp theo nhóm"
$outPath = Join-Path $scriptDir "data.js"

$xlsxCandidates = Get-ChildItem -Path $scriptDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike "~`$*" -and $_.Name -notlike "test_*" } |
    Sort-Object Name

if ($xlsxCandidates.Count -eq 0) {
    Write-Error "Khong tim thay file Excel (.xlsx) nao trong thu muc: $scriptDir"
    exit 1
}
if ($xlsxCandidates.Count -gt 1) {
    Write-Output "CANH BAO: Co $($xlsxCandidates.Count) file .xlsx trong thu muc, dang dung file dau tien: $($xlsxCandidates[0].Name)"
    Write-Output ("  (" + (($xlsxCandidates | ForEach-Object { $_.Name }) -join ", ") + ")")
}
$excelPath = $xlsxCandidates[0].FullName
$excelFileName = $xlsxCandidates[0].Name

$products = @("A456","A789","B456","B789","BCL","CP","CP Nhật (13kg)","CP Nhật (15kg)","CP Nhật (18kg)")

function Get-CellVal($cell) {
    $v = $cell.Value2
    if ($null -eq $v) { return $null }
    if ($v -is [string]) {
        if ($v.Trim() -eq "") { return $null }
        return $v
    }
    return [math]::Round([double]$v, 4)
}

Write-Output "Dang mo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($excelPath, 0, $true)
    $ws = $wb.Worksheets.Item($sheetName)

    # Row 1: ke hoach san luong (cols 15-23) - tinh theo thung
    $kehoach = [ordered]@{}
    for ($i = 0; $i -lt 9; $i++) {
        $c = 15 + $i
        $val = Get-CellVal $ws.Cells.Item(1, $c)
        if ($null -eq $val) { $val = 0 }
        $kehoach[$products[$i]] = $val
    }

    $rows = @()
    $r = 4
    while ($true) {
        $maVT = Get-CellVal $ws.Cells.Item($r, 2)
        if ($null -eq $maVT) {
            # cho phep vai dong trong xen ke, dung han khi gap 3 dong trong lien tiep
            $blankCount = 0
            $rr = $r
            while ($blankCount -lt 3 -and $rr -le ($r + 2)) {
                if ($null -eq (Get-CellVal $ws.Cells.Item($rr, 2))) { $blankCount++ } else { break }
                $rr++
            }
            if ($blankCount -ge 3) { break }
            $r++
            if ($r -gt 500) { break }
            continue
        }

        $tenVT = Get-CellVal $ws.Cells.Item($r, 3)
        $dvt = Get-CellVal $ws.Cells.Item($r, 4)
        $tonNhap = Get-CellVal $ws.Cells.Item($r, 5)
        if ($null -eq $tonNhap) { $tonNhap = 0 }

        $dapUng = [ordered]@{}
        for ($i = 0; $i -lt 9; $i++) {
            $c = 6 + $i
            $dapUng[$products[$i]] = Get-CellVal $ws.Cells.Item($r, $c)
        }
        $nhuCau = [ordered]@{}
        for ($i = 0; $i -lt 9; $i++) {
            $c = 15 + $i
            $v = Get-CellVal $ws.Cells.Item($r, $c)
            if ($null -eq $v) { $v = 0 }
            $nhuCau[$products[$i]] = $v
        }
        $phanBo = [ordered]@{}
        for ($i = 0; $i -lt 9; $i++) {
            $c = 28 + $i
            $phanBo[$products[$i]] = Get-CellVal $ws.Cells.Item($r, $c)
        }

        $tongNhuCau = Get-CellVal $ws.Cells.Item($r, 24)
        $thuaThieuTong = Get-CellVal $ws.Cells.Item($r, 25)
        $danhGia = Get-CellVal $ws.Cells.Item($r, 26)
        $phanTram = Get-CellVal $ws.Cells.Item($r, 27)

        $rows += [PSCustomObject]@{
            maVT            = $maVT
            tenVT           = $tenVT
            dvt             = $dvt
            tonNhap         = $tonNhap
            dapUng          = $dapUng
            nhuCau          = $nhuCau
            phanBo          = $phanBo
            tongNhuCau      = $tongNhuCau
            thuaThieuTong   = $thuaThieuTong
            danhGia         = $danhGia
            phanTramDapUng  = $phanTram
        }
        $r++
    }

    $result = [PSCustomObject]@{
        generatedAt      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        sourceFile       = $excelFileName
        sourceSheet      = $sheetName
        products         = $products
        keHoachSanLuong  = $kehoach
        items            = $rows
    }

    $json = $result | ConvertTo-Json -Depth 8
    $jsContent = "// File du lieu tu dong sinh ra tu Excel - KHONG chinh sua tay`n" +
                 "// Sinh luc: $($result.generatedAt)`n" +
                 "const DASHBOARD_DATA = $json;`n"

    [System.IO.File]::WriteAllText($outPath, $jsContent, [System.Text.Encoding]::UTF8)
    Write-Output "Da cap nhat: $outPath ($($rows.Count) vat tu)"
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
