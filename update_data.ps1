# Script cap nhat du lieu cho Dashboard Kiem soat vat tu
# Cach dung: chuot phai -> "Run with PowerShell", hoac chay trong PowerShell:
#   powershell -ExecutionPolicy Bypass -File update_data.ps1
# Script tu dong doc TAT CA file Excel (.xlsx) trong thu muc nay (moi file la
# 1 "nguon" rieng, vd moi xi nghiep 1-nhieu file), voi dieu kien co sheet
# "Tổng hợp theo nhóm" dung cau truc mau. Ket qua xuat ra data.js de
# dashboard.html hien thi duoi dang cac nut chuyen doi nguon. Doi ten file
# thoai mai, khong anh huong. Sau khi chay xong, mo lai (hoac F5) dashboard.html.

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

function Read-VatTuSheet($ws, $products) {
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

    return [PSCustomObject]@{
        keHoachSanLuong = $kehoach
        items           = $rows
    }
}

Write-Output "Dang mo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$sources = @()
$errors = @()

try {
    foreach ($file in $xlsxCandidates) {
        $label = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        Write-Output "Dang doc: $($file.Name)"
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
            $sheetExists = $false
            foreach ($s in $wb.Worksheets) { if ($s.Name -eq $sheetName) { $sheetExists = $true } }
            if (-not $sheetExists) {
                throw "Khong tim thay sheet '$sheetName' trong file nay."
            }
            $ws = $wb.Worksheets.Item($sheetName)
            $parsed = Read-VatTuSheet $ws $products

            $sources += [PSCustomObject]@{
                label           = $label
                generatedAt     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                sourceFile      = $file.Name
                sourceSheet     = $sheetName
                products        = $products
                keHoachSanLuong = $parsed.keHoachSanLuong
                items           = $parsed.items
            }
        }
        catch {
            $errors += "$($file.Name): $($_.Exception.Message)"
            Write-Output "  LOI, bo qua file nay: $($_.Exception.Message)"
        }
        finally {
            if ($wb) { $wb.Close($false) }
        }
    }
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

if ($sources.Count -eq 0) {
    Write-Error "Khong doc duoc du lieu tu bat ky file nao. Chi tiet: $($errors -join ' | ')"
    exit 1
}

$json = $sources | ConvertTo-Json -Depth 8
# Neu chi co 1 nguon, ConvertTo-Json khong tu boc mang [] - can ep kieu mang
if ($sources.Count -eq 1) { $json = "[$json]" }

$jsContent = "// File du lieu tu dong sinh ra tu Excel - KHONG chinh sua tay`n" +
             "// Sinh luc: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))`n" +
             "// So nguon: $($sources.Count)`n" +
             "const DASHBOARD_DATA_SOURCES = $json;`n"

[System.IO.File]::WriteAllText($outPath, $jsContent, [System.Text.Encoding]::UTF8)
Write-Output "Da cap nhat: $outPath ($($sources.Count) nguon, tong $((($sources | ForEach-Object {$_.items.Count}) | Measure-Object -Sum).Sum) dong vat tu)"
if ($errors.Count -gt 0) {
    Write-Output "CANH BAO: $($errors.Count) file bi loi va da bo qua: $($errors -join ' | ')"
}
