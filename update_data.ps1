# Script cap nhat du lieu cho Dashboard Kiem soat vat tu
# Cach dung: chuot phai -> "Run with PowerShell", hoac chay trong PowerShell:
#   powershell -ExecutionPolicy Bypass -File update_data.ps1
# Script tu dong doc TAT CA file Excel (.xlsx) trong thu muc nay (moi file la
# 1 "nguon" rieng, vd moi xi nghiep 1-nhieu file), voi dieu kien co sheet
# "Tổng hợp theo nhóm" dung cau truc mau. KHONG con co dinh so luong san pham -
# script TU DO TIM tieu de cac bang (dong 2) va TU DONG DOC danh sach san
# pham/pham cap rieng cua tung file (dong 3), nen moi xi nghiep co the co danh
# muc san pham khac nhau (9, 16, hay bao nhieu cung duoc). Ket qua xuat ra
# data.js de dashboard.html hien thi duoi dang cac nut chuyen doi nguon.
# Sau khi chay xong, mo lai (hoac F5) dashboard.html.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sheetName = "Tổng hợp theo nhóm"
$outPath = Join-Path $scriptDir "data.js"

$xlsxCandidates = Get-ChildItem -Path $scriptDir -Filter "*.xlsx" -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike "~`$*" -and $_.Name -notlike "test_*" -and $_.Name -notlike "_backup_*" } |
    Sort-Object Name

if ($xlsxCandidates.Count -eq 0) {
    Write-Error "Khong tim thay file Excel (.xlsx) nao trong thu muc: $scriptDir"
    exit 1
}

function Get-CellVal($cell) {
    $v = $cell.Value2
    if ($null -eq $v) { return $null }
    if ($v -is [string]) {
        if ($v.Trim() -eq "") { return $null }
        return $v
    }
    return [math]::Round([double]$v, 4)
}

# Do tim cac "bang" tren dong tieu de (mac dinh dong 2): moi o khong rong la
# ten 1 bang, do rong bang = khoang cach toi o khong rong ke tiep.
function Get-SectionMap($ws, [int]$headerRow, [int]$maxScanCol) {
    $labels = @()
    for ($c = 6; $c -le $maxScanCol; $c++) {
        $t = $ws.Cells.Item($headerRow, $c).Text
        if ($t -and $t.Trim() -ne "") { $labels += [PSCustomObject]@{ Col = $c; Text = $t.Trim() } }
    }
    $sections = [ordered]@{}
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $startCol = $labels[$i].Col
        $endCol = if ($i + 1 -lt $labels.Count) { $labels[$i + 1].Col - 1 } else { $maxScanCol }
        $sections[$labels[$i].Text] = [PSCustomObject]@{ Start = $startCol; End = $endCol }
    }
    return $sections
}

function Find-SectionKey($sections, [string[]]$mustContain, [string[]]$mustNotContain) {
    if ($null -eq $mustNotContain) { $mustNotContain = @() }
    foreach ($key in $sections.Keys) {
        $ok = $true
        foreach ($m in $mustContain) { if ($key -notlike "*$m*") { $ok = $false; break } }
        if ($ok) {
            foreach ($n in $mustNotContain) { if ($key -like "*$n*") { $ok = $false; break } }
        }
        if ($ok) { return $key }
    }
    return $null
}

function Read-VatTuSheet($ws) {
    $maxScanCol = $ws.UsedRange.Columns.Count
    $sections = Get-SectionMap $ws 2 $maxScanCol

    $capKey = Find-SectionKey $sections @("TỔNG SẢN LƯỢNG")
    $demKey = Find-SectionKey $sections @("NHU CẦU")
    $phanBoKey = Find-SectionKey $sections @("THỪA", "THIẾU", "PHÂN BỔ")
    $tongNhuCauKey = Find-SectionKey $sections @("Tổng nhu cầu")
    $thuaThieuTongKey = Find-SectionKey $sections @("Thừa", "Thiếu") @("PHÂN BỔ", "QUY RA")
    $danhGiaKey = Find-SectionKey $sections @("Đánh giá")
    $pctKey = Find-SectionKey $sections @("đáp ứng") @("TỔNG SẢN LƯỢNG")

    if (-not $capKey -or -not $demKey) {
        throw "Khong tim thay bang 'TONG SAN LUONG' hoac 'NHU CAU' o dong tieu de (dong 2) cua sheet."
    }

    $capStart = $sections[$capKey].Start
    $productCount = $sections[$capKey].End - $capStart + 1
    $demStart = $sections[$demKey].Start
    $phanBoStart = if ($phanBoKey) { $sections[$phanBoKey].Start } else { $null }
    $tongNhuCauCol = if ($tongNhuCauKey) { $sections[$tongNhuCauKey].Start } else { $null }
    $thuaThieuTongCol = if ($thuaThieuTongKey) { $sections[$thuaThieuTongKey].Start } else { $null }
    $danhGiaCol = if ($danhGiaKey) { $sections[$danhGiaKey].Start } else { $null }
    $pctCol = if ($pctKey) { $sections[$pctKey].Start } else { $null }

    # Danh sach san pham/pham cap: doc truc tiep tu dong 3, trong vung bang
    # "Tong san luong" - moi file co the khac nhau, khong con co dinh.
    $products = @()
    for ($i = 0; $i -lt $productCount; $i++) {
        $name = $ws.Cells.Item(3, $capStart + $i).Value2
        if ($null -eq $name) { $name = "" }
        $products += "$name".Trim()
    }

    # Ke hoach san luong: dong 1, dung vi tri voi bang Nhu cau
    $kehoach = [ordered]@{}
    for ($i = 0; $i -lt $productCount; $i++) {
        $val = Get-CellVal $ws.Cells.Item(1, $demStart + $i)
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
        $nhuCau = [ordered]@{}
        $phanBo = [ordered]@{}
        for ($i = 0; $i -lt $productCount; $i++) {
            $dapUng[$products[$i]] = Get-CellVal $ws.Cells.Item($r, $capStart + $i)
            $v = Get-CellVal $ws.Cells.Item($r, $demStart + $i)
            if ($null -eq $v) { $v = 0 }
            $nhuCau[$products[$i]] = $v
            if ($phanBoStart) {
                $phanBo[$products[$i]] = Get-CellVal $ws.Cells.Item($r, $phanBoStart + $i)
            }
            else {
                $phanBo[$products[$i]] = $null
            }
        }

        $tongNhuCau = if ($tongNhuCauCol) { Get-CellVal $ws.Cells.Item($r, $tongNhuCauCol) } else { $null }
        $thuaThieuTong = if ($thuaThieuTongCol) { Get-CellVal $ws.Cells.Item($r, $thuaThieuTongCol) } else { $null }
        $danhGia = if ($danhGiaCol) { Get-CellVal $ws.Cells.Item($r, $danhGiaCol) } else { $null }
        $phanTram = if ($pctCol) { Get-CellVal $ws.Cells.Item($r, $pctCol) } else { $null }

        $rows += [PSCustomObject]@{
            maVT           = $maVT
            tenVT          = $tenVT
            dvt            = $dvt
            tonNhap        = $tonNhap
            dapUng         = $dapUng
            nhuCau         = $nhuCau
            phanBo         = $phanBo
            tongNhuCau     = $tongNhuCau
            thuaThieuTong  = $thuaThieuTong
            danhGia        = $danhGia
            phanTramDapUng = $phanTram
        }
        $r++
    }

    return [PSCustomObject]@{
        products        = $products
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
            $parsed = Read-VatTuSheet $ws

            Write-Output "  -> $($parsed.products.Count) san pham/pham cap: $($parsed.products -join ', ')"

            $sources += [PSCustomObject]@{
                label           = $label
                generatedAt     = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                sourceFile      = $file.Name
                sourceSheet     = $sheetName
                products        = $parsed.products
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
