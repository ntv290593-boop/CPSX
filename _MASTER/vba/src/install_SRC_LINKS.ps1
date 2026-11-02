# Build SRC_LINKS.xlam tu source va cai vao XLSTART (tu nap moi lan mo Excel).
$ErrorActionPreference = 'Stop'
$root = 'D:\COSTING EXCEL\_MASTER\vba'
$src  = Join-Path $root 'src'
$out  = Join-Path $root 'SRC_LINKS.xlam'
$key  = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
$old = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
    $wb  = $xl.Workbooks.Add()
    $vbp = $wb.VBProject
    $vbp.VBComponents.Import((Join-Path $src 'modSrcLinks.bas'))  | Out-Null
    $vbp.VBComponents.Import((Join-Path $src 'clsAppEvents.cls')) | Out-Null
    $vbp.VBComponents.Item($wb.CodeName).CodeModule.AddFromString((Get-Content (Join-Path $src 'ThisWorkbook.txt') -Raw))
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 55)
    $wb.Close($false)
} finally {
    if ($xl) { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
    if ($null -eq $old) { Remove-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $key -Name AccessVBOM -Value $old -Type DWord }
}
# Goi cho nguoi khac (chep tay vao XLSTART, xem HUONG DAN CAI.txt): chi can ban xlam moi nhat
Copy-Item $out (Join-Path $root 'GUI_CHO_NGUOI_KHAC\SRC_LINKS.xlam') -Force
# May da nap add-in qua khoa OPEN thi khong chep them vao XLSTART, neu khong add-in nap 2 lan.
$opt = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' -ErrorAction SilentlyContinue
if ($opt -and ($opt.PSObject.Properties | Where-Object { $_.Name -like 'OPEN*' -and $_.Value -like '*SRC_LINKS.xlam*' })) {
    Write-Output "OK: $out (da nap qua khoa OPEN, khong chep XLSTART)"
    return
}
$xlstart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
if (-not (Test-Path $xlstart)) { New-Item -ItemType Directory -Path $xlstart | Out-Null }
Copy-Item $out (Join-Path $xlstart 'SRC_LINKS.xlam') -Force
Write-Output "OK: $out -> $xlstart"
