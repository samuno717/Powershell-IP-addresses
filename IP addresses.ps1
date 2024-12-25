$excel = New-Object -ComObject Excel.Application
$excel_visible = $false

$path = (Get-Location).Path
$savePath = Join-Path -Path $path -ChildPath "adresy.xlsx"
$outputPath = Join-Path -Path $path -ChildPath "wyniki.xlsx"
Write-Host $outputPath

if (Test-Path -Path $savePath) {
    Remove-Item -Path $savePath -Force -Confirm:$false
}

if (Test-Path -Path $outputPath) {
    Remove-Item -Path $outputPath -Force -Confirm:$false
}

Write-Host "Creating 'adresy.xlsx'..."

$workbook = $excel.Workbooks.add()
$sheet = $workbook.Sheets.Item(1)
$sheet.Name = "IP-Addresses"

$sheet.Cells.Item(1,1).Value = "8.8.8.8"
$sheet.Cells.Item(2,1).Value = "157.240.1.35"
$sheet.Cells.Item(3,1).Value = "142.250.74.174"
$sheet.Cells.Item(4,1).Value = "52.94.236.248"
$sheet.Cells.Item(5,1).Value = "208.80.154.224"


Write-Host "Creating 'wyniki.xlsx'..."

$workbook2 = $excel.Workbooks.add()
$sheet2 = $workbook2.Sheets.Item(1)

$sheet2.Cells.Item(1,1).Value = "Adresy"
$sheet2.Cells.Item(1,2).Value = "Wynik"



for ($i = 1; $i -lt 6; $i++) {
    $IP_address = $sheet.Cells.Item($i,1).Value2
    
    $ping = ping -n 1 $IP_address 
    $sheet2.Cells.Item($i+1,1) = "$IP_address"
    $sheet2.Cells.Item($i+1,2) = "$ping"

}

$workbook2.SaveAs($outputPath)
$workbook2.Close($false)
$workbook.SaveAs($savePath)
$workbook.Close($false)
$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet2) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook2) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null