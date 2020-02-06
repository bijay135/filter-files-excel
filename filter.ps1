$excelFile = Read-Host -Prompt "Input your full excel file path"
$excelSheet = Read-Host -Prompt "Input your excel sheet name"
$metaFolder = ".\MetaFolder"
$filteredFolder = ".\FilteredFolder"

$excel = New-Object -comobject Excel.Application
$workbook = $excel.Workbooks.Open($excelFile)
$worksheet = $workbook.Worksheets.Item($excelSheet)
$workSheet.Name

$excelNames = @()
$i = 2
while ($worksheet.Cells.Item($i, 2).Value() -ne $null) {
    $excelNames += $worksheet.Cells.Item($i, 2).Value().Trim()
    $i++
}
   
$workbook.Close()
$excel.Quit()
# IMPORTANT: clean-up used Com objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Get-ChildItem -Path $metaFolder |
  Where-Object { $excelNames -contains $_.Name } |
  Move-Item -Destination $filteredFolder -verbose

Write-Output "All files successfully moved"
Start-Sleep -s 3