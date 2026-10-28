$ErrorActionPreference = 'Stop'
$outputPath = 'D:\UVG\outputs\doi_soat_san_20261019\Mau_doi_soat_cong_no_san_theo_thang.xlsx'
$excelApp = $null
$excelBook = $null
$qaChecks = [System.Collections.Generic.List[string]]::new()
function Check-Value($actual, $expected, [string]$label) {
  if ($actual -ne $expected) { throw "$label expected $expected, got $actual" }
  $qaChecks.Add($label)
}
try {
  $excelApp = New-Object -ComObject Excel.Application
  $excelApp.Visible = $false
  $excelApp.DisplayAlerts = $false
  $excelApp.EnableEvents = $false
  $excelApp.AskToUpdateLinks = $false
  $excelApp.AutomationSecurity = 3
  $excelBook = $excelApp.Workbooks.Open($outputPath, 0, $true)
  $excelApp.CalculateFullRebuild()
  $summarySheet = $excelBook.Worksheets.Item('Tong_hop')
  $bankSheet = $excelBook.Worksheets.Item('Ngan_hang')
  $payoutSheet = $excelBook.Worksheets.Item('Chuyen_tien')
  Check-Value $summarySheet.Range('H10').Value2 1160000 'Native Excel TikTok balance'
  Check-Value $summarySheet.Range('H11').Value2 1408000 'Native Excel Shopee balance'
  Check-Value $summarySheet.Range('H12').Value2 1406000 'Native Excel Lazada balance'
  Check-Value $summarySheet.Range('K12').Value2 1000 'Native Excel independent difference'
  Check-Value $excelBook.Worksheets.Item('Don_hang').Range('R11').Value2 -2000 'Native Excel order difference'
  foreach ($worksheet in $excelBook.Worksheets) {
    $formulaErrors = $null
    try { $formulaErrors = $worksheet.UsedRange.SpecialCells(-4123,16) } catch { }
    if ($null -ne $formulaErrors) { throw ('Formula errors on ' + $worksheet.Name + ': ' + $formulaErrors.Address()) }
  }
  $qaChecks.Add('No formula errors in all seven sheets after full Excel recalculation')
  $bankSheet.Range('F14').Value2 = [double]1000000
  $excelApp.CalculateFullRebuild()
  Check-Value $summarySheet.Range('K12').Value2 0 'Native Excel changing bank receipt recalculates balance'
  $bankSheet.Range('F14').Value2 = [double]999000
  $bankTable = $bankSheet.ListObjects.Item('T_NH')
  $newRow = $bankTable.ListRows.Add()
  $targetRow = $newRow.Range.Row
  $sourceRow = $targetRow - 1
  [void]$bankSheet.Range("I${sourceRow}:L${sourceRow}").Copy($bankSheet.Range("I${targetRow}:L${targetRow}"))
  $testRecord = @('TikTok Shop','DEMO_TT','NH-EXCEL-TEST','TT-P1',([datetime]'2026-09-20').ToOADate(),[double]1000,'DEMO112','Native Excel table growth test')
  foreach ($i in @(0,1,2,3,6,7)) { $bankSheet.Cells.Item($targetRow,$i+1).Value2 = [string]$testRecord[$i] }
  $bankSheet.Range("E${targetRow}").Value2 = [double]([datetime]'2026-09-20').ToOADate()
  $bankSheet.Range("F${targetRow}").Value2 = [double]1000
  $excelApp.CalculateFullRebuild()
  Check-Value $payoutSheet.Range('J8').Value2 991000 'Native Excel expanded table automatically included in SUMIFS'
  $result = @{ application = $excelApp.Version; checks = $qaChecks.ToArray(); total = $qaChecks.Count }
  $result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath 'D:\UVG\outputs\doi_soat_san_20261019\support\excel_verification.json' -Encoding utf8
  $result | ConvertTo-Json -Depth 4
} finally {
  if ($null -ne $excelBook) { $excelBook.Close($false) }
  if ($null -ne $excelApp) { $excelApp.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelApp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
