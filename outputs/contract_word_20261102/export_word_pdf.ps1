param([Parameter(Mandatory=$true)][string]$InputDocx, [Parameter(Mandatory=$true)][string]$OutputPdf)
$ErrorActionPreference = 'Stop'
$repairWord = $null
$repairDoc = $null
try {
    $repairWord = New-Object -ComObject Word.Application
    $repairWord.Visible = $false
    $repairWord.DisplayAlerts = 0
    $repairWord.AutomationSecurity = 3
    $repairDoc = $repairWord.Documents.Open($InputDocx, $false, $true, $false)
    $repairDoc.ExportAsFixedFormat($OutputPdf, 17)
    Write-Output "Exported Word PDF: $OutputPdf"
} finally {
    if ($repairDoc) { $repairDoc.Close(0) }
    if ($repairWord) {
        $repairWord.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($repairWord) | Out-Null
    }
}
