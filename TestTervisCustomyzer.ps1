ipmo -force powershellorm,terviscustomyzer
$BatchNumber = "20180628-1300"
Set-CustomyzerModuleEnvironment -Name Epsilon
$BatchNumber = Get-CustomyzerApprovalPacklistRecentBatch |
Select-Object -Last 1 -ExpandProperty BatchNumber
Invoke-CustomyzerPackListDocumentsGenerate -BatchNumber $BatchNumber