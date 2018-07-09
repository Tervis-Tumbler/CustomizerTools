ipmo -force powershellorm,terviscustomizer
$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber 20180702-1300
$BatchNumber = "20180628-1300"

$OrderDetailRecords = $PackListLines | Get-CustomyzerApprovalOrderDetail
$PackListLines.orderdetail | Out-Null

$ProjectRecords = $OrderDetailRecords | Get-CustomyzerProject

$PackListLines.orderdetail.project.ProjectImageSource