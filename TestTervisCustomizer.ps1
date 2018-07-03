ipmo -force powershellorm,terviscustomizer
$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber 20180702-1300
$OrderDetailRecords = $PackListLines | Get-CustomyzerApprovalOrderDetail
$ProjectRecords = 