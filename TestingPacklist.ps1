ipmo -Force terviscustomyzer
Set-CustomyzerModuleEnvironment -Name Production
$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber
$PackListLine = $PackListLines | Where-Object {$_.OrderDetail.Order.ERPOrderNumber -eq 11135215 } | Select-Object -First 1
$Project = $PackListLine.OrderDetail.Project

if ($Project.BackgroundID -and $Project.Product_Background.PurchaseFG) {
    $Project.Product_Background.PurchaseFG
} else {
    $Project.Product.FGCode
}	

$Project | Get-CustomyzerProjectFinalFGCode


Get-CustomyzerProduct_Background -ProductID f5197710-719d-4565-8f56-6fbd0ecc26de -BackgroundID "" | measure

$BatchNumber = "20180926-1300" #AR
$BatchNumber = "20181107-1300"


Set-CustomyzerModuleEnvironment -Name Epsilon

$PackListLinesNotInBatch = Get-CustomyzerApprovalPackList -NotInBatch
$PackListLinesNotInBatch.Length
$PackListLinesNotInBatch | Set-CustomyzerApprovalPackList -BatchNumber "NULL"

$BatchNumber = "20181114-1300"
$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber
$PackListLines | Set-CustomyzerApprovalPackList -SentDateUTC (Get-Date).ToUniversalTime()