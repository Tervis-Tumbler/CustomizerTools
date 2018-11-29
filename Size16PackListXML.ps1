$EnvironmentName = "Production"
$DateTime = (Get-Date)

Set-CustomyzerModuleEnvironment -Name $EnvironmentName
Get-CustomyzerApprovalPacklistRecentBatch;

$BatchNumbers = "20181129-1300","20181128-1300"
foreach ($BatchNumber in $BatchNumbers) {
    $PackListLines = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber
    $PackListLinesSorted = Invoke-CustomyzerPackListLinesSort -PackListLines $PackListLines
    $Size16PackListLinesSorted = $PackListLinesSorted |
    Where-Object {
        $_.Orderdetail.Project.Product.Form.Size -eq 16
    }

	$CustomyzerPackListTemporaryFolder = New-CustomyzerPackListTemporaryFolder -BatchNumber $BatchNumber -EnvironmentName $EnvironmentName

    $Parameters = @{
		BatchNumber = $BatchNumber
		PackListLines = $Size16PackListLinesSorted
		Path = $CustomyzerPackListTemporaryFolder.FullName
    }
    $XMLFilePath = New-CustomyzerPackListXML @Parameters -DateTime $DateTime
}

# $MissingOne = $Size16PackListLinesSorted | where {
#     -not $_.OrderDetail.Project.FinalArchedImageLocation
# }
    