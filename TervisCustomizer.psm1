$ModulePath = (Get-Module -ListAvailable TervisCustomizer).ModuleBase

function Invoke-CustomizerSQL {
    param (
		
		[Parameter(Mandatory,ParameterSetName="Parameters")]$TableName,
		[Parameter(Mandatory,ParameterSetName="Parameters")]$Parameters,
		[Parameter(ParameterSetName="Parameters")]$ArbitraryWherePredicate,
        [Parameter(Mandatory,ParameterSetName="SQLCommand")]$SQLCommand
	)
	if (-not $Script:CustomizerConnectionString) {
		$Script:CustomizerConnectionString = Get-PasswordstateMSSQLDatabaseEntryDetails -PasswordID 5366 | ConvertTo-MSSQLConnectionString
	}
    
	if (-not $SQLCommand) {
		$SQLCommand = New-SQLSelect @PSBoundParameters
	}

    Invoke-MSSQL -ConnectionString $Script:CustomizerConnectionString -sqlCommand $SQLCommand -ConvertFromDataRow
}

function Get-CustomizerApprovalOrder {
    param (
        $ERPOrderNumber
    )
    Invoke-CustomizerSQL -SQLCommand @"
Select *
From [Approval].[Order] With (NoLock)
Where [ERPOrderNumber] in ('$ERPOrderNumber')
"@
}

function Get-CustomizerApprovalReviewEventLogLast10InStatusID10 {
    Invoke-CustomizerSQL -SQLCommand @"
Select top 10 *
From Approval.ReviewEventLog With (NoLock)
Where StatusID = 10
order by EventDateUTC desc
"@
}

function Get-CustomizerApprovalOrderLast10InStatusID10 {
    Invoke-CustomizerSQL -SQLCommand @"
Select top 10 *
From [Approval].[Order] With (NoLock)
Where StatusID = 10
order by [Approval].[Order].[SubmitDateUTC] desc
"@
}

function Get-CustomizerApprovalOrderDetail {
    param (
        $ERPOrderNumber
    )
    Invoke-CustomizerSQL -SQLCommand @"
Select *
From [Approval].[Order] With (NoLock)
Join [Approval].[OrderDetail] With (NoLock)
on [Approval].[Order].OrderID = [Approval].[OrderDetail].OrderID
Where [Approval].[Order].[ERPOrderNumber] in ('$ERPOrderNumber')
"@
}

function Get-CustomizerIntegrationsOrdersFromEBS {
    param (
        $ERPOrderNumber
    )
    Invoke-CustomizerSQL -SQLCommand @"
Select * from Integrations.OrdersFromEBS where ERPOrderNumber = '$ERPOrderNumber'
"@
}

function Get-CustomizerOrderNumberInIntegrations.ScheduledOrdersFromEBS {
    param (
        [DateTime]$DateTimeAfterWhichOrderShouldBeFound,
        [DateTime]$DateTimeBeforeWhichOrderShouldBeFound
    )
    $XMLDocuments = Invoke-CustomizerSQL -SQLCommand @"
Select InputXMLPayload 
From integrations.scheduledOrdersFromEBS With (NoLock)
Where LastSyncedDateTime < '$(($DateTimeAfterWhichOrderShouldBeFound).ToString("yyyy-MM-dd"))'
"@
    
}

function Get-CustomizerSQLSession {
    Invoke-CustomizerSQL -SQLCommand @"
SELECT
    c.session_id, c.net_transport, c.encrypt_option,
    s.status,
    c.auth_scheme, s.host_name, s.program_name,
    s.client_interface_name, s.login_name, s.nt_domain,
    s.nt_user_name, s.original_login_name, c.connect_time,
    s.login_time
FROM sys.dm_exec_connections AS c
JOIN sys.dm_exec_sessions AS s
    ON c.session_id = s.session_id
WHERE 
s."host_name" like 'CMAGNUSON10-LT' and
s.program_name like '.Net SqlClient Data Provider'
ORDER BY c.connect_time ASC
"@
}

function Remove-CustomizerSQLSession {
    param (
        [Parameter(Mandatory)]$SessionFromComputerName,
        [Parameter(Mandatory)]$ProgramName
    )
    Invoke-CustomizerSQL -SQLCommand @"
DECLARE @kill varchar(8000) = '';

SELECT @kill = @kill + 'KILL ' + CONVERT(varchar(5), c.session_id) + ';'

FROM sys.dm_exec_connections AS c
JOIN sys.dm_exec_sessions AS s
    ON c.session_id = s.session_id 
WHERE
s."host_name" like '$SessionFromComputerName' and
s."program_name" like '$ProgramName'
ORDER BY c.connect_time ASC
SELECT @kill as varchar
"@
}

function Get-CustomizerInIntegrations.ScheduledOrdersFromEBSPropertiesOfXMLDocument {
    param (
        [Parameter(Mandatory)]$ID
    )
    Invoke-CustomizerSQL -SQLCommand @"
BEGIN TRANSACTION	
declare @ScheduledOrdersData xml 
set @ScheduledOrdersData = (select InputXMLPayload from integrations.scheduledOrdersFromEBS With (NoLock) where ID = '$ID')
;WITH xmlnamespaces( 'http://Tervis.BizTalk.ScheduledOrders.Schemas' as xx)
	SELECT	DISTINCT
	*
	--		[Order].OrderID,
	--		[OrderDetail].OrderDetailID,
	--		ScheduleOrders.[ERPORDERID],
	--		ScheduleOrders.[ERPLINEID],
	--		ScheduleOrders.[SCHEDULENUMBER],
	--		[OrderDetail].OrderQuantity

FROM
		(
		
			SELECT
				Tab.Col.value('@SalesOrderNumber','VARCHAR(40)') AS [ERPORDERID],
				Tab.Col.value('@SalesOrderLineNumber', 'VARCHAR(255)') AS [ERPLINEID],
				Tab.Col.value('@MFGScheduleNumber','VARCHAR(100)') AS [SCHEDULENUMBER],
				Tab.Col.value('@WorkOrderStatus','VARCHAR(30)') AS [STATUS]
			FROM   @ScheduledOrdersData.nodes('/xx:ScheduledOrders/ScheduledOrder') Tab(Col)
		) ScheduleOrders
Commit TRANSACTION
"@
}

function Get-CustomizerIntegrations.ScheduledOrdersFromEBSTop100MostRecent {
    Invoke-CustomizerSQL -SQLCommand @"
select top 100 * from integrations.scheduledOrdersFromEBS With (NoLock)
order by LastSyncedDateTime Desc 
"@
}

function Get-CustomizerIntegrationsIntegrationStatus {
    Invoke-CustomizerSQL -SQLCommand @"
SELECT *
  FROM [Integrations].[IntegrationStatus] With (NoLock)
"@
}

function Test-CustomizerOracleScheduledOrdersAvailable {
@"
SELECT COUNT(*) FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)
"@
}

function Get-CustomizerOracleScheduledOrders {
@"
SELECT 
SALES_ORDER_NUMBER, 
SALES_ORDER_LINE_NUMBER, 
MFG_SCHEDULE_NUMBER, 
WORK_ORDER_STATUS, 
CREATION_DATE "LAST_UPDATE_DATE" 
FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG 
WHERE CREATION_DATE>= SYSDATE - 30/(24*60)
"@
}

function Invoke-CustomizerOracleScheduledOrdersToCustomizer {
@"
USE [customizer_db]
GO

/****** Object:  StoredProcedure [dbo].[ProcessUpdateReleasedScheduleLineItemsData]    Script Date: 3/16/2018 3:33:05 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[ProcessUpdateReleasedScheduleLineItemsData]
	@ScheduledOrdersData xml = null
AS

	UPDATE 
		[Integrations].IntegrationStatus 
	SET	
		[LastSyncDate]=GETDATE()
	WHERE
		[IntegrationTypeToken]='SCHEDULEDORDER'
		
			
	BEGIN TRANSACTION	
	--DECLARE @ScheduledOrdersData xml
	--SET @ScheduledOrdersData =
	--'<ns4:Schedules 	xmlns:ns4="http://Tervis.BizTalk.Schedules.Schemas">
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6019953" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="1"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6018015" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="3"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6017994" MFGScheduleNumber="04/10-A2-A-001" SalesOrderLineNumber="3"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6017963" MFGScheduleNumber="04/09-M2-A-008" SalesOrderLineNumber="2"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6017310" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="2"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6016643" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="1"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6016550" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="1"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6004422" MFGScheduleNumber="04/08-M1-A-097" SalesOrderLineNumber="1"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6004423" MFGScheduleNumber="04/11-A2-A-009" SalesOrderLineNumber="1"/>
	--	<Schedule 	StatusCode="Released" SalesOrderNumber="6018015" MFGScheduleNumber="04/08-M1-A-099" SalesOrderLineNumber="2"/>
	--</ns4:Schedules>'

		--;WITH XmlFile (Contents) AS (
		--SELECT CONVERT (XML, BulkColumn) 
		--FROM OPENROWSET (BULK 'E:\Mizer Schedule Order Info.xml', SINGLE_BLOB) AS XmlData
		--)
		--SELECT @ScheduledOrdersData = Contents
		--FROM XmlFile
	
	--DECLARE @idoc int
	--EXEC sp_xml_preparedocument @idoc OUTPUT, @ScheduledOrdersData, '<root xmlns:ns4="http://Tervis.BizTalk.Schedules.Schemas" />'

	/* Step-1 : Delete anything with a CreatedDateUTC over 2 months old since the current run time from the 
		“Integrations.OrdersFromEBS” table */
	DELETE FROM Integrations.OrdersFromEBS
	WHERE Integrations.OrdersFromEBS.CreatedDateUTC < DATEADD(month, -2, GETUTCDATE())

	/* Step-1 : End */

	CREATE TABLE #Temp_Orders 
	(
		[OrderID]			uniqueidentifier,
		[OrderDetailID]		uniqueidentifier,
		[ERPORDERID]		varchar(40),
		[ERPLINEID]			varchar(255),
		[SCHEDULENUMBER]	varchar(100),
		[OrderQuantity]		int
	)
	
	;WITH xmlnamespaces( 'http://Tervis.BizTalk.ScheduledOrders.Schemas' as xx)
	INSERT INTO #Temp_Orders ([OrderID], [OrderDetailID], [ERPORDERID], [ERPLINEID], [SCHEDULENUMBER], [OrderQuantity])
	SELECT	DISTINCT
			[Order].OrderID,
			[OrderDetail].OrderDetailID,
			ScheduleOrders.[ERPORDERID],
			ScheduleOrders.[ERPLINEID],
			ScheduleOrders.[SCHEDULENUMBER],
			[OrderDetail].OrderQuantity
	FROM
		(
		
			SELECT
				Tab.Col.value('@SalesOrderNumber','VARCHAR(40)') AS [ERPORDERID],
				Tab.Col.value('@SalesOrderLineNumber', 'VARCHAR(255)') AS [ERPLINEID],
				Tab.Col.value('@MFGScheduleNumber','VARCHAR(100)') AS [SCHEDULENUMBER],
				Tab.Col.value('@WorkOrderStatus','VARCHAR(30)') AS [STATUS]
			FROM   @ScheduledOrdersData.nodes('/xx:ScheduledOrders/ScheduledOrder') Tab(Col)
		) ScheduleOrders
		INNER JOIN [Approval].[Order] WITH(nolock)
		ON (
			ScheduleOrders.[ERPORDERID]=[Order].[ERPOrderNumber]
		)
		INNER JOIN [Approval].[OrderDetail] WITH(nolock)
		ON (
			[Order].OrderID = [OrderDetail].OrderID AND
			ScheduleOrders.[ERPLINEID]=[OrderDetail].ERPOrderLineNumber
		)
	WHERE
		ISNULL(ScheduleOrders.[STATUS],'')='Released' AND ISNULL(ScheduleOrders.[SCHEDULENUMBER],'') <> '' AND Approval.[Order].StatusID=9

	UPDATE	
		[Approval].[Order]
	SET		
		[StatusID]=10
	FROM	
		[Approval].[Order] INNER JOIN #Temp_Orders 
		ON (
			[Order].OrderID = #Temp_Orders.OrderID
		)

	UPDATE
		[Approval].[OrderDetail]
	SET
		[StatusID]=10, [ScheduleNumber]=#Temp_Orders.[SCHEDULENUMBER]
	FROM
		[Approval].[OrderDetail] INNER JOIN #Temp_Orders
		ON (
			[OrderDetail].OrderDetailID = #Temp_Orders.[OrderDetailID]
		)

	INSERT INTO Integrations.ScheduledOrdersFromEBS (LastSyncedDateTime, Id, InputXMLPayload) VALUES (GETDATE(), NEWID(), @ScheduledOrdersData)

	INSERT INTO Approval.ReviewEventLog 
	(
		OrderEventLogID, OrderID, StatusID, Comment, EventTypeID, EventBy, EventDateUTC
	)
	SELECT NEWID(), [OrderID], 10, 'Order Scheduled',6,'Integration',GETUTCDATE()
	FROM
		#Temp_Orders

	-- CHECK IF PRODUCT IS PRINTING THROUGH DIRECT PRINT
		
    ;WITH cte as
	(
		SELECT tmpo.[OrderDetailID], tmpo.[SCHEDULENUMBER], tmpo.[OrderQuantity]
		FROM
		#Temp_Orders tmpo 
		join 
		[Approval].[OrderDetail] od on od.OrderDetailID = tmpo.OrderDetailID 
		join 
		Project p on p.ProjectID = od.ProjectID
		join 
		Product pr on pr.ProductID = p.ProductID
		where tmpo.OrderDetailID NOT IN (SELECT OrderDetailID FROM Approval.PackList) 
		and pr.IsDirectPrint = 0
	)
	 INSERT INTO Approval.PackList
	(
		PackListID, OrderDetailID, ScheduleNumber, Quantity, CreatedDateUTC
	)
	SELECT NEWID(), [OrderDetailID], [SCHEDULENUMBER], [OrderQuantity], GETUTCDATE()
	FROM cte
		

 --   INSERT INTO Approval.PackList
	--(
	--	PackListID, OrderDetailID, ScheduleNumber, Quantity, CreatedDateUTC
	--)
	--SELECT NEWID(), [OrderDetailID], [SCHEDULENUMBER], [OrderQuantity], GETUTCDATE()
	--FROM
	--	#Temp_Orders o
	--WHERE o.OrderDetailID NOT IN (SELECT OrderDetailID FROM Approval.PackList)
	
	
	/* Step-2 : Store scheduled orders sent from EBS into the “Integrations.OrdersFromEBS” table  */
	DECLARE @CurrentDateUTC datetime = GETUTCDATE()
	INSERT INTO Integrations.OrdersFromEBS
	(
		Id,
		ERPOrderNumber,
		ERPOrderLineNumber,
		ScheduleNumber,
		CreatedDateUTC
	)
	SELECT newid(), [to].ERPORDERID,[to].ERPLINEID,[to].SCHEDULENUMBER,@CurrentDateUTC
	from #Temp_Orders [to]
	
	/* Step-2 : End */

	/* Step-3 : Compare stored scheduling information against all unscheduled orders */
	
	EXEC [dbo].[ScheduleExistingOrders]
	
	/* End */

COMMIT TRANSACTION;
GO


"@
}

function Get-CustomizerOracleDiagnosticQueries {
@"
SELECT COUNT(*) FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)

SELECT SALES_ORDER_NUMBER, SALES_ORDER_LINE_NUMBER, MFG_SCHEDULE_NUMBER, WORK_ORDER_STATUS, CREATION_DATE "LAST_UPDATE_DATE" FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)

SELECT * FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE Sales_Order_Number in ('10694122')

select 
Sysdate, 
SYSDATE - 30/(24*60), 
30/(24*60),
SYSDATE - 30
from dual

SELECT SALES_ORDER_NUMBER, SALES_ORDER_LINE_NUMBER, MFG_SCHEDULE_NUMBER, WORK_ORDER_STATUS, CREATION_DATE "LAST_UPDATE_DATE" FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)



SELECT * FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG

SELECT COUNT(*) FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)

SELECT SALES_ORDER_NUMBER, SALES_ORDER_LINE_NUMBER, MFG_SCHEDULE_NUMBER, WORK_ORDER_STATUS, CREATION_DATE "LAST_UPDATE_DATE" FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG 
WHERE CREATION_DATE>= SYSDATE - 30/(24*60)


SELECT COUNT(*) FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG TB1
INNER JOIN APPS.OE_ORDER_HEADERS_ALL TB2
ON TB1.SALES_ORDER_NUMBER = TB2.ORDER_NUMBER
WHERE TB1.LAST_UPDATE_DATE IS NULL AND 
TB2.SHIPMENT_PRIORITY_CODE IN ('CD-WEB','CD-ITC','RUSH CD-WEB','RUSH CD-ITC','CS-WEB','CS-ITC')

SELECT SALES_ORDER_NUMBER, SALES_ORDER_LINE_NUMBER, MFG_SCHEDULE_NUMBER, WORK_ORDER_STATUS, CREATION_DATE "LAST_UPDATE_DATE", WIP_ENTITY_ID
FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE SALES_ORDER_NUMBER IN 
(SELECT SALES_ORDER_NUMBER FROM (
SELECT DISTINCT TB1.SALES_ORDER_NUMBER FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG TB1
INNER JOIN APPS.OE_ORDER_HEADERS_ALL TB2
ON TB1.SALES_ORDER_NUMBER = TB2.ORDER_NUMBER
WHERE TB1.LAST_UPDATE_DATE IS NULL AND TB2.SHIPMENT_PRIORITY_CODE IN ('CD-WEB','CD-ITC','RUSH CD-WEB','RUSH CD-ITC','CS-WEB','CS-ITC')) WHERE ROWNUM < 100)


SELECT * FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG where last_update_date is null order by WJ_creation_date desc
"@
}

function Install-CustomizerPackListGeneration {
	azurerm
	
	https://www.nuget.org/packages/DocumentFormat.OpenXml/ or equivelant
	connect-azurermaccount -Subscription "Production and Infrastructure - Microsoft Azure Enterprise"

}

function New-CustomyzerBatchNumber {
	(Get-Date).ToString("yyyyMMdd-HHmm")
}

function New-CustomyzerPacklistBatch {
	$PackListItemsNotOnPackList = Get-CustomyzerApprovalPackList -NotOnPacklist
	$BatchNumber = New-CustomyzerBatchNumber
	$PackListItemsNotOnPackList | Set-CustomyzerApprovalPackList -BatchNumber $BatchNumber
	$BatchNumber
}

function Invoke-CutomyzerPackListProcess {
	$BatchNumber = New-CustomyzerPacklistBatch
	#New-CustomyzerPacklist
	$PackListRecords = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber |
	Add-Member -Force -MemberType ScriptProperty -Name SizeAndFormType -PassThru -Value {
		"$($This.OrderDetail.Project.Product.Form.Size)$($This.OrderDetail.Project.Product.Form.FormType.ToUpper())"
	}

	$PackListRecordsSorted = $PackListRecords |
	Sort-Object -Property {$_.OrderDetail.Project.Product.Form.Size},
		{$_.SizeAndFormType},
		{$_.OrderDetail.Order.ERPOrderNumber},
		{$_.OrderDetail.ERPOrderLineNumber}

	New-CustomyzerPacklistXlsx -BatchNumber $BatchNumber -PackListRecords $PackListRecordsSorted
}

function New-CustomyzerPacklistXlsx {
	param (
		$BatchNumber,
		$PackListRecords
	)

	$RecordToWriteToExcel = foreach ($PackListRecord in $PackListRecords) {
		[PSCustomObject]@{
			FormSize = $PackListRecord.SizeAndFormType
			SalesOrderNumber = $PackListRecord.OrderDetail.Order.ERPOrderNumber
			DesignNumber = $PackListRecord.OrderDetail.ERPOrderLineNumber
			BatchNumber = $PackListRecord.BatchNumber
			Quantity = $PackListRecord.Quantity
			ScheduleNumber = $PackListRecord.ScheduleNumber
		}
	}

	$Excel = Export-Excel -Path $ModulePath\PackListTemplate.xlsx -PassThru
	$PackingListWorkSheet = $Excel.Workbook.Worksheets["PackingList"]

	Set-CustomyzerPackListXlsxHeaderValues -PackingListWorkSheet $PackingListWorkSheet -PackListXlsxLines $RecordToWriteToExcel
	Set-CustomyzerPackListXlsxRowValues -PackingListWorkSheet $PackingListWorkSheet -PackListXlsxLines $RecordToWriteToExcel

	$ExcelFileName = "TervisPackList-$BatchNumber.xlsx"	
	$Excel.SaveAs(".\$ExcelFileName")
}

function Set-CustomyzerPackListXlsxHeaderValues {
	param (
		$PackingListWorkSheet,
		$PackListXlsxLines
	)
	
	$DateTime = Get-Date
	$PackingListWorkSheet.Names["DownloadDate"].value = $DateTime.ToString("MM/dd/yyyy")
	$PackingListWorkSheet.Names["DownloadTime"].value = $DateTime.ToString("hh:mm tt")

	$FormSizeQuantitySums = $PackListXlsxLines |
	Group-Object FormSize |
	Add-Member -MemberType ScriptProperty -Name QuantitySum -PassThru -Force -Value {
		$This.Group |
		Measure-Object -Property Quantity -Sum |
		Select-Object -ExpandProperty Sum
	} |
	Add-Member -MemberType ScriptProperty -Name TotalName -PassThru -Force -Value {
		"Total$($This.Name)"
	} |
	Select-Object -Property QuantitySum, TotalName

	foreach ($FormSizeQuantitySum in $FormSizeQuantitySums) {
		$PackingListWorkSheet.Names[$FormSizeQuantitySum.TotalName].value = $FormSizeQuantitySum.QuantitySum
	}

	$GrandTotal = $FormSizeQuantitySums | Measure-Object -Property QuantitySum -Sum | Select-Object -ExpandProperty Sum
	$PackingListWorkSheet.Names["GrandTotal"].Value = $GrandTotal
}

function Set-CustomyzerPackListXlsxRowValues {
	param (
		$PackingListWorkSheet,
		$PackListXlsxLines
	)
	[int]$StartOfPackListLineDataRowNumber = $PackingListWorkSheet.Names["FirstCellOfPacklistLineData"].FullAddressAbsolute -split "\$" |
	Select-Object -Last 1

	$PropertyNameToColumnLetterMap = @{
		FormSize = "A"
		SalesOrderNumber = "C"
		DesignNumber = "E"
		BatchNumber = "G"
		Quantity = "I"
		ScheduleNumber = "P"
	}

	foreach($Line in $PackListXlsxLines) {
		$RowNumber = $PackListXlsxLines.IndexOf($Line) + $StartOfPackListLineDataRowNumber
		$PropertyNames = $Line.psobject.Properties.name
		
		foreach ($PropertyName in $PropertyNames) {
			$ColumnLetter = $PropertyNameToColumnLetterMap.$PropertyName
			$CellAddress = "$ColumnLetter$RowNumber"
			$PackingListWorkSheet.Cells[$CellAddress].Value = $Line.$PropertyName
		}
	}
}

function New-CustomyzerPackListXML {
	param (
		$BatchNumber,
		$PackListRecords
	)
	$XMLFileName = "TervisPackList-$BatchNumber.xml"
	$DateTime = Get-Date

	New-XMLDocument -AsString -InnerElements {
		New-XMLElement -Name packList -InnerElements {
			New-XMLElement -Name batchNumber -InnerText $BatchNumber
			New-XMLElement -Name batchDate -InnerText $DateTime.ToString("MM/dd/yyyy")
			New-XMLElement -Name batchTime -InnerText $DateTime.ToString("hh:mm tt")
			New-XMLElement -Name orders -InnerElements {
				foreach ($PackListRecord in $PackListRecords) {
					New-XMLElement -Name order -InnerElements {
						New-XMLElement -Name salesOrderNumber -InnerText $PackListRecord.OrderDetail.Order.ERPOrderNumber
						New-XMLElement -Name salesLineNumber -InnerText $PackListRecord.OrderDetail.ERPOrderLineNumber
						New-XMLElement -Name itemQuantity -InnerText $PackListRecord.Quantity
						New-XMLElement -Name size -InnerText $PackListRecord.SizeAndFormType
						New-XMLElement -Name itemNumber -InnerText $PackListRecord.OrderDetail.Project.FinalFGCode
						New-XMLElement -Name scheduleNumber -InnerText $PackListRecord.ScheduleNumber
						New-XMLElement -Name fileName -InnerElements ( New-XMLCDATA -Value $PackListRecord.OrderDetail.Project.FinalArchedImageLocation )
					}
				}
			}
		}
	}


#	var orderXML = new XElement("packList",
#	new XElement("batchNumber", batchNumber),
#	   new XElement("batchDate", creationDate.ToString(ConfigurationManager.AppSettings["DateFormat"].ToString())),
#		new XElement("batchTime", creationDate.ToString("hh:mm tt")),
#		 new XElement("orders",
#from row in packListSpreadsheet.SpreadsheetRows
#select
# new XElement("order",
#	 new XElement("salesOrderNumber", row.SalesOrderNumber),
#		 new XElement("salesLineNumber", row.DesignNumber),
#		 new XElement("itemQuantity", row.Quantity),
#		 new XElement("size", row.FormSize),
#		 new XElement("itemNumber", row.ItemNumber),
#		 new XElement("scheduleNumber", row.ScheduleNumber),
#		 new XElement("fileName", new XCData(row.FileName))
#)));
#
#_log.Debug("XML file generated for batch: " + batchNumber);
#return orderXML.ToString();
}

function New-CustomyzerPackListPurchaseRequisitionCSV {
	param (
		$BatchNumber,
		$PackListRecords
	)

	$RecordToWriteToCSV = foreach ($PackListRecord in $PackListRecords) {
		[PSCustomObject]@{
			ITEM_NUMBER = $PackListRecord.OrderDetail.Project.FinalFGCode
			INTERFACE_SOURCECODE = "MIZER_REQ_IMPORT"
			SALES_ORDER_NO = $PackListRecord.OrderDetail.Order.ERPOrderNumber
			SO_LINE_NO =$PackListRecord.OrderDetail.ERPOrderLineNumber
			QUANTITY = $PackListRecord.Quantity			
			VENDOR_BATCH_NAME = $PackListRecord.BatchNumber
			SCHEDULE_NUMBER = $PackListRecord.ScheduleNumber
		}
	}

	$CSVFileName = "xxmizer_reqimport_$BatchNumber.csv"

	#This should be the simple way to accomplish what is needed but we need to confirm that whatever
	#consumes this CSV can handle the values between the delimiter being in "" as currently they are not
	#$RecordToWriteToCSV |
	#ConvertTo-Csv -Delimiter "|" -NoTypeInformation |
	#Out-File -Encoding ascii -FilePath .\$CSVFileName -Force
	
	$CSVHeader = "ITEM_NUMBER",
	"INTERFACE_SOURCECODE",
	"SALES_ORDER_NO",
	"SO_LINE_NO",
	"QUANTITY",
	"VENDOR_BATCH_NAME",
	"SCHEDULE_NUMBER" -join "|"
	
	$CSVRows = $RecordToWriteToCSV | 
	ForEach-Object { $_.psobject.Properties.value -join "|" }
	
	$CSVHeader,$CSVRows | Out-File -Encoding ascii -FilePath .\$CSVFileName
}

function Get-CustomyzerApprovalPackList {
	param (
		[Switch]$NotOnPacklist,
		$BatchNumber
	)

	if ($NotOnPacklist) {
		$ArbitraryWherePredicateParameter = @{
			ArbitraryWherePredicate = @"
			AND Approval.PackList.BatchNumber IS NULL
			AND Approval.PackList.SentDateUTC IS NULL
"@
		}
	}
	$PSBoundParameters.Remove("NotOnPacklist") | Out-Null

	$SQLCommand = New-SQLSelect -SchemaName Approval -TableName PackList @ArbitraryWherePredicateParameter -Parameters $PSBoundParameters

	Invoke-CustomizerSQL -SQLCommand $SQLCommand |
	Add-Member -MemberType ScriptProperty -Name OrderDetail -PassThru -Force -Value {
		$This | Add-Member -MemberType NoteProperty -Name OrderDetail -Force -Value $($This | Get-CustomyzerApprovalOrderDetail)
		$This.OrderDetail
	}
}

function Get-CustomyzerApprovalOrderDetail {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$OrderDetailID
	)
	begin {
		$ArrayList = New-Object System.Collections.ArrayList
	}
	process {
		$ArrayList.Add($OrderDetailID.Guid) | Out-Null
	}
	end {
		$ParametersParameter = if ($OrderDetailID) {
			@{
				Parameter = @{OrderDetailID = $ArrayList}
			}
		} else { @{} }
		
		$SQLCommand = New-SQLSelect -SchemaName Approval -TableName OrderDetail @ParametersParameter

		Invoke-CustomizerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name Project -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Project -Force -Value $($This | Get-CustomyzerProject)
			$This.Project
		} |
		Add-Member -MemberType ScriptProperty -Name Order -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Order -Force -Value $($This | Get-CustomyzerApprovalOrder)
			$This.Order
		}
	}
}

function Get-CustomyzerProject {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$ProjectID
	)
	begin {
		$ArrayList = New-Object System.Collections.ArrayList
	}
	process {
		$ArrayList.Add($ProjectID.Guid) | Out-Null
	}
	end {
		$ParametersParameter = if ($ProjectID) {
			@{
				Parameter = @{ProjectID = $ArrayList}
			}
		} else { @{} }
		
		$SQLCommand = New-SQLSelect -TableName Project @ParametersParameter

		Invoke-CustomizerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name Product -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Product -Force -Value $($This | Get-CustomyzerProduct)
			$This.Product
		}
	}
}

function Get-CustomyzerProduct {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$ProductID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Product -Parameters $PSBoundParameters
		Invoke-CustomizerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name Form -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Form -Force -Value $($This | Get-CustomyzerForm)
			$This.Form
		}
	}
}

function Get-CustomyzerForm {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$FormID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Form -Parameters $PSBoundParameters
		Invoke-CustomizerSQL -SQLCommand $SQLCommand
	}
}

function Get-CustomyzerApprovalOrder {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$OrderID
	)
	process {
		$SQLCommand = New-SQLSelect -SchemaName Approval -TableName Order -Parameters $PSBoundParameters
		Invoke-CustomizerSQL -SQLCommand $SQLCommand
	}
}

function Invoke-CustomzerHelpSQL {

@"	
SELECT MAX(pl.CreatedDateUTC) AS LastScheduledItemImported, 
MAX(pl.SentDateUTC) AS LastPackListGeneration
FROM Approval.PackList AS pl WITH (NOLOCK)

SELECT *
FROM Approval.PackList AS pl WITH (NOLOCK)
WHERE pl.BatchNumber IS NULL
AND pl.sentdateutc is null
ORDER BY pl.CreatedDateUTC DESC

SELECT top 300 *
FROM Approval.PackList AS pl WITH (NOLOCK)
ORDER BY pl.CreatedDateUTC DESC
"@
}