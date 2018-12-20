$ModulePath = if ($PSScriptRoot) {
	$PSScriptRoot
} else {
	(Get-Module -ListAvailable TervisCustomyzer).ModuleBase
}
. $ModulePath\Definition.ps1

function Get-CustomyzerEnvironment {
	param (
		$EnvironmentName
	)
	$Environment = $CustomyzerEnvionrments |
	Where-Object Name -EQ $EnvironmentName

	$Environment |
	Add-Member -MemberType NoteProperty -Name CustomyzerDBConnectionString -Force -PassThru -Value (
		Get-TervisPasswordstatePassword -Guid $Environment.CustomyzerDatabasePasswordStatePasswordGUID -PropertyMapName MSSQLDatabase -PasswordListID $Environment.PasswordListID |
		ConvertTo-MSSQLConnectionString
	) |
	Add-Member -MemberType NoteProperty -Name EmailAddressToRecieveXLSX -Force -PassThru -Value (
		Get-TervisPasswordstatePassword -Guid $Environment.EmailAddressToRecieveXLSXPasswordStatePasswordGUID -PasswordListID $Environment.PasswordListID |
		Select-Object -ExpandProperty Password
	) |
	Add-Member -MemberType NoteProperty -Name FileShareAccount -Force -PassThru -Value (
		Get-TervisPasswordstatePassword -Guid $Environment.FileShareAccountPasswordStateGUID -PasswordListID $Environment.PasswordListID -AsCredential
	)
}

function Set-CustomyzerModuleEnvironment {
    param (
        [Parameter(Mandatory)]$Name
	)
	$Script:Environment = Get-CustomyzerEnvironment -EnvironmentName $Name
}

function Get-CustomyzerModuleEnvironment {
	if (-not $Script:Environment) {
		throw "Must call Set-CustomyzerModuleEnvironment before using the TervisCustomyzer module"
	}
	$Script:Environment
}

function Invoke-CustomyzerSQL {
    param (
		[Parameter(Mandatory,ParameterSetName="Parameters")]$TableName,
		[Parameter(Mandatory,ParameterSetName="Parameters")]$Parameters,
		[Parameter(ParameterSetName="Parameters")]$ArbitraryWherePredicate,
        [Parameter(Mandatory,ParameterSetName="SQLCommand")]$SQLCommand
	)
	$Environment = Get-CustomyzerModuleEnvironment

	if (-not $SQLCommand) {
		$SQLCommand = New-SQLSelect @PSBoundParameters
	}

    Invoke-MSSQL -ConnectionString $Environment.CustomyzerDBConnectionString -sqlCommand $SQLCommand -ConvertFromDataRow
}


function Get-CustomyzerApprovalReviewEventLogLast10InStatusID10 {
    Invoke-CustomyzerSQL -SQLCommand @"
Select top 10 *
From Approval.ReviewEventLog With (NoLock)
Where StatusID = 10
order by EventDateUTC desc
"@
}

function Get-CustomyzerApprovalOrderLast10InStatusID10 {
    Invoke-CustomyzerSQL -SQLCommand @"
Select top 10 *
From [Approval].[Order] With (NoLock)
Where StatusID = 10
order by [Approval].[Order].[SubmitDateUTC] desc
"@
}

function Get-CustomyzerIntegrationsOrdersFromEBS {
    param (
        $ERPOrderNumber
    )
    Invoke-CustomyzerSQL -SQLCommand @"
Select * from Integrations.OrdersFromEBS where ERPOrderNumber = '$ERPOrderNumber'
"@
}

function Get-CustomyzerOrderNumberInIntegrations.ScheduledOrdersFromEBS {
    param (
        [DateTime]$DateTimeAfterWhichOrderShouldBeFound,
        [DateTime]$DateTimeBeforeWhichOrderShouldBeFound
    )
    $XMLDocuments = Invoke-CustomyzerSQL -SQLCommand @"
Select InputXMLPayload 
From integrations.scheduledOrdersFromEBS With (NoLock)
Where LastSyncedDateTime < '$(($DateTimeAfterWhichOrderShouldBeFound).ToString("yyyy-MM-dd"))'
"@
    
}

function Get-CustomyzerSQLSession {
    Invoke-CustomyzerSQL -SQLCommand @"
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

function Remove-CustomyzerSQLSession {
    param (
        [Parameter(Mandatory)]$SessionFromComputerName,
        [Parameter(Mandatory)]$ProgramName
    )
    Invoke-CustomyzerSQL -SQLCommand @"
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

function Get-CustomyzerInIntegrations.ScheduledOrdersFromEBSPropertiesOfXMLDocument {
    param (
        [Parameter(Mandatory)]$ID
    )
    Invoke-CustomyzerSQL -SQLCommand @"
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

function Get-CustomyzerIntegrations.ScheduledOrdersFromEBSTop100MostRecent {
    Invoke-CustomyzerSQL -SQLCommand @"
select top 100 * from integrations.scheduledOrdersFromEBS With (NoLock)
order by LastSyncedDateTime Desc 
"@
}

function Get-CustomyzerIntegrationsIntegrationStatus {
    Invoke-CustomyzerSQL -SQLCommand @"
SELECT *
  FROM [Integrations].[IntegrationStatus] With (NoLock)
"@
}

function Test-CustomyzerOracleScheduledOrdersAvailable {
@"
SELECT COUNT(*) FROM XXTRVS.XXWIP_DISCRETE_JOBS_MIZER_STG WHERE CREATION_DATE>= SYSDATE - 30/(24*60)
"@
}

function Get-CustomyzerOracleScheduledOrders {
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

function Invoke-CustomyzerOracleScheduledOrdersToCustomyzer {
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

function Get-CustomyzerOracleDiagnosticQueries {
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

function New-CustomyzerBatchNumber {
	(Get-Date).ToString("yyyyMMdd-HHmm")
}



function Get-CustomyzerApprovalPacklistRecentBatch {
	$SQLCommand = @"
SELECT distinct top 10 [BatchNumber]
FROM [Approval].[PackList] (nolock)
Where [CreatedDateUTC] > '$((Get-Date).AddDays(-14).ToString("yyyy-MM-dd"))'
"@
	Invoke-CustomyzerSQL -SQLCommand $SQLCommand | Sort-Object -Property BatchNumber -Descending
}

function Get-CustomyzerApprovalPackList {
	param (
		[Switch]$NotInBatch,
		$BatchNumber
	)

	if ($NotInBatch) {
		$ArbitraryWherePredicateParameter = @{
			ArbitraryWherePredicate = @"
			AND Approval.PackList.BatchNumber IS NULL
			AND Approval.PackList.SentDateUTC IS NULL
"@
		}
	}
	$PSBoundParameters.Remove("NotInBatch") | Out-Null

	$SQLCommand = New-SQLSelect -SchemaName Approval -TableName PackList @ArbitraryWherePredicateParameter -Parameters $PSBoundParameters

	Invoke-CustomyzerSQL -SQLCommand $SQLCommand |
	Add-Member -MemberType ScriptProperty -Name OrderDetail -PassThru -Force -Value {
		$This | Add-Member -MemberType NoteProperty -Name OrderDetail -Force -Value $($This | Get-CustomyzerApprovalOrderDetail)
		$This.OrderDetail
	} |
	Add-Member -Force -MemberType ScriptProperty -Name SizeAndFormType -PassThru -Value {
		"$($This.OrderDetail.Project.Product.Form.Size)$($This.OrderDetail.Project.Product.Form.FormType.ToUpper())"
	}
}

function Set-CustomyzerApprovalPackList {
	param (
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$PackListID,
		$BatchNumber,
		$SentDateUTC
	)
	process {
		$ValueParameters = $PSBoundParameters | ConvertFrom-PSBoundParameters -Property BatchNumber,SentDateUTC -AsHashTable
		$SQLCommand = New-SQLUpdate -SchemaName Approval -TableName PackList -WhereParameters @{PackListID = $PackListID} -ValueParameters $ValueParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand
	}
}

function Get-CustomyzerApprovalOrderDetail {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$OrderDetailID,
		[Parameter(ValueFromPipelineByPropertyName)]$OrderID,
		$ProjectID
	)
	process {
		$SQLCommand = New-SQLSelect -SchemaName Approval -TableName OrderDetail -Parameters $PSBoundParameters

		Invoke-CustomyzerSQL -SQLCommand $SQLCommand |
		Add-TervisMember -MemberType ScriptProperty -Name Project -Force -PassThru -CacheValue -Value {
			$This | Get-CustomyzerProject
		} |
		Add-TervisMember -MemberType ScriptProperty -Name Order -Force -PassThru -CacheValue -Value {
			$This | Get-CustomyzerApprovalOrder
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

		Invoke-CustomyzerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name Product -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Product -Force -Value $($This | Get-CustomyzerProduct)
			$This.Product
		} |
		Add-Member -MemberType ScriptProperty -Name Background -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Background -Force -Value $($This | Get-CustomyzerBackground)
			$This.Background
		} |
		Add-Member -MemberType ScriptProperty -Name Product_Background -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Product_Background -Force -Value $($This | Get-CustomyzerProduct_Background)
			$This.Product_Background
		} |
		Add-Member -MemberType ScriptProperty -Name FinalFGCode -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name FinalFGCode -Force -Value $($This | Get-CustomyzerProjectFinalFGCode)
			$This.FinalFGCode
		}
	}
}

function Get-CustomyzerProjectFinalFGCode {
	param (
		[Parameter(Mandatory,ValueFromPipeLine)]$Project
	)
	process {
		#.BackgroundID test required as otherwise the Product_Background table is queried without BackgroundID in the preciate as opposed to a predicate that includes BackgroundID = ""
		if ($Project.BackgroundID -and $Project.Product_Background.PurchaseFG) {
			$Project.Product_Background.PurchaseFG
		} else {
			$Project.Product.FGCode
		}	
	}
}

function Get-CustomyzerProduct {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$ProductID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Product -Parameters $PSBoundParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name Form -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name Form -Force -Value $($This | Get-CustomyzerForm)
			$This.Form
		}
	}
}

function Get-CustomyzerBackground {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$BackgroundID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Background -Parameters $PSBoundParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand
	}
}

function Get-CustomyzerProduct_Background {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$BackgroundID,
		[Parameter(ValueFromPipelineByPropertyName)]$ProductID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Product_Background -Parameters $PSBoundParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand
	}
}

function Get-CustomyzerForm {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$FormID
	)
	process {
		$SQLCommand = New-SQLSelect -TableName Form -Parameters $PSBoundParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand
	}
}

function Get-CustomyzerApprovalOrder {
	param(
		[Parameter(ValueFromPipelineByPropertyName)]$OrderID,
		[Parameter(ValueFromPipelineByPropertyName)]$ERPOrderNumber
	)
	process {
		$SQLCommand = New-SQLSelect -SchemaName Approval -TableName Order -Parameters $PSBoundParameters
		Invoke-CustomyzerSQL -SQLCommand $SQLCommand |
		Add-Member -MemberType ScriptProperty -Name OrderDetail -Force -PassThru -Value {
			$This | Add-Member -MemberType NoteProperty -Name OrderDetail -Force -Value $($This | Get-CustomyzerApprovalOrderDetail)
			$This.OrderDetail
		}
	}
}

function Invoke-CustomyzerHelpSQL {

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

function Get-CustomyzerPrintURLsFromXML {
	$Content = get-content "C:\TervisPackList-20181101-1300.xml"
	$XML = [xml]$Content
	$XML.packList.orders.order.filename."#cdata-section" | Out-File URLs.txt
}

function Get-CustomyzerWebToPrintImageFileName {
    param (
        [ValidateSet("pdf","png")]$Extension,
        $Size,
        $FormType,
        $ERPOrderNumber,
        $ERPOrderLineNumber,
        [Switch]$WhiteInkImage
    )
    if ($WhiteInkImage -and $Extension -eq "png") {
        "$Size$FormType-$ERPOrderNumber-$ERPOrderLineNumber-WhiteInkOpacityMask.$Extension"
    } elseif ($Extension -eq "png") {
        "$Size$FormType-$ERPOrderNumber-$ERPOrderLineNumber-Color.$Extension"
    } else {
        "$Size$FormType-$ERPOrderNumber-$ERPOrderLineNumber.$Extension"
    }
}

function Get-CustomyzerImageTemplateName {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$Size,
        [ValidateSet("SIP","SWG","WINE","WAV","DWT","DWT","MUG","BEER","DWT","WB")][Parameter(Mandatory,ValueFromPipelineByPropertyName)]$FormType,
        [ValidateSet("Final","Mask","Vignette","Base","Print","FinalWithERPNumber","WhiteInkMask")][Parameter(Mandatory,ValueFromPipelineByPropertyName)]$TemplateType
    )
    DynamicParam {
        if ($TemplateType -eq "Print") {
            New-DynamicParameter -Name PrintTemplateType -ValidateSet "Illustrator","InDesign","Scene7"
        }
    }
    process {
        if (-not $Script:SizeAndFormTypeToImageTemplateNamesIndex) {
            $Script:SizeAndFormTypeToImageTemplateNames |
            Add-Member -MemberType ScriptProperty -Name SizeAndFormTypes -Force -Value {
                $This.FormType |
                ForEach-Object -Process {
                    "$($This.Size)$_"
                }
            }
            $Script:SizeAndFormTypeToImageTemplateNamesIndex = $Script:SizeAndFormTypeToImageTemplateNames |
            New-HashTableIndex -PropertyToIndex SizeAndFormTypes
        }

        $Template = $Script:SizeAndFormTypeToImageTemplateNamesIndex."$Size$FormType".ImageTemplateName.$TemplateType
        if (-not $PSBoundParameters.PrintTemplateType) {
            $Template
        } else {
            $Template."$($PSBoundParameters.PrintTemplateType)"
        }
    }
}

function Get-CustomyzerPrintImageTemplateSizeAndFormType {
    param (
        $PrintImageTemplateName
    )
    if (-not $Script:PrintImageTemplateNameToSizeAndFormTypeIndex) {

        $Script:SizeAndFormTypeToImageTemplateNames |
        Add-Member -MemberType ScriptProperty -Name PrintImageTemplateNames -Force -Value {
            $This.ImageTemplateName.Print.Values | 
            Where-Object {$_}
        }

        $Script:PrintImageTemplateNameToSizeAndFormTypeIndex = $Script:SizeAndFormTypeToImageTemplateNames |
        New-HashTableIndex -PropertyToIndex PrintImageTemplateNames
    }

    $Script:PrintImageTemplateNameToSizeAndFormTypeIndex.$PrintImageTemplateName |
    Select-Object -Property Size, FormType
}