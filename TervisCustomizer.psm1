function Invoke-CustomizerSQL {
    param (
        $SQLCommand
    )
    $CustomizerAccountEntry = Get-PasswordstateMSSQLDatabaseEntryDetails -PasswordID 5366
    $CustomizerSQLCredential = Get-PasswordstateCredential -PasswordID 5366
    Invoke-MSSQL -Server $CustomizerAccountEntry.Host -database $CustomizerAccountEntry.DatabaseName -sqlCommand $SQLCommand -Credential $CustomizerSQLCredential -ConvertFromDataRow
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