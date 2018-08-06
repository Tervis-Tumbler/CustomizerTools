<ns0:ScheduledTaskDWOrder xmlns:ns0="http://Tervis.BizTalk.CustomizerApprover.ProcessedDWOrder"><ErpOrderId>ErpOrderId_0</ErpOrderId><EcsOrderId>EcsOrderId_0</EcsOrderId></ns0:ScheduledTaskDWOrder>



Invoke-WebRequest -Uri http://dlt-customizerservices.tervis.com/DataServices/BizTalk -Body @"
<BtsActionMapping xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <Operation Name="GetDWOrdersForScheduling" Action="http://tempuri.org/IBizTalkService/GetDWOrdersForScheduling" />
</BtsActionMapping>
"@ -Method Post -ContentType "text/xml; charset=utf-8"

