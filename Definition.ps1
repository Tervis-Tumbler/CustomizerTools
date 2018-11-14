$CustomyzerEnvionrments = [PSCustomObject]@{
    Name = "Production"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\PRD-MizerFTP"
    PackListXMLDestinationPath = "\\sgs-tervis-mt\User Mailboxes\SGS-NOKOMIS\To Send"
    RequisitionDestinationPath = "/ebsapps/PRD/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "a59786e2-8468-499f-9881-e44ed3a286b1"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "31c6cd88-fbf7-4547-b2a5-3ed9180bd059"
    PasswordListID = 305
    PasswordStateAPIKeyPasswordGUID = "6ab983d1-49c5-42ca-80b0-da6aa79b60ea"
    FileShareAccountPasswordStateGUID = "10d7227f-5f0b-45f1-9d53-a9279fdcbf74"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt1PM"
},
[PSCustomObject]@{
    Name = "Epsilon"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\EPS-MizerFTP"
    PackListXMLDestinationPath = "\\tervis.prv\applications\Customizer\SGS\Epsilon"
    RequisitionDestinationPath = "/ebsapps/SIT/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "9e22d5cf-96d1-4efb-8d76-57fe1f8720d4"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "b4e466c2-5648-4ca6-ad73-b93e38e1af1f"
    PasswordListID = 310
    PasswordStateAPIKeyPasswordGUID = "8cc4a5e5-62f0-4d73-b268-c4a03ef477a4"
    FileShareAccountPasswordStateGUID = "d5db01dc-c120-4f2e-bea4-5b7505c5db8a"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt2PM"
},
[PSCustomObject]@{
    Name = "Delta"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\DLT-MizerFTP"
    PackListXMLDestinationPath = "\\tervis.prv\applications\Customizer\SGS\Delta"
    RequisitionDestinationPath = "/ebsapps/DEV/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "47d1f398-38e1-4f44-8598-4da0a94be31a"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "dc426f75-b719-43fc-b004-3d02cbbd188a"
    PasswordListID = 309
    PasswordStateAPIKeyPasswordGUID = "786643ab-87a0-414a-97ff-9942c0bd904d"
    FileShareAccountPasswordStateGUID = "4b8f6d82-82e1-4f21-a9ef-f7e83af4efc2"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt2PM"
}
