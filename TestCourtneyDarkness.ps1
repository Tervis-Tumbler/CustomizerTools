$ProjectID = "ea15e71e-3737-40b8-81bf-90564dc483b5"
Set-CustomyzerModuleEnvironment -Name Delta
$Project = Get-CustomyzerProject -ProjectID $ProjectID

$Project.PreviewFlatImageLocation| Out-AdobeScene7UrlPrettyPrint


$ProjectID = "2b09a9bb-c380-416c-be15-3fd1378e914f"
Set-CustomyzerModuleEnvironment -Name Production
$Project30SSProductionExample = Get-CustomyzerProject -ProjectID $ProjectID
$Project30SSProductionExample.PreviewFlatImageLocation | Out-AdobeScene7UrlPrettyPrint

pdfimages -all "Harper New New.pdf" "Harper New New"