
#Importing PnP module for PowerShell
Import-Module PnP.PowerShell

#Enter SharePoint URL
$siteURL= 'https://skywalker57.sharepoint.com/sites/Contosoportal' 

#Connect SharePoint site
Write-Host "Connecting to " $siteURL -ForegroundColor Yellow 

Connect-PnPOnline -Url $siteURL -Interactive

#Get SharePoint online CSOM context 
$clientContext = Get-PnPContext

#Enter list name and content type name
$listName= Read-Host 'Enter List name '
$contentTypeName= Read-Host 'Enter Content type name '

#Get specified content type for current context
$contentType = Get-PnPContentType -List $listName -Identity $contentTypeName

#Enter new form component Id
$newFormComponentId= Read-Host 'Enter New form component Id '
$contentType.NewFormClientSideComponentId = $newFormComponentId;

#Enter edit form component Id
$editFormComponentId= Read-Host 'Enter Edit form component Id '
$contentType.EditFormClientSideComponentId = $editFormComponentId;

#Enter display form component Id 
$displayFormComponentId= Read-Host 'Enter Display form component Id '
$contentType.DisplayFormClientSideComponentId = $displayFormComponentId;

#Update changes to SharePoint
$contentType.Update($false)
$clientContext.ExecuteQuery()
Write-Host "Updated content type successfully!"  -ForegroundColor Cyan  

