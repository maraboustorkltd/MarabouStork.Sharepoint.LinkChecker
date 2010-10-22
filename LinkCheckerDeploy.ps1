param (`
   [string]$url=$(throw "you must specify a url."), `
   [string]$resourcePath=$(throw "you must specify a resource path.") `
  )

#param (`
#     [string]$url="http://localhost:8081", `
#	  [string]$resourcePath="" `
#     )

function main() 
{
    #Include the SharePoint cmdlets
	if ((Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null )
	{
		write-host "Add-PsSnapin Microsoft.Sharepoint.PowerShell"
		Add-PSSnapin Microsoft.SharePoint.PowerShell
	}

	$web = Get-SPWeb($url)	
	
	#Load the custom timer job assembly
	Write-Host Loading Assembly
	[System.Reflection.Assembly]::Load("MarabouStork.Sharepoint.LinkChecker, Version=1.0.0.0, Culture=neutral, PublicKeyToken=5f0a99c6679b2dcc")

	#Restart the timer service and delete the existing timer job
	Write-Host Restart Timer Service....
	Restart-Service SPTimerV4
	
	Create-InvalidUrlsInResources
	
	Persist-LinkCheckerSettings
	
	Create-LinkCheckerTimerJob

	$web.Dispose()
}

function Create-InvalidUrlsInResources()
{
	$listname = "InvalidUrlsInResources"
	
	$userfieldtype = [Microsoft.Sharepoint.SPFieldType]::User
	$textfieldtype = [Microsoft.Sharepoint.SPFieldType]::Text
	$urlfieldtype = [Microsoft.Sharepoint.SPFieldType]::URL
	
	# Drop and recreate the lists
	if($web.Lists[$listname] -ne $null)
	{			
		Write-Host Dropping list $listname		
		$web.Lists[$listname].Delete()
		while($web.Lists[$listname] -ne $null)
		{
			Start-Sleep -Seconds 1
		}
	}
		
	Write-Host Creating List $listname
	$web.Lists.Add($listname,$listname,"GenericList")
	$list = $web.Lists[$listname]
	$list.OnQuickLaunch = 1
	$list.Fields.Add("User", $textfieldtype, 1)
	$list.Fields.Add("Document", $urlfieldtype, 1)
	$list.Fields.Add("Message", $textfieldtype,  0)
	
	#Add the columns
		
	$list.Update()		
		
	$spView = $web.GetViewFromUrl("/Lists/" + $listname + "/AllItems.aspx") 
	$spView.ViewFields.Add( $list.Fields["User"]) 
	$spView.ViewFields.Add( $list.Fields["Document"]) 
	$spView.ViewFields.Add( $list.Fields["Message"]) 
	
	$spView.Update() 

}#

function Create-LinkCheckerTimerJob()
{
	$web = Get-SPWeb($url)	
	$jobName = "MarabouStork.Sharepoint.LinkChecker"

	#Delete the timer job if it already exists
	Get-SPTimerJob | where { $_.TypeName -eq "MarabouStork.Sharepoint.LinkChecker.LinkCheckerTimerJob" } | % { $_.Delete() }
	
	#Create an execution schedule
	[Microsoft.SharePoint.SPSchedule] $schedule = [Microsoft.SharePoint.SPSchedule]::FromString("Every 10 minutes between 0 and 59") 
	
	#Instantiate the custom timer job object 
	[MarabouStork.Sharepoint.LinkChecker.LinkCheckerTimerJob] $linkChecker = New-Object "MarabouStork.Sharepoint.LinkChecker.LinkCheckerTimerJob" ($jobName,(Get-SPWebApplication $url))
	
	$linkChecker.Schedule = $schedule
	$linkChecker.Title = $jobName
	
	$linkChecker.Update($true)
	
	Get-SPTimerJob -Identity $jobName | Start-SPTimerJob
}

function Persist-LinkCheckerSettings()
{
	Write-Host Persisting Timer Job Settings
	
	#Setup persisted parameters to configure link checker 
	$site = $web.Site
	$webApplication = $site.WebApplication
	$configName = "MarabouStork.Sharepoint.LinkChecker.LinkCheckerSettings"
	
	$farm = [Microsoft.SharePoint.Administration.SPFarm]::Local
	$existingSettingsobj = $farm.GetObject($configName, $webApplication.Id, [MarabouStork.Sharepoint.LinkChecker.LinkCheckerPersistedSettings])
    if ($existingSettingsObj -ne $null) { $existingSettingsObj.Delete() }             
	
	#Create a new persisted settings object
	[MarabouStork.Sharepoint.LinkChecker.LinkCheckerPersistedSettings] $settings = New-Object "MarabouStork.Sharepoint.LinkChecker.LinkCheckerPersistedSettings" ($configName, $webApplication)
	$settings.SiteUrl = $url
	$settings.DocLibraries = "Introduction Resources"
	$settings.FieldsToCheck = "my:URL;my:Description"
	$settings.UnpublishInvalidDocs = 1
	$settings.Update();
}

main
