Trap {
	$err = $_.Exception
	write-error $err.Message
	while ( $err.InnerException) 
	{
		$err = $err.InnerException
		Write-Error $err.Message
	}
	break;
}

Set-PSDebug -Strict
$ErrorActionPreference = "stop"

$fullPathIncFileName = $MyInvocation.MyCommand.Definition
$currentScriptName = $MyInvocation.MyCommand.Name
$currentExecutingPath = $fullPathIncFileName.Replace($currentScriptName, "")

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
 
$ScriptBlock = {
	#Functions to Imitate SharePoint 2010 Cmdlets in MOSS 2007
	function global:Get-SPWebApplication($WebUrl)
	 { 
	  return [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($WebUrl)
	 }
	 
	function global:Get-SPSite($Url)
	 {
	    return new-Object Microsoft.SharePoint.SPSite($Url)
	 }
	 
	function global:Get-SPWeb($Url)
	{
	  $site= New-Object Microsoft.SharePoint.SPSite($Url)
	        if($site -ne $null)
	            {
	               $web=$site.OpenWeb();       
	            }
	    return $web
	}
}

try { Add-PSSnapin microsoft.SharePoint.powershell -ErrorAction Stop }
catch
{
	Invoke-Command $ScriptBlock
}

function Global:Test-Log
{
	WriteOut-Log "Oh Snap!"
}

function Global:WriteOut-Log ($msg)
{
	[String]$Local:log = $(Create-Log).ToString()

	Out-File -Append -FilePath $(Create-Log).ToString() -InputObject $msg
}

function Global:Create-Log
{
	<#
		.SYNOPSIS
			Creates a new log file in the default log location
		.DESCRIPTION
			This small internal function will create a log of this execution of the Copy Permissions functions
	#>

	$Local:dirLogLocation = "C:\Logs\"
	$Local:logDateTime = $(Get-Date -Format "yyyy-MM-dd")
	$Local:fileLogName = "Copy-SharePointUserPermissions-$logDateTime.log"
	$Local:logFullPath = $(Join-Path -Path  $Local:dirLogLocation -ChildPath $Local:fileLogName)

	if ((Test-Path $Local:dirLogLocation) -eq 0)
	{
		New-Item -ItemType Directory -Path $Local:dirLogLocation
	}

	if ((Test-Path $Local:logFullPath) -eq 0)
	{
		New-Item -ItemType File -Path $Local:logFullPath
	}

	return $Local:logFullPath.ToString()
}

function global:Copy-UserPermissions
{
	<# 
	.SYNOPSIS 
		Copy SharePoint permissions from user-to-user 
	.DESCRIPTION 
		Allows you to copy SharePoint permissions from one user to another
		on a site-wide/appliation-wide basis
	.NOTES 
		Author     : Larry G. Wapnitsky - larry.wapnitsky@tmnas.com
					 Brendan Thompson <brendan@btsystems.com.au>
	.EXAMPLE 
		Copy-UserPermissions -Url http://site -sourceUsername SourceDom\UserA -destinationUsername DestDOM\UserB
	.PARAMETER Url
	   		Root URL of the site/application on which you will be copying permissions
	.PARAMETER sourceUsername
		Domain and username of sourceUsername user from which you are copying permissions
	.PARAMETER destinationUsername
		Domain and username of destination user to which you are copying permissions
	#>
	
	Param
	(
		[Parameter(Mandatory=$true)]
			[string]$Url,
		[Parameter(Mandatory=$true)]
			[string]$sourceUsername,
		[Parameter(Mandatory=$true)]
			[string]$destinationUsername
	)
	
	$stopwatch = new-object system.diagnostics.stopwatch
	$timeSpan = new-object System.TimeSpan

	$stopwatch.start()
	
	$WebApp = Get-SPWebApplication $Url
	
	[Microsoft.SharePoint.SPSecurity]::RunWithElevatedPrivileges( {
	$SiteCollections = $WebApp.Sites

	foreach ($site in $sitecollections)
	{
		$subsites = @{}
		$site.allwebs | foreach { $subsites.add($(get-spweb $_.Url), $_.Url)} 
		
		$subsiteCounter = 0
	 
		$subsites.GetEnumerator() | foreach {
			$spweb = $_.key
			$subsiteCounter += 1
			
			$timespan = $stopwatch.Elapsed
			$et = "{0:00}:{1:00}:{2:00}" -f $timespan.hours, $timespan.minutes, $timespan.seconds

			write-progress -id 1 -Activity "Processing subsites of $($site.URL)" -CurrentOperation "Analyzing $($spweb.URL)" -PercentComplete (($($subsiteCounter)/$($subsites.Count))*100) -Status "Elapsed Time - $($et)"
			WriteOut-Log "[ACTIVITY] Processing subsites of $($site.URL) -> [CURRENT OPERATION] Analyzing $($spweb.URL)"
			
			
			$srcPerms = $spweb.RoleAssignments | where {$_.member -imatch [regex]::Escape($sourceUsername) }

			WriteOut-Log "ROLE ASSIGNMENT"
			WriteOut-Log "---------------`n"
			
			if ($srcPerms -ne $null) {
				$srcPerms | foreach-Object {
					$RDB = $_.RoleDefinitionBindings
					$RDB | foreach-Object {
						if ($_.Name -ne "Limited Access")
						{
							$spRoleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($($spweb.EnsureUser($destinationUsername)))
							$spRoleAssignment.RoleDefinitionBindings.Add($spWeb.RoleDefinitions[$($_.Name)])
							
							write-progress -id 2 -parentID 1 -Activity "Copying $($sourceUsername) rights for $($destinationUsername)" -status "$($_.Name) on $($_.ParentWeb)"# -Status "Elapsed Time - $($et)"
							WriteOut-Log "`t [ACTIVITY] Copying $($sourceUsername) rights for $($destinationUsername), status $($_.Name) on $($_.ParentWeb)"				

							try {
									$spweb.RoleAssignments.Add($spRoleAssignment)
									$spWeb.Update()
							}
							
							catch 
							{
								write-warning "Rights are inherited from parent - [$($spweb.URL)]" 
								WriteOut-Log "Rights are inherited from parent - [$($spweb.URL)]"
							}
						} 
					}
				}
			}
			
			WriteOut-Log "GROUP ASSIGNMENT"
			WriteOut-Log "--------------`n"

			$GrpColls = @($_.key.SiteGroups, $_.key.Groups)
			foreach ($GrpColl in $GrpColls) {
				$GrpColl | Foreach-Object {

					$Group = $_
					$Group.users | Foreach-Object {
						if (($_.LoginName -imatch [regex]::Escape($sourceUsername)) -and 
								($($Group.Users | Where {$_.LoginName -imatch [regex]::Escape($destinationUsername)}) -eq $null))
						{   
							$timespan_bulk = $stopwatch.Elapsed
							$et_bulk = "{0:00}:{1:00}:{2:00}" -f $timespan_bulk.hours, $timespan_bulk.minutes, $timespan_bulk.seconds
							write-progress -id 2 -parentID 1 -Activity "Copying $($sourceUsername) rights for $($destinationUsername)" -CurrentOperation "Group: $($Group.Name)" -Status "Elapsed Time - $($et)"
							WriteOut-Log "[ACTIVITY] Copying $($sourceUsername) rights for $($destinationUsername) -> [CURRENT OPERATION] Group: $($Group.Name)"
							$Group.AddUser($($spWeb.EnsureUser($destinationUsername)))
						}
						
					} | where {$Group.Users.Count -gt 0}
				}
			} 
		}
	}
	})
	$stopwatch.stop()
}

function global:Copy-UserPermissionsBulk
{
	<# 
	.SYNOPSIS 
		Copy SharePoint permissions from user-to-user in bulk 
	.DESCRIPTION 
		Allows you to specify a CSV file of users and a text file of sites
		to do a mass copy of SharePoint permissions.
	.NOTES 
		Author     : Larry G. Wapnitsky - larry.wapnitsky@tmnas.com,
					 Brendan Thompson <brendan@btsystems.com.au>
	.EXAMPLE 
		Copy-UserPermissionsBulk -UserCSV users.csv -SiteList sitelist.txt
	.PARAMETER UserCSV
	   		User CSV File Format:
		
			"SourceAccount","DestAccount"
			"SourceDom\UserA","DestDOM\UserB"
			
			Please note that the first line is REQUIRED
	.PARAMETER SiteList
		Site File Format:
		
			http://site1
			http://site2
	#> 
 
	Param
	(
		[Parameter(Mandatory=$True,ValueFromPipeline=$true)]
			[ValidateScript({Test-Path $_ -PathType Leaf})] 
			[string]$UserCSV="",
		[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
			[ValidateScript({Test-Path $_ -PathType Leaf})] 
			[string]$SiteList=""
	)
	Process
	{	
		$stopwatch_bulk = new-Object System.Diagnostics.Stopwatch
		$TimeSpan_bulk = new-object system.timeSpan
		
		$SL = $(Get-Content $SiteList)	
		if ($SL.Count -eq $null) {$slCount = 1}
		else { $slCount = $SL.Count}
		
		$UL = $(Import-CSV $UserCSV)

		if ($UL.Count -eq $null) { $ulCount = 1 }
		else { $ulCount = $UL.Count }
		
		$uCount = 1
		$total = $slCount * $ulCount
		
		$stopwatch_bulk.start()
		
		$UL | Foreach-Object {
			$User = $_
			
			$SL | Foreach-Object {
				$timespan_bulk = $stopwatch_bulk.Elapsed
				$et_bulk = "{0:00}:{1:00}:{2:00}" -f $timespan_bulk.hours, $timespan_bulk.minutes, $timespan_bulk.seconds
				
				Write-Progress -ID 10 -Activity "Copying user rights" -CurrentOperation "$($User.SourceAccount) to $($User.DestAccount)" -Status "Bulk Processing Time - $($et_bulk)" -PercentComplete $((($uCount++)/$total)*100)

				try
				{
					Copy-UserPermissions -Url $($_) -sourceUsername $($User.SourceAccount) -destinationUsername $($User.DestAccount)
				}
				catch
				{
					write-warning "Unable to copy permissions from $($User.SourceAccount) to $($User.DestAccount) on $($_)"
				}
			}
		}
	}
}

Write-Host "New Commands added:" -ForegroundColor Green
Write-Host "`tCopy-UserPermissions" 
Write-Host "`tCopy-UserPermissionsBulk"  
Write-Host "`nFor command usage, please use the Get-Help command." -ForegroundColor Yellow
