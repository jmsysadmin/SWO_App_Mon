<#	
	.NOTES
	===========================================================================
	 Created by:   	Jake Muszynski
	 Thwack ID: 	jm_sysadmin
	 Filename:     	SWO_App_Monitor_Automation.ps1        
	===========================================================================
	.SYNOPSIS
		This script automates the creation and assgnment of application monitors 
		for Windows Servers in Orion.
	.DESCRIPTION
		Script requires rights to administer Orion, and the targeted windows servers. 
		Use this script to gather Services and active ports on a list of servers 
		and build a XML application template for use in Orion.
	.LINK
		https://github.com/jmsysadmin
	.PARAMETER OrionServerName
	.PARAMETER Quiet
		Suppress Interactive prompts, use script defaults
	.PARAMETER ServerListSource
		Select where to pull server list from, can be 'CSV' or 'Orion'
	.PARAMETER FilterForProduction 
		Limits servers to nodes marked as 'Production'. Can be set to 'SKIP' to 
		ignore the filter, or set to the name of a custom property where the 
		value says "Production"
	.PARAMETER ApplicationName
		Can be set to 'SKIP' to ignore, or the custom property name that holds the application name on the server
	.PARAMETER TargetedServers
		Choose "All" or "Unmonitored" Servers to include in the list of servers
	.PARAMETER ImportTemplateToOrion
		$true will upload the template to Orion, $false does not
	.PARAMETER AssignTemplateToNode
		If the template is uploaded to Orion, $true will assign it to the node it was created from, $false does not 
#>

Param (
	[String]
	$OrionServerName = 'h2pwswhcowc01.columbuschildrens.net',
	# Assumes that script is running with rights to this server
	
	[Switch]
	$Quiet,
	#Suppress Interactive prompts, use script defaults
	
	[ValidateSet("CSV", "Orion")]
	[String]
	$ServerListSource = "Orion",
	#Can Be 'CSV' or 'Orion'

	[String]
	$FilterForProduction = 'Environment',
	#Can be set to 'SKIP' to ignore the filter, or the name of a custom property where the value says "Production"

	[String]
	$ApplicationName = 'Application',
	#Can be set to 'SKIP' to ignore, or the custom property name that holds the Application on the server

	[ValidateSet("All", "Unmonitored")]
	[String]
	$TargetedServers = 'Unmonitored',
	
	[Boolean]
	$ImportTemplateToOrion = $True,
	
	[Boolean]
	$AssignTemplateToNode = $True
)

Write-Progress -Activity 'Gathering Data' -Status 'Loading functions' -PercentComplete -1 -id 0

Function Start-Log ($prefix) {
	$filename = Get-Date -Format FileDate
	$script:OutputLog = $PSScriptRoot + "\logs\" + $prefix + "_" + $filename + ".csv"
	$First_Line = "sep=" + "`t" # makes csv readable as tab delimited in Excel
	$Column_Header_String = "Timestamp`t Message`t"
	Add-Content -Path $OutputLog -Value $First_Line
	Add-Content -Path $OutputLog -Value $Column_Header_String
	Send-LogMessage -Level '' -message 'Logging started'
}
Function Send-LogMessage ($Level, $message) {
	$line = $null
	#match $level to value for message formating, use default if not matched 
	Switch ($Level) {
		Critical {
			$line = (get-date).ToString("MM/dd/yyyy HH:mm") + "`tCritical: " + $message
			Write-Host $message -ForegroundColor Red
			Break
		}
		Warning  {
			$line = (get-date).ToString("MM/dd/yyyy HH:mm") + "`tWarning: " + $message
			Write-Host $message -ForegroundColor Yellow
			Break
		}
		Success  {
			$line = (get-date).ToString("MM/dd/yyyy HH:mm") + "`tSuccess: " + $message
			Write-Host $message -ForegroundColor Green
			Break
		}
		Default {
			$line = (get-date).ToString("MM/dd/yyyy HH:mm") + "`t" + $message
			Write-Host $message -ForegroundColor White
		}
	}
	Add-Content -Path $OutputLog -Value $line
	$message = $Null
}
Function Get-ScriptDirectory {
	[OutputType([string])]
	Param ()
	If ($null -ne $hostinvocation) {
		Split-Path $hostinvocation.MyCommand.path
	} Else {
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}
Function Get-ApplicationTemplateXMLStart ($inputfile) {
	$ApplicationTemplateXMLStart = Get-Content -path $inputfile -Raw
	Return $ApplicationTemplateXMLStart
}
Function Get-ApplicationTemplateXMLEnd ($inputfile, $TemplateDisplayName, $TemplateDescription) {
	$ApplicationTemplateXMLEnd = Get-Content -path $inputfile -Raw
	$TemplateTimeStamp = $(Get-Date -Format u) -replace ' ', 'T'
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@TEMPLATENAME@@@", $TemplateDisplayName
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@DESCRIPTION@@@", $TemplateDescription
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@CREATEDDATE@@@", $TemplateTimeStamp
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@MODEIFIEDDATE@@@", $TemplateTimeStamp
	$uniqueid = [GUID]::NewGuid().ToString()
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@GUID@@@", $uniqueid
	Return $ApplicationTemplateXMLEnd
}
Function Get-ServiceComponentXML ($inputfile, $ComponentName, $ServiceName) {
	$ComponentXML = Get-Content -path $inputfile -Raw
	$script:ComponentCounter = $ComponentCounter + 1
	$script:ComponentID = $ComponentCounter + 1
	$ComponentXML = $ComponentXML -replace "@@@ComponentID@@@", $ComponentCounter
	$ComponentXML = $ComponentXML -replace "@@@ComponentOrder@@@", $ComponentCounter
	$ComponentXML = $ComponentXML -replace "@@@WindowsServiceMonitor@@@", $ComponentName
	$ComponentXML = $ComponentXML -replace "@@@ServiceName@@@", $ServiceName
	$uniqueid = [GUID]::NewGuid().ToString()
	$ComponentXML = $ComponentXML -replace "@@@GUID@@@", $uniqueid
	Return $ComponentXML
}
Function Get-PortBasedComponentXML ($InputFolder, $PortNumber, $PortDescription) {
	Switch ($PortNumber) {
		80 {
			$InputFile = $InputFolder + '\http.txt'
			$script:ComponentCounter = $ComponentCounter + 1
			$HttpXML = (Get-Content -path $inputfile -Raw) -replace "@@@ComponentOrder@@@", $ComponentCounter
			$uniqueid = [GUID]::NewGuid().ToString()
			$HttpXML = $HttpXML -replace "@@@GUID@@@", $uniqueid
			Return $HttpXML
		}
		443 {
			$script:ComponentCounter = $ComponentCounter + 1
			$HttpsInputFile = $InputFolder + '\https.txt'
			$HttpsXML = (Get-Content -path $HttpsInputFile -Raw) -replace "@@@ComponentOrder@@@", $ComponentCounter
			$uniqueid = [GUID]::NewGuid().ToString()
			$HttpsXML = $HttpsXML -replace "@@@GUID@@@", $uniqueid
			
			#$script:ComponentCounter = $ComponentCounter + 1
			#$SSLCertInputFile = $InputFolder + '\SSLCert.txt'
			#$SSLCertXML = (Get-Content -path $SSLCertInputFile -Raw) -replace "@@@ComponentOrder@@@", $ComponentCounter
			#$uniqueid = [GUID]::NewGuid().ToString()
			#$SSLCertXML = $SSLCertXML -replace "@@@GUID@@@", $uniqueid
			
			$HttpsOut = $HttpsXML # + $SSLCertXML
			Return $HttpsOut
		}
		default {
			$InputFile = $InputFolder + '\Port.txt'
			$script:ComponentCounter = $ComponentCounter + 1
			$PortXML = (Get-Content -path $inputfile -Raw) 
			$PortXML = $PortXML -replace "@@@ComponentOrder@@@", $ComponentCounter
			$PortXML = $PortXML -replace "@@@Port@@@", $PortNumber
			$PortXML = $PortXML -replace "@@@PortDescription@@@", $PortDescription
			$uniqueid = [GUID]::NewGuid().ToString()
			$PortXML = $PortXML -replace "@@@GUID@@@", $uniqueid
			
			Return $PortXML
		}
	}
}
Function Get-TrimmedServiceList ($Services) {
	$NewlyInstalledServices = @()
	ForEach ($Service In $Services) {
		$AddService = $True
		ForEach ($templateservice In $TemplateServiceList) {
			If ($templateservice.Name -like $Service.Name) {
				$AddService = $False
				Break
			} ElseIf ($Service.Name -like $templateservice.Name) {
				$AddService = $False
				Break
			} ElseIf ($Service.DisplayName -like 'Application Host Helper Service') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'ASP.NET State Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'AzureAttestService*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Data Deduplication Volume Shadow Copy Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Connected User Experiences and Telemetry*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Commvault Client Manager Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Commvault Communications Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Commvault Network Daemon*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server Integration Services*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server Browser*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server Agent*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server CEIP*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server Reporting Services*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SQL Server VSS Writer*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Net Driver HPZ12*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Pml Driver HPZ12*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Windows Search*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Net.Msmq Listener Adapter*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Net.Tcp Port Sharing Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Net.Pipe Listener Adapter*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Net.Tcp Listener Adapter*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'DameWare Mini Remote Control*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'ASP.NET State Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Windows Audio*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Windows Update*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'IPsec Policy Agent*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'ControlUp Agent*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Application Experience*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Background Intelligent Transfer Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'DameWare Mini Remote Control*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Java Quick Starter*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'KtmRm for Distributed Transaction Coordinator*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'TCP/IP NetBIOS Helper*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Network List Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Plug and Play*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'IPsec Policy Agent*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Secondary Logon*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Centricity Clinical Gateway (CCG) Service Tools*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Software Licensing*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Tablet PC Input Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Terminal Services*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Desktop Window Manager Session Manager*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Windows Error Reporting Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Windows Update*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Device Association Service*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'ScreenConnect Client*') { 
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SMS AntiMalware') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'CrowdStrike Falcon Sensor Service') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Netlogon') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SNMP Service') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Microsoft Antimalware Service') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SolarWinds Agent') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'SMS Agent Host') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Sync Host_*') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'VMWare*') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Adobe Acrobat Update Service') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Configuration Manager Remote Control') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Performance Counter DLL Host') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'Performance Logs & Alerts') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'cbdhsvc_*') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'CDPUserSvc_*') {
				$AddService = $False
			} ElseIf ($Service.DisplayName -like 'WpnUserService_*') {
				$AddService = $False			
			}
		}
		If ($AddService) {
			$NewlyInstalledServices += $Service
		} 
	}
	Return $NewlyInstalledServices
}

Write-Progress -Activity 'Gathering Data' -Status 'Loading Solarwinds SWIS Module' -PercentComplete -1 -id 0
Start-Log -prefix "SWO_App_Automation_"
Send-LogMessage -Level '' -message $PSCommandPath
Try {
	Import-Module SwisPowerShell
} Catch {
	Send-LogMessage -Level 'Critical' -message "Message: $($Error[0])"
	Exit 1;
}

$HCO_Credentials = Get-Credential
$global:SwisConnection = $null
$SwisConnection = Connect-Swis -Hostname $OrionServerName -credential $HCO_Credentials
$TestQueryResults = Get-SwisData -SwisConnection $SwisConnection -Query "SELECT TOP 10 Caption, IPAddress, ObjectSubType FROM Orion.Nodes"
If ($TestQueryResults.count -ne 10) {
	Send-LogMessage -Level 'Critical' -message "Message: $($Error[0])"
	Exit 1;
}

Write-Progress -Activity 'Gathering Data' -Status 'Initial Script Setup' -PercentComplete -1 -id 0
Send-LogMessage -Level '' -message 'Initial Script Setup'
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Description."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Description."
$YESNOoptions = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
If ($Quiet) {
	$EditExclustionsResponse = 1
} Else {
	$title = "Exclusions to the found services can be added."
	$message = "Do you want to add Service exclusions during this script run?"
	$EditExclustionsResponse = $host.ui.PromptForChoice($title, $message, $YESNOoptions, 1)
}

$ComponentCounter = 0
$ComponentID = 25000
$ScriptFolder = Get-ScriptDirectory
$XMLSourceFolder = $ScriptFolder + '\resources'
$XMLStartFile = $XMLSourceFolder + '\AppFile1.txt'

Write-Progress -Activity 'Gathering Data' -Status 'Loading Server List' -PercentComplete -1 -id 0
Send-LogMessage -Level '' -message 'Loading Server List'
Write-Host $ServerListSource
If ($ServerListSource -eq "CSV") {
	Send-LogMessage -Level '' -message "Importing $ScriptFolder\Servers.csv"
	$Servers = Import-Csv ($ScriptFolder + '\Servers.csv') -Header Server, NodeID, NodeUsedFor, Applications
	$ServersInCSV = ($Servers.server).count
	If (($Servers.server).count -eq 0) {
		Send-LogMessage -Level 'Critical' -message 'No Servers imported from CSV'
		Exit
	} Else {
		Send-LogMessage -Level '' -message "Found $ServersInCSV Server(s) in CSV"
	}
} ElseIf ($ServerListSource -eq "Orion") {
	Send-LogMessage -Level '' -message 'Orion selected as location to get server list from'
	If ($FilterForProduction -like 'SKIP') {
		$ProdColumn = ""
		$ProdGroupBy = ""
	} Else {
		$ProdColumn = ", N.CustomProperties.$FilterForProduction as [NodeUsedFor]"
		$ProdGroupBy = "N.CustomProperties.$FilterForProduction"
	}
	If ($ApplicationName -like 'SKIP') {
		$AppColumn = ""
		$AppGroupBy = ""
	} Else {
		$AppColumn = ", N.CustomProperties.$ApplicationName as [Applications]"
		$AppGroupBy = "N.CustomProperties.$ApplicationName"
	}
	$OptionalColumns = $ProdColumn + $AppColumn
	$OptionalGroupBy = ", $ProdGroupBy, $AppGroupBy"
	$QueryServerList = "SELECT N.NodeID, N.caption as [Server], COUNT(N.Applications.ApplicationID) as [Monitors]$OptionalColumns 
		FROM Orion.Nodes N Where N.Vendor like 'Windows' and N.IsServer = TRUE and $ProdGroupBy like 'Production' and N.Caption not like '%wcxx%' and $AppGroupBy is not Null
		Group By N.NodeID, N.caption$OptionalGroupBy"
	Send-LogMessage -Level '' -message 'Importing Windows Servers from Orion'
	Send-LogMessage -Level '' -message $QueryServerList
	$Servers = Get-SwisData ($SwisConnection) -Query $QueryServerList
	If (($Servers.server).count -eq 0) {
		Write-Progress -Activity 'Exiting Script' -Status 'No Servers Selected' -PercentComplete 100 -id 0
		Send-LogMessage -Level '' -message 'No Servers imported from Orion SWIS Query, Exiting'
		Exit 0
	}	
	Write-Host "Target: $TargetedServers"
	If ($TargetedServers -like 'Unmonitored') {
		Send-LogMessage -Level '' -message 'Filtering Servers from Orion for nodes that do not have Existing Application Monitors'
		$Servers = $Servers | Where-Object -Property Monitors -EQ 0
	}
}

If (($Servers.server).count -eq 0) {
	Write-Progress -Activity 'Exiting Script' -Status 'No Servers Selected' -PercentComplete 100 -id 0
	Send-LogMessage -Level '' -message 'No Production Servers found, Exiting'
	Exit 0
}

If (!$Quiet) {
	Send-LogMessage -Level '' -message 'Interactive Prompt for Servers Selection'
	$Servers = $Servers | Out-GridView -OutputMode Multiple -Title 'Select any Servers you want to monitor applications on. Use Ctrl to pick multiple'
}
$SelectedServerCount = ($Servers.Server).count
$ProgressServerCount = 0
Send-LogMessage -Level '' -message "Number of Servers Selected: $SelectedServerCount"
ForEach ($ServerItem In $Servers) {
	$Server = $ServerItem.Server
	$ProgressServerCount = $ProgressServerCount + 1
	if (!(Test-Connection -ComputerName $Server -Quiet)) {
		Write-Progress -Activity "Connecting to $Server" -Status "$Server cannot be reached, skipping" -PercentComplete -1 -id 1
		Send-LogMessage -Level 'Warning' -message "$Server cannot be reached, skipping"
		continue 
	}
	
	$Application = $ServerItem.Applications
	Write-host $ServerItem

	$OverallPercentComplete = ($ProgressServerCount/$SelectedServerCount) * 100
	If ($OverallPercentComplete -gt 100){$OverallPercentComplete = 100}
	Write-Progress -Activity 'Working server list' -Status "Starting $Server Application Template" -PercentComplete $OverallPercentComplete -id 0
	Send-LogMessage -Level '' -message "Starting $Server Application Template"
	$TemplateOutputXML = $null
	$TemplateOutputXML = (Get-ApplicationTemplateXMLStart -inputfile $XMLStartFile)
	Write-Progress -Activity "Connecting to $Server" -Status "Getting $Server OS Information" -PercentComplete $OverallPercentComplete
	Send-LogMessage -Level '' -message "Getting $Server OS Information"
	$error.clear()
	$OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Server
	If ($error.count -gt 0) {
		Send-LogMessage -Level 'Critical' -message "Message: $($Error[0])"
		$error.clear()
		Continue
	}
	Write-Progress -Activity "Connecting to $Server" -Status "Getting $Server Services" -PercentComplete $OverallPercentComplete
	Send-LogMessage -Level '' -message "Getting $Server Services"
	$ServiceList = Get-Service -ComputerName $Server | Where-Object starttype -like 'Automatic' | Where-Object status -like 'Running'
	$TemplateServiceList = Import-Csv ($ScriptFolder + "\filters\" + $OS.Caption + '_Services.csv')
	Write-Progress -Activity "$Server" -Status "Scrubbing $Server Service list" -PercentComplete -1 -id 1
	Send-LogMessage -Level '' -message "Scrubbing $Server Service list"
	$ServicesSuggestedForMonitoring = Get-TrimmedServiceList -Services $ServiceList
	Switch ($EditExclustionsResponse) {
		0 {
			$ServicesToExclude = $ServicesSuggestedForMonitoring | Out-GridView -OutputMode Multiple -Title 'Select any services that should be excluded from future recommendations. Use Ctrl to pick multiple'
			If ($ServicesToExclude.count -ge 1) {
				$ServicesToExclude.name | Add-Content -Path ($ScriptFolder + "\" + $OS.Caption + '_Services.csv')
				$ServicesSuggestedForMonitoring = Get-TrimmedServiceList -Services $ServicesSuggestedForMonitoring
			}
		}
		1 {
			# Do Nothing for a 'No' 
		}
	}
	If ($Quiet) {
		$SelectedServicesToMonitor = $ServicesSuggestedForMonitoring
	} Else {
		$SelectedServicesToMonitor = $ServicesSuggestedForMonitoring | Out-GridView -OutputMode Multiple -Title 'Select any services that should monitored for this template. Use Ctrl to pick multiple'
		Send-LogMessage -Level '' -message "Manaully Selected Sevices to Include in Monitor"
	}
	$ReportServicesHTML = $SelectedServicesToMonitor | ConvertTo-Html -Property DisplayName, Name, Status, StartType -Fragment -PreContent "<h2>Monitored Services</h2>"
	$ServiceXMLSource = $XMLSourceFolder + '\Service.txt'
	Write-Progress -Activity "$Server" -Status "Adding Services to XML" -PercentComplete $OverallPercentComplete
	Send-LogMessage -Level '' -message "Adding Services to XML"
	If (($SelectedServicesToMonitor.Name).count -eq 1) {
		$TemplateOutputXML = $TemplateOutputXML + (Get-ServiceComponentXML -inputfile $ServiceXMLSource -ComponentName $SelectedServicesToMonitor.DisplayName -ServiceName $SelectedServicesToMonitor.Name)
		$ServiceDisaplayName = $ServicesSuggestedForMonitoring.DisplayName
		Send-LogMessage -Level '' -message "Adding Service: $ServiceDisaplayName"
	} ElseIf (($SelectedServicesToMonitor.Name).count -gt 1) {
		ForEach ($SelectedService In $SelectedServicesToMonitor) {
			$TemplateOutputXML = $TemplateOutputXML + (Get-ServiceComponentXML -inputfile $ServiceXMLSource -ComponentName $SelectedService.DisplayName -ServiceName $SelectedService.Name)
			$ServiceDisaplayName = $SelectedService.DisplayName
			Send-LogMessage -Level '' -message "Adding Service: $ServiceDisaplayName"
		}
	}
	$PortComponentsXML = $null
	$PortCount = 0
	#$ports = 20, 21, 22, 23, 25, 53, 69, 80, 88, 389, 443, 445, 389, 1433, 1434, 3306, 5060, 5432
	$ports = 20, 21, 80, 443, 389, 1433, 1434, 3306, 5060, 5432
	Send-LogMessage -Level '' -message "Checking Common Ports"
	Write-Progress -Activity "$Server" -Status "Checking Common Ports" -PercentComplete $OverallPercentComplete
	$ReportPorts = @()
	$Ports | ForEach-Object {
		Write-Progress -Activity "$Server" -Status "Checking Common Ports" -PercentComplete $OverallPercentComplete
		$port = $_; If (Test-NetConnection -ComputerName $Server -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue) {
			$PortCount = $PortCount + 1
			Switch ($port) {
				20   {
					$PortDescription = 'FTP'
				}
				21   {
					$PortDescription = 'FTP'
				}
				22   {
					$PortDescription = 'SSH'
				}
				23   {
					$PortDescription = 'Telnet'
				}
				25   {
					$PortDescription = 'SMTP'
				}
				53   {
					$PortDescription = 'DNS'
				}
				69   {
					$PortDescription = 'TFTP'
				}
				80   {
					$PortDescription = 'Http'
				}
				88   {
					$PortDescription = 'Kerberos'
				}
				389  {
					$PortDescription = 'LDAP'
				}
				443  {
					$PortDescription = 'Https'
				}
				445  {
					$PortDescription = 'Microsoft Directory Services'
				}
				389  {
					$PortDescription = 'LDAP'
				}
				1433 {
					$PortDescription = 'Microsoft SQL'
				}
				1434 {
					$PortDescription = 'Microsoft SQL'
				}
				3306 {
					$PortDescription = 'MySQL'
				}
				5060 {
					$PortDescription = 'SIP'
				}
				5432 {
					$PortDescription = 'PostgreSQL'
				}
				default {
					$PortDescription = 'Application'
				}
			}
			$PortComponentsXML = $PortComponentsXML + (Get-PortBasedComponentXML -InputFolder $XMLSourceFolder -PortNumber $Port -PortDescription $PortDescription)
			$ReportPorts += [pscustomobject]@{ 'Port Description' = $PortDescription; 'Port' = $Port
			}
			Send-LogMessage -Level 'Success' -message "Adding Port monitor for $Port $PortDescription"
		} Else {
			Send-LogMessage -Level 'Warning' -message "$Port is not listening, monitor skipped"
		}
	}
	$ReportPortsHTML = $ReportPorts |  ConvertTo-Html  -Fragment -PreContent "<h2>Monitored Ports</h2>"
	$TemplateOutputXML = $TemplateOutputXML + $PortComponentsXML
	$XMLEndFile = $XMLSourceFolder + '\AppFile2.txt'
	If ($quiet) {
		If ($Application.length -gt 1) {
			$OutTemplateDisplayName = $Application
		} Else {
			$OutTemplateDisplayName = (Get-Date -Format FileDate) + '-' + $server
		}	
		$OutTemplateDescription = Get-Date
	} Else {
		$OutTemplateDisplayName = Read-Host -Prompt 'What do you want to name this monitoring template?'
		$OutTemplateDescription = Read-Host -Prompt 'What do you want to set as the template description?'
	}
	
	$FoundItemCount = 0
	$FoundItemCount = (($SelectedServicesToMonitor.count) + ($ReportPorts.count))
	If ($FoundItemCount -gt 0) {
		Send-LogMessage -Level 'Success' -message "Template has $FoundItemCount services and ports for $Server" 
		$ItemsFound = $True 
	} else {
		Send-LogMessage -Level 'Warning' -message "Template has $FoundItemCount services and ports, no template or assignment will happen for $Server" 	
		$ItemsFound = $False 
	}	
	
	$TemplateOutputXML = $TemplateOutputXML + (Get-ApplicationTemplateXMLEnd -inputfile $XMLEndFile -TemplateDisplayName $OutTemplateDisplayName -TemplateDescription $OutTemplateDescription)
	[IO.Path]::GetinvalidFileNameChars() | ForEach-Object { $OutTemplateDisplayName = $OutTemplateDisplayName.Replace($_, "-") }
	Send-LogMessage -Level '' -message "$OutTemplateDisplayName.apm-template file archived to the templates directory"
	$CreatedTemplateFile = "$ScriptFolder\templates\$OutTemplateDisplayName-$Server.apm-template"
	Write-Progress -Activity "$Server" -Status "Creating Template XML File" -PercentComplete ($PortCount/$ports.Count) -id 1
	
	If (Test-Path $CreatedTemplateFile) { Remove-Item $CreatedTemplateFile }
	If($ItemsFound) { New-Item -Path $CreatedTemplateFile -Value $TemplateOutputXML }
	If ($ImportTemplateToOrion -and $ItemsFound) {
		Send-LogMessage -Level '' -message "$OutTemplateDisplayName will be imported to Orion"
		$ImportResult = Invoke-SwisVerb -SwisConnection ($SwisConnection) Orion.APM.ApplicationTemplate ImportTemplate @($TemplateOutputXML)
		$ApplicationTemplateId = $ImportResult.InnerText
		
		Write-host "ApplicationTemplateID: $applicationTemplateId"

		[int]$NodeID = $ServerItem.NodeID
		Write-host "NodeID: $nodeID"

		If ($AssignTemplateToNode -and ($nodeID -gt 0)) {
			Send-LogMessage -Level '' -message 'Assigning Application Monitor with inherited credentials'
			$CredentialSetId = -3
			
			$ApplicationId = (Invoke-SwisVerb -SwisConnection ($SwisConnection) "Orion.APM.Application" "CreateApplication" @($NodeID, $applicationTemplateId, $CredentialSetId, "false")).InnerText
			Write-host "Assignment output: $ApplicationId"
			If ($ApplicationId -eq -1) {
				Send-LogMessage -Level '' -message "Application wasn't created. Likely the template is already assigned to node and the skipping of duplications are set to 'true'."
				continue
			} Else {
				Send-LogMessage -Level '' -message "Application created with ID '$ApplicationId'."
			}
		} ElseIf (!($nodeID -gt 0)) {
			Send-LogMessage -Level 'Warning' -message "NodeID Not set, template will not be assigned to Server"
		}
	}
	$SytleHeaderFile = $XMLSourceFolder + '\styleheader.txt'
	$ImportedStyleHeader = Get-Content -path $SytleHeaderFile -Raw
	$HTMLReport = ConvertTo-HTML -Body "<h1>$Server</h1> $ReportServicesHTML $ReportPortsHTML " -Head $ImportedStyleHeader -Title "$Server Report of Found Application Monitors"
	$HTMLReportPath = "$ScriptFolder\reports\$OutTemplateDisplayName-$Server.html"
	If (Test-Path $HTMLReportPath) { Remove-Item $HTMLReportPath }
	$HTMLReport | Out-File $HTMLReportPath
	Send-LogMessage -Level '' -message "Starting Next Server" 
	$ProgressServerCount = $ProgressServerCount + 1
}
Send-LogMessage -Level '' -message "Script Completed"
Write-Progress -Activity "All Servers completed" -Status "Checking Common Ports" -PercentComplete 100 -Completed