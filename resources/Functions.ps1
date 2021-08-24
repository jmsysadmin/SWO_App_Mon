Function Get-ScriptDirectory {
	[OutputType([string])]
	Param ()
	If ($null -ne $hostinvocation) {
		Split-Path $hostinvocation.MyCommand.path
	} Else {
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

Function Get-ApplicationTemplateXMLStart ([System.String]$inputfile) {
	$ApplicationTemplateXMLStart = Get-Content -path $inputfile -Raw
	Return $ApplicationTemplateXMLStart
}

Function Get-ApplicationTemplateXMLEnd ([System.String]$inputfile, $TemplateDisplayName, $TemplateDescription ) {
	$ApplicationTemplateXMLEnd = Get-Content -path $inputfile -Raw
	$TemplateTimeStamp = get-date -Format u
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@TEMPLATENAME@@@", $TemplateDisplayName
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@DESCRIPTION@@@", $TemplateDescription
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@CREATEDDATE@@@", $TemplateTimeStamp
	$ApplicationTemplateXMLEnd = $ApplicationTemplateXMLEnd -replace "@@@MODEIFIEDDATE@@@", $TemplateTimeStamp
	Return $ApplicationTemplateXMLEnd
}

Function Get-ServiceComponentXML ([System.String]$inputfile, [System.String]$ComponentName, [System.String]$ServiceName, [System.String]$ComponentOrder) {
	$ComponentXML = Get-Content -path $inputfile -Raw
	$ComponentXML = $ComponentXML -replace "@@@ComponentOrder@@@" , $ComponentOrder
	$ComponentXML = $ComponentXML -replace "@@@WindowsServiceMonitor@@@" , $ComponentName
	$ComponentXML = $ComponentXML -replace "@@@ServiceName@@@" , $ServiceName
	Return $ComponentXML
}

function Get-PortBasedComponentXML ($InputFolder, $PortNumber, $ComponentOrder) {
	If ($SSLCertComponentAdded) { $ComponentOrder+=1}

	switch ($PortNumber) {
		80 {
			$InputFile = $InputFolder + '\http.txt'
			$ComponentXML = (Get-Content -path $inputfile -Raw) -replace "@@@ComponentOrder@@@", $ComponentOrder
		}
		443 {
			$HttpsInputFile = $InputFolder + '\https.txt'
			$SSLCertInputFile = $InputFolder + '\SSLCert.txt'
			$HttpsXML = (Get-Content -path $HttpsInputFile -Raw) -replace "@@@ComponentOrder@@@", $ComponentOrder
			$SSLCertComponentAdded = $True
			$SSLCertXML = (Get-Content -path $SSLCertInputFile -Raw) -replace "@@@ComponentOrder@@@", ($ComponentOrder+1)
		}
		default {
			$InputFile = $InputFolder + '\Port.txt'
			$PortXML = (Get-Content -path $inputfile -Raw) -replace "@@@ComponentOrder@@@", $ComponentOrder
			$PortXML = (Get-Content -path $inputfile -Raw) -replace "@@@Port@@@", $PortNumber
		}
	}
}

$global:SSLCertComponentAdded = $false


