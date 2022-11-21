param ([Parameter(Mandatory)]$settingsXmlFile)

$ErrorActionPreference  = "Stop";

$settingsXmlFileResolved = Resolve-Path $settingsXmlFile
Write-Host "Reading configuration file `"$settingsXmlFileResolved`""
[xml]$settingsXml = Get-Content -Path "$settingsXmlFileResolved"

$ExportFolderPath =		$settingsXml.settings.documentTemplates.exportFolderPath
$FetchXml =				$settingsXml.settings.documentTemplates.fetchXmlFilter."#cdata-section"
$connectionString = 	$settingsXml.settings.environment.connectionString
$env =					$settingsXml.settings.environment.orgName

if ([string]::IsNullOrWhitespace($ExportFolderPath))
{
	Write-Host "Skip exporting document templates from $env" -ForegroundColor DarkYellow
}
else 
{
	$ExportFolderPath =	Resolve-Path $ExportFolderPath

	$CmdLetsDllFile = Resolve-Path "$PSScriptRoot\DataverseCmdlets.dll"
	Write-Host "Importing cmdlets from `"$CmdLetsDllFile`""
	Import-Module "$CmdLetsDllFile" #-Verbose
	
	Write-Host "Exporting document templates from $env"	
	Write-Host "Export-DocumentTemplates `
		-ConnectionString       `"$connectionString`" `
		-ExportFolder			`"$ExportFolderPath`" `
		-FetchXmlFilter         `"$FetchXml`" `
		-Verbose"

	Export-DocumentTemplates `
		-ConnectionString       "$connectionString" `
		-ExportFolder			"$ExportFolderPath" `
		-FetchXmlFilter         "$FetchXml" `
		-Verbose
}