param ([Parameter(Mandatory)]$settingsXmlFile)

$ErrorActionPreference  = "Stop";

# Read the configuration file.
$settingsXmlFileResolved = Resolve-Path $settingsXmlFile
Write-Host "Checking file paths in configuration file `"$settingsXmlFileResolved`""
[xml]$settingsXml = Get-Content -Path $settingsXmlFileResolved

$solutionName = 		$settingsXml.settings.solution.solutionName
$env =					$settingsXml.settings.environment.orgName
$connectionString = 	$settingsXml.settings.environment.connectionString
$connectionOwner = 		$settingsXml.settings.workflowSettings.cloudFlowSettings.connectionOwner.domainname
$connectionReference = 	$settingsXml.settings.workflowSettings.cloudFlowSettings.connectionReferences.connectionReference

# Get the list of connection references from the configuration file.
$connectionReferencesToUse = @();
foreach ($cr in $connectionReference)
{
	$item = @{ForConnectionType=$cr.ForConnectionType;ConnectionReferenceName=$cr.ConnectionReferenceName};
    $connectionReferencesToUse += $item
}
$cntFound = $connectionReferencesToUse.count

if ($cntFound -gt 0)
{	
	# Load the cmdlets.
	$CmdLetsDllFile = Resolve-Path "$PSScriptRoot\DataverseCmdlets.dll"
	Write-Host "Importing cmdlets from `"$CmdLetsDllFile`""
	Import-Module "$CmdLetsDllFile" # -Verbose
	
	$connectionReferencesToUseStr = $connectionReferencesToUse | Out-String
	Write-Host "Fixing connection references inside solution $($solutionName) from $env"

	Write-Host "Update-SolutionConnectionReferenceUsage `
		-ConnectionString          $connectionString `
		-SolutionName              $solutionName `
		-ConnectionOwner           $connectionOwner `
		-Verbose `
		-ConnectionReferencesToUse $connectionReferencesToUseStr"
		
	Update-SolutionConnectionReferenceUsage `
		-ConnectionString          "$connectionString" `
		-SolutionName              "$solutionName" `
		-ConnectionOwner           "$connectionOwner" `
		-ConnectionReferencesToUse $connectionReferencesToUse `
		-Verbose
}
else 
{
	Write-Host "Skip fixing connection references inside solution $($solutionName) from $env" -ForegroundColor DarkYellow
}