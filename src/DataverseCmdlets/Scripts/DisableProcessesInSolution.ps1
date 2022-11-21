param ([Parameter(Mandatory)]$settingsXmlFile)
$ErrorActionPreference  = "Stop";
[xml]$settingsXml = Get-Content -Path $settingsXmlFile

Import-Module "$PSScriptRoot\DataverseCmdlets.dll"  -Verbose

Disable-ProcessesInSolution  `
-ConnectionString   $settingsXml.settings.environment.connectionString `
-SolutionName 		$settingsXml.settings.solution.solutionName `
-ExclusionPattern 	$settingsXml.settings.workflowSettings.workFlowDeactivationRule.pattern `	
-Verbose