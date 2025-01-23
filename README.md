# DataverseCmdlets

This package contains a number of powershell cmdlets for working with Dataverse in our ALM process.

Note: a number of these can now be achieved with the Power Apps CLI (https://www.nuget.org/packages/Microsoft.PowerApps.CLI). For example the Set-OrgSettings cmdlet can now be done using `pac env update-settings`.

## importing the module into powershell

```
$CmdLetsDllFile = Resolve-Path "$PSScriptRoot\DataverseCmdlets.dll"
Write-Host "Importing cmdlets from `"$CmdLetsDllFile`""
Import-Module "$CmdLetsDllFile" # -Verbose
```

## the cmdlets

All cmdlets take the following mandatory parameters:

`-ConnectionString`

## Azure DevOps 



### Disable-ProcessesInSolution

### Enable-CloudFlowsInSolution

### Enable-SLAsInSolution

### Export-DocumentTemplates

### Import-DocumentTemplates

### Initialize-Queue

### Remove-OptionSetValue

### Set-DuplicateDetectionRulesPublished

### Set-OrgSettings

### Set-ProcessesActiveOrDraft

### Set-UserSettings

### Update-DefaultBusinessUnit

### Update-SolutionConnectionReferenceUsage