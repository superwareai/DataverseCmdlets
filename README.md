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

### Disable-ProcessesInSolution
Disable all workflows (category=0) in a solution, except for those flows that contain in their name text provided in the exclusion pattern parameter

Parameters
```
-SolutionName
-ExclusionPattern 
```

### Enable-CloudFlowsInSolution
Enables all cloud flows in a solution, except for those where the text in the exclusion pattern is found anywhere in the name of that flow.
The ConnectionOwner parameter is supplied, the cmdlet will impersonate that user when activating the cloud flows (see https://www.develop1.net/public/post/2021/04/01/connection-references-with-alm-mind-the-gap)

Parameters
```
-SolutionName
-ExclusionPattern 
-ConnectionOwner  (this is the "domainname" of "systemuser")
```

### Enable-SLAsInSolution
Enables all SLAs that a part of the solution

Parameters
```
-SolutionName
```

### Export-DocumentTemplates
This cmdlet does what the MscrmTools.DocumentTemplatesMover xrm toolbox plugin does.

Parameters
```
-ExportFolder (where to export and unpack the document template xml/metadata)
-FetchXmlFilter (fetchxml conditions to limit which document templates are exported)
```

FetchXmlFilter is used in the fetchxml below, so supply an appropriate filter condition 

```
<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
  <entity name=""documenttemplate"">
    <all-attributes />
    <order attribute=""documenttype"" descending=""false"" />
    <order attribute=""name"" descending=""false"" />
    {FetchXmlFilter}
  </entity>
</fetch>
```

### Import-DocumentTemplates
This cmdlet does what the MscrmTools.DocumentTemplatesMover xrm toolbox plugin does.  Give a folder that contains the document template as created by the Export-DocumentTemplates, imports them to the dataverse environment (as determined by the connection string).

Parameters
```
-Folder (where to document template xml/metadata exist)
```


### Initialize-Queue
Create a queue, given an id, name, type (public/private), and initial email address, in a Dataverse environment

Parameters
```
-Id (guid of the queue)
-Name
-Type [Public|Private]
-InitialEmailAddress 
```

### Remove-OptionSetValue
When you want to remove a value from an optionset (choice) that comes from a manged solution (such as a Dynamics base/core solution).  You can't just customise the choice in your solution, as solution merge behaviour will never remove that choice.

Parameters
```
-OptionSet (name of option set)
-OptionName 
-OptionValue 
-EntityName (name of the table, if the optionset is not global, otherwise leave out) 
```

### Set-DuplicateDetectionRulesPublished
Publish the duplicate detection rules that are in a solution

Parameters
```
-SolutionName
```

### Set-OrgSettings
Please use power apps CLI instead.

Changes some org settings
Parameters
```
-AuditLog [Disable|Enable]
-PluginTraceLog [Disable|Enable]
-DataverseSearch [Disable|Enable]
-Locale [en_au] only at this stage
-RecalculateSLA
-InactivityTimeoutEnabled
-InactivityTimeoutInMins
-InactivityTimeoutReminderInMins
```

### Set-ProcessesActiveOrDraft
Turn a process (worklow) on (Active) or off (Draft)

Parameters
```
-ProcessId
-ProcessName
-ProcessStatus [Active|Draft]
```

### Set-UserSettings
Sets the default dashboard for all that have the supplied security role.  Only does this for users without a default dashboard already/previously set.

Parameters
```
-SecurityRoleId
-DefaultDashboardId
-TimeZone
-Locale
```

### Update-DefaultBusinessUnit
Changes the name of the default Business unit in the give Dataverse environment, and assigns the security role to the team of the default BU

Parameters
```
-BusinessUnitName
-SecurityRoleName
```

### Update-SolutionComponentOwner
Changes the owner of the nominated components (only workflows and cloud flows at this stage) in a given solution.

The new owner must be an application user

Parameters
```
-SolutionName
-ApplicationUserAppId (the client if of the entra application registration that was created into an application user)
-ComponentTypeToUpdate [WorkFlow|ModernFlow]
```

### Update-SolutionConnectionReferenceUsage
In a given solution, change each connection type to use the nominated connection reference.

This is used when there are many developers in a single environment using and because the connection is not/cannot be shared, a new connection reference gets created when they modify someone else cloud flow. 

This cmdlet is usually used before exporting a solution for source control.  It will go through the cloud flows of a solution and find all connection references and change the flow to only use the one connection refernce.

Not recommended if you require a cloud flow to use multiple connection reference of the same type within a cloud flow (.e.g if for some reason you want a different connections with different credentials)

Parameters
```
-SolutionName
-ConnectionReferencesToUse
-ConnectionOwner (If supplied, then this script will impersonate the systemuser when connecting to Dataverse. See https://www.develop1.net/public/post/2021/04/01/connection-references-with-alm-mind-the-gap)
```

ConnectionReferencesToUse parameter is an array of 

`@{ForConnectionType="";ConnectionReferenceName=""};`

ForConnectionType is the type of the connection (e.g. shared_commondataserviceforapps) and ConnectionReferenceName is the object name (not display name) of the connection reference you want to keep.


## Azure DevOps 

If you are using the [Power Platform Build Tools](https://github.com/microsoft/powerplatform-build-tools) in Azure DevOps, you don't have to maintain Dataverse credentials separately in order to construct a connection string to pass to the cmdlets. Instead, if you add a "Power Platform Set Connection Variables" task you can get access to the connection string in a later Powershell task like this:

`$(connectionVariables.BuildTools.DataverseConnectionString)` 

where the "connectionVariables" prefix is the "Reference Name" you enter in the "Power Platform Set Connection Variables" tasks "Output Variables" section.


