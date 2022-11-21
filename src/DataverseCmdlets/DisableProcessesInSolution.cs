using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;

namespace DataverseCmdlets
{
    [Cmdlet(VerbsLifecycle.Disable, "ProcessesInSolution")]
    public class DisableProcessesInSolution : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "If the name of the process contains this string, then it wil not be activated.")]
        public string ExclusionPattern { get; set; }

        protected override void CustomLogic()
        {
			if (string.IsNullOrWhiteSpace(ExclusionPattern))
			{
				WriteObject($"ExclusionPattern cannot be blank or empty");
				return;
			}

			//Activate flows
			var fetchCloudFlow = new FetchExpression($@"<fetch>
			<entity name='workflow'>
			<all-attributes />
			<filter>
				<condition attribute='category' operator='eq' value='0' />  <!-- Modern Flow: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/workflow?view=dynamics-ce-odata-9 -->
			</filter>
			<link-entity name='solutioncomponent' from='objectid' to='workflowid'>
				<link-entity name='solution' from='solutionid' to='solutionid'>
				<filter>
					<condition attribute='uniquename' operator='eq' value='{SolutionName}' />
				</filter>
				</link-entity>
			</link-entity>
			</entity>
			</fetch>");

			var flows = service.RetrieveMultiple(fetchCloudFlow).Entities;

			foreach (var f in flows)
			{
				var currentStateCode = ((OptionSetValue)f["statecode"]).Value;//Draft = 0, Activated = 1
				var processName = (string)f["name"];

				if (!processName.Contains(ExclusionPattern.Trim()))
				{
					WriteObject($"Process will not be turned off {processName} {f.Id}");
					continue;
                }

				if (currentStateCode == 0)
				{
					WriteObject($"Process already turned off {processName} {f.Id}");
					continue;
				}
				
				try
				{
					WriteObject($"Turning off process {processName} {f.Id}");
					var update = new Entity(f.LogicalName, f.Id);
					update["statecode"] = new OptionSetValue(0);
					update["statuscode"] = new OptionSetValue(1);
					service.Update(update);
				}
				catch (Exception ex)
				{
					WriteWarning($"ERROR TURNING OFF THE PROCESS {ex.Message}");
					//WriteError(new ErrorRecord(ex, $"ERROR TURNING OFF THE PROCESS", default(ErrorCategory), null));
				}
			}
		}

        protected override void EndProcessing()
        {
            service?.Dispose();
            base.EndProcessing();
        }
    }
}
