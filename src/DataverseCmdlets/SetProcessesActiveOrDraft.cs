using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;
using System.ServiceModel;

namespace DataverseCmdlets
{
    /// <summary>
    /// Turn a process (worklow) on (Active) or off (Draft)
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "ProcessesActiveOrDraft")]
    public class SetProcessesActiveOrDraft : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true, HelpMessage = "Guid of the process.  Corresponds to the workflowid attribute of the workflow table")]
        public Guid ProcessId { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "Name of the process/workflow/cloud flow.  Doesn't have to match")]
        public string ProcessName { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "Desired state")]
        public ProcessStatusEnum ProcessStatus { get; set; }

        [Parameter(Mandatory = false, HelpMessage = "Type of process e.g. Action, Workflow, Modern Flow, Business Process Flow, Business Rule")]
        public string ProcessType { get; set; } = "workflow";

        protected override void CustomLogic()
        {
			var targetStateCode = ProcessStatus == ProcessStatusEnum.Draft ? 0 : 1;
			var turn = targetStateCode == 0 ? "OFF" : "ON";

			try
			{
				var existing = service.Retrieve("workflow", ProcessId, new ColumnSet("statecode", "name"));
				var existingName = (string)existing["name"];
				var existingStateCode = ((OptionSetValue)existing["statecode"]).Value;

				if (targetStateCode != existingStateCode)
				{
					WriteVerbose($"TURN {turn} {existingName}");
					var updated = new Entity("workflow", ProcessId);

					if (turn == "OFF")
					{
						updated["statecode"] = new OptionSetValue(0);
						updated["statuscode"] = new OptionSetValue(1);
					}
					else if (turn == "ON")
					{
						updated["statecode"] = new OptionSetValue(1);
						updated["statuscode"] = new OptionSetValue(2);
					}

					service.Update(updated);
				}
				else
				{
					WriteVerbose($"ALREADY {turn} {existingName}");
				}

			}
			catch (Exception ex)
			{
				WriteWarning($"ERROR TURNING {turn} {ProcessType.ToUpper()} {ProcessName}");
				WriteWarning(ex.Message);
				WriteWarning("--------------------------------------------------------------------");
			}
		}

        public enum ProcessStatusEnum
        {
            Active,
            Draft
        }

    }
}
