using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets
{
	/// <summary>
	/// Enable all cloud flows in the nominated solution in a Dataverse environment
	/// </summary>
	[Cmdlet(VerbsLifecycle.Enable, "CloudFlowsInSolution")]
	public class EnableCloudFlowsInSolution : CmdletWithCrmServiceClientBase
    {
		[Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

		[Parameter(Mandatory = false, HelpMessage ="If the name of the cloud flow contains this string, then it wil not be activated.")]
		public string ExclusionPattern { get; set; }

		[Parameter(Mandatory = false, HelpMessage = "If supplied, then this script will impersonate the systemuser when connecting to Dataverse. See https://www.develop1.net/public/post/2021/04/01/connection-references-with-alm-mind-the-gap")]
		public string ConnectionOwner { get; set; }

		protected override void EndProcessing()
		{
			service?.Dispose();
			base.EndProcessing();
		}

        protected override void CustomLogic()
        {
			Guid? connectionOwnerId = null;
			if (!string.IsNullOrWhiteSpace(ConnectionOwner))
			{
				var qSu = new QueryExpression("systemuser");
				qSu.Criteria.AddCondition("domainname", ConditionOperator.Equal, ConnectionOwner);
				connectionOwnerId = service.RetrieveMultiple(qSu).Entities.FirstOrDefault()?.Id;
				if (connectionOwnerId is null)
				{
					WriteWarning($"Systemuser {ConnectionOwner} could not be found");
				}
			}

			if (connectionOwnerId.HasValue)
			{
                service.CallerId = connectionOwnerId.Value; //why do this? because: https://www.develop1.net/public/post/2021/04/01/connection-references-with-alm-mind-the-gap
			}

			//Activate flows
			var fetchCloudFlow = new FetchExpression($@"<fetch>
			<entity name='workflow'>
			<all-attributes />
			<filter>
				<condition attribute='category' operator='eq' value='5' />  <!-- Modern Flow: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/workflow?view=dynamics-ce-odata-9 -->
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
				var flowName = (string)f["name"];
				var setStatecode = 1;
				var onOff = "ON";

				if(!string.IsNullOrWhiteSpace(ExclusionPattern) && flowName.Contains(ExclusionPattern.Trim()))
                {
					if (currentStateCode == 0)
					{
						WriteObject($"Cloud flow already turned off {flowName} {f.Id}");
						continue;
					}

					setStatecode = 0;
					onOff = "OFF";
				}

				if(currentStateCode == 1 && setStatecode == 1)
                {
					WriteObject($"Cloud flow already turned on {flowName} {f.Id}");
					continue;
				}

				try
				{
					WriteObject($"Turning {onOff} cloud flow {flowName} {f.Id}");
					var update = new Entity(f.LogicalName, f.Id);
					update["statecode"] = new OptionSetValue(setStatecode);
					service.Update(update);
				}
				catch (Exception ex)
				{
					WriteWarning($"ERROR TURNING {onOff} THE CLOUD FLOW {ex.Message}");
					//WriteError(new ErrorRecord(ex, $"ERROR TURNING {onOff} THE CLOUD FLOW", default(ErrorCategory), null));
				}
			}
		}
    }
}
