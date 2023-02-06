using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;

namespace DataverseCmdlets
{
    /// <summary>
    /// Publish Duplicate Detection Rules in a solution
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "DuplicateDetectionRulesPublished")]
    public class SetDuplicateDetectionRulesPublished : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        private EntityCollection GetDuplicateDetectionRules(IOrganizationService service)
        {
            var fetchDupeRulesInSolution = new FetchExpression($@"<fetch>
  <entity name=""duplicaterule"">
    <attribute name=""duplicateruleid"" />
    <attribute name=""name"" />
    <link-entity name=""solutioncomponent"" from=""objectid"" to=""duplicateruleid"">
      <link-entity name=""solution"" from=""solutionid"" to=""solutionid"">
        <filter>
          <condition attribute=""uniquename"" operator=""eq"" value=""{SolutionName}"" />
        </filter>
      </link-entity>
    </link-entity>
  </entity>
</fetch>");

            return service.RetrieveMultiple(fetchDupeRulesInSolution);
        }

        protected override void CustomLogic()
        {
			try
			{
                EntityCollection rules = GetDuplicateDetectionRules(service);
                WriteVerbose("Obtained " + rules.Entities.Count.ToString() + " duplicate detection rules.");

                if (rules.Entities.Count >= 1)
                {
                    // Create an ExecuteMultipleRequest object.
                    ExecuteMultipleRequest request = new ExecuteMultipleRequest()
                    {
                        // Assign settings that define execution behavior: don't continue on error, don't return responses. 
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        // Create an empty organization request collection.
                        Requests = new OrganizationRequestCollection()
                    };

                    //Create a collection of PublishDuplicateRuleRequests, and execute them in one batch
                    foreach (Entity entity in rules.Entities)
                    {
                        PublishDuplicateRuleRequest publishReq = new PublishDuplicateRuleRequest { DuplicateRuleId = entity.Id };
                        request.Requests.Add(publishReq);
                    }

                    service.Execute(request);
                }
                else
                {
                    WriteVerbose("Execution cancelled, as there are no duplicate detection rules to publish.");
                    return;
                }
            }
			catch (Exception ex)
			{
				WriteWarning($"ERROR while publishing duplicate detection rules");
				WriteWarning(ex.Message);
				WriteWarning("--------------------------------------------------------------------");
			}
		}
    }
}
