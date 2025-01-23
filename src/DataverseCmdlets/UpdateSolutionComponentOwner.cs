using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets;

[Cmdlet(VerbsData.Update, "SolutionComponentOwner")]
public class UpdateSolutionComponentOwner : CmdletWithCrmServiceClientBase
{
    [Parameter(Mandatory = true)]
    public string SolutionName { get; set; }

    [Parameter(Mandatory = true)]
    public Guid ApplicationUserAppId { get; set; }

	[Parameter(Mandatory = true)]	
	public ComponentType ComponentTypeToUpdate { get; set; }

    protected override void CustomLogic()
    {
        var qSystemUser = new QueryExpression("systemuser");
        qSystemUser.ColumnSet.AddColumn("fullname");
        qSystemUser.Criteria.AddCondition("applicationid", ConditionOperator.Equal, ApplicationUserAppId);

        var systemusers = service.RetrieveMultiple(qSystemUser).Entities;

        var systemuser = systemusers.FirstOrDefault();
        var name = systemuser["fullname"];

        WriteObject($"Checking to change owner of {ComponentTypeToUpdate} to {name}:");



        var query = new QueryExpression("workflow");
        query.ColumnSet.AddColumns("ownerid", "name", "statecode");
        var query_solutioncomponent = query.AddLink("solutioncomponent", "workflowid", "objectid");
        var query_solutioncomponent_solution = query_solutioncomponent.AddLink("solution", "solutionid", "solutionid");
        query_solutioncomponent_solution.LinkCriteria.AddCondition("uniquename", ConditionOperator.Equal, SolutionName);


		if(ComponentTypeToUpdate == ComponentType.ModernFlow)
		{
            query.Criteria.AddCondition("clientdata", ConditionOperator.NotNull);
            query.Criteria.AddCondition("category", ConditionOperator.Equal, 5);
        }

        if (ComponentTypeToUpdate == ComponentType.Workflow)
        {
            query.Criteria.AddCondition("category", ConditionOperator.Equal, 0);
        }

        var cloudFlows = service.RetrieveMultiple(query)
            .Entities
            .ToList();

		foreach(var cf in cloudFlows)
		{
			var ownerid = (EntityReference)cf["ownerid"];
            WriteObject($"{cf["name"]}:");

            if (ownerid.Id == systemuser.Id)
            {
                WriteObject($" - already reassigned");
                continue;
            }

            bool reactivate = false;
            if (ComponentTypeToUpdate == ComponentType.Workflow && (((OptionSetValue)cf["statecode"]).Value == 1))  //active
            {
                WriteObject($" - deactivating");
                var deactivate = new Entity(cf.LogicalName, cf.Id);
                deactivate["statecode"] = new OptionSetValue(0);
                service.Update(deactivate);
                reactivate = true;
            }

			var update = new Entity(cf.LogicalName, cf.Id);
			update["ownerid"] = new EntityReference("systemuser", systemuser.Id);

            service.Update(update);
            WriteObject($" - re-assigned");

            if (reactivate)
            {
                WriteObject($" - activating");
                var activate = new Entity(cf.LogicalName, cf.Id);
                activate["statecode"] = new OptionSetValue(1);
                service.Update(activate);
            }

            WriteObject("");
		}
    }
}
