using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets
{
	/// <summary>
	/// Enable all SLAs that are part of the nominated solution in a Dataverse environment
	/// </summary>
	[Cmdlet(VerbsLifecycle.Enable, "SLAsInSolution")]
    public class EnableSLAsInSolution : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        protected override void CustomLogic()
        {
			var qS = new QueryExpression("solution");
			qS.Criteria.AddCondition("uniquename", ConditionOperator.Equal, SolutionName);
			var solution = service.RetrieveMultiple(qS).Entities.SingleOrDefault();

			if (solution == null)
			{
				var em = $"Solution {SolutionName} does not exist";
				var err = new ErrorRecord(new Exception(em), em, ErrorCategory.ResourceUnavailable, null);
				WriteError(err);
				return;
			}

			var qSc = new QueryExpression("solutioncomponent");
			qSc.Criteria.AddCondition("solutionid", ConditionOperator.Equal, solution.Id);
			qSc.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 152);
			qSc.ColumnSet.AllColumns = true;
			var slas = service.RetrieveMultiple(qSc).Entities;

			if (!slas.Any())
			{
				WriteVerbose($"No SLAs found in solution {SolutionName}");
				return;
			}

			foreach (var s in slas)
			{
				var slaId = (Guid)s["objectid"];
				var sla = new Entity("sla", slaId);
				sla["statecode"] = new OptionSetValue(1);
				service.Update(sla);
				WriteVerbose($"SLA {slaId} activated");
			}

			var defaultSlaId = (Guid)slas.First()["objectid"];
			var defaultSla = new Entity("sla", defaultSlaId);
			defaultSla["isdefault"] = true;
			service.Update(defaultSla);
			WriteVerbose($"SLA {defaultSlaId} made default");
		}
    }
}
