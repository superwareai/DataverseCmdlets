using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets
{
    /// <summary>
    /// Changes the name of the default Business unit in the give Dataverse environment, and assigns the security role to the team of the default BU
    /// </summary>
    [Cmdlet(VerbsData.Update, "DefaultBusinessUnit")]
    public class UpdateDefaultBusinessUnit : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true)]
        public string BusinessUnitName { get; set; }

        [Parameter(Mandatory = true)]
        public string SecurityRoleName { get; set; }

        protected override void CustomLogic()
        {
            //get the default Business Unit;
            var qBu = new QueryExpression("businessunit");
            qBu.ColumnSet.AllColumns = true;
            qBu.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Null);

            var businessUnit = service.RetrieveMultiple(qBu).Entities.SingleOrDefault();

            if ((string)businessUnit["name"] == BusinessUnitName)
            {
                WriteObject($"Default business unit name is already {BusinessUnitName}");
            }
            else
            {
                var update = new Entity(businessUnit.LogicalName, businessUnit.Id);
                update["name"] = "IPU Team";
                service.Update(update);
                WriteObject($"Default business unit name updated to {BusinessUnitName}");
            }

            //Get the default team for the parent business unit. The name and membership for default team are inherited from their parent business unit.
            var qT = new QueryExpression("team");
            qT.ColumnSet.AllColumns = true;
            qT.Criteria.AddCondition("isdefault", ConditionOperator.Equal, true);
            qT.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessUnit.Id);

            var team = service.RetrieveMultiple(qT).Entities.SingleOrDefault();

            //get the security role 

            var qR = new QueryExpression("role");
            qR.ColumnSet.AllColumns = true;
            qR.Criteria.AddCondition("name", ConditionOperator.Equal, SecurityRoleName);

            var securityRole = service.RetrieveMultiple(qR).Entities.SingleOrDefault();
            try
            {
                //associate it to the team
                service.Associate(
                           team.LogicalName,
                           team.Id,
                           new Relationship("teamroles_association"),
                           new EntityReferenceCollection() { new EntityReference(securityRole.LogicalName, securityRole.Id) });

                WriteObject($"Security role {SecurityRoleName} associated to team {team["name"]}");
            }
            catch (Exception ex)
            {
                WriteObject(ex);
            }
        }
    }
}
