using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets
{
    /// <summary>
    /// In a solution in a Dataverse Solution, change each connection type to use the nominated connection reference
    /// </summary>
    [Cmdlet(VerbsData.Update, "SolutionConnectionReferenceUsage")]
	public class UpdateSolutionConnectionReferenceUsage : CmdletWithCrmServiceClientBase
    {
		[Parameter(Mandatory = true)]
		public string SolutionName { get; set; }

		[Parameter(Mandatory = true)]
		public ConnectionReferenceToUse[] ConnectionReferencesToUse { get; set; }

		[Parameter(Mandatory = false, HelpMessage = "If supplied, then this script will impersonate the systemuser when connecting to Dataverse. See https://www.develop1.net/public/post/2021/04/01/connection-references-with-alm-mind-the-gap")]
		public string ConnectionOwner { get; set; }


		protected override void CustomLogic()
        {
			var fetchWorkFlows = new FetchExpression($@"<fetch>
				<entity name='workflow' >
					<attribute name='name' />
					<attribute name='clientdata' />
					<attribute name='ownerid' />
					<filter>
						<condition attribute='clientdata' operator='not-null'/>
						<condition attribute='category' operator='eq' value='5'/>
					</filter>
					<link-entity name='solutioncomponent' from='objectid' to='workflowid' >
						<link-entity name='solution' from='solutionid' to='solutionid' >
						<filter>
							<condition attribute='uniquename' operator='eq' value='{SolutionName}' />
						</filter>
						</link-entity>
					</link-entity>
					</entity>
				</fetch>");

			var cloudFlows = service.RetrieveMultiple(fetchWorkFlows).Entities.OrderBy(c => (string)c["name"]);

			Guid? connectionOwnerId = null;
			if (!string.IsNullOrWhiteSpace(ConnectionOwner))
			{
				var qSu = new QueryExpression("systemuser");
				qSu.Criteria.AddCondition("domainname", ConditionOperator.Equal, ConnectionOwner);
				connectionOwnerId = service.RetrieveMultiple(qSu).Entities.FirstOrDefault()?.Id;
				if(connectionOwnerId is null)
                {
					WriteWarning($"Systemuser {ConnectionOwner} could not be found");
				}
			}


			foreach (var cloudFlow in cloudFlows)
			{
				WriteObject($"CHECKING CLOUD FLOW: {cloudFlow["name"]}");

				var clientdataJson = cloudFlow["clientdata"] as string;

				JObject clientdata = JsonConvert.DeserializeObject<JObject>(clientdataJson);
				JObject connectionReferences = clientdata["properties"]["connectionReferences"] as JObject;

				var tobeupdated = false;

				foreach (var prop in connectionReferences.Properties())
				{
					var key = prop.Name;
					var connectionReference = connectionReferences[key] as JObject;

					var apiName = (string)connectionReference["api"]["name"];
					var connectionName = (string)connectionReference["connection"]["nsame"];
					var connectionReferenceLogicalName = (string)connectionReference["connection"]["connectionReferenceLogicalName"];

					WriteObject($"- checking {apiName}");

					foreach (var crto in this.ConnectionReferencesToUse)
					{
						if (apiName == crto.ForConnectionType)
						{
							if (connectionName != null)
							{
								(connectionReference["connection"] as JObject).Property("name").Remove();
								WriteObject($"- removed connection {connectionName}");
								tobeupdated = true;
							}

							if (connectionReferenceLogicalName != crto.ConnectionReferenceName)
							{
								connectionReference["connection"]["connectionReferenceLogicalName"] = crto.ConnectionReferenceName;
								WriteObject($"- set connectionReferenceLogicalName to {crto.ConnectionReferenceName}");
								tobeupdated = true;
							}
						}
					}
				}

				WriteObject($"- cloud flow {(tobeupdated ? "will be updated" : "okay") }");
				WriteObject("");

				if (tobeupdated)
				{
					if (connectionOwnerId.HasValue)
					{
						service.CallerId = connectionOwnerId.Value;
					}

					var updatedClientData = JsonConvert.SerializeObject(clientdata, Formatting.Indented);
					
					var updatedCloudFlow = new Entity("workflow", cloudFlow.Id);
					updatedCloudFlow["clientdata"] = updatedClientData;
					
					try
					{
						service.Update(updatedCloudFlow);
                    }
                    catch (Exception ex)
                    {
						WriteError(new ErrorRecord(ex, "ERROR FIXING CONNECTION REFERENCE ON THE CLOUD FLOW", default(ErrorCategory), null));
						WriteWarning(connectionReferences.ToString());
					}

				}
			}

			WriteObject($"");
		}

		public class ConnectionReferenceToUse
		{
			public string ForConnectionType { get; set; }
			public string ConnectionReferenceName { get; set; }
		}
    }
}
