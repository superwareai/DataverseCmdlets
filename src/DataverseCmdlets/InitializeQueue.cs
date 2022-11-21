using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;

namespace DataverseCmdlets
{
    /// <summary>
    /// Create a queue, given an id, name, type (public/private), and initial email address, in a Dataverse environment
    /// </summary>
    [Cmdlet(VerbsData.Initialize, "Queue")]
    public class InitializeQueue : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true)]
        public Guid Id { get; set; }

        [Parameter(Mandatory = true)]
        public string Name { get; set; }

        [Parameter(Mandatory = true)]
        public QueueTypeEnum Type { get; set; }

        [Parameter(Mandatory = true)]
        public string InitialEmailAddress { get; set; }

        protected override void CustomLogic()
        {
            WriteObject($"Making sure queue {Name} with id {Id} exists.");

            var queue = new Entity("queue", Id);
            queue["name"] = Name;
            queue["queueviewtype"] = new OptionSetValue((int)Type);

            var upsert = new UpsertRequest
            {
                Target = queue
            };

            service.Execute(upsert);

            var updatedQueue = service.Retrieve("queue", Id, new ColumnSet("emailaddress"));

            if(!updatedQueue.Contains("emailaddress"))
            {
                updatedQueue["emailaddress"] = InitialEmailAddress;
                service.Update(updatedQueue);
            }
        }
    }
    
    /// <summary>
    /// queueviewtype
    /// </summary>
    public enum QueueTypeEnum
    {
        Public = 0,
        Private = 1
        
    }
}
