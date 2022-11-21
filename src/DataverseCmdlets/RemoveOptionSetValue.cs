using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Management.Automation;

namespace DataverseCmdlets
{
    [Cmdlet(VerbsCommon.Remove, "OptionSetValue")]
    public class RemoveOptionSetValue : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = false)]
        public string EntityName{ get; set; }

        [Parameter(Mandatory = true)]
        public string OptionSet { get; set; }

        [Parameter(Mandatory = true)]
        public string OptionName { get; set; }

        [Parameter(Mandatory = true)]
        public int OptionValue { get; set; }

        protected override void CustomLogic()
        {
            var dovr = new DeleteOptionValueRequest
            {
                Value = OptionValue,
            };

            // https://docs.microsoft.com/en-us/previous-versions/dynamicscrm-2016/developers-guide/gg327594(v=crm.8)#remarks
            // Use the OptionSetName property when working with global option sets.
            // For local option sets in a picklist attribute, use the EntityLogicalName and AttributeLogicalName properties.

            if (!string.IsNullOrWhiteSpace(EntityName))
            {
                dovr.EntityLogicalName = EntityName;
                dovr.AttributeLogicalName = OptionSet;

                WriteObject($"Removing {OptionName} (value {OptionValue}) from optionset {OptionSet} on entity {EntityName}");
            }
            else
            {
                dovr.OptionSetName = OptionSet;
                WriteObject($"Removing {OptionName} (value {OptionValue}) from optionset {OptionSet}");
            }

            try
            {
                var dovresp = (DeleteOptionValueResponse)service.Execute(dovr);
                WriteObject("- removed");
            }
            catch (Exception ex)
            {
                WriteObject("- optionset value doesn't exist - probably already removed?");
            }
            //catch(Exception ex)
            //{
            //    WriteError(new ErrorRecord(ex, "TRYING TO DELETE OPTIONSET VALUE", default(ErrorCategory), null));
            //}
        }
    }
}












