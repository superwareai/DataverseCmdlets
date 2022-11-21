using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Management.Automation;

namespace DataverseCmdlets
{
    /// <summary>
    /// Chagne some orgSettings in a Dataverse environment, like enabling plugin trace and audit log, locale
    /// </summary>
    [Cmdlet(VerbsCommon.Set, "OrgSettings")]
    public class SetOrgSettings : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = false)]
        public StatusEnum? AuditLog { get; set; }

        [Parameter(Mandatory = false)]
        public StatusEnum? PluginTraceLog { get; set; }

        [Parameter(Mandatory = false)]
        public StatusEnum? DataverseSearch { get; set; }

        [Parameter(Mandatory = false)]
        public LocaleEnum? Locale { get; set; }

        [Parameter(Mandatory = false)]
        public bool? RecalculateSLA { get; set; }

        [Parameter(Mandatory = false)]
        public bool? InactivityTimeoutEnabled { get; set; }

        [Parameter(Mandatory = false)]
        public int InactivityTimeoutInMins { get; set; }

        [Parameter(Mandatory = false)]
        public int InactivityTimeoutReminderInMins { get; set; }

        protected override void EndProcessing()
        {
            service?.Dispose();
            base.EndProcessing();
        }

        protected override void CustomLogic()
        {
            WriteObject("STARTING");

            if (!AuditLog.HasValue && !PluginTraceLog.HasValue && !Locale.HasValue && !DataverseSearch.HasValue && !RecalculateSLA.HasValue)
            {
                WriteObject("NOTHING TO CHANGE");
                return;
            }

            var qGetOrg = new QueryExpression("organization");
            qGetOrg.ColumnSet.AllColumns = true;
            var existingOrgSettings = service.RetrieveMultiple(qGetOrg).Entities.SingleOrDefault();

            var org = new Entity("organization", existingOrgSettings.Id);

            if (AuditLog != null)
            {
                WriteObject("SETTING AuditLog");
                var val = AuditLog.Value == StatusEnum.Enable;
                if((bool)existingOrgSettings["isauditenabled"] != val)
                {
                    org["isauditenabled"] = val;
                    // "isuseraccessauditenabled": true,
                    // "isreadauditenabled": true
                }
            }

            if (DataverseSearch != null)
            {
                WriteObject("SETTING DataverseSearch");
                var val = DataverseSearch.Value == StatusEnum.Enable;

                if (!existingOrgSettings.Contains("isexternalsearchindexenabled") || (bool)existingOrgSettings["isexternalsearchindexenabled"] != val || !existingOrgSettings.Contains("newsearchexperienceenabled") || (bool)existingOrgSettings["newsearchexperienceenabled"] != val)
                {
                    org["isexternalsearchindexenabled"] = val;
                    org["newsearchexperienceenabled"] = val;
                    //org["relevancesearchmodifiedon"] = DateTime.UtcNow;
                }
            }

            if (Locale != null)
            {
                WriteObject("SETTING Locale");
                org["localeid"] = (int)Locale;

                if (Locale == LocaleEnum.en_au)
                {
                    org["negativecurrencyformatcode"] = 1;
                    org["dateformatstring"] = "d/MM/yyyy";
                    org["longdateformatcode"] = 1;
                    org["dateformatcode"] = 3;
                }
            }

            if (PluginTraceLog != null)
            {
                WriteObject("SETTING PluginTraceLog");
                org["plugintracelogsetting"] = PluginTraceLog == StatusEnum.Enable ? new OptionSetValue(2) : new OptionSetValue(0);
            }

            if (RecalculateSLA != null)
            {
                WriteObject("SETTING RecalculateSLA");
                if ((bool)existingOrgSettings["recalculatesla"] != RecalculateSLA.Value)
                {
                    org["recalculatesla"] = RecalculateSLA.Value;
                }
            }

            org["inactivitytimeoutinmins"] = InactivityTimeoutInMins;
            org["inactivitytimeoutreminderinmins"] = InactivityTimeoutReminderInMins;
            org["inactivitytimeoutenabled"] = InactivityTimeoutEnabled;

            WriteObject("The following organisation entity attributes will be changed:");
            WriteObject(org.Attributes);
            service.Update(org);
            WriteObject("DONE");
        }
    }

    
}
