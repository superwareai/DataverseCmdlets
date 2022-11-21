using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;

namespace DataverseCmdlets
{
	/// <summary>
	/// Sets the default dashboard for each user based on their security role in a Dataverse environment.  Only does this for users without a default dashboard already/previously set.
	/// </summary>
	[Cmdlet(VerbsCommon.Set, "UserSettings")]
    public class SetUserSettings : CmdletWithCrmServiceClientBase
    {
		[Parameter(Mandatory = true)]
		public Guid SecurityRoleId { get; set; }

		[Parameter(Mandatory = false)]
        public Guid? DefaultDashboardId { get; set; }

		[Parameter(Mandatory = false)]
		public TimeZoneEnum? TimeZone{ get; set; }

		[Parameter(Mandatory = false)]
		public LocaleEnum? Locale { get; set; }

		protected override void CustomLogic()
        {
			var role = service.Retrieve("role", SecurityRoleId, new ColumnSet("name"));
			Entity dashboard = null;
            if (DefaultDashboardId.HasValue)
            {
				dashboard= service.Retrieve("systemform", DefaultDashboardId.Value, new ColumnSet("name"));
			}
			
			var qUsersWithRole = new QueryExpression("systemuserroles");
			qUsersWithRole.ColumnSet.AddColumn("systemuserid");
			qUsersWithRole.Criteria.AddCondition("roleid", ConditionOperator.Equal, SecurityRoleId);
			var users = service.RetrieveMultiple(qUsersWithRole).Entities;

			if (dashboard != null)
			{
				WriteObject($"All users with role {role["name"]}, if they don't already have a default dashboard, will have their's set to {dashboard["name"]}");
			}

			if(TimeZone.HasValue)
            {
				WriteObject($"All users with role {role["name"]} will have their timezone changed to {TimeZone.Value}");
			}

			if (Locale.HasValue)
			{
				WriteObject($"All users with role {role["name"]} will have their Local changed to {Locale.Value}");
			}

			foreach (var u in users)
			{
				var user = service.Retrieve("systemuser", (Guid)u["systemuserid"], new ColumnSet("fullname"));
				var userSettings = service.Retrieve("usersettings", (Guid)u["systemuserid"], new ColumnSet(true));

				var userUpdate = new Entity("usersettings", userSettings.Id);

				if (dashboard != null)
				{
					if (!userSettings.Contains("defaultdashboardid") || userSettings["defaultdashboardid"] == null)
					{
						//updating user default dashboard
						WriteObject($"The user {user["fullname"]} will have their default dashboard set.");
						userUpdate["defaultdashboardid"] = DefaultDashboardId;
					}
					else
					{
						WriteObject($"The user {user["fullname"]} already has a default dashboard.");
					}
				}

                if (TimeZone.HasValue)
                {
					userUpdate["timezonecode"] = (int)TimeZone.Value;
				}

                if (Locale.HasValue)
                {
					userUpdate["localeid"] = (int)Locale.Value;
				}

				service.Update(userUpdate);
			}
		}
    }

}
