using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Management.Automation;
using System.Net;

namespace DataverseCmdlets
{
	/// <summary>
	/// Base class for Cmdlets that want to connect to a Dataverse environment
	/// </summary>
	public abstract class CmdletWithCrmServiceClientBase : Cmdlet
    {
		[Parameter(Mandatory = true)]
		public string ConnectionString { get; set; }

		protected CrmServiceClient service = null;

		protected override void ProcessRecord()
        {
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

			try
			{
				service = new CrmServiceClient(ConnectionString);
			}
			catch (Exception ex)
			{
				WriteError(new ErrorRecord(ex, "PLEASE CHECK THE CONNECTIONSTRING", ErrorCategory.ConnectionError, null));
				return;
			}

			if (!service.IsReady)
			{
				var err = new ErrorRecord(service.LastCrmException, service.LastCrmError, ErrorCategory.ResourceUnavailable, null);
				WriteError(err);
				return;
			}

			CustomLogic();

            base.ProcessRecord();
        }

		protected override void EndProcessing()
		{
			service?.Dispose();
			base.EndProcessing();
		}

		protected abstract void CustomLogic();
    }
}
