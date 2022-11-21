using System;
using System.Management.Automation;
using System.Net;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace DataverseCmdlets
{
	/// <summary>
	/// Base class for Cmdlets that want to connect to a Dataverse environment
	/// </summary>
	public abstract class CmdletWithCrmServiceClientBase : Cmdlet
    {
		[Parameter(Mandatory = true)]
		public string ConnectionString { get; set; }

		protected ServiceClient service = null;

		protected override void ProcessRecord()
        {
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

			try
			{
				service = new ServiceClient(ConnectionString);
			}
			catch (Exception ex)
			{
				WriteError(new ErrorRecord(ex, "PLEASE CHECK THE CONNECTIONSTRING", ErrorCategory.ConnectionError, null));
				return;
			}

			if (!service.IsReady)
			{
				var err = new ErrorRecord(service.LastException, service.LastError, ErrorCategory.ResourceUnavailable, null);
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
