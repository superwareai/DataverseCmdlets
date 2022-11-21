using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Management.Automation;
using System.Runtime.Serialization;
using System.Xml;

namespace DataverseCmdlets
{
    [Cmdlet(VerbsData.Export, "DocumentTemplates")]
    public class ExportDocumentTemplates : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true, HelpMessage ="The folder to export the document templates to")]
        public string ExportFolder { get; set; }

        [Parameter(Mandatory = true, HelpMessage = "The fetchxml for finding/filtering the document templates to export")]
        public string FetchXmlFilter { get; set; }

        protected override void CustomLogic()
        {
            var settings = new XmlWriterSettings()
            {
                Indent = true,
                IndentChars = "\t"
            };
            DataContractSerializer dc = new DataContractSerializer(typeof(Entity));

            var fetchXml = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
  <entity name=""documenttemplate"">
    <all-attributes />
    <order attribute=""documenttype"" descending=""false"" />
    <order attribute=""name"" descending=""false"" />
    {FetchXmlFilter}
  </entity>
</fetch>";

            var q = new FetchExpression(fetchXml);
            var docs = service.RetrieveMultiple(q).Entities;

            foreach (var doc in docs)
            {
                var filename = $"{doc.Id} - {doc["name"]}";

                WriteObject($"Exporting document template {filename}");
                var associatedentitytypecode = (string)doc["associatedentitytypecode"];
                var entityTypeCode = GetEntityTypeCode(service, associatedentitytypecode);
                doc["associatedentitytypecode"] = $"{associatedentitytypecode}/{entityTypeCode}";

                using (var writer = XmlWriter.Create($@"{ExportFolder}\{filename}.xml", settings))
                {
                    var documentbody = (string)doc["content"];
                    byte[] bytes = Convert.FromBase64String(documentbody);
                    System.IO.File.WriteAllBytes($@"{ExportFolder}\{filename}.docx", bytes);
                   
                    doc.Attributes.Remove("content");
                    doc.Attributes.Remove("createdby");
                    doc.Attributes.Remove("createdon");
                    doc.Attributes.Remove("modifiedby");
                    doc.Attributes.Remove("modifiedon");
                    doc.Attributes.Remove("organizationid");
                    doc.FormattedValues.Clear();
                    dc.WriteObject(writer, doc);
                }
            }
        }

        public static int? GetEntityTypeCode(IOrganizationService service, string entity)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();

            request.LogicalName = entity;
            request.EntityFilters = EntityFilters.Entity;

            RetrieveEntityResponse response = (RetrieveEntityResponse)service.Execute(request);
            EntityMetadata metadata = response.EntityMetadata;

            return metadata.ObjectTypeCode;
        }
    }
}
