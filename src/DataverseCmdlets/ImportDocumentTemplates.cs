using DocumentFormat.OpenXml.Packaging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DataverseCmdlets
{
    [Cmdlet(VerbsData.Import, "DocumentTemplates")]
    public class ImportDocumentTemplates : CmdletWithCrmServiceClientBase
    {
        [Parameter(Mandatory = true, HelpMessage = "The folder that contains the document templates that were exported using the Export-DocumentTemplates cmdlet")]
        public string Folder { get; set; }
        protected override void CustomLogic()
        {
            var settings = new XmlWriterSettings()
            {
                Indent = true,
                IndentChars = "\t"
            };

            DataContractSerializer dc = new DataContractSerializer(typeof(Entity));

			var di = new System.IO.DirectoryInfo(Folder);
			var xmlFiles = di.GetFiles("*.xml");

			foreach (var xmlFile in xmlFiles)
			{
				using (var stream = xmlFile.OpenRead())
				{
					var documentTemplateToUpload = (Entity)dc.ReadObject(stream);

					var file = xmlFile.FullName.Replace("xml", "docx");

					var bytes = System.IO.File.ReadAllBytes(file);
					var documentbody = Convert.ToBase64String(bytes);
					documentTemplateToUpload["content"] = documentbody;

                    Console.WriteLine($"Upserting {documentTemplateToUpload.Id} {documentTemplateToUpload["name"]}");

                    byte[] content = Convert.FromBase64String(documentTemplateToUpload.GetAttributeValue<string>("content"));
                    MemoryStream contentStream = new MemoryStream();
                    contentStream.Write(content, 0, content.Length);
                    contentStream.Position = 0;

                    var associatedentitytypecodeParts = ((string)documentTemplateToUpload["associatedentitytypecode"]).Split('/');

                    try
                    {
                        #region this code is adapted from https://github.com/MscrmTools/MscrmTools.DocumentTemplatesMover/blob/master/MsCrmTools.DocumentTemplatesMover/TemplatesManager.cs

                        var entityTypeCode = associatedentitytypecodeParts[0];
                        var oldEnityTypeCode = associatedentitytypecodeParts[1];
                        var newEnityTypeCode = ExportDocumentTemplates.GetEntityTypeCode(service, entityTypeCode);
                        documentTemplateToUpload["associatedentitytypecode"] = entityTypeCode;

                        string toFind = string.Format("{0}/{1}", entityTypeCode, oldEnityTypeCode);
                        string replaceWith = string.Format("{0}/{1}", entityTypeCode, newEnityTypeCode);

                        Console.WriteLine($"Replacing {toFind} with {replaceWith} in document");

                        using (var doc = WordprocessingDocument.Open(contentStream, true, new OpenSettings { AutoSave = true }))
                        {
                            // crm keeps the etc in multiple places; parts here are the actual merge fields
                            doc.MainDocumentPart.Document.InnerXml = doc.MainDocumentPart.Document.InnerXml.Replace(toFind, replaceWith);

                            // next is the actual namespace declaration
                            doc.MainDocumentPart.CustomXmlParts.ToList().ForEach(a =>
                            {
                                using (StreamReader reader = new StreamReader(a.GetStream()))
                                {
                                    var xml = XDocument.Load(reader);

                                    // crappy way to replace the xml, couldn't be bothered figuring out xml root attribute replacement...
                                    var crappy = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" + xml.ToString();

                                    if (crappy.IndexOf(toFind) > -1) // only replace what is needed
                                    {
                                        crappy = crappy.Replace(toFind, replaceWith);

                                        using (var stream2 = new MemoryStream(Encoding.UTF8.GetBytes(crappy)))
                                        {
                                            a.FeedData(stream2);
                                        }
                                    }
                                }
                            });
                        }
                        documentTemplateToUpload["content"] = Convert.ToBase64String(contentStream.ToArray());
                        #endregion

                    
						var upsert = new UpsertRequest
						{
							Target = documentTemplateToUpload
						};

						
						Console.WriteLine();
						service.Execute(upsert);
					}
					catch (Exception ex)
					{
						WriteError(new ErrorRecord(ex, "TRYING TO IMPORT DOCUMENT TEMPLATE", default(ErrorCategory), documentTemplateToUpload));
					}
				}
			}
		}
    }
}
