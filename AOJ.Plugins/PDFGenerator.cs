using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json;
using System;
using System.Activities;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Text;

namespace AOJ.Plugins
{
    public class PDFGenerator : CodeActivity
    {
        [Input("ContactIds")]
        public InArgument<string> ContactIds { get; set; }

        [Output("PDF")]
        public OutArgument<string> PDF { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            ITracingService log = context.GetExtension<ITracingService>();
            IOrganizationService svc = context.GetExtension<IOrganizationServiceFactory>().CreateOrganizationService(workflowContext.UserId);

            log.Trace("Step 1");
            var ids = ContactIds.Get(context).Split(',');
            log.Trace("Step 2");

            var notes = svc.RetrieveMultiple(new FetchExpression($@"
                <fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                    <entity name=""annotation"">
                    <all-attributes />
                    <filter type=""and"">
                        <condition attribute=""mimetype"" operator=""eq"" value=""application/pdf"" />
                    </filter>
                    <link-entity name=""contact"" from=""contactid"" to=""objectid"" link-type=""inner"" alias=""aa"">
                        <filter type=""or"">
                        {ids.Select(x => $"<condition attribute=\"contactid\" operator=\"eq\" value=\"{x}\" />")}                            
                        </filter>
                    </link-entity>
                    </entity>
                </fetch>
                "));
            log.Trace("Step 3");

            try
            {
                // Create a MemoryStream to hold the merged PDF
                using (MemoryStream mergedStream = new MemoryStream())
                {
                    log.Trace("Step 4");
                    // Create a Document and PdfCopy object
                    Document mergedDocument = new Document();

                    log.Trace("Step 5");
                    PdfWriter writer = PdfWriter.GetInstance(mergedDocument, mergedStream);
                    mergedDocument.Open();

                    log.Trace("Step 6");
                    foreach (var note in notes.Entities)
                    {
                        byte[] pdfBytes = Convert.FromBase64String(note["documentbody"].ToString());
                        using (var pdfStream = new MemoryStream(pdfBytes))
                        {
                            var pdfReader = new PdfReader(pdfStream);

                            for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                            {
                                var page = writer.GetImportedPage(pdfReader, i);
                                mergedDocument.Add(Image.GetInstance(page));
                            }
                        }
                    }
                    log.Trace("Step 9");
                    mergedDocument.Close();

                    // Update the target entity with the merged PDF
                    byte[] mergedContent = mergedStream.ToArray();

                    log.Trace("Step 10");
                    PDF.Set(context, Convert.ToBase64String(mergedContent));
                    log.Trace("Step 11");
                }
            }
            catch (Exception ex)
            {
                log.Trace(ex.Message);
                log.Trace(ex.StackTrace);
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}
