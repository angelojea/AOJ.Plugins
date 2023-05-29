using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;
using System.Xml.Linq;

namespace AOJ.ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var url = "https://orga0410328.crm.dynamics.com/";
            var user = "admin@CRMbc508129.onmicrosoft.com";
            var pwd = "abb7qay3PS";

            using (var svc = new CrmServiceClient(
                $@"AuthenticationType=Office365; url={url}; UserName={user}; Password={pwd};"))
            {
                var ids = new List<string>() { "5ca50030-cdfd-ed11-8f6d-000d3a9d67f2" };

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

                try
                {
                    // Create a Document and PdfCopy object
                    Document mergedDocument = new Document();
                    MemoryStream outputStream = new MemoryStream();
                    PdfWriter writer = PdfWriter.GetInstance(mergedDocument, outputStream);

                    mergedDocument.Open();

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

                    mergedDocument.Close();

                    // Update the target entity with the merged PDF
                    byte[] mergedContent = outputStream.ToArray();

                    //Console.WriteLine(Convert.ToBase64String(mergedContent));

                    using (FileStream fileStream = new FileStream("C:\\Users\\angelojea\\Desktop\\output\\Resume.pdf", FileMode.OpenOrCreate))
                    {
                        byte[] bytes = outputStream.ToArray();
                        fileStream.Write(bytes, 0, bytes.Length);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message, ex);
                }
            }
        }
    }
}
