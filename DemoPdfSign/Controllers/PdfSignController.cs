using Aspose.Words;
using DemoPdfSign.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Office.Interop.Word;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using DataTable = System.Data.DataTable;
using Document = Microsoft.Office.Interop.Word.Document;

namespace DemoPdfSign.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfSignController : ControllerBase
    {
        private IHostingEnvironment environment;
        public PdfSignController(IHostingEnvironment _environment)
        {

            environment = _environment;
        }

        [HttpPost("sign")]
        public async Task<string> sign(PdfSignRequest pdfSignRequest)
        {
            try
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(@"D:\dev\DemoPdfSign\DemoPdfSign\Resources\Template\POC_Template.docx");

                var imageUrl = saveImage(pdfSignRequest.image);

                string[] fieldNames = new string[] { "SupplierName", "SupplierPocId", "OrderCode", "OrderDate", "OrderTime", 
                    "Signature", 
                    "CustomerName", "CustomerSupplierCode", "CustomerManufactureCode", "CustomerAddress", "CustomerPhoneNumber",
                    "BillingName", "BillingAddress", "TaxCode", "Comment" 
                };
                Object[] fieldValues = new Object[] { "ThanhEtn", "ETN001", "S001", "08/07/2022", "15:30:23", imageUrl, 
                    "Etn001", "S001", "M001", "xxx yyy", "000011112222",
                    "Bill 001", "uuuuuuu", "t0001", ""
                };
                doc.MailMerge.Execute(fieldNames, fieldValues);

                var savePath = environment.WebRootPath;
                var filePath = @"Template\POC_bill.pdf";

                if (!Directory.Exists(Path.Combine(savePath,"Template")))
                {
                    Directory.CreateDirectory(Path.Combine(savePath, "Template"));
                }

                doc.Save(Path.Combine(savePath, filePath));

                return filePath;
            }catch(Exception ex)
            {
                return ex.Message;
            }
        }

        [HttpPost("sign2")]
        public async Task<string> sign2(PdfSignRequest pdfSignRequest)
        {
            // create a temporary ServerTextControl
            try
            {
                var imageUrl = saveImage(pdfSignRequest.image);

                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = @"D:\dev\DemoPdfSign\DemoPdfSign\Resources\Template\POC_Template.docx";
                Object oHeader = "FirstName, LastName, Address, CityStateZip";

                string[] fieldNames = new string[] { "SupplierName", "SupplierPocId", "OrderCode", "OrderDate", "OrderTime",
                    "Signature",
                    "CustomerName", "CustomerSupplierCode", "CustomerManufactureCode", "CustomerAddress", "CustomerPhoneNumber",
                    "BillingName", "BillingAddress", "TaxCode", "Comment"
                };
                Object[] fieldValues = new Object[] { "ThanhEtn", "ETN001", "S001", "08/07/2022", "15:30:23", 
                    imageUrl,
                    "Etn001", "S001", "M001", "xxx yyy", "000011112222",
                    "Bill 001", "uuuuuuu", "t0001", ""
                };

                Object oFalse = false;
                Selection wrdSelection;
                MailMerge wrdMailMerge;
                MailMergeFields wrdMergeFields;
                Application wrdApp = new Application();
                _Document wrdDoc;
                //wrdApp.Visible = true;

                //Microsoft.Office.Interop.Word.Document wrdDoc = new Microsoft.Office.Interop.Word.Document();

                wrdDoc = wrdApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                wrdDoc.Select();

                //var wrdDoc = wrdApp.Documents.Add(oTemplatePath);

                wrdSelection = wrdApp.Selection;
                wrdMailMerge = wrdDoc.MailMerge;

                wrdDoc.MailMerge.CreateDataSource(ref oTemplatePath, ref oMissing, ref oMissing, ref oHeader, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Insert merge data.
                wrdSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                wrdMergeFields = wrdMailMerge.Fields;
                wrdMergeFields.Add(wrdSelection.Range, "SupplierName");
                wrdSelection.TypeText("ThanhEtn");

                // Perform mail merge.
                wrdMailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                wrdMailMerge.Execute(ref oFalse);

                var savePath = environment.WebRootPath;
                var filePath = @"Template\POC_bill.docx";

                if (!Directory.Exists(Path.Combine(savePath, "Template")))
                {
                    Directory.CreateDirectory(Path.Combine(savePath, "Template"));
                }

                wrdDoc.SaveAs(Path.Combine(savePath, filePath));

                return filePath;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private string saveImage(string imageUrl)
        {
            var dataUri = imageUrl;
            var encodedImage = dataUri.Split(',')[1];
            var decodedImage = Convert.FromBase64String(encodedImage);


            var savePath = Path.Combine(environment.WebRootPath, @"signature\sign001.png");


            if (!Directory.Exists(Path.Combine(environment.WebRootPath, "signature")))
            {
                Directory.CreateDirectory(Path.Combine(environment.WebRootPath, "signature"));
            }

            System.IO.File.WriteAllBytes(savePath, decodedImage);

            return savePath;
        }
    }
}
