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
            Document wrdDoc = new Microsoft.Office.Interop.Word.Document();
            Application wrdApp = new Application();

            try
            {
                var imageUrl = saveImage(pdfSignRequest.image);

                Object oMissing = System.Reflection.Missing.Value;
                Object oTemplatePath = @"C:\dev\PdfSign\DemoPdfSign\Resources\Template\POC_Template.docx";
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
                //wrdApp.Visible = true;

                //Microsoft.Office.Interop.Word.Document wrdDoc = new Microsoft.Office.Interop.Word.Document();

                wrdDoc = wrdApp.Documents.Add(oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                foreach (Field myMergeField in wrdDoc.Fields)
                {

                    Microsoft.Office.Interop.Word.Range rngFieldCode = myMergeField.Code;
                    String fieldText = rngFieldCode.Text;
                    // ONLY GETTING THE MAILMERGE FIELDS
                    if (fieldText.StartsWith(" MERGEFIELD"))
                    {
                        // THE TEXT COMES IN THE FORMAT OF
                        // MERGEFIELD  MyFieldName  \\* MERGEFORMAT
                        // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"
                        Int32 endMerge = fieldText.IndexOf("\\");
                        Int32 fieldNameLength = fieldText.Length - endMerge;
                        String fieldName = fieldText.Substring(11, endMerge - 11);
                        // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE
                        fieldName = fieldName.Trim();
                        // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//
                        // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE
                        if (fieldName == "SupplierName")
                        {
                            myMergeField.Select();
                            wrdApp.Selection.TypeText("ThanhEtn");
                        }

                        if (fieldName == "Image:Signature")
                        {
                            myMergeField.Select();
                            wrdApp.Selection.InlineShapes.AddPicture(imageUrl);
                        }
                    }
                }

                var savePath = environment.WebRootPath;
                var filePath = @"Template\POC_bill.pdf";

                if (!Directory.Exists(Path.Combine(savePath, "Template")))
                {
                    Directory.CreateDirectory(Path.Combine(savePath, "Template"));
                }

                wrdDoc.ExportAsFixedFormat(Path.Combine(savePath, filePath), Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                return filePath;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                wrdDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                wrdApp.Quit();
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
