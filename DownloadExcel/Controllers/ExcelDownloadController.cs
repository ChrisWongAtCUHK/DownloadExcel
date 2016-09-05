using ExcelUtility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace DownloadExcel.Controllers
{
    public class ExcelDownloadController : ApiController
    {
        // GET api/ExcelDownload?firstName=Chris&lastName=Wong
        [System.Web.Http.HttpGet]
        public HttpResponseMessage Get(string firstName, string lastName)
        {
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);

            try
            {
                var path = System.Web.HttpContext.Current.Server.MapPath("~/Content/Contacts.xlsx"); ;
                FileStream fs = File.OpenRead(path);

                // FileStream to MemoryStream, to overcome the FileAccess
                MemoryStream ms = new MemoryStream();
                byte[] bytes = new byte[fs.Length];
                fs.Read(bytes, 0, (int)fs.Length);
                ms.Write(bytes, 0, (int)fs.Length);

                Dictionary<string, string> updateDictionary = new Dictionary<string, string>();
                updateDictionary.Add("A2", firstName);
                updateDictionary.Add("B2", lastName);

                MemoryStream outputMs = MemoryStreamExcelHelper.UpdateExcel(ms, "Sheet1", updateDictionary);

                result.Content = new ByteArrayContent(outputMs.ToArray());
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentDisposition.FileName = Path.GetFileName(path);
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                result.Content.Headers.ContentLength = outputMs.Length;
            }
            catch (Exception ex)
            {

                string message = ex.Message;
            }

            return result;
        }
    }
}
