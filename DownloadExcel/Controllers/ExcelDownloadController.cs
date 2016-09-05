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
        // GET api/ExcelDownload
        [System.Web.Http.HttpGet]
        public HttpResponseMessage Get()
        {
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);

            try
            {
                var path = System.Web.HttpContext.Current.Server.MapPath("~/Content/Contacts.xlsx"); ;
                FileStream fs = File.OpenRead(path);

                result.Content = new StreamContent(fs);
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                result.Content.Headers.ContentDisposition.FileName = Path.GetFileName(path);
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                result.Content.Headers.ContentLength = fs.Length;
            }
            catch (Exception ex)
            {

                string message = ex.Message;
            }

            return result;
        }
    }
}
