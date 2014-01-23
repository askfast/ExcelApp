using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

namespace ASKFastOfficeAppWeb.App
{
    /// <summary>
    /// Summary description for Handler1
    /// </summary>
    public class Handler1 : IHttpHandler
    {
        static String baseUrl = "http://shravan1.askfastmarket1.appspot.com/products/broadcastnew/stream?username=apptestoneline&password=eadeb77d8fba90b42b32b7de13e8aaa6&useClipboard=true";
        public void ProcessRequest(HttpContext context)
        {
            String payload = null;
            using (var reader = new StreamReader(context.Request.InputStream))
            {
                payload = reader.ReadToEnd();
            }
            WebClient askFast = new WebClient();
            //var response = askFast.UploadData(baseUrl, Encoding.ASCII.GetBytes(payload) );
            context.Response.ContentType = "text/plain";
            //context.Response.Write(response);
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}