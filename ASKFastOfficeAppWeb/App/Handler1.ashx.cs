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
        static String baseUrl = "http://char-a-lot.appspot.com/rpc";

        public void ProcessRequest(HttpContext context)
        {
            String payload = null;
            using (var reader = new StreamReader(context.Request.InputStream))
            {
                payload = reader.ReadToEnd();
            }
            WebClient askFast = new WebClient();
            var response = askFast.UploadData(baseUrl, Encoding.ASCII.GetBytes(payload) );
            context.Response.ContentType = "text/plain";
            context.Response.Write(response);
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