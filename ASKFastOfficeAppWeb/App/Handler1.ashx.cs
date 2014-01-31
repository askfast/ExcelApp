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
        static String appID = "EXCEL_OFFICE_APP";
        static String baseUrl = "http://shravan1.askfastmarket1.appspot.com/products/broadcastnew/stream?username=apptestoneline&password=eadeb77d8fba90b42b32b7de13e8aaa6";
        //static string fetchResponseURL = "http://shravan1.askfastmarket1.appspot.com/resource/examples/clipboard?username=apptestoneline&password=eadeb77d8fba90b42b32b7de13e8aaa6";
        static string fetchResponseURL = "http://127.0.0.1:8888/resource/examples/clipboard?username=apptestoneline&password=eadeb77d8fba90b42b32b7de13e8aaa6";
        public void ProcessRequest(HttpContext context)
        {
            var response = "";
            var instanceId = HttpUtility.UrlEncode(context.Request.LogonUserIdentity.Name 
                + "_" + context.Request.LogonUserIdentity.Owner.Value);
            var appIdParameter = "appId=" + appID + ":" + instanceId;
            if (context.Request.HttpMethod.Equals("POST"))
            {
                String payload = null;
                foreach (var queryKey in context.Request.QueryString.AllKeys)
                {
                    baseUrl += "&" + queryKey + "=" + context.Request.QueryString[queryKey];
                }
                //add the combination of appId:instanceId as a query parameter for having this as a unique request
                baseUrl += "&" + appIdParameter;
                using (var reader = new StreamReader(context.Request.InputStream))
                {
                    payload = reader.ReadToEnd();
                }
                WebClient askFast = new WebClient();
                var responseByteArray = askFast.UploadData(baseUrl, Encoding.ASCII.GetBytes(payload));
                response = System.Text.Encoding.Default.GetString(responseByteArray);
            }
            else if(context.Request.HttpMethod.Equals("GET"))
            {
                //add the combination of appId:instanceId as a query parameter for having this as a unique request
                fetchResponseURL += "&clipboardKey=" + appID + "&instanceId=" + instanceId;
                foreach (var queryKey in context.Request.QueryString.AllKeys)
                {
                    fetchResponseURL += "&" + queryKey + "=" + context.Request.QueryString[queryKey];
                }
                WebClient askFast = new WebClient();
                var responseByteArray = askFast.DownloadData(fetchResponseURL);
                response = System.Text.Encoding.Default.GetString(responseByteArray);
            }
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