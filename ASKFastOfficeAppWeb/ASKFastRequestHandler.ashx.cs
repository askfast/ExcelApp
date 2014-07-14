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
    /// Summary description for ASKFastRequestHandler
    /// </summary>
    public class ASKFastRequestHandler : IHttpHandler
    {
        public static readonly String MARKETPLACE_PATH = "https://askfastmarket.appspot.com";
        //public static readonly String MARKETPLACE_PATH = "http://127.0.0.1:8888";
        public static readonly String appID = "EXCEL_OFFICE_APP";
        public static readonly String baseUrl = MARKETPLACE_PATH + "/products/broadcastnew/stream";
        public static readonly String fetchResponseURL = MARKETPLACE_PATH + "/resource/examples/clipboard";
        public void ProcessRequest(HttpContext context)
        {
            var response = "";
            var instanceId = HttpUtility.UrlEncode(context.Request.LogonUserIdentity.Name
                + "_" + context.Request.LogonUserIdentity.Owner.Value);
            try
            {
                WebClient askFast = new WebClient();
                foreach (var queryKey in context.Request.QueryString.AllKeys)
                {
                    switch (queryKey)
                    {
                        case "appId":
                            askFast.QueryString.Add(queryKey, appID + ":" + instanceId);
                            break;
                        case "clipboardKey":
                            askFast.QueryString.Add(queryKey, appID);
                            break;
                        case "instanceId":
                            askFast.QueryString.Add(queryKey, instanceId);
                            break;
                        default:
                            askFast.QueryString.Add(queryKey, context.Request.QueryString.Get(queryKey));
                            break;
                    }
                }
                if (context.Request.Headers.Get("X-SESSION_ID") != null)
                {
                    askFast.Headers.Add("X-SESSION_ID", context.Request.Headers.Get("X-SESSION_ID"));
                }
                //POST operation for performing outbound broadcasts
                if (context.Request.HttpMethod.Equals("POST"))
                {
                    String payload = null;
                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        payload = reader.ReadToEnd();
                    }
                    response = askFast.UploadString(MARKETPLACE_PATH + context.Request.PathInfo, payload);
                }
                //GET oepration to fetch the reports and login
                else if (context.Request.HttpMethod.Equals("GET"))
                {
                    response = askFast.DownloadString(MARKETPLACE_PATH + context.Request.PathInfo);
                }
            }
            catch (WebException ex)
            {
                var errorResponse = (HttpWebResponse)ex.Response;
                context.Response.StatusDescription = errorResponse.StatusDescription;
                context.Response.StatusCode = (int)errorResponse.StatusCode;
                response = "Error: " + ex.Message;
                Console.Write(response);
            }
            context.Response.ContentType = "application/json";
            context.Response.Write(response);
        }

        public bool IsReusable
        {
            get
            {
                return true;
            }
        }
    }
}