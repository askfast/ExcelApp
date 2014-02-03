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
        public static readonly String MARKETPLACE_PATH = "http://askfastmarket1.appspot.com";
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
                    askFast.QueryString.Add(queryKey, context.Request.QueryString.Get(queryKey));
                }
                if (context.Request.Headers.Get("X-SESSION_ID") != null)
                {
                    askFast.Headers.Add("X-SESSION_ID", context.Request.Headers.Get("X-SESSION_ID"));
                }
                if (context.Request.HttpMethod.Equals("POST"))
                {
                    String payload = null;
                    //add the combination of appId:instanceId as a query parameter for having this as a unique request
                    askFast.QueryString.Add("appId", appID + ":" + instanceId);
                    using (var reader = new StreamReader(context.Request.InputStream))
                    {
                        payload = reader.ReadToEnd();
                    }
                    var responseByteArray = askFast.UploadData(baseUrl, Encoding.ASCII.GetBytes(payload));
                    response = System.Text.Encoding.Default.GetString(responseByteArray);
                }
                else if (context.Request.HttpMethod.Equals("GET"))
                {
                    String requestURL = "";
                    if (context.Request.PathInfo.Equals("/login"))
                    {
                        requestURL = MARKETPLACE_PATH + "/login";
                    }
                    else
                    {
                        //add the combination of appId:instanceId as a query parameter for having this as a unique request
                        askFast.QueryString.Add("clipboardKey", appID);
                        askFast.QueryString.Add("instanceId", instanceId);
                        requestURL = fetchResponseURL;
                    }
                    var responseByteArray = askFast.DownloadData(requestURL);
                    response = System.Text.Encoding.Default.GetString(responseByteArray);
                }
            }
            catch (Exception ex)
            {
                context.Response.StatusCode = 500;
                response = "Error: " + ex.Message;
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