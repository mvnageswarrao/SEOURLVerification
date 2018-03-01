using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using System.Threading;
using System.Net;

namespace ConsoleApplication1
{
    class W3CValidityCheckResult
    {
        public bool IsValid { get; set; }
        public int WarningsCount { get; set; }
        public int ErrorsCount { get; set; }
        public string Body { get; set; }

        private static AutoResetEvent _w3cValidatorBlock = new AutoResetEvent(true);

        private static void ResetBlocker(object state)
        {
                // Ensures that W3C Validator service is not called more than once a second
                Thread.Sleep(1000);
                _w3cValidatorBlock.Set();
        }

        public static W3CValidityCheckResult ReturnsValidHtml(string url)
        {
                var result = new W3CValidityCheckResult();
                WebHeaderCollection w3cResponseHeaders = new WebHeaderCollection();

                using (var wc = new WebClient())
                {
                    string html = GetPageHtml(wc, url);

                    // Send to W3C validator
//                    string w3cUrl = "http://validator.w3.org/check";
                    string w3cUrl = "http://validator.w3.org/";

                    wc.Encoding = System.Text.Encoding.UTF8;
                    var values = new NameValueCollection();
                    values.Add("fragment", html);
                    values.Add("prefill", "0");
                    values.Add("group", "0");
                    values.Add("doctype", "inline");

                    try
                    {
                        _w3cValidatorBlock.WaitOne();
                        byte[] w3cRawResponse = wc.UploadValues(w3cUrl, values);
                        result.Body = Encoding.UTF8.GetString(w3cRawResponse);
                        w3cResponseHeaders.Add(wc.ResponseHeaders);
                    }
                    finally
                    {
                        ThreadPool.QueueUserWorkItem(ResetBlocker); // Reset on background thread
                    }
                }

                // Extract result from response headers
                int warnings = -1;
                int errors = -1;
                int.TryParse(w3cResponseHeaders["X-W3C-Validator-Warnings"], out warnings);
                int.TryParse(w3cResponseHeaders["X-W3C-Validator-Errors"], out errors);
                string status = w3cResponseHeaders["X-W3C-Validator-Status"];

                result.WarningsCount = warnings;
                result.ErrorsCount = errors;
                result.IsValid = (!String.IsNullOrEmpty(status) && status.Equals("Valid", StringComparison.InvariantCultureIgnoreCase));

                return result;
       }

       private static string GetPageHtml(WebClient wc, string url)
       {
                // Pretend to be Firefox 3 so that ASP.NET renders compliant HTML
                wc.Headers["User-Agent"] = "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US; rv:1.9.0.1) Gecko/2008070208 Firefox/3.0.1 (.NET CLR 3.5.30729)";
                wc.Headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                wc.Headers["Accept-Language"] = "en-au,en-us;q=0.7,en;q=0.3";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                // allows for validation of SSL conversations
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                // Read page HTML
                string html = "";
                WebClient web = new WebClient();
                web.UseDefaultCredentials = true;
                html = null;
                html = web.DownloadString(url);
                web.Dispose();
                return html;
       }
    }

       
}
