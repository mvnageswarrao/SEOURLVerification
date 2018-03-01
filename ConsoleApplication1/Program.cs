using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Web;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Collections.Specialized;
using ConsoleApplication1;
using System.Reflection;

namespace URLResponse
{
    static class Program
    {
        public static string XMLName { get; set; }
        public static void Main(string[] args)
        {
            const int maxMenuItems = 15;
            int selector = 0;
            bool cmdLineParam = false;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // allows for validation of SSL conversations
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            if (args.Length > 0)
            {
                int i = 0;
                selector = Convert.ToInt32(args[i]);
            }

            while (selector != maxMenuItems)
            {

                if (args.Length == 0)
                {
                    Console.Clear();
                    DrawTitle();
                    DrawMenu(maxMenuItems);
                    //Console.WriteLine(selector);
                    cmdLineParam = int.TryParse(Console.ReadLine(), out selector);
                }
                else
                {
                    cmdLineParam = true;
                }
                if (cmdLineParam)
                {
                    switch (selector)
                    {
                        case 1:
                            StringBuilder strLog = new StringBuilder();
                            string sitemapURLfile = ConfigurationSettings.AppSettings["sitemapURLfile"].ToString();
                            string logFileSEO = initiateLogFile(selector);
                            writeconsole(1, logFileSEO);
                            System.IO.StreamWriter logfileName = new System.IO.StreamWriter(logFileSEO, true);

                            try
                            {
                                strLog.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL" + '\t' + "PageTitle" + '\t' + "PageTitleAlex" + '\t' + "PageTitleStatus" + '\t' + "MetaDescription" + '\t' + "MetaDescriptionAlex" + '\t' + "MetaDescriptionStatus" + '\t' + "MetaKeyword" + '\t' + "PriMetaKeywordAlex" + '\t' + "SecMetaKeywordAlex" + '\t' + "TrtMetaKeywordAlex" + '\t' + "PriMetaKeywordStatus" + '\t' + "SecMetaKeywordStatus" + '\t' + "TrtMetaKeywordStatus" + '\t' + "H1" + '\t' + "H1Alex" + '\t' + "H1Status");
                                
                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(sitemapURLfile);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking URL : " + str);
                                    checkURLstatus(str, logFileSEO, strLog, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(logfileName, strLog.ToString());
                            }
                            writeconsole(0, logFileSEO);
                            var name = Console.ReadLine();
                            break;

                        case 16:
                            string Dira = ConfigurationSettings.AppSettings["dir"].ToString();
                            string sitemapURLfilea = ConfigurationSettings.AppSettings["sitemapURLfile"].ToString();
                            //string URLa = "http://www.villaplus.com/algarve/villas/praia-d-oura/villa-diniz";

                                DataTable sliderImgData = new DataTable();
                                DataTable ActivitiesImg = new DataTable();
                                DataTable ToursData = new DataTable();
                                DataTable FloorPlansData = new DataTable();
                                DataTable PhotosTabData = new DataTable();
                                DataTable[] table = new DataTable[5];
                                StringBuilder strLoga = new StringBuilder();

                            try
                            {
                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(sitemapURLfilea);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string URLa = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    URLa = n.InnerText;
                                    int slashcount = 0;
                                    foreach (char c in URLa) 
                                    {
                                        if (c == '/')
                                        {
                                            slashcount++;
                                        }
                                    }

                                    if (slashcount == 6)
                                    {
                                        var hwa = new HtmlWeb();
                                        HtmlDocument doca = hwa.Load(URLa);
                                        string inputa = doca.DocumentNode.OuterHtml;

                                        string VillaID = getVillaID(inputa);
                                        string logFileImgChk = "VillaID_" + VillaID + "_IMGVillaplus" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm");
                                        string FileFullPatha = Dira + logFileImgChk;
                                        table[0] = getImageDataDetails(URLa, VillaID, inputa);
                                        table[1] = ActImageDataDetails(URLa, VillaID, inputa);
                                        table[2] = ToursDetails(URLa, VillaID, inputa);
                                        table[3] = FloorPlanDetails(URLa, VillaID, inputa);
                                        table[4] = PhotostabDetails(URLa, VillaID, inputa);
                                        ExportToExcel(table, FileFullPatha);
                                    }
                                }
                            }
                            finally
                            {
                                //actionLogFile(logfileName, strLog.ToString());
                            }
                        break;                           
                        
                        case 2:
                            string platform = ConfigurationSettings.AppSettings.Get("platform");
                            int count = 0;
                            if (args.Length > 0)
                            {
                                if (platform == "")
                                {
                                    string[] arrPlatform = { "D", "T", "M" };
                                    string dirForxmlFile = ConfigurationSettings.AppSettings.Get("dir");
                                    string whichXML = ConfigurationSettings.AppSettings.Get("whichXML");
                                    Boolean successXML = getXMLandSave(whichXML, dirForxmlFile);
                                    
                                    DataTable dtStatusStatsTable = new DataTable();
                                    dtStatusStatsTable.Columns.Add("platform");
                                    dtStatusStatsTable.Columns.Add("status");
                                    dtStatusStatsTable.Columns.Add("column");

                                    foreach (string platformFrmArr in arrPlatform)
                                    {
                                        platform = platformFrmArr;
                                        StringBuilder strLogurl = new StringBuilder();
                                        //string newUrlChecksfile = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                                        string logFileURLMapping = initiateLogFile(selector, platform, whichXML);
                                        //writeconsole(1, logFileURLMapping);
                                        string line = string.Empty;
                                        System.IO.StreamWriter newUrllog = new System.IO.StreamWriter(logFileURLMapping, true);
                                        if (File.Exists(dirForxmlFile + whichXML))
                                        {
                                            //System.IO.StreamReader file5 = new System.IO.StreamReader(newUrlChecksfile);
                                            try
                                            {
                                                //strLogurl.AppendLine("http://www.villaplus.com/" + whichXML);
                                                if (!successXML)
                                                {
                                                    //strLogurl.AppendLine("Ran test with old " + whichXML + " file");

                                                }
                                                strLogurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");
                                                System.Xml.XmlDocument sitemapdoc = new System.Xml.XmlDocument();
                                                sitemapdoc.Load(dirForxmlFile + whichXML);
                                                System.Xml.XmlElement xmlroot = sitemapdoc.DocumentElement;
                                                System.Xml.XmlNodeList urllst = xmlroot.GetElementsByTagName("loc");
                                                string strURL = string.Empty;
                                                foreach (System.Xml.XmlNode n in urllst)
                                                {
                                                    strURL = n.InnerText;
                                                    switch (platform)
                                                    {
                                                        case "D":
                                                            break;
                                                        case "T":
                                                            strURL = strURL.Replace("www", "t");
                                                            break;
                                                        case "M":
                                                            strURL = strURL.Replace("www", "m");
                                                            break;
                                                        default:
                                                            break;
                                                    }
                                                    count = urllst.Count;
                                                    //Console.WriteLine("Checking URL : " + strURL);
                                                    checkURLstatus(strURL, logFileURLMapping, strLogurl, selector);
                                                }
                                            }
                                            finally
                                            {
                                                actionLogFile(newUrllog, strLogurl.ToString());
                                            }
                                            //writeconsole(0, logFileURLMapping);
                                            //var name1 = Console.ReadLine();
                                        }
                                        else
                                        {
                                            //Console.WriteLine("XML not found.. exiting");
                                            //strLogurl.AppendLine("XML " + whichXML + " not found... Exiting the Test");
                                            actionLogFile(newUrllog, strLogurl.ToString());
                                        }

                                        dtStatusStatsTable = getUrlStatusStats(logFileURLMapping, platform, dtStatusStatsTable);
                                        saveInXls(logFileURLMapping);
                                    }
                                    string htmlFileName = createHTMLReport(dtStatusStatsTable, successXML, whichXML, count);

                                    if (args.Length > 0)
                                    {
                                        break;
                                    }
                                }

                                else
                                {
                                    string[] arrPlatform = {platform};
                                    string dirForxmlFile = ConfigurationSettings.AppSettings.Get("dir");
                                    string whichXML = ConfigurationSettings.AppSettings.Get("whichXML");
                                    Boolean successXML = getXMLandSave(whichXML, dirForxmlFile);
                                    DataTable dtStatusStatsTable = new DataTable();
                                    dtStatusStatsTable.Columns.Add("platform");
                                    dtStatusStatsTable.Columns.Add("status");
                                    dtStatusStatsTable.Columns.Add("column");

                                    foreach (string platformFrmArr in arrPlatform)
                                    {
                                        platform = platformFrmArr;
                                        StringBuilder strLogurl = new StringBuilder();
                                        //string newUrlChecksfile = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                                        string logFileURLMapping = initiateLogFile(selector, platform, whichXML);
                                        //writeconsole(1, logFileURLMapping);
                                        string line = string.Empty;
                                        System.IO.StreamWriter newUrllog = new System.IO.StreamWriter(logFileURLMapping, true);
                                        if (File.Exists(dirForxmlFile + whichXML))
                                        {
                                            //System.IO.StreamReader file5 = new System.IO.StreamReader(newUrlChecksfile);
                                            try
                                            {
                                                //strLogurl.AppendLine("http://www.villaplus.com/" + whichXML);
                                                if (!successXML)
                                                {
                                                    //strLogurl.AppendLine("Ran test with old " + whichXML + " file");

                                                }
                                                strLogurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");
                                                System.Xml.XmlDocument sitemapdoc = new System.Xml.XmlDocument();
                                                sitemapdoc.Load(dirForxmlFile + whichXML);
                                                System.Xml.XmlElement xmlroot = sitemapdoc.DocumentElement;
                                                System.Xml.XmlNodeList urllst = xmlroot.GetElementsByTagName("loc");
                                                string strURL = string.Empty;
                                                foreach (System.Xml.XmlNode n in urllst)
                                                {
                                                    strURL = n.InnerText;
                                                    switch (platform)
                                                    {
                                                        case "D":
                                                            break;
                                                        case "T":
                                                            strURL = strURL.Replace("www", "t");
                                                            break;
                                                        case "M":
                                                            strURL = strURL.Replace("www", "m");
                                                            break;
                                                        case "ST":
                                                            strURL = strURL.Replace("www", "staging1");
                                                            break;
                                                        default:
                                                            break;
                                                    }
                                                    count = urllst.Count;
                                                    //Console.WriteLine("Checking URL : " + strURL);
                                                    checkURLstatus(strURL, logFileURLMapping, strLogurl, selector);
                                                }
                                            }
                                            finally
                                            {
                                                actionLogFile(newUrllog, strLogurl.ToString());
                                            }
                                            //writeconsole(0, logFileURLMapping);
                                            //var name1 = Console.ReadLine();
                                        }
                                        else
                                        {
                                            //Console.WriteLine("XML not found.. exiting");
                                            //strLogurl.AppendLine("XML " + whichXML + " not found... Exiting the Test");
                                            actionLogFile(newUrllog, strLogurl.ToString());
                                        }

                                        dtStatusStatsTable = getUrlStatusStats(logFileURLMapping, platform, dtStatusStatsTable);
                                        saveInXls(logFileURLMapping);
                                    }
                                    string htmlFileName = createHTMLReport(dtStatusStatsTable, successXML, whichXML, count);

                                    if (args.Length > 0)
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                StringBuilder strLogurl = new StringBuilder();
                                string newUrlChecksfile = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                                string logFileURLMapping = initiateLogFile(selector);
                                writeconsole(1, logFileURLMapping);
                                System.IO.StreamWriter newUrllog = new System.IO.StreamWriter(logFileURLMapping, true);
                                System.IO.StreamReader file5 = new System.IO.StreamReader(newUrlChecksfile);
                                string line = "";
                                try
                                {
                                    strLogurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");

                                    while ((line = file5.ReadLine()) != null)
                                    {
                                        Console.WriteLine("Checking URL : " + line);
                                        checkURLstatus(line, logFileURLMapping, strLogurl, selector);
                                    }
                                }
                                finally
                                {
                                    actionLogFile(newUrllog, strLogurl.ToString());
                                }
                                writeconsole(0, logFileURLMapping);
                                var name1 = Console.ReadLine();
                            }
                            break;

                        case 3:
                            StringBuilder strLogGA = new StringBuilder();
                            string sitemapURLfileGA = ConfigurationSettings.AppSettings["sitemapURLfile"].ToString();
                            string logFileGA = initiateLogFile(selector);
                            writeconsole(1, logFileGA);
                            System.IO.StreamWriter galogfileName = new System.IO.StreamWriter(logFileGA, true);

                            try
                            {
                                strLogGA.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL" + '\t' + "GAStatus" + '\t' + "GATrackingCodeStatus");

                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(sitemapURLfileGA);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking GA in URL : " + str);
                                    checkURLstatus(str, logFileGA, strLogGA, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(galogfileName, strLogGA.ToString());
                            }
                            writeconsole(0, logFileGA);
                            var name2 = Console.ReadLine();
                            break;

                        case 4:
                            StringBuilder strPPLog = new StringBuilder();
                            string PPChecksfile = ConfigurationSettings.AppSettings["PPUrlList"].ToString();
                            string logFilePPURLs = initiateLogFile(selector);
                            writeconsole(1, logFilePPURLs);
                            string ppURLline = string.Empty;
                            System.IO.StreamWriter PPmissingimagechecklog = new System.IO.StreamWriter(logFilePPURLs, true);
                            System.IO.StreamReader file1 = new System.IO.StreamReader(PPChecksfile);

                            try
                            {
                                strPPLog.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");

                                while ((ppURLline = file1.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking URL : " + ppURLline);
                                    checkURLstatus(ppURLline, logFilePPURLs, strPPLog, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(PPmissingimagechecklog, strPPLog.ToString());
                            }

                            writeconsole(0, logFilePPURLs);
                            var name4 = Console.ReadLine();
                            break;
                        default:
                            if (selector != maxMenuItems)
                            {
                                ErrorMessage();
                            }
                            break;

                        case 5:
                            StringBuilder strPCIurl = new StringBuilder();
                            string PCICHKSURL = ConfigurationSettings.AppSettings["PCICHKSURL"].ToString();
                            string logFilePCI = initiateLogFile(selector);
                            writeconsole(1, logFilePCI);
                            string[] Char = null;
                            Char = ConfigurationSettings.AppSettings["Char"].ToString().Split(',');
                            string line5 = string.Empty;
                            System.IO.StreamWriter newPCILog = new System.IO.StreamWriter(logFilePCI, true);
                            System.IO.StreamReader newfile6 = new System.IO.StreamReader(PCICHKSURL);

                            try
                            {
                                strPCIurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Description" + '\t' + "Result");

                                while ((line5 = newfile6.ReadLine()) != null)
                                {
                                    foreach (string str in Char)
                                    {
                                        string urlToCheck = line5 + str;
                                        Console.WriteLine("Checking URL : " + urlToCheck);
                                        checkURLstatus(urlToCheck, logFilePCI, strPCIurl, selector);
                                    }
                                }
                            }
                            finally
                            {
                                actionLogFile(newPCILog, strPCIurl.ToString());
                            }
                            writeconsole(0, logFilePCI);
                            var name6 = Console.ReadLine();
                            break;

                        case 6:
                            StringBuilder strCanonicalTags = new StringBuilder();
                            string chksCanonicalTags = ConfigurationSettings.AppSettings["URLsCHKCanonicalTags"].ToString();
                            string logFileCanonical = initiateLogFile(selector);
                            writeconsole(1, logFileCanonical);
                            System.IO.StreamWriter canonicallogfileName = new System.IO.StreamWriter(logFileCanonical, true);

                            string strToreplace = "";
                            try
                            {
                                strCanonicalTags.AppendLine("URL" + '\t' + "CANONICALTagsStatus" + '\t' + "ResponseURL" + '\t' + "CanonicalURLCheckFromPageSouce" + '\t' + "CanonicalURLsStatus" + '\t' + "AlternateStatus" + '\t' + "AlternateURL" + '\t' + "AlternateURLStatus");

                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(chksCanonicalTags);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking Canonical Tags in URL : " + str);
                                    checkURLstatus(str, logFileCanonical, strCanonicalTags, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(canonicallogfileName, strCanonicalTags.ToString());
                            }
                            writeconsole(0, logFileCanonical);
                            var name7 = Console.ReadLine();
                            break;


                        case 7:
                            StringBuilder strPageViewTags = new StringBuilder();
                            string chksPageViewTags = ConfigurationSettings.AppSettings["URLsCHKPageViewTags"].ToString();
                            string logFilePageView = initiateLogFile(selector);
                            writeconsole(1, logFilePageView);
                            System.IO.StreamWriter PageViewlogfileName = new System.IO.StreamWriter(logFilePageView, true);

                            try
                            {
                                strPageViewTags.AppendLine("URL" + '\t' + "CreatePageViewTagsStatus" + '\t' + "SetClientIDTagsStatus" + '\t' + "Eluminate.js FileCheckStatus");

                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(chksPageViewTags);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking Page View Tags in URL : " + str);
                                    checkURLstatus(str, logFilePageView, strPageViewTags, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(PageViewlogfileName, strPageViewTags.ToString());
                            }
                            writeconsole(0, logFilePageView);
                            var name8 = Console.ReadLine();
                            break;


                        case 8:
                            StringBuilder strNoIndexNoFollowTags = new StringBuilder();
                            string chksNoINdexNoFollowTags = ConfigurationSettings.AppSettings["URLsCHKNoIndexNoFollowTags"].ToString();
                            string logFileNoIndexNoFollow = initiateLogFile(selector);
                            writeconsole(1, logFileNoIndexNoFollow);
                            System.IO.StreamWriter NoIndexNoFollowlogfileName = new System.IO.StreamWriter(logFileNoIndexNoFollow, true);

                            try
                            {
                                strNoIndexNoFollowTags.AppendLine("URL" + '\t' + "NoIndexNoFollowTagsStatus");

                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(chksNoINdexNoFollowTags);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking No Index No Follow Tags in URL : " + str);
                                    checkURLstatus(str, logFileNoIndexNoFollow, strNoIndexNoFollowTags, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(NoIndexNoFollowlogfileName, strNoIndexNoFollowTags.ToString());
                            }
                            writeconsole(0, logFileNoIndexNoFollow);
                            var name9 = Console.ReadLine();
                            break;

                        case 9:
                            StringBuilder strLogCheckUsingCrawlerurl = new StringBuilder();
                            string CheckUsingCrawlerFile = ConfigurationSettings.AppSettings["CheckUsingCrawler"].ToString();
                            string logCheckUsingCrawler = initiateLogFile(selector);
                            writeconsole(1, logCheckUsingCrawler);
                            System.IO.StreamWriter CheckUsingCrawlerfileName = new System.IO.StreamWriter(logCheckUsingCrawler, true);

                            try
                            {
                                //strLogCheckUsingCrawlerurl.AppendLine("URL" + '\t' + "CreatePageViewTagsStatus" + '\t' + "SetClientIDTagsStatus" + '\t' + "Eluminate.js FileCheckStatus");

                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                doc.Load(CheckUsingCrawlerFile);
                                System.Xml.XmlElement root = doc.DocumentElement;
                                System.Xml.XmlNodeList lst = root.GetElementsByTagName("loc");

                                string str = string.Empty;
                                foreach (System.Xml.XmlNode n in lst)
                                {
                                    str = n.InnerText;
                                    Console.WriteLine("Checking Page : " + str);
                                    checkURLstatus(str, logCheckUsingCrawler, strLogCheckUsingCrawlerurl, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(CheckUsingCrawlerfileName, strLogCheckUsingCrawlerurl.ToString());
                            }
                            writeconsole(0, logCheckUsingCrawler);
                            var name10 = Console.ReadLine();
                            break;

                        case 10:
                            StringBuilder strFormHeader = new StringBuilder();
                            string newUrlChecksfile10 = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                            string logFileURLMapping10 = initiateLogFile(selector);
                            writeconsole(1, logFileURLMapping10);
                            string line10 = string.Empty;
                            System.IO.StreamWriter newFormHeader = new System.IO.StreamWriter(logFileURLMapping10, true);
                            System.IO.StreamReader file10 = new System.IO.StreamReader(newUrlChecksfile10);

                            try
                            {
                                strFormHeader.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");

                                while ((line10 = file10.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking URL : " + line10);
                                    checkURLstatus(line10, logFileURLMapping10, strFormHeader, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(newFormHeader, strFormHeader.ToString());
                            }
                            writeconsole(0, logFileURLMapping10);
                            if (args.Length > 0)
                            {
                                break;
                            }
                            var name11 = Console.ReadLine();
                            break;

                        case 11:
                            StringBuilder strUserAgent = new StringBuilder();
                            string userAgentfile = ConfigurationSettings.AppSettings["userAgentfile"].ToString();
                            string logFileURLMapping11 = initiateLogFile(selector);
                            writeconsole(1, logFileURLMapping11);
                            string line11 = string.Empty;
                            System.IO.StreamWriter newFormHeader1 = new System.IO.StreamWriter(logFileURLMapping11, true);
                            System.IO.StreamReader file11 = new System.IO.StreamReader(userAgentfile);
                            string urlBrowserinfo = "http://staging1.villaplus.com/browserInfo.aspx ";

                            try
                            {
                                strUserAgent.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL");

                                while ((line11 = file11.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking for User Agent : " + line11);
                                    checkURLstatus(urlBrowserinfo, selector, line11);
                                    break;
                                }
                            }
                            finally
                            {
                                actionLogFile(newFormHeader1, strUserAgent.ToString());
                            }
                            writeconsole(0, logFileURLMapping11);
                            Console.WriteLine("Done...");
                            if (args.Length > 0)
                            {
                                break;
                            }
                            var name12 = Console.ReadLine();
                            break;

                        case 12:
                            StringBuilder strDFAtagurl = new StringBuilder();
                            string dfaTagUrlChecksfile = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                            string dfaTaglogFile = initiateLogFile(selector);
                            writeconsole(1, dfaTaglogFile);
                            System.IO.StreamWriter dfaTagUrllog = new System.IO.StreamWriter(dfaTaglogFile, true);
                            System.IO.StreamReader file12 = new System.IO.StreamReader(dfaTagUrlChecksfile);
                            string line12 = "";
                            try
                            {
                                strDFAtagurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL" + '\t' + "DFATag");

                                while ((line12 = file12.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking URL : " + line12);
                                    checkURLstatus(line12, dfaTaglogFile, strDFAtagurl, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(dfaTagUrllog, strDFAtagurl.ToString());
                            }
                            writeconsole(0, dfaTaglogFile);
                            var name13 = Console.ReadLine();
                            break;

                        case 13:
                            StringBuilder strCriteoTagurl = new StringBuilder();
                            string CriteoTagUrlChecksfile = ConfigurationSettings.AppSettings["newUrlChecksfile"].ToString();
                            string CriteoTaglogFile = initiateLogFile(selector);
                            writeconsole(1, CriteoTaglogFile);
                            System.IO.StreamWriter CriteoTagUrllog = new System.IO.StreamWriter(CriteoTaglogFile, true);
                            System.IO.StreamReader file13 = new System.IO.StreamReader(CriteoTagUrlChecksfile);
                            string line13 = "";
                            try
                            {
                                strCriteoTagurl.AppendLine("URL" + '\t' + "Status" + '\t' + "Result" + '\t' + "ResponseURL" + '\t' + "CriteoTag");

                                while ((line13 = file13.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking URL : " + line13);
                                    checkURLstatus(line13, CriteoTaglogFile, strCriteoTagurl, selector);
                                }
                            }
                            finally
                            {
                                actionLogFile(CriteoTagUrllog, strCriteoTagurl.ToString());
                            }
                            writeconsole(0, CriteoTaglogFile);
                            var name14 = Console.ReadLine();
                            break;

                        case 14:
                            string logFileW3CCheck = initiateLogFile(selector);
                            writeconsole(1, logFileW3CCheck);
                            System.IO.StreamWriter logFileW3CChecklog = new System.IO.StreamWriter(logFileW3CCheck, true);

                            string W3CCheckUrls = ConfigurationSettings.AppSettings["W3CCheckUrls"].ToString();
                            System.IO.StreamReader newfile7 = new System.IO.StreamReader(W3CCheckUrls);

                            string line6 = string.Empty;
                            StringBuilder strW3CCheck = new StringBuilder();
                            try
                            {
                                strW3CCheck.AppendLine("URL" + '\t' + "Valid" + '\t' + "Warnings" + '\t' + "Error");

                                while ((line6 = newfile7.ReadLine()) != null)
                                {
                                    Console.WriteLine("Checking URL : " + line6);
                                    var result = new W3CValidityCheckResult();
                                    result = W3CValidityCheckResult.ReturnsValidHtml(line6);
                                    strW3CCheck.AppendLine(line6 + '\t' + result.IsValid.ToString() + '\t' + result.WarningsCount.ToString() + '\t' + result.ErrorsCount.ToString());
                                }
                            }
                            finally
                            {
                                actionLogFile(logFileW3CChecklog, strW3CCheck.ToString());
                            }
                            writeconsole(0, logFileW3CCheck);
                            var name15 = Console.ReadLine();
                            break;

                        case 15:
                            string Dir = ConfigurationSettings.AppSettings["dir"].ToString();
                            string Filename = ConfigurationSettings.AppSettings["HRefResponseCheck"].ToString();
                            string FileFullPath = Dir + Filename;
                            string[] AllURLS = System.IO.File.ReadAllLines(Filename);
                            foreach (string URL in AllURLS)
                            {
                                var hw = new HtmlWeb();
                                HtmlDocument doc = hw.Load(URL);
                                string input = doc.DocumentNode.OuterHtml;
                                List<string> href = ExtractAllAHrefs(input);
                                checkURLstatusForHrefs(URL, href);
                                logging("*************HRefs Verification for \"" + URL + "\" is Completed*************" + DateTime.Now.ToString() + "\r\n");
                                saveHRefLogsInXls(Dir + "HrefResults.txt");
                            }                        
                            break;

                        case 17:
                            Console.WriteLine("Exiting...Press any key to exit. ");
                            Console.ReadKey();
                            Environment.Exit(0);
                            break;
                    }
                }
                else
                {
                    ErrorMessage();
                }
                if (args.Length == 0)
                {
                    //Console.ReadKey();
                }
                else
                {
                    break;
                }
            }
        }
        private static List<string> ExtractAllAHrefs(string input)
        {
            List<string> hrefs = new List<string>();
            Regex regex = new Regex(" href=\"(.*?)\"");
            MatchCollection match;
            match = regex.Matches(input);
            foreach(Match a in match)
            {
                hrefs.Add(a.ToString());
            }
            return hrefs;
        }
        private static void checkURLstatusForHrefs(string URL, List<string> hrefs)
        {
            string PlatformtoTest = ConfigurationSettings.AppSettings.Get("PlatformtoCheck");
            string strUserAgent = string.Empty;
            string ResponseUri = string.Empty;
            string StatusCode = string.Empty;
            logging("*************HRefs Verification for \"" + URL + "\" is Started*************" + DateTime.Now.ToString() + "\r\n");
            foreach (string str in hrefs)
            {
                    string RefStr = str.Remove(0, 7);
                    RefStr = RefStr.Replace("\"", "");
                    if (RefStr.Contains(".villaplus.com") || RefStr.Contains("www.") || RefStr.Contains("https://") || RefStr.Contains("http://"))
                    {
                        int findindex = RefStr.IndexOf(".com");
                        RefStr = RefStr.Remove(0, findindex + 4);
                    }
                    string URI = "http://" + PlatformtoTest + ".villaplus.com" + RefStr;
                    if (URI.Contains("javascript") || URI.Contains("visa") || URI.Contains("quick") || URI.Contains("news") || URI.Contains(".."))
                    {

                    }
                    else
                    {
                        HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(URI);
                        request.UserAgent = strUserAgent;
                        request.Method = "HEAD";
                        request.AllowAutoRedirect = true;

                        try
                        {
                           HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                            switch ((int)response.StatusCode)
                            {
                                case 200:
                                    {
                                        response.Close();
                                        //HttpWebResponse response1 = (HttpWebResponse)request.GetResponse();
                                        //response1.Close();
                                        
                                        break;
                                    }
                                case 301:
                                    {
                                        response.Close();
                                        //HttpWebRequest request1 = (HttpWebRequest)HttpWebRequest.Create(str);
                                        //request1.AllowAutoRedirect = true;
                                        //HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                                        //response1.Close();
                                        break;
                                    }

                                default:
                                    {
                                        response.Close();
                                        //HttpWebResponse response1 = (HttpWebResponse)request.GetResponse();
                                        //response1.Close();
                                        break;
                                    }
                                    
                            }
                            ResponseUri = response.ResponseUri.ToString();
                            StatusCode = response.StatusCode.ToString();
                            Console.WriteLine(ResponseUri + " " + StatusCode);
                        }

                        catch (WebException ex)
                        {
                            if (ex.Status == (WebExceptionStatus.ProtocolError))
                            {
                                //response = ((HttpWebResponse)ex.Response);

                                //strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + (string)response.StatusDescription + '\t' + "FAIL");
                                ResponseUri = URI;
                                StatusCode = ex.Message.ToString();
                                Console.WriteLine(ResponseUri + StatusCode);
                            }
                            else
                            {
                                //strLog.AppendLine(str + '\t' + ex.Response + '\t' + "FAIL");
                            }
                        }
                       
                        logging(URL, ResponseUri, StatusCode);                        
                    }
                }
        }
        private static void logging(string Msg)
        {
            string strAppPath = ConfigurationSettings.AppSettings["dir"].ToString();
            string logfileName = "HrefResults_" + DateTime.Now.ToString("dd-MM-yyyy");

            FileInfo f = new FileInfo(strAppPath + logfileName + ".txt");
            StreamWriter w = f.AppendText();
            w.WriteLine(Msg);
            w.Close();
        }
        private static void loggingDt(DataTable Msg)
        {
            string strAppPath = ConfigurationSettings.AppSettings["dir"].ToString();
            string logfileName = "IMGCHK_" + DateTime.Now.ToString("dd-MM-yyyy");

            FileInfo f = new FileInfo(strAppPath + logfileName + ".txt");
            StreamWriter w = f.AppendText();
            w.WriteLine(Msg);
            w.Close();
        }
        private static void saveHRefLogsInXls(string logFileforXlsName)
        {
            Excel.Workbook MyBook = null;
            Excel.Application MyApp = null;
            Excel.Worksheet MySheet = null;

            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Add();
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];
                int lastRow = 0;

                System.IO.StreamReader file5 = new System.IO.StreamReader(logFileforXlsName);
                string line = "";
                while ((line = file5.ReadLine()) != null)
                {
                    int j = 1;
                    lastRow += 1;
                    if (!line.Contains("***"))
                    {
                        String[] Lines = line.Split(new char[] { ' ' }, 3);

                        for (int i = 0; i < Lines.Length; i++)
                        {
                            String lineNew = Lines[i];
                            MySheet.Cells[lastRow, j] = lineNew;
                            j = j + 1;
                        }
                    }
                    else
                    {
                        String lineNew = line;
                        MySheet.Cells[lastRow, j] = lineNew;
                        j = j + 1;
                    }
                }

                MyBook.SaveAs(logFileforXlsName.Replace(".log", string.Empty) + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution, true,
                    Missing.Value, Missing.Value, Missing.Value);
                MyBook.Close(null, null, null);
                MyApp.Quit();
            }
            catch (Exception ex)
            {
                MyBook.Close(null, null, null);
                MyApp.Quit();
            }

        }
        private static void logging(string URL, string ResponseUri, string StatusCode)
        {
            string strAppPath = ConfigurationSettings.AppSettings["dir"].ToString();
            string logfileName = "HrefResults_" + DateTime.Now.ToString("dd-MM-yyyy");
            
            FileInfo f = new FileInfo(strAppPath + logfileName + ".txt");
            StreamWriter w = f.AppendText();
            string logMsg = URL + " " + ResponseUri + " " + StatusCode;
            w.WriteLine(logMsg);
            w.Close();
        }

        //private static List<string> ExtractAllImgrefs(string input)
        //{
        //    Hashtable ImageRef = new Hashtable();
        //    List<string> Imgrefs = new List<string>();
        //    Regex regex = new Regex(" href=\"(.*?)\"");
        //    MatchCollection match;
        //    match = regex.Matches(input);
        //    foreach (Match a in match)
        //    {
        //        //ImageRef.Add();
        //        Imgrefs.Add(a.ToString());
        //    }
        //    return Imgrefs;
        //}
        private static string getVillaID(string input)
        {
            string villaID = "";
            string regexvillaID = @"((VillaID[\s]=[\s])(.*?);)";
            Regex exvillaID = new Regex(regexvillaID, RegexOptions.IgnoreCase);
            villaID = exvillaID.Match(input).Value.Trim().ToString();
            string[] villaIDarr = villaID.Split('=');
            villaID = villaIDarr[1];
            villaID = villaID.Replace(";", "").Trim();
            return villaID;
        }
        private static DataTable getImageDataDetails(string url, string villaid, string page)
        {
            DataTable sliderImgData = new DataTable();
            List<string> hrefLst = new List<string> { };
            List<string> idLst = new List<string> { };
            List<string> altLst = new List<string> { };
            List<string> titleLst = new List<string> { };
            List<string> srcLst = new List<string> { };
            List<string> hrefStatus = new List<string> { };

            sliderImgData.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("VILLAID"),
                new DataColumn("URL"),
                new DataColumn("ID"),
                new DataColumn("HREF"),
                new DataColumn("TITLE"),
                new DataColumn("SRC"),
                new DataColumn("ALT"),
                new DataColumn("HREFStatus"),
            });

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(page);

            //HtmlNodeCollection imgnode = htmlDoc.DocumentNode.SelectNodes("//img");

            HtmlNodeCollection imgnode = htmlDoc.DocumentNode.SelectNodes("//div[@class='c-panel']//div[@class='thumbnails']//img");
            HtmlNodeCollection anode = htmlDoc.DocumentNode.SelectNodes("//div[@class='c-panel']//div[@class='thumbnails']//a");
            
            if (anode != null)
            {
                foreach (HtmlNode alinkNode in anode)
                {
                    HtmlAttribute link = alinkNode.Attributes["href"];
                    hrefLst.Add(link == null ? " " : link.Value);
                    string hrefstatus = Program.Ping(link.Value);
                    hrefStatus.Add(hrefstatus);
                }
            }

            if (imgnode != null)
            {
                foreach (HtmlNode linkNode in imgnode)
                {
                    //HtmlAttribute link = linkNode.Attributes["href"];
                    HtmlAttribute id = linkNode.Attributes["id"];
                    HtmlAttribute alt = linkNode.Attributes["alt"];
                    HtmlAttribute title = linkNode.Attributes["title"];
                    HtmlAttribute src = linkNode.Attributes["src"];
                    idLst.Add(id == null ? " " : id.Value);
                    altLst.Add(alt == null ? " " : alt.Value);
                    titleLst.Add(title == null ? " " : title.Value);
                    srcLst.Add(src == null ? " " : src.Value);
                }
                for (int i = 0; i < idLst.Count; i++)
                {
                    sliderImgData.Rows.Add(villaid, url, idLst[i], hrefLst[i], titleLst[i], srcLst[i], altLst[i],hrefStatus[i]);
                }
                foreach (DataRow dr in sliderImgData.Rows)
                {
                    Console.WriteLine(dr[0].ToString() + '\t' +
                        dr[1].ToString() + '\t' +
                        dr[2].ToString() + '\t' +
                        dr[3].ToString() + '\t' +
                        dr[4].ToString() + '\t' +
                        dr[5].ToString() + '\t' +
                        dr[6].ToString() + '\t' +
                        dr[7].ToString()
                        );
                }
            }
            return sliderImgData;
        }
        private static DataTable ActImageDataDetails(string url, string villaid, string page)
    {
            DataTable ActivitiesImgData = new DataTable();
            List<string> Act_idLst = new List<string> { };
            List<string> Act_titleLst = new List<string> { };
            List<string> Act_SrcLst = new List<string> { };
            List<string> Act_AltLst = new List<string> { };

            ActivitiesImgData.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("VILLAID"),
                new DataColumn("URL"),
                new DataColumn("ID"),
                new DataColumn("TITLE"),
                new DataColumn("SRC"),
                new DataColumn("ALT"),
            });

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(page);
            HtmlNodeCollection Actimgnode = htmlDoc.DocumentNode.SelectNodes("//div[@class='WhiteBkgColor dvActivity dvActivityIE7 overflow-auto']//img");
            if (Actimgnode != null)
            {
                foreach (HtmlNode linkNode in Actimgnode)
                {
                    //HtmlAttribute link = linkNode.Attributes["href"];
                    HtmlAttribute id = linkNode.Attributes["id"];
                    HtmlAttribute title = linkNode.Attributes["title"];
                    HtmlAttribute src = linkNode.Attributes["src"];
                    HtmlAttribute alt = linkNode.Attributes["alt"];
                    Act_idLst.Add(id == null ? " " : id.Value);
                    Act_titleLst.Add(title==null ? " " :title.Value);
                    Act_SrcLst.Add(src == null ? " ": src.Value);
                    Act_AltLst.Add(alt == null ? " ": alt.Value);
                }
            }
            for (int i = 0; i < Act_idLst.Count; i++)
            {
                ActivitiesImgData.Rows.Add(villaid, url, Act_idLst[i], Act_titleLst[i], Act_SrcLst[i], Act_AltLst[i]);
            }

            foreach (DataRow dr in ActivitiesImgData.Rows)
            {
                Console.WriteLine(dr[0].ToString() + '\t' +
                    dr[1].ToString() + '\t' +
                    dr[2].ToString() + '\t' +
                    dr[3].ToString() + '\t' +
                    dr[4].ToString() + '\t' +
                    dr[5].ToString() 
                    );
            }
            return ActivitiesImgData;
    }
        private static DataTable ToursDetails(string url, string villaid, string page)
        {
            DataTable ToursData = new DataTable();
            List<string> Tour_SrcLst = new List<string> { };

            ToursData.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("VILLAID"),
                new DataColumn("URL"),
                new DataColumn("SRC"),
            });

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(page);
            HtmlNodeCollection Toursnode = htmlDoc.DocumentNode.SelectNodes("//div[@class='dvTourStatus']//option");
            if (Toursnode != null)
            {
                foreach (HtmlNode linkNode in Toursnode)
                {
                    HtmlAttribute Src = linkNode.Attributes["value"];
                    Tour_SrcLst.Add(Src == null ? " " : Src.Value);
                }
            }
            for (int i = 0; i < Tour_SrcLst.Count; i++)
            {
                ToursData.Rows.Add(villaid, url, Tour_SrcLst[i]);
            }

            foreach (DataRow dr in ToursData.Rows)
            {
                Console.WriteLine(dr[0].ToString() + '\t' +
                    dr[1].ToString() + '\t' +
                    dr[2].ToString()
                   );
            }
            return ToursData;
        }
        private static DataTable FloorPlanDetails(string url, string villaid, string page)
        {
            DataTable FloorPlanData = new DataTable();
            List<string> Flr_idLst = new List<string> { };
            List<string> Flr_titleLst = new List<string> { };
            List<string> Flr_SrcLst = new List<string> { };
            List<string> Flr_AltLst = new List<string> { };

            FloorPlanData.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("VILLAID"),
                new DataColumn("URL"),
                new DataColumn("ID"),
                new DataColumn("TITLE"),
                new DataColumn("SRC"),
                new DataColumn("ALT"),
            });

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(page);
            HtmlNodeCollection FlrPlannode = htmlDoc.DocumentNode.SelectNodes("//div[@id='MainContent_dvFloorPlan']//img");
            if (FlrPlannode != null)
            {
                foreach (HtmlNode linkNode in FlrPlannode)
                {
                    HtmlAttribute id = linkNode.Attributes["ID"];
                    HtmlAttribute title = linkNode.Attributes["title"];
                    HtmlAttribute src = linkNode.Attributes["src"];
                    HtmlAttribute alt = linkNode.Attributes["alt"];
                    Flr_idLst.Add(id == null ? " " : id.Value);
                    Flr_titleLst.Add(title == null ? " " : title.Value);
                    Flr_SrcLst.Add(src == null ? " " : src.Value);
                    Flr_AltLst.Add(alt == null ? " " : alt.Value);
                }
            }
            for (int i = 0; i < Flr_idLst.Count; i++)
            {
                FloorPlanData.Rows.Add(villaid, url, Flr_idLst[i], Flr_titleLst[i], Flr_SrcLst[i], Flr_AltLst[i]);
            }

            foreach (DataRow dr in FloorPlanData.Rows)
            {
                Console.WriteLine(dr[0].ToString() + '\t' +
                    dr[1].ToString() + '\t' +
                    dr[2].ToString() + '\t' +
                    dr[3].ToString() + '\t' +
                    dr[4].ToString() + '\t' +
                    dr[5].ToString()
                   );
            }
            return FloorPlanData;
        }
        private static DataTable PhotostabDetails(string url, string villaid, string page)
        {
            DataTable PhotosData = new DataTable();
            List<string> Ph_hrefLst = new List<string> { };
            List<string> Ph_idLst = new List<string> { };
            List<string> Ph_titleLst = new List<string> { };
            List<string> Ph_SrcLst = new List<string> { };
            List<string> Ph_AltLst = new List<string> { };
            List<string> Ph_dataoriginal = new List<string> { };
            List<string> hrefStatus = new List<string> { };
            
            PhotosData.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("VILLAID"),
                new DataColumn("URL"),
                new DataColumn("HREF"),
                new DataColumn("ID"),
                new DataColumn("TITLE"),
                new DataColumn("SRC"),
                new DataColumn("ALT"),
                new DataColumn("DATAOriginal"),
                new DataColumn("HREFStatus"),

            });

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(page);
            HtmlNodeCollection PhotosAnode = htmlDoc.DocumentNode.SelectNodes("//div[@id='PhotosTab']//a");
            HtmlNodeCollection PhotosIMGnode = htmlDoc.DocumentNode.SelectNodes("//div[@id='PhotosTab']//img ");

            if (PhotosAnode != null)
            {
                foreach (HtmlNode linkNode in PhotosAnode)
                {
                    HtmlAttribute id = linkNode.Attributes["ID"];
                    HtmlAttribute href = linkNode.Attributes["HREF"];
                    HtmlAttribute title = linkNode.Attributes["title"];
                    string hrefstatus = Program.Ping(href.Value);

                    Ph_idLst.Add(id == null ? " " : id.Value);
                    Ph_hrefLst.Add(href == null ? " " : href.Value);
                    Ph_titleLst.Add(title == null ? " " : title.Value);
                    hrefStatus.Add(hrefstatus);  
                }
            }

            if (PhotosIMGnode != null)
            {
                foreach (HtmlNode linkNode in PhotosIMGnode)
                {
                    HtmlAttribute id = linkNode.Attributes["ID"];
                    HtmlAttribute src = linkNode.Attributes["src"];
                    HtmlAttribute alt = linkNode.Attributes["alt"];
                    HtmlAttribute dataoriginal = linkNode.Attributes["data-original"];
                    Ph_SrcLst.Add(src == null ? " " : src.Value);
                    Ph_AltLst.Add(alt == null ? " " : alt.Value);
                    Ph_dataoriginal.Add(dataoriginal == null ? " " : dataoriginal.Value);
                }
            }
            for (int i = 0; i < Ph_idLst.Count; i++)
            {
                PhotosData.Rows.Add(villaid, url, Ph_hrefLst[i], Ph_idLst[i], Ph_titleLst[i], Ph_SrcLst[i], Ph_AltLst[i], Ph_dataoriginal[i], hrefStatus[i]);
            }
                    
            foreach (DataRow dr in PhotosData.Rows)
            {
                Console.WriteLine(dr[0].ToString() + '\t' +
                    dr[1].ToString() + '\t' +
                    dr[2].ToString() + '\t' +
                    dr[3].ToString() + '\t' +
                    dr[4].ToString() + '\t' +
                    dr[5].ToString() + '\t' +
                    dr[6].ToString() + '\t' +
                    dr[7].ToString() + '\t' +
                    dr[8].ToString()
                   );
            }
             
            return PhotosData;
        }
                
        private static void ErrorMessage()
        {
            Console.WriteLine("Typing error, press key to continue.");
        }

        private static Boolean getXMLandSave(string xmlFileName, string dirForxmlFile)
        {

            Stream objStream;
            StreamReader objSR;
            System.Text.Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
            
            string xmlFileNameURL = "http://www.villaplus.com/" + xmlFileName;
            try
            {
                HttpWebRequest wrquest = (HttpWebRequest)WebRequest.Create(xmlFileNameURL);
                HttpWebResponse getresponse = null;
                getresponse = (HttpWebResponse)wrquest.GetResponse();

                objStream = getresponse.GetResponseStream();
                objSR = new StreamReader(objStream, encode, true);
                string strResponse = objSR.ReadToEnd();
                File.WriteAllText(dirForxmlFile + xmlFileName, strResponse, Encoding.UTF8);
                return true;
            }
            catch (WebException ex)
            {
                return false;
            }
        }

        private static void DrawStarLine()
        {
            Console.WriteLine("=================================================================");
        }

        private static void DrawTitle()
        {
            DrawStarLine();
            Console.WriteLine("+++   URL,Page,Tags Verification   +++");
            DrawStarLine();
        }

        private static void DrawMenu(int maxitems)
        {
            DrawStarLine();
            Console.WriteLine(" 1.  PageTitle, MetaDiscription, MetaKeyWord");
            Console.WriteLine(" 2.  HttpResponse");
            Console.WriteLine(" 3.  Google Analytics Tags");
            Console.WriteLine(" 4.  Missing Images [Product Page - Activity Tab]");
            Console.WriteLine(" 5.  PCI URLs Check");
            Console.WriteLine(" 6.  Canonical Tags");
            Console.WriteLine(" 7.  Page View Tags");
            Console.WriteLine(" 8.  No Index No Follow Tags");
            Console.WriteLine(" 9.  UpperCase URLs ");
            Console.WriteLine(" 10. Form Header");
            Console.WriteLine(" 11. User Agents");
            Console.WriteLine(" 12. DoubleClick Floodlight Tag");
            Console.WriteLine(" 13. Criteo Tag");
            Console.WriteLine(" 14. W3C Check");
            Console.WriteLine(" 15. HrefResponseCheck");
            // more here
            Console.WriteLine(" 16. Exit");
            DrawStarLine();
            Console.WriteLine("NOTE : Please update the config file before selecting any option.", maxitems);
            DrawStarLine();
            Console.WriteLine("Make your choice: type 1, 2,... or 13 for exit", maxitems);
            DrawStarLine();
        }

        private static void checkURLstatus(string str, int selector, string strUserAgent)
        {
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(str);
            //request.UserAgent = @"Mozilla/5.0 (Linux; U; Android 4.0.3; ko-kr; LG-L160L Build/IML74K) AppleWebkit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30";
            request.UserAgent = strUserAgent;
            request.Method = "GET";
            //request.AllowAutoRedirect = false;
            request.AllowAutoRedirect = true;
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                switch ((int)response.StatusCode)
                {
                    case 200:
                        {
                            response.Close();
                            //HttpWebResponse response1 = (HttpWebResponse)request.GetResponse();
                            //response1.Close();
                            break;
                        }
                    case 301:
                        {
                            response.Close();
                            //HttpWebRequest request1 = (HttpWebRequest)HttpWebRequest.Create(str);
                            //request1.AllowAutoRedirect = true;
                            //HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                            //response1.Close();
                            break;
                        }

                    default:
                        {
                            response.Close();
                            //HttpWebResponse response1 = (HttpWebResponse)request.GetResponse();
                            //response1.Close();
                            break;
                        }
                }
            }

            catch (WebException ex)
            {
                if (ex.Status == (WebExceptionStatus.ProtocolError))
                {
                    var response = ((HttpWebResponse)ex.Response);

                    //strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + (string)response.StatusDescription + '\t' + "FAIL");

                }
                else
                {
                    //strLog.AppendLine(str + '\t' + ex.Response + '\t' + "FAIL");
                }
            }

        }

        private static void checkURLstatus(string str, string logfile, StringBuilder strLog, int selector)
        {
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //// allows for validation of SSL conversations
            //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(str);
            //request.UserAgent = @"Mozilla/5.0 (Linux; U; Android 4.0.3; ko-kr; LG-L160L Build/IML74K) AppleWebkit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30";
            //request.UserAgent = @"Mozilla/5.0 (Linux; U; Android 2.2; en-gb; GT-P1000 Build/FROYO) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1";
            request.Method = "GET";
            request.AllowAutoRedirect = false;
            //request.AllowAutoRedirect = true;
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                switch ((int)response.StatusCode)
                {
                    case 200:
                        {
                            response.Close();
                            HttpWebResponse response1 = (HttpWebResponse)request.GetResponse();
                            response1.Close();
                            if (selector == 9)
                            {
                                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc = web.Load(str);
                                List<string> hrefTagsLst = ExtractAllAHrefTags(doc);
                                string hrefTagsStr = string.Join("/\n", hrefTagsLst.ToArray());
                                strLog.AppendLine("####### Main URL : " + str);
                                strLog.AppendLine("Upper Case URLs - ");
                                strLog.AppendLine(hrefTagsStr);
                            }
                            else if (selector == 1 || selector == 3 || selector == 4 || selector == 6 || selector == 7 || selector == 8 || selector == 10 || selector == 12 || selector == 13)
                            {
                                chkTitleMtDescMtKeywGA(str, (int)response.StatusCode, response.StatusCode.ToString(), logfile, strLog, selector);
                            }
                            else
                            {
                                strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + response.StatusCode.ToString() + '\t' + response1.ResponseUri.ToString());
                            }
                            break;
                        }
                    case 301:
                        {
                            response.Close();
                            HttpWebRequest request1 = (HttpWebRequest)HttpWebRequest.Create(str);
                            request1.AllowAutoRedirect = true;
                            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                            response1.Close();
                            if (selector == 9)
                            {
                                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                                HtmlAgilityPack.HtmlDocument doc = web.Load(str);
                                List<string> hrefTagsLst = ExtractAllAHrefTags(doc);
                                string hrefTagsStr = string.Join("/\n", hrefTagsLst.ToArray());
                                strLog.AppendLine("####### Main URL : " + str);
                                strLog.AppendLine("Upper Case URLs - ");
                                strLog.AppendLine(hrefTagsStr);
                            }
                            if (selector == 1 || selector == 3 || selector == 4 || selector == 6 || selector == 7 || selector == 8 || selector == 10 || selector == 12 || selector == 13)
                            {
                                str = response1.ResponseUri.ToString();
                                chkTitleMtDescMtKeywGA(str, (int)response.StatusCode, response.StatusCode.ToString(), logfile, strLog, selector);
                            }
                            else
                            {
                                strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + response.StatusCode.ToString() + '\t' + response1.ResponseUri.ToString());

                            }
                            break;
                        }

                    default:
                        {
                            response.Close();
                            HttpWebRequest request1 = (HttpWebRequest)HttpWebRequest.Create(str);
                            request1.AllowAutoRedirect = true;
                            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                            response1.Close();
                            if (selector == 1 || selector == 6 || selector == 7 || selector == 8 || selector == 10)
                            {
                                chkTitleMtDescMtKeywGA(str, (int)response.StatusCode, response.StatusCode.ToString(), logfile, strLog, selector);
                            }
                            else
                            {
                                strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + response.StatusCode.ToString() + '\t' + response1.ResponseUri.ToString());
                            }
                            break;
                        }
                }
            }

            catch (WebException ex)
            {
                if (ex.Status == (WebExceptionStatus.ProtocolError))
                {
                    var response = ((HttpWebResponse)ex.Response);

                    strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + (string)response.StatusDescription + '\t' + "FAIL");

                }
                else
                {
                    strLog.AppendLine(str + '\t' + ex.Response + '\t' + "FAIL");
                }
            }

        }

        private static void chkTitleMtDescMtKeywGA(string str, int responseCode, string responseStr, string logfile, StringBuilder strLog, int selector)
        {

            string page = string.Empty;
            string pageTitle = string.Empty;
            string Message = string.Empty;
            string gaContent = string.Empty;
            string gatrccode = string.Empty;
            string cnContent = string.Empty;
            string formHeaderContent = string.Empty;
            string dfaTagContent = string.Empty;
            string criteoTagContent = string.Empty;

            string mappedPageTitle = string.Empty;
            string mappedMetaDesc = string.Empty;
            string mappedMetaKw_Pri = string.Empty;
            string mappedMetaKw_Sec = string.Empty;
            string mappedMetaKw_Trt = string.Empty;
            string mappedHeader = string.Empty;

            string line = string.Empty;
            string metaDescription = string.Empty;
            string metaKeyword = string.Empty;
            string header = string.Empty;

            string regexTitle = string.Empty;
            string regexMetadesc = string.Empty;
            string regexMetakw = string.Empty;
            string regexheader = string.Empty;
            string regexGA = string.Empty;
            string regexGAtrccode = string.Empty;
            string regexCN = string.Empty;
            string regexFormHeader = string.Empty;
            string regexDfaTag = string.Empty;
            string regexDfaTagOld = string.Empty;
            string regexCriteoTag = string.Empty;

            string regexPP = string.Empty;
            string ppContent = string.Empty;
            string ppStatus = string.Empty;
            string pvContent = string.Empty;
            string pvStatus = string.Empty;
            string pvURLStatus = string.Empty;

            string pageTitleStatus = "NOT MAPPED";
            string metaDescStatus = string.Empty;
            string metaKwPriStatus = string.Empty;
            string metaKwSecStatus = string.Empty;
            string metaKwTrtStatus = string.Empty;
            string headerStatus = string.Empty;
            string gaStatus = string.Empty;
            string gatrccodeStatus = string.Empty;
            string cnStatus = string.Empty;
            string cnURLStatus = string.Empty;
            string regexPV = string.Empty;
            string regexIF = string.Empty;
            string ifContent = string.Empty;
            string ifStatus = string.Empty;
            // string exPV = string.Empty; 

            string urlCanonicalMatch = "FAIL";
            string regexAT = string.Empty;
            string cnAlternateTagStatus = string.Empty;
            string urlCanonicalALTMatch = "FAIL";
            string cnaltContent = string.Empty;
            string urlToCheckREp = string.Empty;
            string strToreplace = string.Empty;



            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(str);
            //request.UserAgent = @"Mozilla/5.0 (Linux; U; Android 4.0.3; ko-kr; LG-L160L Build/IML74K) AppleWebkit/534.30 (KHTML, like Gecko) Version/4.0 Mobile Safari/534.30";
            request.Method = "GET";
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                string URLheader = response.Headers.ToString();
                regexTitle = @"(?<=<title.*>)([\s\S]*)(?=</title>)";
                regexMetadesc = @"<meta[\s]+[^>]*?name[\s]?=[\s""']+(description)[\s""']+content[\s]?=[\s""']+(.*?)[""']+.*?>";
                regexMetakw = @"<meta[\s]+[^>]*?name[\s]?=[\s""']+(keywords)[\s""']+content[\s]?=[\s""']+(.*?)[""']+.*?>";
                //regexheader = @"(?<=<h1>)([\s\S]*)(?=</h1>)";
                regexheader = @"(?<=<h1.*>)([\s\S]*)(?=</h1>)";



                if (response.Headers["Content-Type"].StartsWith("text/html"))
                {
                    // Download the page
                    WebClient web = new WebClient();
                    web.UseDefaultCredentials = true;
                    page = null;
                    page = web.DownloadString(str);
                    web.Dispose();

                    //Uncomment to test from local file source
                    //System.IO.StreamReader file12 = new System.IO.StreamReader("E:\\00_Work\\01_QA\\04_Test Automation\\FrmShailesh\\HomePageSource.txt");
                    //page = file12.ReadToEnd().ToString();

                    if (selector == 13)
                    {
                        regexCriteoTag = @"<script[\s]src[\s]?=[\s""]/scripts/criteo_ld.js[\s""][\s]type[\s]?=[\s""]text/javascript[\s""]></script>";
                        Regex exCriteoTag = new Regex(regexCriteoTag, RegexOptions.IgnoreCase);
                        criteoTagContent = exCriteoTag.Match(page).Value.Trim().ToString();

                        if (criteoTagContent == "")
                        {
                            criteoTagContent = "NOT FOUND";
                        }
                        else
                        {
                            criteoTagContent = criteoTagContent.Replace("\n", string.Empty);
                            criteoTagContent = criteoTagContent.Replace("\r", string.Empty);
                            criteoTagContent = criteoTagContent.Replace("\t", string.Empty);
                        }

                        strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + criteoTagContent);
                    }

                    else if (selector == 12)
                    {
                        regexDfaTag = @"<iframe[\s]+src=[\s""']+(http:)+(//3702205.fls.doubleclick.net/activityi)+([\s\S]*)</iframe>";
                        regexDfaTagOld = @"<iframe[\s]+src=[\s""']+(http:)+(//fls.doubleclick.net/activityi)+([\s\S]*)</iframe>";
                        Regex exDfaTag = new Regex(regexDfaTag, RegexOptions.IgnoreCase);
                        Regex exDfaTagOld = new Regex(regexDfaTagOld, RegexOptions.IgnoreCase);
                        dfaTagContent = exDfaTag.Match(page).Value.Trim().ToString();

                        if (dfaTagContent == "")
                        {
                            dfaTagContent = exDfaTagOld.Match(page).Value.Trim().ToString();
                            if (dfaTagContent == "")
                            {
                                dfaTagContent = "NOT FOUND";
                            }
                            else
                            {
                                dfaTagContent = dfaTagContent.Replace("\n", string.Empty);
                                dfaTagContent = dfaTagContent.Replace("\r", string.Empty);
                                dfaTagContent = dfaTagContent.Replace("\t", string.Empty);
                            }
                        }
                        strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + dfaTagContent);
                    }

                    else if (selector == 10)
                    {
                        //regexFormHeader = @"<form[\s]+[^>]*?method[\s]?=[\s""']+(post)[\s""']+action[\s]?=[\s""']+Headerpage.aspx[\s]+(.*?)+.*?>";
                        //regexFormHeader = @"<form[\s]+method[\s]?=[\s""']+(post)[\s""'][\s]+action[\s]?=[\s""']+Headerpage.aspx(.*?)>";
                        //regexFormHeader = @"<form[\s]+method[\s]?=[\s""']+(post)[\s""'][\s]+action[\s]?=[\s""']+(.*?)>";
                        regexFormHeader = @"aa=nitin";
                        Regex exFormHeader = new Regex(regexFormHeader, RegexOptions.IgnoreCase);
                        formHeaderContent = exFormHeader.Match(page).Value.Trim().ToString();

                        strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + formHeaderContent);
                    }

                    else if (selector == 3)
                    {
                        regexGA = "'stats\\.g\\.doubleclick\\.net\\/dc\\.js";
                        Regex exGA = new Regex(regexGA, RegexOptions.IgnoreCase);
                        gaContent = exGA.Match(page).Value.Trim().ToString();

                        if (gaContent != "" || gaContent != null)
                        {
                            if (gaContent.Contains("dc.js"))
                            {
                                gaStatus = "FOUND";
                            }
                        }
                        else
                        {
                            gaStatus = "NOT FOUND";
                        }

                        strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + gaStatus);
                    }

                    else if (selector == 6)
                    {
                        regexCN = @"<link[\s]+rel[\s]?=[\s""]+(canonical)[\s""]+href[\s]?=.*?/>";
                        regexAT = @"<link[\s]rel[\s]?=[\s""](alternate)[\s""]+media[\s]?=[\s""]only screen and \(max-width: 640px\)[\""]+ href[\s]?=.*?/>";
                        Regex exCN = new Regex(regexCN, RegexOptions.IgnoreCase);
                        Regex exALT = new Regex(regexAT, RegexOptions.IgnoreCase);
                        cnContent = exCN.Match(page).Value.Trim().ToString();
                        cnaltContent = exALT.Match(page).Value.Trim().ToString();

                        string urlToCheck = "";
                        if (cnContent != "" || cnContent != null)
                        {
                            if (cnContent.Contains("canonical"))
                            {
                                cnStatus = "FOUND";
                                int inx = cnContent.IndexOf("href=");
                                cnContent = cnContent.Substring(inx, cnContent.Length - inx);
                                cnContent = cnContent.Replace("/>", "");
                                cnContent = cnContent.Replace("href=", "");
                                cnContent = cnContent.Replace("\"", "");
                                cnContent = cnContent.TrimEnd();
                                int newIndx = response.ResponseUri.ToString().LastIndexOf('/');
                                urlToCheck = response.ResponseUri.ToString();
                                if (cnContent == urlToCheck)
                                {
                                    urlCanonicalMatch = "PASS";
                                }
                            }

                            else
                            {
                                cnStatus = "NOT FOUND";
                                
                            }

                            if (cnaltContent != "" || cnaltContent != null)
                            {
                                if (cnaltContent.Contains("alternate"))
                                {
                                    cnAlternateTagStatus = "FOUND";
                                    cnaltContent = cnaltContent.Replace("\n", string.Empty);
                                    int inxAlt = cnaltContent.IndexOf("href=");
                                    cnaltContent = cnaltContent.Substring(inxAlt);
                                    cnaltContent = cnaltContent.Replace("/>", "");
                                    cnaltContent = cnaltContent.Replace("href=", "");
                                    cnaltContent = cnaltContent.Replace("\"", "");
                                    cnaltContent = cnaltContent.TrimEnd();
                                    int newIndxAlt = response.ResponseUri.ToString().LastIndexOf('/');
                                    urlToCheck = response.ResponseUri.ToString();
                                    strToreplace = urlToCheck.Substring(0, urlToCheck.IndexOf('.'));
                                    urlToCheckREp = urlToCheck.Replace(strToreplace, "http://m");
                                    if (cnaltContent == urlToCheckREp)
                                    {
                                        urlCanonicalALTMatch = "PASS";
                                    }

                                }

                                else
                                {
                                    
                                    cnAlternateTagStatus = "NOT FOUND";
                                }
                            }

                            strLog.AppendLine(str + '\t' + cnStatus + '\t' + cnContent + '\t' + urlToCheck + '\t' + urlCanonicalMatch + '\t' + cnAlternateTagStatus + '\t' + urlToCheckREp + '\t' + urlCanonicalALTMatch);
                        }
                    }

                
                    else if (selector == 8)
                    {
                        //regexIF = @"<meta[\s]name[\s]?=[\s""]((?i)robots)[\s""]+content[\s]?=[\s""](noindex, nofollow, noarchive)[\s""] />";
                        regexIF = @"<meta[\s]name[\s]?=[\s""](robots)[\s""]+content[\s]?=.*?/>";
                        Regex exIF = new Regex(regexIF, RegexOptions.IgnoreCase);
                        ifContent = exIF.Match(page).Value.Trim().ToString();

                        if (ifContent != "" || ifContent != null)
                        {
                            if (ifContent.Contains("noindex"))
                            {
                                if (ifContent.Contains("nofollow"))
                                {
                                    if (ifContent.Contains("noarchive"))
                                    {
                                        ifStatus = "FOUND";
                                    }
                                }
                            }
                            else
                            {
                                ifStatus = "NOT FOUND";
                            }

                        }
                        strLog.AppendLine(str + '\t' + ifStatus);
                    }


                    else if (selector == 7)
                    {
                        string[] regexPVT = new string[3];
                        regexPVT[0] = @"<script[\s]type[\s]?=[\s""]text/javascript[\s""]>cmCreatePageviewTag(.*?);</script>";
                        regexPVT[1] = @"<script[\s]type[\s]?=[\s""]text/javascript[\s""]>cmSetClientID(.*?);</script>";
                        regexPVT[2] = @"<script[\s]type[\s]?=[\s""]text/javascript[\s""][\s]src[\s]?=[\s""]//libs.coremetrics.com/eluminate.js[\s""]></script>";

                        pvStatus = CheckTag(regexPVT[0], "cmCreatePageviewTag", str);
                        pvStatus += CheckTag(regexPVT[1], "cmSetClientID", str);
                        pvStatus += CheckTag(regexPVT[2], "eluminate.js", str);
                        strLog.AppendLine(str + '\t' + pvStatus + '\t');
                    }

                    else if (selector == 4)
                    {
                        regexPP = @"(?<=<div.*id=[\s""']+(ActivitiesTab)[\s""'].*?>)([\s\S]*)(?=<div.*id=[\s""']+(ReviewsTab)[\s""'].*?>)";
                        Regex exPP = new Regex(regexPP, RegexOptions.IgnoreCase);
                        ppContent = exPP.Match(page).Value.Trim().ToString();

                        if (ppContent != "" || ppContent != null)
                        {
                            ppStatus = "OK";
                            if (ppContent.Contains("noImageS.gif"))
                            {
                                ppStatus = "MISSING";
                            }
                        }

                        if (ppStatus == "MISSING")
                        {
                            strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + ppStatus + '\t' + ppContent.Replace("\r", string.Empty).Replace("\n", string.Empty));
                        }
                        else
                        {
                            strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + ppStatus);
                        }

                    }
                    else
                    {
                        // Extract the title
                        Regex exTitle = new Regex(regexTitle, RegexOptions.IgnoreCase);
                        pageTitle = exTitle.Match(page).Value.Trim().ToString();
                        if (pageTitle != "" || pageTitle != null)
                        {
                            if (pageTitle.Contains("</title>"))
                            {
                                pageTitle = pageTitle.Substring(0, pageTitle.IndexOf("</title>"));
                            }
                        }
                        else
                        {
                            pageTitle = "NOT FOUND";
                        }

                        Regex exMetadesc = new Regex(regexMetadesc, RegexOptions.IgnoreCase);
                        metaDescription = exMetadesc.Match(page).Value.Trim().ToString();
                        Regex exMetakw = new Regex(regexMetakw, RegexOptions.IgnoreCase);
                        metaKeyword = exMetakw.Match(page).Value.Trim().ToString();
                        Regex exheader = new Regex(regexheader, RegexOptions.IgnoreCase);
                        header = exheader.Match(page).Value.Trim().ToString();

                        if (metaDescription.Contains("content"))
                        {
                            if (metaDescription != string.Empty)
                            {
                                metaDescription = metaDescription.Substring(metaDescription.IndexOf("content="), metaDescription.Length - metaDescription.IndexOf("content=") - 2);
                                string[] metaDescriptionArr = metaDescription.Split('=');
                                metaDescription = metaDescriptionArr[1];
                                metaDescription = metaDescription.TrimStart('"');
                                metaDescription = metaDescription.Remove(metaDescription.Length - 2);
                            }
                            else
                            {
                                metaDescription = "NOT FOUND";
                            }
                        }
                        else
                        {
                            metaDescription = "NOT FOUND";
                        }

                        if (metaKeyword.Contains("content"))
                        {
                            if (metaKeyword != string.Empty)
                            {
                                metaKeyword = metaKeyword.Substring(metaKeyword.IndexOf("content="), metaKeyword.Length - metaKeyword.IndexOf("content=") - 2);
                                string[] metaKeywordArr = metaKeyword.Split('=');
                                metaKeyword = metaKeywordArr[1];
                            }
                            else
                            {
                                metaKeyword = "NOT FOUND";
                            }
                        }
                        else
                        {
                            metaKeyword = "NOT FOUND";
                        }

                        if (header != string.Empty)
                        {
                            if (header.Contains("span id"))
                            {
                                string[] headerArr1 = header.Split('>');
                                header = headerArr1[1];
                                string[] headerArr2 = header.Split('<');
                                header = headerArr2[0];
                            }
                            if (header.Contains("ctl00_pageHeading"))
                            {

                            }
                        }
                        else
                        {
                            header = "NOT FOUND";
                        }
                        if (pageTitle != "NOT FOUND")
                        {
                            string URLTitleMappingFile = ConfigurationSettings.AppSettings["urlTitleMappingfile"].ToString();
                            System.IO.StreamReader file = new System.IO.StreamReader(URLTitleMappingFile);
                            while ((line = file.ReadLine()) != null)
                            {
                                if (line.Contains(response.ResponseUri.ToString()))
                                {
                                    string[] arr = line.Split('\t');
                                    if (arr[0] == response.ResponseUri.ToString())
                                    {
                                        mappedPageTitle = arr[1];
                                        mappedMetaDesc = arr[2];
                                        mappedMetaKw_Pri = arr[3];
                                        mappedMetaKw_Sec = arr[4];
                                        mappedMetaKw_Trt = arr[5];
                                        mappedMetaKw_Trt = arr[5];
                                        mappedHeader = arr[6];
                                        break;
                                    }
                                }
                            }
                            file.Close();

                            if (mappedPageTitle != string.Empty)
                            {
                                pageTitleStatus = "FAIL";
                                if (pageTitle.Trim() == mappedPageTitle.Trim())
                                {
                                    pageTitleStatus = "PASS";
                                }
                            }

                            if (metaDescription != "NOT FOUND")
                            {
                                metaDescStatus = "NOT MAPPED";
                                if (mappedMetaDesc != string.Empty)
                                {
                                    metaDescStatus = "FAIL";
                                    if (metaDescription.Trim() == mappedMetaDesc.Trim())
                                    {
                                        metaDescStatus = "PASS";
                                    }
                                }
                            }

                            if (metaKeyword != "NOT FOUND")
                            {
                                metaKwPriStatus = "NOT MAPPED";
                                metaKwSecStatus = "NOT MAPPED";
                                metaKwTrtStatus = "NOT MAPPED";
                                if (mappedMetaKw_Pri != string.Empty)
                                {
                                    metaKwPriStatus = "FAIL";
                                    if (metaKeyword.Trim().Contains(mappedMetaKw_Pri.Trim()))
                                    {
                                        metaKwPriStatus = "PASS";
                                    }
                                }

                                if (mappedMetaKw_Sec != string.Empty)
                                {
                                    metaKwSecStatus = "FAIL";
                                    if (metaKeyword.Trim().Contains(mappedMetaKw_Sec.Trim()))
                                    {
                                        metaKwSecStatus = "PASS";
                                    }
                                }

                                if (mappedMetaKw_Trt != string.Empty)
                                {
                                    metaKwTrtStatus = "FAIL";
                                    if (metaKeyword.Trim().Contains(mappedMetaKw_Trt.Trim()))
                                    {
                                        metaKwTrtStatus = "PASS";
                                    }
                                }
                            }

                            if (header != "NOT FOUND")
                            {
                                headerStatus = "FAIL";
                                if (mappedHeader.Trim() == header.Trim())
                                {
                                    headerStatus = "PASS";
                                }
                            }
                            strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + pageTitle + '\t' + mappedPageTitle + '\t' + pageTitleStatus + '\t' + metaDescription + '\t' + mappedMetaDesc + '\t' + metaDescStatus + '\t' + metaKeyword + '\t' + mappedMetaKw_Pri + '\t' + mappedMetaKw_Sec + '\t' + mappedMetaKw_Trt + '\t' + metaKwPriStatus + '\t' + metaKwSecStatus + '\t' + metaKwTrtStatus + '\t' + header + '\t' + mappedHeader + '\t' + headerStatus);
                        }
                        else
                        {
                            pageTitleStatus = "NOT VERIFIED";
                            metaDescStatus = "NOT VERIFIED";
                            metaKwPriStatus = "NOT VERIFIED";
                            metaKwSecStatus = "NOT VERIFIED";
                            metaKwTrtStatus = "NOT VERIFIED";
                            strLog.AppendLine(str + '\t' + responseCode + '\t' + responseStr + '\t' + response.ResponseUri.ToString() + '\t' + pageTitle + '\t' + mappedPageTitle + '\t' + pageTitleStatus + '\t' + metaDescription + '\t' + mappedMetaDesc + '\t' + metaDescStatus + '\t' + metaKeyword + '\t' + mappedMetaKw_Pri + '\t' + mappedMetaKw_Sec + '\t' + mappedMetaKw_Trt + '\t' + metaKwPriStatus + '\t' + metaKwSecStatus + '\t' + metaKwTrtStatus);
                        }
                    }
                }
            }


        }

        private static string CheckTag(string regexPVT, string tagName, string str)
        {
            string pvContent = string.Empty;
            string pvStatus = string.Empty;
            string page = string.Empty;
            WebClient web = new WebClient();
            web.UseDefaultCredentials = true;
            page = null;
            page = web.DownloadString(str);
            web.Dispose();

            Regex exPV = new Regex(regexPVT, RegexOptions.IgnoreCase);
            pvContent = exPV.Match(page).Value.Trim().ToString();

            if (pvContent != "" || pvContent != null)
            {
                if (pvContent.Contains(tagName))
                {
                    pvStatus = tagName + " PRESENT\t";

                }

                else
                {
                    pvStatus = tagName + " NOT PRESENT\t";
                }
            }
            return pvStatus;

        }

        private static void writeconsole(int startend, string logfilename)
        {
            if (startend == 1)
            {
                Console.WriteLine("--------------------------------------------");
                Console.WriteLine("Log File Name : " + logfilename);
                Console.WriteLine("--------------------------------------------");
                Console.WriteLine("Test Started : " + DateTime.Now.ToString());
                Console.WriteLine("--------------------------------------------");
            }
            else
            {
                Console.WriteLine("--------------------------------------------");
                Console.WriteLine("Test Complete : " + DateTime.Now.ToString());
                Console.WriteLine("--------------------------------------------");
                Console.Write("Press any key to close...");
                Console.Out.Flush();
            }
        }

        private static string initiateLogFile(int selector)
        {
            string initiateLogFile = ConfigurationSettings.AppSettings.Get("logfile") + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".log";
            string initiateXLFile = ConfigurationSettings.AppSettings.Get("logfile") + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") + ".xls";
            //string initiateLogFile = ConfigurationSettings.AppSettings.Get("dir") + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".log";
            //string initiateXLFile = ConfigurationSettings.AppSettings.Get("dir") + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") + ".xls";
            string dirForLogFile = ConfigurationSettings.AppSettings.Get("dir");
            if (selector == 1)
            {
                initiateLogFile = dirForLogFile + "SEO_" + initiateLogFile;
            }
            else if (selector == 2)
            {
                initiateLogFile = dirForLogFile + "URLMapping_" + initiateLogFile;
            }
            else if (selector == 3)
            {
                initiateLogFile = dirForLogFile + "GA_" + initiateLogFile;
            }
            else if (selector == 4)
            {
                initiateLogFile = dirForLogFile + "PPURLs_" + initiateLogFile;
            }
            else if (selector == 5)
            {
                initiateLogFile = dirForLogFile + "PCIURLs_" + initiateLogFile;
            }
            else if (selector == 6)
            {
                initiateLogFile = dirForLogFile + "CanonicalTags_" + initiateLogFile;
            }
            else if (selector == 7)
            {
                initiateLogFile = dirForLogFile + "PageViewTags_" + initiateLogFile;
            }
            else if (selector == 8)
            {
                initiateLogFile = dirForLogFile + "NoIndexNoFollowTags_" + initiateLogFile;
            }
            else if (selector == 9)
            {
                initiateLogFile = dirForLogFile + "CheckUpperCaseUrls_" + initiateLogFile;
            }
            else if (selector == 10)
            {
                initiateLogFile = dirForLogFile + "FormHeaderStatus_" + initiateLogFile;
            }
            else if (selector == 11)
            {
                initiateLogFile = dirForLogFile + "UserAgentStatus_" + initiateLogFile;
            }
            else if (selector == 12)
            {
                initiateLogFile = dirForLogFile + "DFATagStatus_" + initiateLogFile;
            }
            //else if (selector == 12)
            //{
            //    initiateLogFile = dirForLogFile + "PagehrefStatus_" + initiateLogFile;
            //}
            else if (selector == 13)
            {
                initiateLogFile = dirForLogFile + "CriteoTagStatus_" + initiateLogFile;
            }
            else if (selector == 14)
            {
                initiateLogFile = dirForLogFile + "W3CCheckfStatus_" + initiateLogFile;
            }
            else if (selector == 15)
            {
                initiateLogFile = dirForLogFile + "HrefsStatus_" + initiateLogFile;
            }
            else if (selector == 16)
            {
                initiateLogFile = dirForLogFile + "ImgVillapluschk_" + initiateXLFile;
            }
            return initiateLogFile;
        }

        private static string initiateLogFile(int selector, string platform, string whichXML)
        {
            string initiateLogFile = ConfigurationSettings.AppSettings.Get("logfile") + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".log";
            string dirForLogFile = ConfigurationSettings.AppSettings.Get("dir");
            if (selector == 2)
            {
                if (whichXML.Contains("product"))
                {
                    initiateLogFile = dirForLogFile + platform + "_ProductmapURLsCheck_" + initiateLogFile;
                }
                if (whichXML.Contains("sitemap"))
                {
                    initiateLogFile = dirForLogFile + platform + "_SitemapURLsCheck_" + initiateLogFile;
                }

            }
            return initiateLogFile;
        }

        private static void actionLogFile(StreamWriter genericLogFile, string file)
        {
            genericLogFile.WriteLine(file);
            genericLogFile.Close();
            genericLogFile.Dispose();

        }

        private static List<string> ExtractAllAHrefTags(HtmlAgilityPack.HtmlDocument htmlSnippet)
        {
            List<string> hrefTags = new List<string>();

            foreach (HtmlAgilityPack.HtmlNode link in htmlSnippet.DocumentNode.SelectNodes("//a[@href]"))
            {
                HtmlAgilityPack.HtmlAttribute att = link.Attributes["href"];
                if (att.Value.StartsWith("/") || att.Value.Contains("villaplus"))
                {
                    bool check = hasUpperCase(att.Value.ToString());
                    if (check == false)
                    {
                        hrefTags.Add(att.Value + "-UPPERCASE ");
                    }
                    else
                    {
                        hrefTags.Add(att.Value);
                    }
                }
            }

            return hrefTags;
        }

        private static bool hasUpperCase(string str)
        {
            if (string.IsNullOrEmpty(str))
                return false;
            for (int i = 0; i < str.Length; i++)
            {
                if (char.IsLower(str[i]))
                    return true;
            }
            return false;
        }

        private static DataTable getUrlStatusStats(string logFileURLMapping, string platform, DataTable dtStatusStatsTable)
        {

            string delimiter = "\t";
            DataTable dtResultTable = new DataTable();
            StreamReader s = new StreamReader(logFileURLMapping);
            string[] columns = s.ReadLine().Split(delimiter.ToCharArray());
            foreach (string col in columns)
            {
                bool added = false;
                string next = "";
                int i = 0;
                while (!added)
                {
                    string columnname = col + next;
                    columnname = columnname.Replace("#", "");
                    columnname = columnname.Replace("'", "");
                    columnname = columnname.Replace("&", "");

                    if (!dtResultTable.Columns.Contains(columnname))
                    {
                        dtResultTable.Columns.Add(columnname);
                        added = true;
                    }
                    else
                    {
                        i++;
                        next = "_" + i.ToString();
                    }
                }
            }

            string AllData = s.ReadToEnd();

            string[] rows = AllData.Split("\r\n".ToCharArray());

            foreach (string r in rows)
            {
                if( r != "")
                {
                    string[] items = r.Split(delimiter.ToCharArray());
                    dtResultTable.Rows.Add(items);
                }
            }

            var distinctResultStatus = (from row in dtResultTable.AsEnumerable()
                                        select row.Field<string>("Result")).Distinct();

            var distinctStatus = (from row in dtResultTable.AsEnumerable()
                                 select row.Field<string>("Status")).Distinct();

            foreach (var name in distinctStatus) 
            {
                if (name != "")
                {
                    var cnt = dtResultTable
                                .AsEnumerable()
                                .Where(p => p.Field<string>("Status") == name)
                                .Count();
                    string[] statusStats = { platform, name, cnt.ToString() };
                    dtStatusStatsTable.Rows.Add(statusStats);
                }
            }

            foreach (var name in distinctResultStatus)
            {
                if (name == "FAIL")
                {
                    var cnt = dtResultTable
                                .AsEnumerable()
                                .Where(p => p.Field<string>("Result") == name)
                                .Count();
                    string[] statusStats = { platform, name, cnt.ToString() };
                    dtStatusStatsTable.Rows.Add(statusStats);
                }
            }

            return dtStatusStatsTable;
        }

        private static string createHTMLReport(DataTable dtStatusStatsTable, Boolean successXML, string whichXML, int count)
        {
            string htmlReportFileName = "htmlReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm") + ".html";
            string dirForLogFile = ConfigurationSettings.AppSettings.Get("dir");
            htmlReportFileName = dirForLogFile + htmlReportFileName;

            StringBuilder sb = new StringBuilder();

            sb = topHTML(sb, successXML, whichXML, count);
            //sb = middleHTML(sb, "Desktop");

            bool addMiddleHTMLforD = true;
            bool addMiddleHTMLforT = true;
            bool addMiddleHTMLforM = true;

            foreach (DataRow row in dtStatusStatsTable.Rows) // Loop over the rows.
            {
                if(row.ItemArray[0].ToString() == "D")
                {
                    if (addMiddleHTMLforD)
                    {
                        sb = middleHTML(sb, "Desktop");
                        addMiddleHTMLforD = false;
                    }
                    sb.AppendLine("<tr><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[1] + "</td><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[2] + "</td></tr>");
                }
            }

            sb = middleStartHTML(sb);
            //sb = middleHTML(sb, "Tablet");

            foreach (DataRow row in dtStatusStatsTable.Rows) // Loop over the rows.
            {
                if (row.ItemArray[0].ToString() == "T")
                {
                    if (addMiddleHTMLforT)
                    {
                        sb = middleHTML(sb, "Tablet");
                        addMiddleHTMLforT = false;
                    }
                    sb.AppendLine("<tr><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[1] + "</td><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[2] + "</td></tr>");
                }
            }

            sb = middleStartHTML(sb);
            //sb = middleHTML(sb, "Mobile");

            foreach (DataRow row in dtStatusStatsTable.Rows) // Loop over the rows.
            {
                if (row.ItemArray[0].ToString() == "M")
                {
                    if (addMiddleHTMLforM)
                    {
                        sb = middleHTML(sb, "Mobile");
                        addMiddleHTMLforM = false;
                    }
                    sb.AppendLine("<tr><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray; font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[1] + "</td><td bgcolor=\"#FFFFFF\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-size: 12px;'>" + row.ItemArray[2] + "</td></tr>");
                }
            }

            sb = endHTML(sb);

            StreamWriter sw = new StreamWriter(htmlReportFileName , true);
            sw.Write(sb.ToString());
            sw.Close();

            return htmlReportFileName;
        }

        private static StringBuilder middleHTML(StringBuilder sb, string platform)
        {
            sb.AppendLine("<table width=\"175\" border='0' cellpadding=\"0\" cellspacing=\"0\">");
            sb.AppendLine("<tr>");
            sb.AppendLine("<td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;font-weight: bold;'>" + platform + " URL Status</td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<table width=\"300\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" bordercolor=\"#000066\" bgcolor=\"#000066\">");
            sb.AppendLine("<tr>");
            sb.AppendLine("<td align=\"left\" valign=\"top\"><table width=\"300\" border='0' cellpadding=\"3\" cellspacing=\"1\">");
            sb.AppendLine("<tr>");
            sb.AppendLine("<td width=\"146\" height=\"0\" bgcolor=\"#CFF1FE\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 12px;'>HTTP Status Code</td>");
            sb.AppendLine("<td width=\"139\" height=\"0\" bgcolor=\"#CFF1FE\" style = 'border: 1px solid gray;font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 12px;'>Number of Occurrences</td>");
            sb.AppendLine("</tr>");

            return sb;
        }

        private static StringBuilder topHTML(StringBuilder sb, Boolean successXML, string whichXML, int count)
        {
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table border='0' cellpadding=\"0\" cellspacing=\"0\" style=\"width: 600px\">");
            sb.AppendLine("<tr><td style = 'color: #000000;font-family: Arial,Helvetica,sans-serif;font-size: 20px;font-weight: bold;'>Website Sitemap URLs Status Report </td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<br>");
            sb.AppendLine("<table border='0' cellpadding=\"0\" cellspacing=\"0\" style=\"width: 600px\">");

            if (!successXML)
            {
                sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;font-weight: bold;'>Used Website sitemap.xml from location : local " + whichXML + "</td>");
            }
            else
            {
                sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;font-weight: bold;'> Used Website sitemap.xml from location : " + "http://www.villaplus.com/" + whichXML + "</td>");
            }

            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<br>");
            sb.AppendLine("<table border='0' cellpadding=\"0\" cellspacing=\"0\" style=\"width: 600px\">");
            sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;font-weight: bold;'>Total URLs Checked : " + count + "</td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<br>");
            return sb;
        }

        private static StringBuilder middleStartHTML(StringBuilder sb)
        {
            sb.AppendLine("</table></td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<br>");

            return sb;
         }

        private static StringBuilder endHTML(StringBuilder sb)
        {
            sb.AppendLine("</table></td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("<br>");
            sb.AppendLine("<table border='0' cellpadding=\"0\" cellspacing=\"10\" style=\"width: 600px\">");
            sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;font-weight: bold;'>Please refer attatched reports for more details. </td>");
            sb.AppendLine("</tr>");
            //sb.AppendLine("<br>");
            sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;'>Please do not reply to this message via e-mail. This e-mail is auto generated. </td>");
            sb.AppendLine("<tr><td style = 'color: #000000; font-family: Arial,Helvetica,sans-serif;font-size: 13px;'>If any queries, please contact testing@ash-software.com </td>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb;
         }

        private static void saveInXls(string logFileforXlsName)
        {
            Excel.Workbook MyBook = null;
            Excel.Application MyApp = null;
            Excel.Worksheet MySheet = null;

            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Add();
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];
                int lastRow = 0;

                System.IO.StreamReader file5 = new System.IO.StreamReader(logFileforXlsName);
                string line = "";
                while ((line = file5.ReadLine()) != null)
                {
                    int j = 1;
                    lastRow += 1;
                    String[] Lines = line.Split('\t');
                    for (int i = 0; i < Lines.Length; i++)
                        {
                            String lineNew = Lines[i];
                            MySheet.Cells[lastRow, j] = lineNew;
                            j = j + 1;
                        }
                }
                MyBook.SaveAs(logFileforXlsName.Replace(".log", string.Empty) + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution, true,
                    Missing.Value, Missing.Value, Missing.Value);
                MyBook.Close(null, null, null);
                MyApp.Quit();
            }
            catch (Exception ex)
            {
                MyBook.Close(null, null, null);
                MyApp.Quit();
            }

        }

        //public static class My_DataTable_Extensions
        //{

            // Export DataTable into an excel file with field names in the header line
            // - Save excel file without ever making it visible if filepath is given
            // - Don't save excel file, just make it visible if no filepath is given
        private static void ExportToExcel(this DataTable[] Tables, string ExcelFilePath)
            {
                // load excel, and create a new workbook
                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                int sheetNum = 1;
                foreach (DataTable Tbl in Tables)
                {
                    if (sheetNum > 1)
                    {
                        excelApp.DisplayAlerts = false;
                        Excel.Workbook xlWorkBook = excelApp.Workbooks.Open(ExcelFilePath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        Excel.Sheets workSheet = xlWorkBook.Worksheets;
                        var xlNewSheet = (Excel.Worksheet)workSheet.Add(workSheet[1],
                        Type.Missing, Type.Missing, Type.Missing);
                        xlNewSheet.Name = "newsheet" + sheetNum;
                        xlNewSheet.Cells[1, 1] = "New sheet content";
                    }   
                    
                    try
                    {
                        if (Tbl == null || Tbl.Columns.Count == 0)
                            throw new Exception("ExportToExcel: Null or empty input table!\n");
                        
                        Excel._Worksheet workSheet = excelApp.ActiveSheet;
                        if(sheetNum==1)
                        {
                            workSheet.Name = "SliderImages";
                        }
                        else if(sheetNum==2)
                        {
                            workSheet.Name = "ActivitiesImages";

                        }else if(sheetNum==3)
                        {
                            workSheet.Name = "Tours";

                        }else if(sheetNum==4)
                        {
                            workSheet.Name = "FloorPlan";

                        }
                        else if(sheetNum==5)
                        {
                            workSheet.Name = "PhotosTabImages";

                        }                      
                        // column headings
                        for (int i = 0; i < Tbl.Columns.Count; i++)
                        {
                            workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
                        }

                        // rows
                        for (int i = 0; i < Tbl.Rows.Count; i++)
                        {
                            // to do: format datetime values before printing
                            for (int j = 0; j < Tbl.Columns.Count; j++)
                            {
                                workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                            }
                        }

                        // check fielpath
                        if (ExcelFilePath != null && ExcelFilePath != "")
                        {
                            try
                            {
                                workSheet.SaveAs(ExcelFilePath);
                                excelApp.Quit();
                                sheetNum++;
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                    + ex.Message);
                            }
                        }
                        else    // no filepath is given
                        {
                            excelApp.Visible = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: \n" + ex.Message);
                    }
                }
        }
        private static string Ping(string url)
        {
            string strUserAgent = string.Empty;
            string PlatformtoTest = ConfigurationSettings.AppSettings.Get("PlatformtoCheck");
            string FinalURL = "http://" + PlatformtoTest + ".villaplus.com" + url;
            string HrefStatuscode = string.Empty;

            System.Net.HttpWebRequest wreq;
            System.Net.HttpWebResponse wresp;
            System.IO.Stream mystream;
            System.Drawing.Bitmap bmp;

            bmp = null;
            mystream = null;
            wresp = null;
            try
            {
                wreq = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(FinalURL);
                wreq.AllowWriteStreamBuffering = true;

                wresp = (HttpWebResponse)wreq.GetResponse();

                if ((mystream = wresp.GetResponseStream()) != null)
                    bmp = new System.Drawing.Bitmap(mystream);
                HrefStatuscode = wresp.StatusCode.ToString();


                //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(FinalURL);
                //request.UserAgent = strUserAgent;
                //request.Timeout = 300;
                //request.AllowAutoRedirect = true;
                //request.Method = "GET";

                //HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                //HrefStatuscode = response.StatusCode.ToString();

                //int Statuscode = (int)response.StatusCode;
                //string HrefStatuscode = Convert.ToString(Statuscode);
            }
            catch (WebException ex)
            {
                if (ex.Status == (WebExceptionStatus.ProtocolError))
                {
                    //response = ((HttpWebResponse)ex.Response);

                    //strLog.AppendLine(str + '\t' + (int)response.StatusCode + '\t' + (string)response.StatusDescription + '\t' + "FAIL");
                    HrefStatuscode = ex.Message.ToString();
                }
                else
                {
                    HrefStatuscode = ex.Message.ToString();
                }
            }

            return HrefStatuscode;

            //catch()
            //{
            //    return "FAIL";
            //}
        }

        //public static void GetPictureSize(string url)
        //{
        //    //string strUserAgent = string.Empty;
        //    string PlatformtoTest = ConfigurationSettings.AppSettings.Get("PlatformtoCheck");
        //    string FinalURL = "http://" + PlatformtoTest + ".villaplus.com" + url;
        //    string HrefStatuscode = string.Empty;

        //    System.Net.HttpWebRequest wreq;
        //    System.Net.HttpWebResponse wresp;
        //    System.IO.Stream mystream;
        //    System.Drawing.Bitmap bmp;

        //    bmp = null;
        //    mystream = null;
        //    wresp = null;
        //    try
        //    {
        //        wreq = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(FinalURL);
        //        wreq.AllowWriteStreamBuffering = true;

        //        wresp = (HttpWebResponse)wreq.GetResponse();

        //        if ((mystream = wresp.GetResponseStream()) != null)
        //            bmp = new System.Drawing.Bitmap(mystream);
        //    }
        //    catch (Exception er)
        //    {
        //        //err = er.Message;
        //        return;
        //    }
        //    finally
        //    {
        //        if (mystream != null)
        //            mystream.Close();

        //        if (wresp != null)
        //            wresp.Close();
        //    }
        //}


    }

}