using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Xml;
using System.Xml.Serialization;
using HtmlAgilityPack;
using log4net;


namespace VisaBulletinNotification
{
    public class VisaBulletinPageParser
    {
        private const string DefaultVisaBulletWebUrl = "https://travel.state.gov/content/visas/en/law-and-policy/bulletin.html";

        private const string UiIdOfBulletins = "recent_bulletins";
        private const string UiClassOfCurrent = "current";
        private const string UiClassOfNext = "coming_soon";

        public List<BulletinPageLink> ParseBulletinPageLinks()
        {
            var visaBulletinHomeUrl = new Uri(DefaultVisaBulletWebUrl);
            var httpClient = new HttpClient();
            var html = httpClient.GetStringAsync(visaBulletinHomeUrl).Result;
            HtmlDocument visaBulletinHomePage = new HtmlDocument();
            visaBulletinHomePage.LoadHtml(html);

            List<BulletinPageLink> pageLinks = new List<BulletinPageLink>();

            string month = string.Empty;
            string year = string.Empty;
            string url = string.Empty;
            string type = string.Empty;


            var bulletins = visaBulletinHomePage.GetElementbyId(UiIdOfBulletins);
            var elements = bulletins.SelectNodes("li");

            //var elements = bulletins.SelectNodes("li[@class='" + UiClassOfCurrent + "']");
            if (elements != null && elements.Count > 0)
            {
                foreach (var ele in elements)
                {
                    var links = ele.SelectNodes("div/a");
                    if (links != null && links.Count > 0)
                    {
                        var dateText = links[0].InnerText;
                        if (dateText.Contains("Coming Soon"))
                        {
                            type = "next";
                        }
                        else
                        {
                            var regex = new System.Text.RegularExpressions.Regex("[a-zA-Z]{3,}");
                            var matches = regex.Matches(dateText);
                            if (matches.Count == 1)
                            {
                                month = matches[0].Captures[0].Value;
                            }

                            regex = new System.Text.RegularExpressions.Regex("[0-9]{4}");
                            matches = regex.Matches(dateText);
                            if (matches.Count == 1)
                            {
                                year = matches[0].Captures[0].Value;
                            }

                            type = "current";
                        }

                        var attributes = links[0].Attributes;
                        if (attributes.Contains("href"))
                        {
                            var path = attributes["href"].Value;
                            if (!string.IsNullOrEmpty(path))
                            {
                                if (!path.Contains(Uri.UriSchemeHttp) && !path.Contains(Uri.UriSchemeHttps))
                                {
                                    url = visaBulletinHomeUrl.OriginalString.Replace(visaBulletinHomeUrl.LocalPath, path);
                                }
                                else
                                {
                                    url = path;
                                }

                            }
                        }
                        else
                        {
                            if (type == "next")
                            {
                                url = "";
                                month = "";
                                year = "";
                            }
                        }

                        var pagelink = new BulletinPageLink();
                        pagelink.TargetBulletinMonth = month;
                        pagelink.TargetBulletinYear = year;
                        pagelink.TargetBulletinType = type;
                        pagelink.TargetBulletinUrl = url;
                        pageLinks.Add(pagelink);
                    }
                }
            }


            return pageLinks;
        }

        public List<VisaCutOffDate> ParseBulletin(Uri visaBulletinPageUrl)
        {
            List<VisaCutOffDate> bulletins = new List<VisaCutOffDate>();

            if (visaBulletinPageUrl == null)
            {
                return bulletins;
            }

            var httpClient = new HttpClient();
            var html = httpClient.GetStringAsync(visaBulletinPageUrl).Result;
            HtmlDocument visaBulletinPage = new HtmlDocument();
            visaBulletinPage.LoadHtml(html);

            // get the month 
            var header = visaBulletinPage.DocumentNode.SelectSingleNode("//div[@id='main']/h1");
            string _visaMonth = ClearText(header.InnerHtml);
            _visaMonth = _visaMonth.Replace("Visa Bulletin For ", string.Empty);
            var parts = _visaMonth.Split(new[] { ' ' });
            _visaMonth = parts[0];
            string _visaYear = parts[1];

            // build all tables
            var sections = visaBulletinPage.DocumentNode.SelectNodes("//div[@class='simple_richtextarea section']");
            int counts = sections.Count;
            for (var i = 0; i < counts; i++)
            {
                var section = sections[i];
                var tables = section.ChildNodes.Where(node => node.Name == "table");
                if (tables != null && tables.Count() > 0)
                {
                    string _dataType = "";

                    bool findInCurrentSection = true;
                    int indexOfFinal = -1;
                    int indexOfFilling = -1;

                    if (!section.InnerHtml.Contains("<u>FINAL ACTION DATES") &&
                        !section.InnerHtml.Contains("<u>DATES FOR FILING"))
                    {
                        if (i - 1 >= 0 && i - 1 <= counts - 1)
                        {
                            var previous = sections[i - 1];

                            //A. &nbsp;<u>FINAL ACTION DATES
                            //B.&nbsp; <u>DATES FOR FILING

                            indexOfFinal = previous.InnerHtml.IndexOf("<u>FINAL ACTION DATES");
                            indexOfFilling = previous.InnerHtml.IndexOf("<u>DATES FOR FILING");
                            findInCurrentSection = false;
                        }
                    }
                    else
                    {
                        indexOfFinal = section.InnerHtml.IndexOf("<u>FINAL ACTION DATES");
                        indexOfFilling = section.InnerHtml.IndexOf("<u>DATES FOR FILING");
                        findInCurrentSection = true;
                    }

                    if (indexOfFinal == -1 && indexOfFilling == -1)
                    {
                        continue;
                    }

                    if (indexOfFinal > 0 && indexOfFilling > 0)
                    {
                        // if current section only contains table, then use the larger index
                        if (section.SelectNodes("table").Count == 1)
                        {
                            if (findInCurrentSection == false)
                            {
                                if (indexOfFinal > indexOfFilling)
                                {
                                    _dataType = "final";
                                }
                                else
                                {
                                    _dataType = "filling";
                                }
                            }
                            else
                            {
                                if (indexOfFilling < indexOfFinal)
                                {
                                    _dataType = "filing";
                                }
                                else
                                {
                                    _dataType = "final";
                                }
                            }
                        }
                    }
                    else if (indexOfFinal > 0)
                    {
                        _dataType = "final";
                    }
                    else if (indexOfFilling > 0)
                    {
                        _dataType = "filing";
                    }
                    else if (indexOfFilling < indexOfFinal)
                    {
                        _dataType = "filing";
                    }

                    // previous section tells purpose of the table, like whether it is final action dates, or filling dates


                    foreach (var table in tables)
                    {
                        HtmlNodeCollection rows;
                        if (table.OuterHtml.Contains("tbody"))
                        {
                            rows = table.SelectNodes("tbody/tr");
                        }
                        else
                        {
                            rows = table.SelectNodes("tr");
                        }


                        // first cell contains the 
                        if (rows != null)
                        {

                            string _visaSponser = "";
                            int rsize = rows.Count;

                            var headercells = rows[0].SelectNodes("td");
                            if (headercells == null)
                            {
                                continue;
                            }

                            if (headercells[0].InnerHtml.Contains("Family"))
                            {
                                _visaSponser = "family";
                            }
                            else if (headercells[0].FirstChild.InnerHtml.Contains("Employ"))
                            {
                                _visaSponser = "employment";
                            }


                            for (int r = 1; r < rsize; r++)
                            {
                                var datacells = rows[r].SelectNodes("td");
                                var csize = datacells.Count;
                                var _visaType = datacells[0].InnerHtml;
                                for (int c = 1; c < csize; c++)
                                {
                                    var _visaArea = headercells[c].InnerHtml;
                                    var _visaDate = datacells[c].InnerHtml;

                                    VisaCutOffDate cutOffDate = new VisaCutOffDate()
                                    {
                                        VisaYear = ClearText(_visaYear),
                                        VisaMonth = ClearText(_visaMonth),
                                        VisaSponser = ClearText(_visaSponser),
                                        VisaType = ClearText(_visaType),
                                        VisaArea = ClearText(_visaArea),
                                        VisaDate = ClearText(_visaDate),
                                        DateType = ClearText(_dataType)
                                    };


                                    bulletins.Add(cutOffDate);
                                }
                            }

                        }
                    }
                }
            }

            return bulletins;
        }

        private string ClearText(string raw)
        {
            var decoded = WebUtility.HtmlDecode(raw);

            return decoded.Replace("<b>", string.Empty).Replace("</b>", string.Empty).Replace("<br>", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);
        }
    }

    public class VisaBulletinWatcher
    {
        public void SaveAndNotifyBulletins(string rootPath)
        {
            try
            {
                var smtphost = ConfigurationManager.AppSettings.Get("smtphost");
                var smtpport = ConfigurationManager.AppSettings.Get("smtpport");
                var sender = ConfigurationManager.AppSettings.Get("sender");
                var password = ConfigurationManager.AppSettings.Get("password");
                var receiver = ConfigurationManager.AppSettings.Get("receiver");

                if (string.IsNullOrWhiteSpace(smtphost)
                    || string.IsNullOrWhiteSpace(smtpport)
                    || string.IsNullOrWhiteSpace(sender)
                    || string.IsNullOrWhiteSpace(password)
                    || string.IsNullOrWhiteSpace(receiver))
                {
                    Logger.LogError("configuration error. 'smtphost' or 'smtpport' or 'sender' or 'password' or 'receiver' is null or empty");
                    return;
                }

                var parser = new VisaBulletinPageParser();
                var pageLinks = parser.ParseBulletinPageLinks();
                if (pageLinks.Count == 0)
                {
                    Logger.LogInfo("No Bulletin Page Links Were Found.");
                    return;
                }

                var nonEmptyLinks = pageLinks.Where(link => !string.IsNullOrEmpty(link.TargetBulletinUrl));

                // save available bulletins
                if (nonEmptyLinks.Count() > 0)
                {
                    foreach (var pageLink in nonEmptyLinks)
                    {
                        var file = this.SaveBulletins(rootPath, pageLink);
                        if (!string.IsNullOrEmpty(file))
                        {
                            var title = string.Format("Visa Bulletin - {0},{1}", pageLink.TargetBulletinMonth,
                                pageLink.TargetBulletinYear);
                            var content = this.FormatMailContent(File.ReadAllText(file));

                            this.NotifyBulletin(smtphost, int.Parse(smtpport), sender, receiver, title, content, password);
                        }
                    }
                }

                else
                {
                    Logger.LogInfo("urls of all builletin page links are empty.");
                }
            }
            catch (Exception e)
            {
                Logger.LogError(e.ToString());
            }
        }

        public bool CheckNextBulletin()
        {
            var month = DateTime.UtcNow.Month;
            var year = DateTime.UtcNow.Year;

            var parser = new VisaBulletinPageParser();
            var pageLinks = parser.ParseBulletinPageLinks();
            if (pageLinks.Count == 0)
            {
                Logger.LogInfo("no bulletin page links are found");
                return false;
            }

            var emptyLinks = pageLinks.Where(link => string.IsNullOrEmpty(link.TargetBulletinUrl));
            var nonEmptyLinks = pageLinks.Where(link => !string.IsNullOrEmpty(link.TargetBulletinUrl));

            // one available, another not available means next bulletin is not there
            if (nonEmptyLinks.Count() == 1 && emptyLinks.Count() == 1)
            {
                Logger.LogInfo("next bulletin not available @ " + DateTime.UtcNow);
                return false;
            }

            // all links are there, so next bulletin available
            if (emptyLinks.Count() == 0 && nonEmptyLinks.Count() == 2)
            {
                Logger.LogInfo("next bulletin is available @ " + DateTime.UtcNow);
                return true;
            }

            return false;
        }

        private void NotifyBulletin(string smtphost, int smtpport, string from, string to, string title, string content, string password)
        {
            try
            {
                MailMessage mail = new MailMessage(from, to, title, content);
                mail.IsBodyHtml = true;
                SmtpClient smptClient = new SmtpClient()
                {
                    Host = smtphost,
                    Port = smtpport,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(from, password)
                };
                smptClient.Send(mail);
            }
            catch (Exception e)
            {
                Logger.LogInfo(e.ToString());
            }
        }

        private string FormatMailContent(string rawContent)
        {
            XmlSerializer reader = new XmlSerializer(typeof(VisaCutOffDate[]));
            var bulletins = (VisaCutOffDate[])reader.Deserialize(new StringReader(rawContent));

            // sort bulletins by sponser, then datetype
            HtmlDocument doc = new HtmlDocument();
            var table = doc.CreateElement("table");
            var tbody = doc.CreateElement("tbody");
            table.ChildNodes.Add(tbody);
            var headerrow = doc.CreateElement("tr");
            tbody.ChildNodes.Add(headerrow);

            var thsponser = doc.CreateElement("th");
            thsponser.InnerHtml = "Sponser Type";
            headerrow.ChildNodes.Add(thsponser);

            var thcharttype = doc.CreateElement("th");
            thcharttype.InnerHtml = "Date Type";
            headerrow.ChildNodes.Add(thcharttype);

            var thvisatype = doc.CreateElement("th");
            thvisatype.InnerHtml = "Visa Type";
            headerrow.ChildNodes.Add(thvisatype);

            var thvisaarea = doc.CreateElement("th");
            thvisaarea.InnerHtml = "Visa Area";
            headerrow.ChildNodes.Add(thvisaarea);

            var thvisadate = doc.CreateElement("th");
            thvisadate.InnerHtml = "Visa Date";
            headerrow.ChildNodes.Add(thvisadate);

            // famil or employment
            var sponserGrouped = bulletins.GroupBy(x => x.VisaSponser);
            foreach (var sponserGroup in sponserGrouped)
            {
                var sponser = sponserGroup.Key;
                var groups = sponserGroup.ToArray();

                // final or filling
                var dateTypeGrouped = groups.GroupBy(x => x.DateType);

                // sponser - dateType - visaType - visaArea
                foreach (var dateTypeGroup in dateTypeGrouped)
                {
                    var chartType = dateTypeGroup.Key;
                    var cutOffDates = dateTypeGroup.ToArray();

                    foreach (var cutOffDate in cutOffDates)
                    {
                        var datarow = doc.CreateElement("tr");
                        tbody.ChildNodes.Add(datarow);

                        var tdsponser = doc.CreateElement("td");
                        tdsponser.InnerHtml = sponser;
                        datarow.ChildNodes.Add(tdsponser);

                        var tdcharttype = doc.CreateElement("td");
                        tdcharttype.InnerHtml = chartType;
                        datarow.ChildNodes.Add(tdcharttype);

                        var tdvisatype = doc.CreateElement("td");
                        tdvisatype.InnerHtml = cutOffDate.VisaType;
                        datarow.ChildNodes.Add(tdvisatype);

                        var tdvisaarea = doc.CreateElement("td");
                        tdvisaarea.InnerHtml = cutOffDate.VisaArea;
                        datarow.ChildNodes.Add(tdvisaarea);


                        var tdvisadate = doc.CreateElement("td");
                        tdvisadate.InnerHtml = cutOffDate.VisaDate;
                        datarow.ChildNodes.Add(tdvisadate);
                    }
                }
            }

            var html = doc.CreateElement("html");
            var body = doc.CreateElement("body");
            html.ChildNodes.Add(body);
            body.ChildNodes.Add(table);

            return /*WebUtility.HtmlEncode(*/html.OuterHtml/*)*/;
        }

        private string SaveBulletins(string rootPath, BulletinPageLink pageLink)
        {
            if (string.IsNullOrEmpty(pageLink.TargetBulletinYear) ||
                              string.IsNullOrEmpty(pageLink.TargetBulletinMonth) ||
                              string.IsNullOrEmpty(pageLink.TargetBulletinUrl))
            {
                return null;
            }


            var filename = string.Format("Visa-Bulletin-{0}-{1}.xml", pageLink.TargetBulletinYear,
            pageLink.TargetBulletinMonth);
            var filepath = Path.Combine(rootPath, filename);

            if (File.Exists(filepath))
            {
                Logger.LogInfo("Bulletin file already exist. " + filepath);
                return null;
            }
            var parser = new VisaBulletinPageParser();
            var bulletins = parser.ParseBulletin(new Uri(pageLink.TargetBulletinUrl));
            XmlSerializer writer = new XmlSerializer(typeof(VisaCutOffDate[]));
            FileStream file = File.Create(filepath);
            writer.Serialize(file, bulletins.ToArray());
            file.Close();
            
            Logger.LogInfo("Bulletin file saved at " + filepath);

            return new FileInfo(filepath).FullName;
        }
    }

    public static class Logger
    {
        private static ILog log = LogManager.GetLogger("VisaBulletinNotification");

        public static void LogInfo(string message)
        {
            log.Info(message);
        }

        public static void LogError(string message)
        {
            log.Error(message);
        }
    }


    [Serializable]
    public class VisaCutOffDate
    {
        public string VisaYear { get; set; }
        public string VisaMonth { get; set; }
        public string DateType { get; set; }
        public string VisaSponser { get; set; }
        public string VisaArea { get; set; }
        public string VisaType { get; set; }
        public string VisaDate { get; set; }
    }
    
    public class BulletinPageLink
    {
        public string TargetBulletinType { get; set; }
        public string TargetBulletinMonth { get; set; }
        public string TargetBulletinYear { get; set; }
        public string TargetBulletinUrl { get; set; }
    }

}
