using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Net;
using System.IO;
using System.Threading;

using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Administration; 
using Microsoft.SharePoint;

namespace MarabouStork.Sharepoint.LinkChecker
{
    public class LinkChecker
    {
        static Regex urlRegEx = new Regex(@"http(s)?://([\w+?\.\w+])+([a-zA-Z0-9\~\!\@\#\$\%\^\&amp;\*\(\)_\-\=\+\\\/\?\.\:\;\'\,]*)?|]*?HREF\s*=\s*[""']?([^'"" >]+?)[ '""]?[^>]*?>");

        static readonly string userAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)"; // Do we need to set a user agent for the scraper.
        static string baseSiteUrl = "http://localhost:8081";
        static bool shouldUnpublish = true;
        static bool debug = false;

        /// <summary>
        ///     Determine if any of the documents in the test libraries contain references
        ///     to invalid urls
        /// </summary>
        /// <param name="modsSinceDate">
        ///     Used to retrieve documents that have been modified since the date provided
        /// </param>
        public static void ValidateDocumentUrls(string siteUrl, SPWebApplication webApplication, DateTime modsSinceDate)
        {
            baseSiteUrl = siteUrl;

            var farm = Microsoft.SharePoint.Administration.SPFarm.Local;
            var settings = farm.GetObject("Nhs.Evidence.Arms.LinkCheckerSettings", webApplication.Id, typeof(LinkCheckerPersistedSettings)) as LinkCheckerPersistedSettings;
            
            if (settings != null)
            {
                // Determine the list of document libraries to check
                string[] docLibs = settings.DocLibraries.Split(new char[] { ',', ';' });
                List<string> docLibsToCheck = new List<string>();
                foreach (string docLib in docLibs) docLibsToCheck.Add(string.Format("{0}/{1}", baseSiteUrl, docLib.Replace(" ", "%20")));

                // Determine the list of docment fields to check
                List<String> fieldsToCheck = new List<string>(settings.FieldsToCheck.Split(new char[] { ',', ';' }));

                // Should we unpublish documents that contain invalid urls?
                shouldUnpublish = settings.UnpublishInvalidDocs;

                //TODO: The CAML doesnt really need to worry about the modified date I dont think, although this could be used to make
                //      the checkatron more efficient if it is ran quite frequently.
                var query = "<Where>" +
                            "   <Gt>" +
                            "       <FieldRef Name='Modified' IncludeTimeValue='TRUE'/>" +
                            "       <Value Type='DateTime'>{0}</Value>" +
                            "   </Gt>" +
                            "</Where>";

                var retrvDt = SPUtility.CreateISO8601DateTimeFromSystemDateTime(modsSinceDate);
                var caml = new SPQuery { DatesInUtc = true, Query = string.Format(query, retrvDt) };

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    var spSite = new SPSite(baseSiteUrl);

                    using (var spWeb = spSite.OpenWeb())
                    {
                        // Check all documents in each of the document libraries
                        foreach (string docLibToCheck in docLibsToCheck)
                        {
                            try
                            {
                                var list = spWeb.GetList(docLibToCheck);
                                if (list != null)
                                {
                                    var documents = list.GetItems(caml);

                                    if (documents != null)
                                    {
                                        foreach (SPListItem document in documents)
                                        {
                                            var file = document.File;

                                            MemoryStream inStream = null;

                                            if (file.Level != SPFileLevel.Published)
                                            {
                                                foreach (SPFileVersion version in file.Versions)
                                                {
                                                    if ((version.Level == SPFileLevel.Published) && (version.IsCurrentVersion))
                                                    {
                                                        inStream = new MemoryStream(version.OpenBinary());
                                                        break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                inStream = new MemoryStream(file.OpenBinary());
                                            }

                                            // If we do not have a file stream at this point then the current document
                                            // is still in draft (unpublished) mode.
                                            if (inStream != null)
                                            {
                                                XmlDocument docContents = new XmlDocument();
                                                using (var sr = new StreamReader(inStream))
                                                {
                                                    docContents.Load(sr);
                                                }

                                                // Now that we have a published version of a document check the urls it contains
                                                if (debug)
                                                    Console.WriteLine("Checking Document : " + document.Name);
                                                CheckDocURls(document, docContents, fieldsToCheck);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                //TODO: There has been some sort of error. We should log it.
                            }
                        }
                    }
                });
            }
        }

        /// <summary>
        ///     Checks the validity of all of the urls that can be identified within the document provided
        /// </summary>
        /// <param name="doc">
        ///     The content of the document
        /// </param>
        /// <param name="docListItem">
        ///     The <see cref=" cref="SPListItem"/> instance representing the document in the document library
        /// </param> 
        /// <param name="fieldsToCheck">
        ///     The list of fields that may contain urls that need to be validated
        /// </param>
        private static void CheckDocURls(SPListItem docListItem, XmlDocument doc, List<string> fieldsToCheck)
        {
            XmlNamespaceManager man = new XmlNamespaceManager(doc.NameTable);
            man.AddNamespace("my", doc.DocumentElement.NamespaceURI);

            string textToCheck = string.Empty;

            foreach (string fieldToCheck in fieldsToCheck)
            {
                string xPath = "//my:myFields/" + ((fieldToCheck.StartsWith("my:")) ? fieldToCheck : "my:" + fieldToCheck);
                string fieldValue = RemoveCDATA(GetElementValue(doc, xPath, man));

                if (!string.IsNullOrEmpty(fieldValue.Trim())) textToCheck += " " + fieldValue;
            }

            textToCheck = textToCheck.Trim();

            bool docUrlsValid = true;
            string validationMessage = string.Empty;

            MatchCollection allUrls = urlRegEx.Matches(textToCheck);
            foreach (Match url in allUrls)
            {
                if (!ValidateUrl(url.Value, out validationMessage))
                {
                    docUrlsValid = false;
                    break;
                }
            }

            if (!docUrlsValid)
            {
                CreateUserWorkItem(docListItem, doc, validationMessage);
            }
        }

        /// <summary>
        ///     Creates a new item in the user workitems list to highlight to the relevant user that 
        ///     they need to review the failed document
        /// </summary>
        /// <param name="docUrl"></param>
        /// <param name="doc"></param>
        /// <param name="validationMessage"></param>
        private static void CreateUserWorkItem(SPListItem docListItem, XmlDocument doc, string validationMessage)
        {
            var query = "<Where>" +
                            "<Eq>" +
                                "<FieldRef Name='Document' />" +
                                "<Value Type='URL'>{0}</Value>" +
                            "</Eq>" +
                        "</Where>";

            var caml = new SPQuery { DatesInUtc = true, Query = string.Format(query, docListItem.Url) };

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                var spSite = new SPSite(baseSiteUrl);

                using (var spWeb = spSite.OpenWeb())
                {
                    var list = spWeb.GetList(baseSiteUrl + "/Lists/InvalidUrlsInResources");

                    var existingListItem = list.GetItems(caml);
                    if ((existingListItem == null) || ((existingListItem != null) && (existingListItem.Count == 0)))
                    {
                        if (list != null)
                        {
                            SPListItem newItem = list.AddItem();

                            newItem["User"] = docListItem["Author"];
                            newItem["Title"] = "Invalid Url in " + docListItem.Title;
                            newItem["Document"] = baseSiteUrl + "/" + docListItem.Url;
                            newItem["Message"] = validationMessage;

                            newItem.Update();

                            if (shouldUnpublish)
                                docListItem.File.UnPublish("The document was unpublished because it contains an invalid link. The link checker returned the following message \r\n\r\n" + validationMessage);
                        }
                        else
                        {
                            //TODO: Cannot find the requested workitems list
                        }
                    }
                }
            });

            if (debug)
                Console.WriteLine("The document was unpublished because it contains an invalid link. The link checker returned the following message \r\n\r\n" + validationMessage);
        }

        /// <summary>
        ///     Ensure that the url provided is valid and resolves to a web page or resource
        /// </summary>
        /// <param name="url">
        ///     The url to validate
        /// </param>
        /// <returns>
        ///     True where the url is valis; otherwise false;
        /// </returns>
        private static bool ValidateUrl(string url)
        {
            string message;
            return ValidateUrl(url, out message);
        }

        /// <summary>
        ///     Ensure that the url provided is valid and resolves to a web page or resource
        /// </summary>
        /// <param name="url">
        ///     The url to validate
        /// </param>
        /// <param name="message">
        ///     Optionally return the reason why the url failed.
        /// </param>
        /// <returns>
        ///     True where the url is valis; otherwise false;
        /// </returns>
        private static bool ValidateUrl(string url, out string message)
        {
            if (debug)
                Console.WriteLine("Checking Url : " + url);

            bool retVal = true;

            message = string.Empty;

            Uri uri;
            if (Uri.TryCreate(url, UriKind.Absolute, out uri))
            {
                // Try to fetch the page from the given URL, in case of any error return null string
                try
                {
                    checklink(url, ref message, ref retVal);
                }
                catch (WebException ex)
                {
                    //The Exception handler seems to be triggered regardless
                    if (ex.Status == WebExceptionStatus.Timeout)
                    {
                        try
                        {
                            Thread.Sleep(new TimeSpan(0, 0, 10));

                            checklink(url, ref message, ref retVal);
                        }
                        catch (WebException innerEx)
                        {
                            if (innerEx.Status == WebExceptionStatus.Timeout)
                            {
                                ex = null; // Clear the original exception
                            }
                            else
                                ex = innerEx;
                        }
                    }
                    if (ex != null)
                    {
                        message = "An error occured  while querying the specified url: " + ex.Message + " Please review the url " + url;
                        retVal = false;
                    }
                }
            }
            else
            {
                message = "The URL was not well formatted.";
                retVal = false;
            }

            return retVal;
        }

        private static void checklink(string url, ref string message, ref bool retVal)
        {
            // Cookies ?
            HttpWebRequest objRequest = (HttpWebRequest)System.Net.HttpWebRequest.Create(url);
            CookieContainer cookieContainer = new CookieContainer();

            if (!String.IsNullOrEmpty(userAgent))
            {
                objRequest.UserAgent = userAgent;
            }

            objRequest.CookieContainer = cookieContainer;
            objRequest.AllowAutoRedirect = true;
            objRequest.MaximumAutomaticRedirections = 5;
            objRequest.Method = "HEAD";
            HttpWebResponse objResponse = null;

            objResponse = (HttpWebResponse)objRequest.GetResponse();

            // In case of page not found error, return null string
            if (objResponse.StatusCode != HttpStatusCode.OK)
            {
                retVal = false;
                message = objResponse.StatusDescription + " Please review the url " + url;
            }
            
            objResponse.Close();
        }

        /// <summary>
        ///     Removes CDATA tag from around the string provided
        /// </summary>
        /// <param name="stringToChange"></param>
        /// <returns></returns>
        private static string RemoveCDATA(string stringToChange)
        {
            if (string.IsNullOrEmpty(stringToChange)) return string.Empty;

            var retVal = stringToChange;

            Regex cdataregex = new Regex("<!\\[[cC][dD][aA][tT][aA]\\[|\\]\\]>|&lt;!\\[[cC][dD][aA][tT][aA]\\[|\\]\\]&gt;|\\]\\]>|\\]\\]&gt;");
            while (cdataregex.IsMatch(retVal)) retVal = cdataregex.Replace(retVal, string.Empty);

            return retVal;
        }

        /// <summary>
        ///     Safely extracts a named value from the xml document provided
        /// </summary>
        /// <param name="document"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetElementValue(XmlNode document, string xpath, XmlNamespaceManager nsmgr)
        {
            var result = string.Empty;

            var item = document.SelectSingleNode(xpath, nsmgr);

            if (item != null)
            {
                result = item.InnerText;
            }

            return result;
        }
    }
}
