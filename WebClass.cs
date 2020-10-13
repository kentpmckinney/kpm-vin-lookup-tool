//-----------------------------------------------------------------------
// <copyright file="WebClass.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Linq;
using System.Collections.Generic;
using System.Net;
using System.Xml;

namespace VehicleInformationLookupTool
{
    public class WebClass : IWebClass
    {
        public bool NhtsaServiceIsWorking(string uri)
        {
            /* This method assumes that it should get XML results with a Message node in it */
            
            const string testvin = "JH4TB2H26CC000000";
            var vinUri = uri.Replace("{VIN}", testvin);

            var rawXmlString = default(string);
            using (var web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var messageNode = xmlDoc.SelectNodes(@"//Message");

            return messageNode != null && messageNode[0].InnerText.StartsWith("Results returned successfully");
        }


        public List<string> GetVinDataRow(string uri, string lookupVin, string xpath, bool autoCorrect, bool discardInvalid)
        {
            /* This method is called from an alternate thread */

            var vinUri = uri.Replace("{VIN}", lookupVin);

            var rawXmlString = default(string);
            using (var web = new WebClient())
            {
                try
                {
                    rawXmlString = web.DownloadString(vinUri);
                }
                catch (WebException)
                {
                    ;
                }
            }
            if (rawXmlString is null)
                return new List<string>();

            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var nodes = xmlDoc.SelectNodes(xpath);

            var messageNode = xmlDoc.SelectNodes(@"//Message");
            var message = messageNode?[0]?.InnerText ?? string.Empty;

            var errorNode = xmlDoc.SelectNodes(@"//ErrorCode");
            var error = errorNode?[0]?.InnerText ?? string.Empty;
            
            var suggestedVinNode = xmlDoc.SelectNodes(@"//SuggestedVIN");
            var suggestedVin = suggestedVinNode?[0]?.InnerText ?? string.Empty;

            var vinNode = xmlDoc.SelectNodes(@"//VIN");
            var vin = vinNode?[0]?.InnerText ?? string.Empty;

            /* Logic to auto-correct VIN number */
            var vinWasAutoCorrected = false;
            var originalVin = string.Empty;
            if (autoCorrect)
            {
                if (vinNode != null && !string.IsNullOrWhiteSpace(suggestedVin))
                {
                    originalVin = vin;
                    vinNode[0].InnerText = suggestedVin;
                    vinWasAutoCorrected = true;
                }
            }

            /* Logic to discard invalid VIN data */
            if (discardInvalid)
            {
                if (message == "Invalid URL" || error.StartsWith("11"))
                {
                    return new List<string>() { "Discarded" };
                }
            }
            
            var vinItems = (from XmlNode node in nodes select node?.InnerText).ToList();
            vinItems.Add(message);
            vinItems.Add(vinWasAutoCorrected.ToString());
            vinItems.Add(originalVin);

            return vinItems;
        }


        public List<string> GetVinColumnHeaders(string uri, string xpath)
        {
            /* This method assumes that header fields are the same for all results */

            const string testvin = "JH4TB2H26CC000000";
            var vinUri = uri.Replace("{VIN}", testvin);

            var rawXmlString = default(string);
            using (var web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var nodes = xmlDoc.SelectNodes(xpath);

            if (nodes is null)
                return default;

            var headerList = new List<string>();
            foreach (XmlNode node in nodes)
            {
                var columnName = node?.Name ?? "";
                headerList.Add(columnName);
            }
            headerList.Add("MessageFromServer");
            headerList.Add("AutoCorrectedVIN");
            headerList.Add("OriginalVIN");

            return headerList;
        }

        
        public bool IsConnectedToInternet()
        {
            try
            {
                using (var web = new WebClient())
                {
                    using (var stream = web.OpenRead(@"http://www.google.com"))
                    {
                        if (stream != null && stream.CanRead)
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }
    }
}