//-----------------------------------------------------------------------
// <copyright file="WebClass.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Linq;
using System.Windows.Forms;

namespace VehicleInformationLookupTool
{
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// Encapsulates web access
    /// </summary>
    public class WebClass : IWebClass
    {
        /// <summary>
        /// Determines whether the NHTSA web service at the provided uri is working
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <returns> A boolean that if true means the web service is working properly </returns>
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

        /// <summary>
        /// Gets a single row of data for the specified vin
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="vin"> The vehicle's VIN number </param>
        /// <param name="xpath"> Specifies what nodes to retrieve from the XML response </param>
        /// <returns> A string list with the column values for the specified vin number </returns>
        public List<string> GetVinDataRow(string uri, string vin, string xpath, bool autoCorrect, bool discardInvalid)
        {
            /* This method is called from an alternate thread */

            var vinUri = uri.Replace("{VIN}", vin);

            var rawXmlString = default(string);
            using (var web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var nodes = xmlDoc.SelectNodes(xpath);

            var messageNode = xmlDoc.SelectNodes(@"//Message");
            var message = messageNode?[0]?.InnerText ?? "";

            var errorNode = xmlDoc.SelectNodes(@"//ErrorCode");
            var error = errorNode?[0]?.InnerText ?? "";
            
            var suggestedVinNode = xmlDoc.SelectNodes(@"//SuggestedVIN");
            var suggestedVin = suggestedVinNode?[0]?.InnerText ?? "";

            /* Logic to auto-correct VIN number */
            var vinWasAutoCorrected = false;
            if (autoCorrect)
            {
                if (error.StartsWith("2") || error.StartsWith("3") || error.StartsWith("4"))
                {
                    var vinNode = xmlDoc.SelectNodes(@"//VIN");
                    if (vinNode != null)
                    {
                        vinNode[0].InnerText = suggestedVin;
                        vinWasAutoCorrected = true;
                    }
                }
            }

            /* Logic to discard invalid VIN data */
            // TODO: double check this logic
            if (discardInvalid)
            {
                if (message == "Invalid URL" || error.StartsWith("11"))
                {
                    return null;
                }
            }
            
            var vinItems = (from XmlNode node in nodes select node?["Value"]?.InnerText).ToList();
            vinItems.Add(message);
            vinItems.Add(vinWasAutoCorrected.ToString());

            return vinItems;
        }

        /// <summary>
        /// Gets a list of column headers for vin data
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="xpath"> Specifies what nodes to retrieve from the XML response </param>
        /// <returns> A string list with column header text for the items of vin data returned </returns>
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

            if (nodes == null)
                return default;

            var headerList = new List<string>();
            foreach (XmlNode node in nodes)
            {
                var columnName = node?["Variable"]?.FirstChild.InnerText ?? "";
                headerList.Add(columnName);
            }
            headerList.Add("MessageFromServer");
            headerList.Add("AutoCorrectedVIN");

            return headerList;
        }

        /// <summary>
        /// Check for Internet access
        /// </summary>
        /// <returns> A boolean value indicating whether www.google.com is reachable </returns>
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