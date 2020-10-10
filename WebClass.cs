//-----------------------------------------------------------------------
// <copyright file="WebClass.cs" company="N/A">
//     Copyright (c) 2016 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Windows;
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
        public bool NHTSAServiceIsWorking(string uri)
        {
            /* This method assumes that it should get XML results with a Message node in it */
            
            const string TESTVIN = "JH4TB2H26CC000000";
            string vinUri = uri.Replace("{VIN}", TESTVIN);

            string rawXmlString = string.Empty;
            using (WebClient web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            XmlNodeList messageNode = xmlDoc.SelectNodes(@"//Message");

            return messageNode[0].InnerText.StartsWith("Results returned successfully");
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

            string vinUri = uri.Replace("{VIN}", vin);

            string rawXmlString = string.Empty;
            using (WebClient web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            XmlNodeList nodes = xmlDoc.SelectNodes(xpath);

            XmlNodeList messageNode = xmlDoc.SelectNodes(@"//Message");
            string message = messageNode[0]?.InnerText ?? "";

            XmlNodeList errorNode = xmlDoc.SelectNodes(@"//ErrorCode");
            string error = "";
            if (errorNode[0] != null && errorNode[0].InnerText != null)
            {
                error = errorNode[0].InnerText;
            }

            XmlNodeList suggestedVinNode = xmlDoc.SelectNodes(@"//SuggestedVIN");
            string suggestedVin = "";
            if (suggestedVinNode[0] != null && suggestedVinNode[0].InnerText != null)
            {
                suggestedVin = suggestedVinNode[0].InnerText;
            }

            /* Logic to auto-correct VIN number */
            bool vinWasAutoCorrected = false;
            if (autoCorrect)
            {
                if (error.StartsWith("2") || error.StartsWith("3") || error.StartsWith("4"))
                {
                    XmlNodeList vinNode = xmlDoc.SelectNodes(@"//VIN");
                    vinNode[0].InnerText = suggestedVin;
                    vinWasAutoCorrected = true;
                }
            }

            /* Logic to discard invalid VIN data */
            if (discardInvalid)
            {
                if ((message == "Invalid URL") || error.StartsWith("11"))
                {
                    return null;
                }
            }

            List<string> vinItems = new List<string>();
            for (int i = 0; i < nodes.Count; i++)
            {
                vinItems.Add(nodes[i].InnerText);
            }

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

            const string TESTVIN = "JH4TB2H26CC000000";
            string vinUri = uri.Replace("{VIN}", TESTVIN);

            string rawXmlString = string.Empty;
            using (WebClient web = new WebClient())
            {
                rawXmlString = web.DownloadString(vinUri);
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            XmlNodeList nodes = xmlDoc.SelectNodes(xpath);

            List<string> headerList = new List<string>();
            for (int i = 0; i < nodes.Count; i++)
            {
                headerList.Add(nodes[i].Name);
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
                using (WebClient web = new WebClient())
                {
                    using (Stream stream = web.OpenRead(@"http://www.google.com"))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
