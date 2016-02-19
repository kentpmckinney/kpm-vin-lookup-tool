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

            return messageNode[0].InnerText == "Results returned successfully";
        }

        /// <summary>
        /// Gets a single row of data for the specified vin
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="vin"> The vehicle's VIN number </param>
        /// <param name="xpath"> Specifies what nodes to retrieve from the XML response </param>
        /// <returns> A string list with the column values for the specified vin number </returns>
        public List<string> GetVinDataRow(string uri, string vin, string xpath)
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

            List<string> vinItems = new List<string>();
            for (int i = 0; i < nodes.Count; i++)
            {
                vinItems.Add(nodes[i].InnerText);
            }
            
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
