//-----------------------------------------------------------------------
// <copyright file="WebClass.cs">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Linq;
using System.Collections.Generic;
using System.Net;
using System.Xml;
using System.Net.Http;
using System.Threading;

namespace VehicleInformationLookupTool
{
    public class WebClass : IWebClass
    {
        private readonly HttpClient _client = new HttpClient();

        public WebClass()
        {
            /* Configure network connection settings */
            ServicePointManager.DefaultConnectionLimit = 4;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.CheckCertificateRevocationList = false;
            ServicePointManager.ReusePort = true;
            ServicePointManager.UseNagleAlgorithm = false;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;
        }

        private bool IsConnectedToInternet()
        {
            try
            {
                _ = _client.GetAsync(@"http://www.google.com").Result;
                return true;
            }
            catch
            {
                return false;
            }

            return false;
        }


        public bool IsApiAccessible(string uri, CancellationToken token)
        {
            /* Perform a GET against the API and store the raw XML result */
            const string testVin = "JH4TB2H26CC000000";
            var rawXmlString = ApiRequest(uri, testVin, token);
            if (string.IsNullOrWhiteSpace(rawXmlString) || rawXmlString == "exception")
            {
                return false;
            }

            /* De-serialize the XML and get the Message node */
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var messageNode = xmlDoc.SelectNodes(@"//Message");

            /* Return a success boolean based on the value of the Message node */
            return messageNode != null && messageNode[0].InnerText.StartsWith("Results returned successfully");
        }


        public List<string> GetVinDataRow(string uri, string lookupVin, string xpath, bool autoCorrect, bool discardInvalid, CancellationToken token)
        {
            /* This method is called from an alternate thread */

            /* Perform a GET against the API and store the raw XML result */
            var rawXmlString = ApiRequest(uri, lookupVin, token);
            if (string.IsNullOrWhiteSpace(rawXmlString) || rawXmlString == "exception")
            {
                return default;
            }

            /* De-serialize the XML and get the relevant nodes */
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var nodes = xmlDoc.SelectNodes(xpath);
            if (nodes == null || nodes.Count <= 0)
            {
                return default;
            }

            var messageNode = xmlDoc.SelectNodes(@"//Message");
            var message = messageNode?[0]?.InnerText ?? string.Empty;

            var errorNode = xmlDoc.SelectNodes(@"//ErrorCode");
            var error = errorNode?[0]?.InnerText ?? string.Empty;
            
            var suggestedVinNode = xmlDoc.SelectNodes(@"//SuggestedVIN");
            var suggestedVin = suggestedVinNode?[0]?.InnerText ?? string.Empty;

            var vinNode = xmlDoc.SelectNodes(@"//VIN");
            var vin = vinNode?[0]?.InnerText ?? string.Empty;

            xmlDoc = null;

            /* Optionally Auto-correct VIN numbers */
            var vinWasAutoCorrected = false;
            var originalVin = string.Empty;
            if (autoCorrect)
            {
                if (vinNode != null && !string.IsNullOrWhiteSpace(suggestedVin))
                {
                    /* Limit auto-correcting the VIN to errors indicating a problem with a single digit */
                    if (error.StartsWith("2") || error.StartsWith("3") || error.StartsWith("4"))
                    {
                        originalVin = vin;
                        vinNode[0].InnerText = suggestedVin;
                        vinWasAutoCorrected = true;
                    }
                }
            }

            /* Optionally Discard invalid VIN data */
            if (discardInvalid)
            {
                if (message == "Invalid URL" || error.StartsWith("11"))
                {
                    return default;
                }
            }
            
            /* Get the values for the data row from the API result */
            var vinItems = (from XmlNode node in nodes select node.InnerText).ToList();
            
            /* Add additional columns to the data */
            vinItems.Add(message);
            vinItems.Add(vinWasAutoCorrected.ToString());
            vinItems.Add(originalVin);

            return vinItems;
        }


        public List<string> GetVinColumnHeaders(string uri, string xpath, CancellationToken token)
        {
            /* Perform a GET against the API and store the raw XML result */
            const string testVin = "JH4TB2H26CC000000";
            var rawXmlString = ApiRequest(uri, testVin, token);
            if (string.IsNullOrWhiteSpace(rawXmlString) || rawXmlString == "exception")
            {
                return default;
            }

            /* De-serialize the XML and get the relevant nodes */
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawXmlString);
            var nodes = xmlDoc.SelectNodes(xpath);
            if (nodes == null || nodes.Count <= 0)
            {
                return default;
            }

            /* Get the column names for the data row from the API result */
            var headerList = (from XmlNode node in nodes select node.Name).ToList();

            /* Add additional column names to the data */
            headerList.Add("MessageFromServer");
            headerList.Add("AutoCorrectedVIN");
            headerList.Add("OriginalVIN");

            return headerList;
        }


        public string ApiRequest(string uriString, string vinNumber, CancellationToken token, int attempts = 3)
        {
            var uri = uriString.Replace("{VIN}", vinNumber);
            while (true)
            {
                try
                {
                    if (token.IsCancellationRequested)
                    {
                        return string.Empty;
                    }
                    attempts--;
                    return _client.GetAsync(uri, token).Result.Content.ReadAsStringAsync().Result;
                }
                catch when (attempts > 0)
                {
                    while (!IsConnectedToInternet())
                    {
                        if (token.IsCancellationRequested)
                        {
                            return string.Empty;
                        }
                        Thread.Sleep(30000);
                    }
                }
            }
        }
    }
}