//-----------------------------------------------------------------------
// <copyright file="IWebClass.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace VehicleInformationLookupTool
{
    using System.Collections.Generic;

    /// <summary>
    /// Declares the public interface between this class and the rest of the application
    /// </summary>
    public interface IWebClass
    {
        /// <summary>
        /// Determines whether the NHTSA web service at the provided uri is working
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="token"> A cancellation token that this method should respect </param>
        /// <returns> A boolean that if true means the web service is working properly </returns>
        bool IsApiAccessible(string uri, CancellationToken token);

        /// <summary>
        /// Gets a list of column headers for vin data
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="xpath"> Specifies what nodes to retrieve from the XML response </param>
        /// <param name="token"> A cancellation token that this method should respect </param>
        /// <returns> A string list with column header text for the items of vin data returned </returns>
        List<string> GetVinColumnHeaders(string uri, string xpath, CancellationToken token);

        /// <summary>
        /// Gets a single row of data for the specified vin
        /// </summary>
        /// <param name="uri"> The fully-qualified location of the web service </param>
        /// <param name="vin"> The vehicle's VIN number </param>
        /// <param name="xpath"> Specifies what nodes to retrieve from the XML response </param>
        /// <param name="autoCorrect"> Specifies whether to auto-correct a VIN number </param>
        /// <param name="discardInvalid"> Specifies whether to discard data rows with invalid VIN numbers </param>
        /// <param name="token"> A cancellation token that this method should respect </param>
        /// <returns> A string list with the column values for the specified vin number </returns>
        List<string> GetVinDataRow(string uri, string vin, string xpath, bool autoCorrect, bool discardInvalid, CancellationToken token);
    }
}
