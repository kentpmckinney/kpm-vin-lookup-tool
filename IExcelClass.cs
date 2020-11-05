//-----------------------------------------------------------------------
// <copyright file="IExcelClass.cs">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.Data;

namespace VehicleInformationLookupTool
{
    public interface IExcelClass
    {
        /// <summary>
        /// Open an Excel or CSV file
        /// </summary>
        /// <param name="fileName"> The full path of the file </param>
        void OpenFile(string fileName);


        /// <summary>
        /// Close the currently open file
        /// </summary>
        void CloseFile();


        /// <summary>
        /// Clear stored data
        /// </summary>
        void ClearData();


        /// <summary>
        /// Get a list of sheet names in the currently open file
        /// </summary>
        /// <returns> A string array of worksheet names </returns>
        List<string> GetSheetNames();


        /// <summary>
        /// Get a list of column names in the currently open file
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet (only applies to Excel files) </param>
        /// <returns> A string array of column names </returns>
        List<string> GetColumnNames(string sheetName);


        /// <summary>
        /// Get a DataTable from a worksheet
        /// </summary>
        /// <param name="worksheetIndex"> The index number of the worksheet </param>
        /// <returns> The DataTable with data for the specified worksheet </returns>
        DataTable GetDataTable(int worksheetIndex);


        /// <summary>
        /// Get a worksheet in the loaded file that heuristically seems to contain a column with VIN numbers
        /// </summary>
        /// <returns> The index of the worksheet </returns>
        int SheetLikelyToContainVins();


        /// <summary>
        /// Get a column within a worksheet that heuristically seems to contain VIN numbers
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> The index of the column </returns>
        int ColumnLikelyToContainVins(string sheetName);


        /// <summary>
        /// Save data to an Excel file
        /// </summary>
        /// <param name="saveFileName"> The name of the file </param>
        /// <param name="data"> A DataTable reference </param>
        /// <returns> True or false whether the operation is successful </returns>
        bool SaveExcelFile(string saveFileName, in DataTable data);


        /// <summary>
        /// Save data to a CSV file
        /// </summary>
        /// <param name="saveFileName"> The name of the file </param>
        /// <param name="data"> A DataTable reference </param>
        /// <returns> True or false whether the operation is successful </returns>
        bool SaveCsvFile(string saveFileName, in DataTable data);
    }
}