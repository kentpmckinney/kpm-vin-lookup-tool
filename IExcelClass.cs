//-----------------------------------------------------------------------
// <copyright file="IExcelClass.cs">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System.Collections.Generic;
    using System.Data;

    public interface IExcelClass
    {
        /// <summary>
        /// Opens an Excel file
        /// </summary>
        /// <param name="fileName"> The file name </param>
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
        /// Get a list of column names given the name of a worksheet
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> A string array of column names </returns>
        List<string> GetColumnNames(string sheetName);


        /// <summary>
        /// Get a DataTable object given the index number of a worksheet
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
        /// <returns> True or false whether the operation is successful or not </returns>
        bool SaveExcelFile(string saveFileName, in DataTable data);


        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveFileName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        bool SaveCsvFile(string saveFileName, in DataTable data);
    }
}