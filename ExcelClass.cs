//-----------------------------------------------------------------------
// <copyright file="ExcelClass.cs" company="N/A">
//     Copyright (c) 2016 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using Excel;
    using OpenXmlPackaging;
    using VehicleInformationLookupTool;

    /// <summary>
    /// Encapsulates Excel functionality
    /// </summary>
    public class ExcelClass : IExcelClass, IDisposable
    {
        /// <summary>
        /// Typical names of columns in an Excel worksheet which would likely contain VIN numbers
        /// </summary>
        /// <remarks>
        /// Not case sensitive
        /// </remarks>
        private static readonly List<string> VININDICATORS = new List<string> { "vin", "vehicleidentificationnumber" };

        /// <summary>
        /// Instance of the third-party class used to read Excel files
        /// </summary>
        private IExcelDataReader excelDataReader;

        /// <summary>
        /// Instance of the class which allows access to the Excel file itself
        /// </summary>
        private FileStream fileStream;

        /// <summary>
        /// Instance of a DataSet that the IExcelDataReader class transfers data to
        /// </summary>
        private DataSet data;

        /// <summary>
        /// Properly dispose of data
        /// </summary>
        public void Dispose()
        {
            if (this.excelDataReader != null)
            {
                this.excelDataReader.Dispose();
                this.excelDataReader = null;
            }

            if (this.fileStream != null)
            {
                this.fileStream.Dispose();
                this.fileStream = null;
            }

            if (this.data != null)
            {
                this.data.Dispose();
                this.data = null;
            }
        }

        /// <summary>
        /// Close the currently open file
        /// </summary>
        public void CloseFile()
        {
            this.excelDataReader.Close();
            this.fileStream.Close();
        }

        /// <summary>
        /// Clear stored data
        /// </summary>
        public void ClearData()
        {
            this.data.Clear();
        }

        /// <summary>
        /// Get a list of column names given the name of a worksheet
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> A string array of column names </returns>
        public List<string> GetColumnNames(string sheetName)
        {
            List<string> columnNames = new List<string>();
            foreach (DataTable table in this.data.Tables)
            {
                if (table.TableName == sheetName)
                {
                    foreach (DataColumn column in table.Columns)
                    {
                        columnNames.Add(column.ColumnName);
                    }
                }
            }

            return columnNames;
        }

        /// <summary>
        /// Get a list of sheet names in the currently open file
        /// </summary>
        /// <returns> A string array of worksheet names </returns>
        public List<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (DataTable table in this.data.Tables)
            {
                sheetNames.Add(table.TableName);
            }

            return sheetNames;
        }

        /// <summary>
        /// Indicates whether the Excel file is valid or not
        /// </summary>
        /// <returns> A boolean indicating whether the file is valid </returns>
        public bool IsValidFile()
        {
            bool valid = this.excelDataReader.IsValid;
            return valid;
        }

        /// <summary>
        /// Opens an Excel file
        /// </summary>
        /// <param name="fileName"> The file name </param>
        public void OpenFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                MessageBox.Show("File Not Found", fileName);
            }

            this.fileStream = File.Open(fileName, FileMode.Open, FileAccess.Read);

            if (Path.GetExtension(fileName) == ".xls")
            {
                this.excelDataReader = ExcelReaderFactory.CreateBinaryReader(this.fileStream, ReadOption.Loose);
            }
            else
            {
                this.excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(this.fileStream);
            }

            this.excelDataReader.IsFirstRowAsColumnNames = true;
            this.data = this.excelDataReader.AsDataSet();
        }

        /// <summary>
        /// Get a DataTable object given the index number of a worksheet
        /// </summary>
        /// <param name="worksheetIndex"> The index number of the worksheet </param>
        /// <returns> The DataTable with data for the specified worksheet </returns>
        public DataTable GetDataTable(int worksheetIndex)
        {
            DataTable table = this.data.Tables[worksheetIndex];
            return table;
        }

        /// <summary>
        /// Get a worksheet in the loaded file that heuristically seems to contain a column with VIN numbers
        /// </summary>
        /// <returns> The index of the worksheet </returns>
        public int SheetLikelyToContainVINs()
        {
            if (this.IsValidFile())
            {
                List<string> sheets = this.GetSheetNames();
                for (int i = 0; i < sheets.Count; i++)
                {
                    List<string> columnNames = this.GetColumnNames(sheets[i]);
                    for (int j = 0; j < columnNames.Count; j++)
                    {
                        if (VININDICATORS.Contains(columnNames[j], StringComparer.OrdinalIgnoreCase))
                        {
                            return i;
                        }
                    }
                }
            }

            return 0;
        }

        /// <summary>
        /// Get a column within a worksheet that heuristically seems to contain VIN numbers
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> The index of the column </returns>
        public int ColumnLikelyToContainVINs(string sheetName)
        {
            List<string> columnNames = this.GetColumnNames(sheetName);
            for (int i = 0; i < columnNames.Count; i++)
            {
                if (VININDICATORS.Contains(columnNames[i], StringComparer.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return 0;
        }

        /// <summary>
        /// Save data to an Excel file
        /// </summary>
        /// <param name="saveFileName"> The name of the file </param>
        /// <param name="data"> A DataTable reference </param>
        /// <returns> True or false whether the operation is successful or not </returns>
        public bool SaveExcelFile(string saveFileName, DataTable data)
        {
            try
            {
                using (SpreadsheetDocument doc = new SpreadsheetDocument(saveFileName))
                {
                    /* Create a worksheet */
                    Worksheet sheet = doc.Worksheets.Add("Vehicle Information Lookup Tool");

                    /* Import the DataTable into the worksheet */
                    sheet.ImportDataTable(data, "A1", true);
                    
                    /* 
                      The OpenXmlPackaging documentation states:
                      "No need to explicitly save/close the excel file.
                       When you come out of the using statement, the file gets automatically saved!"
                    */
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
