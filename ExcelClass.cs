//-----------------------------------------------------------------------
// <copyright file="ExcelClass.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
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
    using OfficeOpenXml;
    using System.Text;
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
        private static readonly List<string> VinIndicators = new List<string> { "vin", "vehicleidentificationnumber" };

        /// <summary>
        /// Instance of the third-party class used to read Excel files
        /// </summary>
        private IExcelDataReader _excelDataReader;

        /// <summary>
        /// Instance of the class which allows access to the Excel file itself
        /// </summary>
        private FileStream _fileStream;

        /// <summary>
        /// Instance of a DataSet that the IExcelDataReader class transfers _data to
        /// </summary>
        private DataSet _data;

        /// <summary>
        /// Properly dispose of _data
        /// </summary>
        public void Dispose()
        {
            _excelDataReader?.Dispose();
            _excelDataReader = null;

            _fileStream?.Dispose();
            _fileStream = null;
            
            _data?.Dispose();
            _data = null;
        }

        /// <summary>
        /// Close the currently open file
        /// </summary>
        public void CloseFile()
        {
            _excelDataReader?.Close();
            _fileStream?.Close();
        }

        /// <summary>
        /// Clear stored _data
        /// </summary>
        public void ClearData() =>
            _data?.Clear();
        
        /// <summary>
        /// Get a list of column names given the name of a worksheet
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> A string array of column names </returns>
        public List<string> GetColumnNames(string sheetName) =>
            (from DataTable table in _data.Tables where table.TableName == sheetName from DataColumn column in table.Columns select column.ColumnName).ToList();
        
        /// <summary>
        /// Get a list of sheet names in the currently open file
        /// </summary>
        /// <returns> A string array of worksheet names </returns>
        public List<string> GetSheetNames() => 
            (from DataTable table in _data.Tables select table.TableName).ToList();
        
        /// <summary>
        /// Indicates whether the Excel file is valid or not
        /// </summary>
        /// <returns> A boolean indicating whether the file is valid </returns>
        public bool IsValidFile() => 
            _excelDataReader?.IsValid == true;
        
        /// <summary>
        /// Opens an Excel file
        /// </summary>
        /// <param name="fileName"> The file name </param>
        public void OpenFile(string fileName)
        {
            if (!File.Exists(fileName))
            {
                MessageBox.Show(fileName, "File Not Found");
            }

            try
            {
                _fileStream = File.Open(fileName, FileMode.Open, FileAccess.Read);

                _excelDataReader = Path.GetExtension(fileName) == ".xls"
                    ? ExcelReaderFactory.CreateBinaryReader(_fileStream, ReadOption.Loose)
                    : ExcelReaderFactory.CreateOpenXmlReader(_fileStream);

                _excelDataReader.IsFirstRowAsColumnNames = true;
                _data = _excelDataReader?.AsDataSet();
            }
            catch (IOException)
            {
                MessageBox.Show("The file could not be opened, possiby because it is open in another program:\n\n" + fileName, "Unable to Open File");
            }
        }

        /// <summary>
        /// Get a DataTable object given the index number of a worksheet
        /// </summary>
        /// <param name="worksheetIndex"> The index number of the worksheet </param>
        /// <returns> The DataTable with _data for the specified worksheet </returns>
        public DataTable GetDataTable(int worksheetIndex) =>
            _data.Tables[worksheetIndex];
        
        /// <summary>
        /// Get a worksheet in the loaded file that heuristically seems to contain a column with VIN numbers
        /// </summary>
        /// <returns> The index of the worksheet </returns>
        public int SheetLikelyToContainVins()
        {
            if (!IsValidFile())
                return 0;

            var sheets = GetSheetNames();
            for (var i = 0; i < sheets.Count; i++)
            {
                var columnNames = GetColumnNames(sheets[i]);
                if (columnNames.Any(name => VinIndicators.Contains(name, StringComparer.OrdinalIgnoreCase)))
                {
                    return i;
                }
            }

            return 0;
        }

        /// <summary>
        /// Get a column within a worksheet that heuristically seems to contain VIN numbers
        /// </summary>
        /// <param name="sheetName"> The name of the worksheet </param>
        /// <returns> The index of the column </returns>
        public int ColumnLikelyToContainVins(string sheetName)
        {
            var columnNames = this.GetColumnNames(sheetName);
            for (var i = 0; i < columnNames.Count; i++)
            {
                if (VinIndicators.Contains(columnNames[i], StringComparer.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return 0;
        }

        /// <summary>
        /// Save _data to an Excel file
        /// </summary>
        /// <param name="saveFileName"> The name of the file </param>
        /// <param name="data"> A DataTable reference </param>
        /// <returns> True or false whether the operation is successful or not </returns>
        public bool SaveExcelFile(string saveFileName, DataTable data)
        {
            try
            {
                /* Use EPPlus to save the Excel file */
                using (var excel = new ExcelPackage(new FileInfo(saveFileName)))
                {
                    /* add a worksheet */
                    var worksheet = excel.Workbook.Worksheets.Add("Vehicle Information Lookup Tool");

                    /* Load datatable into the worksheet */
                    worksheet.SelectedRange.LoadFromDataTable(data, true);

                    /* Autofit columns for all cells */
                    worksheet.Cells.AutoFitColumns(0);

                    /* Save the file */
                    excel.Save();
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveFileName"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool SaveCsvFile(string saveFileName, DataTable data)
        {
            try
            {
                using (var writer = new StreamWriter(saveFileName))
                {
                    var numColumns = data.Columns.Count;
                    var lastColumn = numColumns - 1;
                    const string comma = ",";

                    for (var r = 0; r < data.Rows.Count; r++)
                    {
                        var line = new StringBuilder();
                        var values = data.Rows[r].ItemArray;

                        for (var c = 0; c < data.Columns.Count; c++)
                        {
                            var value = values[c].ToString();
                            value = value.Replace(',',' ');

                            if (c == lastColumn)
                            {
                                line.Append(value);
                            }
                            else
                            {
                                line.Append(value + comma);
                            }
                        }

                        writer.WriteLine(line);
                        writer.Flush();
                    }
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