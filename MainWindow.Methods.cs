//-----------------------------------------------------------------------
// <copyright file="MainWindow.Methods.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Linq;

namespace VehicleInformationLookupTool
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Win32;

    /// <summary>
    /// Helper methods for MainWindow
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>
        /// Opens the default _web browser to the specified site
        /// </summary>
        /// <param name="site"> The _web site address </param>
        private static void LaunchWebBrowser(string site)
        {
            try
            {
                Process.Start(site);
            }
            catch (Exception)
            {
                var message = "Unable to open the site in a _web browser automatically.";
                message += " Please visit the following _web page manually:\n\n" + site;
                message += "\n\nPossible Causes:\n";
                message += "   This application cannot access the Internet\n";
                message += "   The _web page has moved to a different location\n";
                message += "   Configuration setting or problem with the operating system\n";
                MessageBox.Show(message);
            }
        }

        /// <summary>
        /// Consumes the contents of the text box containing VINs converting it to a more usable form
        /// </summary>
        /// <param name="textContainingVINs"> Raw text input </param>
        /// <returns> A DataTable with the VINs in a single column </returns>
        private static DataTable VinTextToDataTable(string textContainingVins)
        {
            /*
              Requirement: the user must only use commas, semicolons, and newlines to delineate VINs (as indicated in the UI)
              This application does not attempt to validate VIN numbers
            */

            /* Normalize the text so that there is one VIN per line */
            textContainingVins = textContainingVins.Replace(";", "\n");
            textContainingVins = textContainingVins.Replace(",", "\n");
            textContainingVins = textContainingVins.Replace(Environment.NewLine, "\n");
            textContainingVins = textContainingVins.Replace("\n\n\n", "\n");
            textContainingVins = textContainingVins.Replace("\n\n", "\n");
            textContainingVins = textContainingVins.Replace(" ", string.Empty);

            /* Convert the text into a list of strings */
            var vinList = textContainingVins.Split('\n');

            /* Create a DataTable */
            var table = new DataTable();

            /* Add columns */
            table.Columns.Add("VIN");

            /* Add rows */
            foreach (var vin in vinList)
            {
                if (!string.IsNullOrEmpty(vin))
                {
                    table.Rows.Add(vin);
                }
            }

            return table;
        }

        /// <summary>
        /// Checks the registry to see if this is the first time that the program has run in the current user context
        /// </summary>
        /// <returns> True or false </returns>
        private bool IsFirstRun()
        {
            if (ReadRegistryString("FirstRun") == "False")
            {
                return false;
            }
            
            WriteRegistryString("FirstRun", "False");
            return true;
        }

        /// <summary>
        /// Reads the registry value that indicates whether the user has agreed to the EULA
        /// </summary>
        /// <returns> True or false </returns>
        private bool UserHasAgreedToEula() => 
            ReadRegistryString("AgreedEULA") == true.ToString();
        
        /// <summary>
        /// Writes the registry value to indicate whether the user has agreed to the EULA
        /// </summary>
        /// <param name="state"> True or false </param>
        private void SetUserAgreedToEula(bool state) =>
            WriteRegistryString("AgreedEULA", state.ToString());
        
        /// <summary>
        /// Reads a string value from the registry key HKCU\SOFTWARE\VehicleInformationLookupTool
        /// </summary>
        /// <param name="valueName"> The name of the registry value </param>
        /// <returns> A string that is empty or which contains the specified registry value </returns>
        private static string ReadRegistryString(string valueName)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                var hkcu = Registry.CurrentUser;
                var software = hkcu?.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                var key = software?.OpenSubKey("VehicleInformationLookupTool", RegistryKeyPermissionCheck.ReadWriteSubTree);

                /* Attempt to get the named value */
                var valueString = key?.GetValue(valueName) as string ?? string.Empty;

                key?.Close();
                return valueString;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Writes a string value to the registry key {HKCU\SOFTWARE\VehicleInformationLookupTool}\valueName
        /// </summary>
        /// <param name="valueName"> The name of the registry value </param>
        /// <param name="valueString"> The string value </param>
        private static void WriteRegistryString(string valueName, string valueString)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                var hkcu = Registry.CurrentUser;
                var software = hkcu?.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                var key = software?.OpenSubKey("VehicleInformationLookupTool", RegistryKeyPermissionCheck.ReadWriteSubTree);

                /* If the key was not found attempt to create it */
                if (key is null)
                {
                    key = software?.CreateSubKey("VehicleInformationLookupTool"); 
                    key?.SetValue(valueName, valueString);
                    key?.Close();
                }
            }
            catch (Exception)
            {
                ; // Do nothing
            }
        }

        /// <summary>
        /// Prompts the user for the name of an Excel file to open
        /// </summary>
        /// <returns> A string that is empty or which contains the name of a file </returns>
        private static string PromptOpenExcelFileName()
        {
            var dialog = new OpenFileDialog()
            {
                DefaultExt = ".xlsx",
                Filter = "Excel files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All files (*.*)|*.*"
            };

            /* Display the dialog and return the text entered by the user */
            return dialog?.ShowDialog() == true
                ? dialog.FileName
                : string.Empty;
        }

        /// <summary>
        /// Prompts the user for the name of an Excel file to which to save data
        /// </summary>
        /// <returns> A string that is empty or which contains the name of a file </returns>
        private string PromptSaveExcelFileName()
        {
            var dialog = Page6CreateNewExcelFileRadioButton.IsChecked == true
                ? new SaveFileDialog()
                {
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    OverwritePrompt = true,
                    Filter = "Excel files (*.xlsx)|*.xlsx"
                }
                : new SaveFileDialog()
                {
                    DefaultExt = ".csv",
                    AddExtension = true,
                    OverwritePrompt = true,
                    Filter = "CSV files (*.csv)|*.csv"
                };

            /* Display the dialog and return the text entered by the user */
            return dialog.ShowDialog() == true
                ? dialog.FileName
                : string.Empty;
        }

        /// <summary>
        /// Extract a list of vin numbers from a DataGrid
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <param name="columnName"> The name of the source column </param>
        /// <returns> String array containing vin numbers </returns>
        private static List<string> GetVinList(in DataGrid grid, string columnName)
        {
            var columnIndex = GetDataGridColumnIndex(grid, columnName);
            var rows = grid.ItemsSource;

            return (from DataRowView row in rows select row[columnIndex] as string).ToList();
        }

        /// <summary>
        /// Get the index of a DataGrid column given the column's string name
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <param name="columnName"> The name of the source column </param>
        /// <returns> Column index </returns>
        private static int GetDataGridColumnIndex(in DataGrid grid, string columnName)
        {
            for (var i = 0; i < grid.Columns.Count; i++)
            {
                var columnHeader = grid.Columns[i].Header as string;
                if (columnHeader == columnName)
                {
                    return i;
                }
            }

            return 0;
        }

        /// <summary>
        /// Add a row of vin data values to a DataTable
        /// </summary>
        /// <param name="table"> A reference to the target DataTable </param>
        /// <param name="dataValues"> A string array of values </param>
        private static void AddVinRowToDataTable(in DataTable table, in List<string> dataValues)
        {
            if (table is null || dataValues is null)
                return;

            var row = table.NewRow();
            for (var i = 0; i < dataValues.Count; i++)
            {
                row[i] = dataValues[i];
            }

            table.Rows.Add(row);
        }

        /// <summary>
        /// Get the column headers of a DataGrid
        /// </summary>
        /// <param name="datagrid"> A reference to the source DataGrid </param>
        /// <returns> String array </returns>
        private static List<string> GetDataGridColumnNames(in DataGrid datagrid) =>
            datagrid?.Columns.Select(column => column.Header.ToString()).ToList();

        /// <summary>
        /// Get the data values in a DataGrid at the specified row
        /// </summary>
        /// <param name="datagrid"> A reference to the source DataGrid </param>
        /// <param name="rowNumber"> Index of the row </param>
        /// <returns> String array containing data values </returns>
        private static List<string> GetDataGridRowValues(in DataGrid datagrid, int rowNumber)
        {
            var dataValues = new List<string>();
            var row = datagrid?.Items[rowNumber] as DataRowView;

            for (var i = 0; i < datagrid?.Columns.Count; i++)
            {
                dataValues.Add(row?[i] as string);
            }

            return dataValues;
        }

        /// <summary>
        /// Ensure that the order of items in the DataGrid follows the specified list
        /// </summary>
        /// <param name="datagrid"> Source DataGrid reference </param>
        /// <param name="columnToOrder"> Source column name </param>
        /// <param name="orderedStringList"> String list in the correct order </param>
        private void OrderGridViewItems(in DataGrid datagrid, string columnToOrder, in List<string> orderedStringList)
        {
            ////THIS FUNCTION DROPS CORRECTED VIN NUMBERS

            /* Create a new DataTable */
            var orderedData = new DataTable();

            /* Add columns */
            var columnNames = GetDataGridColumnNames(datagrid);
            foreach (var name in columnNames)
            {
                orderedData.Columns.Add(name);
            }

            /* Add rows */
            var columnIndex = GetDataGridColumnIndex(datagrid, columnToOrder);
            foreach (var orderedVinNumber in orderedStringList)
            {
                foreach (DataRowView row in datagrid.Items)
                {
                    if (orderedVinNumber == row[columnIndex] as string)
                    {
                        orderedData.ImportRow(row.Row);
                        break;
                    }
                }
            }

            /* Replace the old ItemsSource with the new DataTable */
            datagrid.ItemsSource = orderedData.DefaultView;
        }
    }
}