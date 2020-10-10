//-----------------------------------------------------------------------
// <copyright file="MainWindow.Methods.cs" company="N/A">
//     Copyright (c) 2016 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics;
    using System.Windows;
    using System.Windows.Controls;
    using Microsoft.Win32;

    /// <summary>
    /// Helper methods for MainWindow
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Opens the default web browser to the specified site
        /// </summary>
        /// <param name="site"> The web site address </param>
        private void LaunchWebBrowser(string site)
        {
            try
            {
                Process.Start(site);
            }
            catch (Exception)
            {
                string message = "Unable to open the site in a web browser automatically.";
                message += " Please visit the following web page manually:\n\n" + site;
                message += "\n\nPossible Causes:\n";
                message += "   This application cannot access the Internet\n";
                message += "   The web page has moved to a different location\n";
                message += "   Configuration setting or problem with the operating system\n";
                MessageBox.Show(message);
            }
        }

        /// <summary>
        /// Consumes the contents of the text box containing VINs converting it to a more usable form
        /// </summary>
        /// <param name="textContainingVINs"> Raw text input </param>
        /// <returns> A DataTable with the VINs in a single column </returns>
        private DataTable VinTextToDataTable(string textContainingVINs)
        {
            /*
              Requirement: the user must only use commas, semicolons, and newlines to delineate VINs (as indicated in the UI)
              This application does not attempt to validate VIN numbers
            */

            /* Normalize the text so that there is one VIN per line */
            textContainingVINs = textContainingVINs.Replace(";", "\n");
            textContainingVINs = textContainingVINs.Replace(",", "\n");
            textContainingVINs = textContainingVINs.Replace(Environment.NewLine, "\n");
            textContainingVINs = textContainingVINs.Replace("\n\n\n", "\n");
            textContainingVINs = textContainingVINs.Replace("\n\n", "\n");
            textContainingVINs = textContainingVINs.Replace(" ", string.Empty);

            /* Convert the text into a list of strings */
            string[] vinList = textContainingVINs.Split('\n');

            /* Create a DataTable */
            DataTable table = new DataTable();

            /* Add columns */
            table.Columns.Add("VIN");

            /* Add rows */
            foreach (string vin in vinList)
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
            if (this.ReadRegistryString("FirstRun") == "False")
            {
                return false;
            }
            
            this.WriteRegistryString("FirstRun", "False");
            return true;
        }

        /// <summary>
        /// Reads the registry value that indicates whether the user has agreed to the EULA
        /// </summary>
        /// <returns> True or false </returns>
        private bool UserHasAgreedToEULA()
        {
            bool state = this.ReadRegistryString("AgreedEULA") == true.ToString();
            return state;
        }

        /// <summary>
        /// Writes the registry value to indicate whether the user has agreed to the EULA
        /// </summary>
        /// <param name="state"> True or false </param>
        private void SetUserAgreedToEULA(bool state)
        {
            this.WriteRegistryString("AgreedEULA", state.ToString());
        }

        /// <summary>
        /// Reads a string value from the registry key HKCU\SOFTWARE\VehicleInformationLookupTool
        /// </summary>
        /// <param name="valueName"> The name of the registry value </param>
        /// <returns> A string that is empty or which contains the specified registry value </returns>
        private string ReadRegistryString(string valueName)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                RegistryKey hkcu = Registry.CurrentUser;
                RegistryKey software = hkcu.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                RegistryKey key = software.OpenSubKey("VehicleInformationLookupTool", RegistryKeyPermissionCheck.ReadWriteSubTree);

                if (key != null)
                {
                    /* Attempt to get the named value */
                    string valueString = key.GetValue(valueName) as string;

                    key.Close();
                    return valueString;
                }
                
                return string.Empty;
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
        private void WriteRegistryString(string valueName, string valueString)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                RegistryKey hkcu = Registry.CurrentUser;
                RegistryKey software = hkcu.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                RegistryKey key = software.OpenSubKey("VehicleInformationLookupTool", RegistryKeyPermissionCheck.ReadWriteSubTree);

                /* If the key was not found attempt to create it */
                if (key == null)
                {
                    key = software.CreateSubKey("VehicleInformationLookupTool");
                }

                if (key != null)
                {
                    /* Attempt to create the named value */
                    key.SetValue(valueName, valueString);

                    key.Close();
                }
            }
            catch (Exception)
            {
                // Do nothing
            }
        }

        /// <summary>
        /// Prompts the user for the name of an Excel file to open
        /// </summary>
        /// <returns> A string that is empty or which contains the name of a file </returns>
        private string PromptOpenExcelFileName()
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                DefaultExt = ".xlsx",
                Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
            };

            /* Display the dialog */ 
            bool? result = dialog.ShowDialog();

            /* Get the selected file name */
            string fileName = string.Empty;
            if (result == true)
            {
                fileName = dialog.FileName;
            }

            return fileName;
        }

        /// <summary>
        /// Prompts the user for the name of an Excel file to which to save data
        /// </summary>
        /// <returns> A string that is empty or which contains the name of a file </returns>
        private string PromptSaveExcelFileName()
        {
            SaveFileDialog dialog;

            if (page6CreateNewExcelFileRadioButton.IsChecked == true)
            {
                dialog = new SaveFileDialog()
                {
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    OverwritePrompt = true,
                    Filter = "Excel files (*.xlsx)|*.xlsx" //"Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
                };
            }
            else
            {
                dialog = new SaveFileDialog()
                {
                    DefaultExt = ".csv",
                    AddExtension = true,
                    OverwritePrompt = true,
                    Filter = "CSV files (*.csv)|*.csv"
                };
            }

            /* Display the dialog */
            bool? result = dialog.ShowDialog();

            /* Get the selected file name */
            string fileName = string.Empty;
            if (result == true)
            {
                fileName = dialog.FileName;
            }

            return fileName;
        }

        /// <summary>
        /// Extract a list of vin numbers from a DataGrid
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <param name="columnName"> The name of the source column </param>
        /// <returns> String array containing vin numbers </returns>
        private List<string> GetVinList(DataGrid grid, string columnName)
        {
            int columnIndex = this.GetDataGridColumnIndex(grid, columnName);

            List<string> vinList = new List<string>();
            IEnumerable rows = grid.ItemsSource;
             foreach (DataRowView row in rows)
            {
                vinList.Add(row[columnIndex] as string);
            }
            
            return vinList;
        }

        /// <summary>
        /// Get the index of a DataGrid column given the column's string name
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <param name="columnName"> The name of the source column </param>
        /// <returns> Column index </returns>
        private int GetDataGridColumnIndex(DataGrid grid, string columnName)
        {
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                string columnHeader = grid.Columns[i].Header as string;
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
        private void AddVinRowToDataTable(DataTable table, List<string> dataValues)
        {
            if (dataValues == null)
            {
                return;
            }

            DataRow row = table.NewRow();
            for (int i = 0; i < dataValues.Count; i++)
            {
                row[i] = dataValues[i];
            }

            table.Rows.Add(row);
        }
        
        /// <summary>
        /// Get the column headers of a DataGrid
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <returns> String array </returns>
        private List<string> GetDataGridColumnNames(DataGrid grid)
        {
            List<string> headers = new List<string>();

            for (int i = 0; i < grid.Columns.Count; i++)
            {
                string header = grid.Columns[i].Header.ToString();
                headers.Add(header);
            }

            return headers;
        }

        /// <summary>
        /// Get the data values in a DataGrid at the specified row
        /// </summary>
        /// <param name="grid"> A reference to the source DataGrid </param>
        /// <param name="rowNumber"> Index of the row </param>
        /// <returns> String array containing data values </returns>
        private List<string> GetDataGridRowValues(DataGrid grid, int rowNumber)
        {
            List<string> dataValues = new List<string>();
            DataRowView row = (DataRowView)grid.Items[rowNumber];

            for (int i = 0; i < grid.Columns.Count; i++)
            {
                dataValues.Add(row[i] as string);
            }

            return dataValues;
        }

        /// <summary>
        /// Ensure that the order of items in the DataGrid follows the specified list
        /// </summary>
        /// <param name="grid"> Source DataGrid reference </param>
        /// <param name="columnToOrder"> Source column name </param>
        /// <param name="orderedStringList"> String list in the correct order </param>
        private void OrderGridViewItems(DataGrid grid, string columnToOrder, List<string> orderedStringList)
        {
            ////THIS FUNCTION DROPS CORRECTED VIN NUMBERS

            /* Create a new DataTable */
            DataTable orderedData = new DataTable();

            /* Add columns */
            List<string> columnNames = this.GetDataGridColumnNames(grid);
            foreach (string name in columnNames)
            {
                orderedData.Columns.Add(name);
            }

            /* Add rows */
            int columnIndex = this.GetDataGridColumnIndex(grid, columnToOrder);
            foreach (string orderedVinNumber in orderedStringList)
            {
                foreach (DataRowView row in grid.Items)
                {
                    if (orderedVinNumber == row[columnIndex] as string)
                    {
                        orderedData.ImportRow(row.Row);
                        break;
                    }
                }
            }

            /* Replace the old ItemsSource with the new DataTable */
            grid.ItemsSource = orderedData.DefaultView;
        }
    }
}