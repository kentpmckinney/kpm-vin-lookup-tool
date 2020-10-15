//-----------------------------------------------------------------------
// <copyright file="MainWindow.Methods.cs">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System.Linq;
using System.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace VehicleInformationLookupTool
{
    public partial class MainWindow
    {
        private static void LaunchWebBrowser(string site)
        {
            try
            {
                Process.Start(site);
            }
            catch (Exception)
            {
                var message = new StringBuilder("Unable to open the site in a _web browser automatically.");
                message.Append(" Please visit the following _web page manually:\n\n" + site);
                message.Append("\n\nPossible Causes:\n");
                message.Append("   This application cannot access the Internet\n");
                message.Append("   The _web page has moved to a different location\n");
                message.Append("   Configuration setting or problem with the operating system\n");
                MessageBox.Show(message.ToString());
            }
        }


        private static DataTable VinTextToDataTable(ref string textContainingVins)
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
                if (!string.IsNullOrWhiteSpace(vin))
                {
                    table.Rows.Add(vin);
                }
            }

            return table;
        }


        private static bool IsFirstRun()
        {
            if (ReadRegistryString("FirstRun") == "False")
            {
                return false;
            }
            
            WriteRegistryString("FirstRun", "False");
            return true;
        }


        private static bool UserHasAgreedToEula() => 
            ReadRegistryString("AgreedEULA") == true.ToString();
        

        private static void SetUserAgreedToEula(bool state) =>
            WriteRegistryString("AgreedEULA", state.ToString());
        

        private static string ReadRegistryString(string valueName)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                var hkcu = Registry.CurrentUser;
                var software = hkcu?.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                var key = software?.OpenSubKey("VehicleInformationLookupTool",
                    RegistryKeyPermissionCheck.ReadWriteSubTree);

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


        private static void WriteRegistryString(string valueName, string valueString)
        {
            try
            {
                /* Attempt to open the key HKCU\SOFTWARE\VehicleInformationLookupTool */
                var hkcu = Registry.CurrentUser;
                var software = hkcu?.OpenSubKey("SOFTWARE", RegistryKeyPermissionCheck.ReadWriteSubTree);
                var key = software?.OpenSubKey("VehicleInformationLookupTool",
                    RegistryKeyPermissionCheck.ReadWriteSubTree);

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
                /* Ignore */
            }
        }


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


        private static List<string> GetVinList(in DataGrid grid, string columnName)
        {
            grid.ThrowIfNullOrEmpty();
            columnName.ThrowIfNullOrEmpty();

            var columnIndex = GetDataGridColumnIndex(grid, columnName);
            var rows = grid.ItemsSource;

            return (from DataRowView row in rows select row[columnIndex] as string).ToList();
        }


        private static int GetDataGridColumnIndex(in DataGrid grid, string columnName)
        {
            grid.ThrowIfNullOrEmpty();
            columnName.ThrowIfNullOrEmpty();

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


        private static void AddVinRowToDataTable(in DataTable table, in List<string> dataValues)
        {
            table.ThrowIfNullOrEmpty();
            dataValues.ThrowIfNullOrEmpty();

            var row = table.NewRow();
            if (row.Table.Columns.Count != dataValues.Count)
            {
                return;
            }

            for (var i = 0; i < row.Table.Columns.Count; i++)
            {
                row[i] = dataValues[i] ?? string.Empty;
            }

            table.Rows.Add(row);
        }


        private static List<string> GetDataGridColumnNames(in DataGrid datagrid)
        {
            datagrid.ThrowIfNullOrEmpty();

            return datagrid.Columns.Select(column => column.Header.ToString()).ToList();
        }


        private static List<string> GetDataGridRowValues(in DataGrid datagrid, int rowNumber)
        {
            datagrid.ThrowIfNullOrEmpty();

            var dataValues = new List<string>();
            var row = datagrid.Items[rowNumber] as DataRowView;

            if (row == null)
            {
                return new List<string>();
            }

            for (var i = 0; i < datagrid.Columns.Count; i++)
            {
                dataValues.Add(row[i] as string);
            }

            return dataValues;
        }


        private static void OrderGridViewItems(in DataGrid datagrid, string columnToOrder, in List<string> orderedStringList)
        {
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
                if (string.IsNullOrWhiteSpace(orderedVinNumber))
                {
                    continue;
                }

                foreach (DataRowView row in datagrid.Items)
                {
                    if (row is null)
                    {
                        continue;
                    }

                    var vinValue = row[columnIndex]?.ToString().ToLower() ?? string.Empty;
                    var originalVin = row["OriginalVIN"]?.ToString().ToLower() ?? string.Empty;
                    if (orderedVinNumber.ToLower() == vinValue || orderedVinNumber.ToLower() == originalVin)
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