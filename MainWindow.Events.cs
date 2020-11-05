//-----------------------------------------------------------------------
// <copyright file="MainWindow.Events.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using VehicleInformationLookupTool.Properties;

namespace VehicleInformationLookupTool
{
    public partial class MainWindow
    {
        private void Window_Closing(object sender, CancelEventArgs e) =>
            Settings.Default.Save();


        private void ShutdownApplication(object sender, RoutedEventArgs e)
        {
            /* Cancel download if in progress */
            _downloadCancellationSource.Cancel(false);

            /* Perform the shutdown */
            Application.Current.Shutdown();
        }


        private void GoToNextPage(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() => TabControl.SelectedIndex++));
        }
        

        private void GoToPreviousPage(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() => TabControl.SelectedIndex--));
        }
        

        private void Hyperlink_Click(object sender, RoutedEventArgs e) =>
            LaunchWebBrowser((e.Source as Hyperlink)?.NavigateUri.OriginalString);


        private void Help_Click(object sender, RoutedEventArgs e) =>
            HelpClass.ShowTopic((sender as FrameworkElement)?.Tag as string);


        private void AboutPageTabLoaded(object sender, RoutedEventArgs e)
        {
            /* Show the about page on the first run of the application and skip otherwise */
            if (!IsFirstRun())
            {
                GoToNextPage(null, null);
            }

            /* Skip the EULA page if the user has previously agreed to the terms */
            if (UserHasAgreedToEula())
            {
                EulaCheckBox.IsChecked = true;
                GoToNextPage(null, null);
            }
        }


        private void EulaPageCheckBox_Click(object sender, RoutedEventArgs e)
        {
            SetUserAgreedToEula(EulaCheckBox.IsChecked == true);
            EulaPageNext.IsEnabled = EulaCheckBox.IsChecked == true;
        }


        private void Page1Browse_Click(object sender, RoutedEventArgs e)
        {
            /* Programmatically check the radio button to import from Excel */
            Page1SourceExcelRadioButton.IsChecked = true;

            /* Prompt to select the file */
            var fileName = PromptOpenExcelFileName();
            if (string.IsNullOrWhiteSpace(fileName) || !File.Exists(fileName))
            {
                return;
            }

            /* Open the file which remains open until the user exits the application or clicks Next */
            _excel.OpenFile(fileName);

            /* Get sheet names */
            var sheets = _excel.GetSheetNames();
            foreach (var sheet in sheets)
            {
                Page1SheetComboBox.Items.Add(sheet);
            }

            /* Display the open file, enable selection of a worksheet, and provide next step hint */
            Page1FileNameTextBox.Text = fileName;
            Page1SheetComboBox.IsEnabled = true;
            Page1Step1ComboHint.Visibility = Visibility.Visible;

            /* Look for a sheet with a column that is likely to contain VIN numbers and set to combo box to display that sheet */
            Page1SheetComboBox.SelectedIndex = _excel.SheetLikelyToContainVins();
        }


        private void Page1SheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedSheet = Page1SheetComboBox.SelectedItem.ToString();
            if (string.IsNullOrWhiteSpace(selectedSheet))
            {
                return;
            }

            var columnNames = _excel.GetColumnNames(selectedSheet);
            if (columnNames == null)
            {
                return;
            }

            foreach (var name in columnNames)
            {
                Page1ColumnComboBox.Items.Add(name);
            }

            /* Enable selection of a column and provide next step hint*/
            Page1ColumnComboBox.IsEnabled = true;
            Page1Step1ComboHint.Visibility = Visibility.Hidden;
            Page1Step2ComboHint.Visibility = Visibility.Visible;

            /* Attempt to auto-select a column */
            Page1ColumnComboBox.SelectedIndex = _excel.ColumnLikelyToContainVins(selectedSheet);
        }


        private void Page1ColumnComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Page1Next.IsEnabled = true;
            Page1Step2ComboHint.Visibility = Visibility.Hidden;
            Page1Hint.Visibility = Visibility.Hidden;
        }


        private void Page1TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Page1Hint.Visibility = Visibility.Hidden;
            Page1Next.IsEnabled = true;

            /* Programmatically check the radio button indicating text source */
            Page1SourceTextRadioButton.IsChecked = true;
        }


        private void Page1Next_Click(object sender, RoutedEventArgs e)
        {
            _navigateDirection = Direction.Forward;

            /* Set button state */
            Page2ValidCheckBox.IsChecked = false;
            Page2Next.IsEnabled = false;

            /* Clear highlighting in the scenario where the column of interest has changed */
            Page2DataGrid.ItemsSource = null;

            /* Set data binding for the DataGrid on the next page */
            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                var worksheetIndex = Page1SheetComboBox.SelectedIndex;
                Page2DataGrid.ItemsSource = _excel.GetDataTable(worksheetIndex).DefaultView;
                _excel.CloseFile();
            }
            else
            {
                var text = Page1TextBox.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                Page2DataGrid.ItemsSource = VinTextToDataTable(ref text).DefaultView;
            }

            GoToNextPage(null, null);
        }


        private void Page1_Selected(object sender, RoutedEventArgs e)
        {
            if (_navigateDirection is Direction.Backward && Page2DataGrid.ItemsSource != null)
            {
                Page2DataGrid.ItemsSource = default;
                Page2DataGrid.Items.Clear();
            }
        }


        private void Page2_Selected(object sender, RoutedEventArgs e)
        {
            if (_navigateDirection is Direction.Backward && Page3DataGrid.ItemsSource != null)
            {
                Page3DataGrid.ItemsSource = null;
                Page3DataGrid.Items.Clear();
            }
        }


        private void Page2DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e is null)
            {
                return;
            }

            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                /* In this scenario the column being automatically generated matches the selected column */
                if (e.Column.Header.ToString() == Page1ColumnComboBox.SelectedItem.ToString())
                {
                    e.Column.CellStyle = new Style(typeof(DataGridCell));
                    e.Column.CellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Colors.LightYellow)));
                    e.Column.CellStyle.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Colors.Black)));
                }
            }
            else
            {
                /* In this scenario there is only one column and it is named VIN */
                if (e.Column.Header.ToString().ToLower() == "vin")
                {
                    e.Column.CellStyle = new Style(typeof(DataGridCell));
                    e.Column.CellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Colors.LightYellow)));
                    e.Column.CellStyle.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Colors.Black)));
                }
            }
        }


        private void Page2DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e is null)
            {
                return;
            }

            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }


        private void Page2ValidCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (Page2ValidCheckBox.IsChecked == true)
            {
                Page2Next.IsEnabled = true;
                Page2Hint.Visibility = Visibility.Hidden;
            }
            else
            {
                Page2Next.IsEnabled = false;
                Page2Hint.Visibility = Visibility.Visible;
            }
        }


        private void Page2Next_Click(object sender, RoutedEventArgs e)
        {
            _navigateDirection = Direction.Forward;
            GoToNextPage(null, null);
        }


        private void Page3_Selected(object sender, RoutedEventArgs e)
        {
            /* Runs only when the page is visible to the user and they will be able to see UI changes on the page */

            /* Run only when ItemsSource is null */
            if (Page3DataGrid.ItemsSource is null)
            {
                DownloadVinData();
            }
        }


        private async void DownloadVinData()
        {
            /* Set button state */
            Page3Next.IsEnabled = false;
            Page3Previous.IsEnabled = false;
            Page3ClipboardCopyButton.IsEnabled = false;
            Page3CancelDownloadButton.IsEnabled = true;

            var uri = "http://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{VIN}?format=xml";
            var xpath = "/Response/Results/DecodedVINValues/*";

            /* Populate the list of VIN numbers */
            var vinList = Page1SourceExcelRadioButton.IsChecked == true
                ? GetVinList(Page2DataGrid, Page1ColumnComboBox.Text)
                : GetVinList(Page2DataGrid, "VIN");

            /* Clear existing columns and rows from the DataTable */
            _vinData.Columns.Clear();
            _vinData.Rows.Clear();

            /* Add columns to the DataTable */
            var columnNames = _web.GetVinColumnHeaders(uri, xpath, _downloadCancellationToken);
            if (columnNames is null)
            {
                return;
            }
            foreach (var name in columnNames)
            {
                if (name is null)
                {
                    continue;
                }
                _vinData.Columns.Add(name);
            }

            /* Establish data binding between VinData and Page4DataGrid */
            /* This should happen after the columns have been added and before any rows are added*/
            Page3DataGrid.ItemsSource = _vinData.DefaultView;

            /* Set up the progress bar */
            Page3ProgressBar.Maximum = vinList.Count;
            Page3ProgressBar.Value = 0;
            Page3DownloadStatus.Text = "Status: Downloading...";

            /* Temporarily change the text of the Cancel button at the bottom to "Exit" for clarity */
            Page4Cancel.Content = "Exit";

            var autoCorrect = Page2AutoCorrectVinCheckBox?.IsChecked == true;
            var discardInvalid = Page2DiscardInvalidVinCheckBox?.IsChecked == true;

            /* Run download tasks in parallel */
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var semaphoreSlim = new SemaphoreSlim(1, 4);
            var tasks = new List<Task>();
            foreach (var vinNumber in vinList)
            {
                try
                {
                    await semaphoreSlim.WaitAsync(_downloadCancellationToken);
                    tasks.Add(
                        /* Start a new task to download data for the current VIN number and when done pass the result to the next step */
                        Task<List<string>>.Factory
                            .StartNew(() => _web.GetVinDataRow(uri, vinNumber, xpath, autoCorrect, discardInvalid, _downloadCancellationToken),
                                _downloadCancellationToken)
                            .ContinueWith((task) =>
                                {
                                    /* Proceed only if the user has not requested cancellation */
                                    if (_downloadCancellationToken.IsCancellationRequested == false)
                                    {
                                        if (task.Result != null)
                                        {
                                            AddVinRowToDataTable(_vinData, task.Result);
                                        }

                                        Page3ProgressBar.Value++;
                                    }

                                    semaphoreSlim.Release();
                                }, _downloadCancellationToken, TaskContinuationOptions.RunContinuationsAsynchronously,
                                scheduler));
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }

            /* Update the user interface when all tasks have completed */ 
            /* Creates a new cancellation token so that it runs even if the user clicks cancel */
            await Task.Factory.ContinueWhenAll(
                tasks.ToArray(),
                t => UpdateUserInterfaceDownloadComplete(),
                new CancellationToken(), TaskContinuationOptions.PreferFairness, scheduler);
        }


        private void UpdateUserInterfaceDownloadComplete()
        {
            /* Ensure that the items in Page4DataGrid are in the correct order since
            the asynchronous and simultaneous downloading puts them out of order*/
            Page3DownloadStatus.Text = "Status: Re-ordering Results...";

            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                var vinList = GetVinList(Page2DataGrid, Page1ColumnComboBox.Text);
                OrderGridViewItems(Page3DataGrid, Page1ColumnComboBox.Text, vinList);
            }
            else
            {
                var vinList = GetVinList(Page2DataGrid, "VIN");
                OrderGridViewItems(Page3DataGrid, "VIN", vinList);
            }

            //Refresh the data view
            Page3DataGrid.Items.Refresh();

            /* Change the text of the Cancel button at the bottom back to "Cancel" */
            Page3Cancel.Content = "Cancel";

            /* Update the UI to reflect that the download is complete */
            Page3CancelDownloadButton.IsEnabled = false;
            Page3Hint.Visibility = Visibility.Hidden;
            Page3Next.IsEnabled = true;
            Page3Previous.IsEnabled = true;
            Page3Previous.IsEnabled = true;
            Page3DownloadStatus.Text = "Status: Completed";
            Page3ClipboardCopyButton.IsEnabled = true;
        }


        private void Page3CancelDownload_Click(object sender, RoutedEventArgs e) =>
            _downloadCancellationSource.Cancel(true);


        private void Page3ClipboardCopyButton_Click(object sender, RoutedEventArgs e)
        {
            var clipboardText = new StringBuilder();
            var newline = Environment.NewLine;
            const char delimiter = '\t';

            /* Add a header row to clipboardText */
            var columnNames = GetDataGridColumnNames(Page3DataGrid);
            if (columnNames is null)
            {
                return;
            }

            for (var i = 0; i < columnNames.Count; i++)
            {
                clipboardText.Append(columnNames[i]); //var lastColumn = columnNames.Count
                if (i != columnNames.Count)
                {
                    clipboardText.Append(delimiter);
                }
            }
            clipboardText.Append(newline);

            /* Add data rows to clipboardText */
            var numRows = Page3DataGrid.Items.Count;
            var numColumns = Page3DataGrid.Columns.Count;
            if (numRows <= 0 || numColumns <= 0)
            {
                return;
            }
            var lastColumn = numColumns - 1;
            for (var r = 0; r < numRows; r++)
            {
                var columnValues = GetDataGridRowValues(Page3DataGrid, r);
                if (columnValues is null || columnValues.Count <= 0)
                {
                    continue;
                }
                for (var c = 0; c < numColumns; c++)
                {
                    var value = columnValues[c] ?? string.Empty;
                    value = value.Replace(delimiter, ' ');
                    clipboardText.Append(value);
                    clipboardText.Append(c == lastColumn ? newline : delimiter.ToString());
                }
            }
            
            /* Copy to the clipboard */
            Clipboard.SetText(clipboardText.ToString(), TextDataFormat.Text);
        }


        private void Page3Previous_Click(object sender, RoutedEventArgs e)
        {
            _navigateDirection = Direction.Backward;
            GoToPreviousPage(null, null);
        }


        private void Page3Next_Click(object sender, RoutedEventArgs e)
        {
            _navigateDirection = Direction.Forward;

            var columnNames = GetDataGridColumnNames(Page3DataGrid);
            if (columnNames is null || columnNames.Count <= 0)
            {
                return;
            }
            foreach (var name in columnNames)
            {
                var box = new CheckBox()
                {
                    Content = name,
                    IsChecked = true
                };
                Page4ListView.Items.Add(box);
            }

            GoToNextPage(null, null);
        }


        private void Page4CheckAllButton_Click(object sender, RoutedEventArgs e)
        {
            var buttonText = Page4CheckAllButton.Content as string;
            var isButtonCheckAll = buttonText == "Check All";

            /* Check or uncheck all checkboxes */
            foreach (CheckBox item in Page4ListView.Items)
            {
                item.IsChecked = isButtonCheckAll;
            }

            /* Toggle the text on the button to its logical opposite */
            Page4CheckAllButton.Content = isButtonCheckAll ? "Uncheck All" : "Check All";
        }


        private void Page5Browse_Click(object sender, RoutedEventArgs e)
        {
            var fileName = PromptSaveExcelFileName(Page5CreateNewExcelFileRadioButton.IsChecked == true);

            Page5NewExcelFileTextBox.Text = string.IsNullOrWhiteSpace(fileName)
                ? string.Empty
                : fileName;

            /* Enable or disable the Save button */
            if (string.IsNullOrWhiteSpace(Page5NewExcelFileTextBox.Text))
            {
                Page5Save.IsEnabled = false;
            }
            else
            {
                Page5Save.IsEnabled = true;
            }
        }


        private void Page5Save_Click(object sender, RoutedEventArgs e)
        {
            var saveFileName = Page5NewExcelFileTextBox.Text;

            /* Abort if the file name is an empty string or whitespace */
            if (string.IsNullOrWhiteSpace(saveFileName))
            {
                return;
            }

            var filteredTable = new DataTable();

            /* Add column headers to the DataTable */
            var columnNames = GetDataGridColumnNames(Page3DataGrid);
            if (columnNames is null || columnNames.Count <= 0)
            {
                return;
            }
            foreach (var name in columnNames)
            {
                foreach (CheckBox item in Page4ListView.Items)
                {
                    if (item is null)
                    {
                        continue;
                    }

                    if (item.Content as string == name && item.IsChecked == true)
                    {
                        filteredTable.Columns.Add(name);
                    }
                }
            }
            
            /* Add data rows to the worksheet for each row in the source DataGrid */
            var numRows = Page3DataGrid.Items.Count;
            for (var row = 0; row < numRows; row++)
            {
                var numCheckedColumns = 0;
                var filteredRow = filteredTable.NewRow();

                /* For each column in the source DataGrid via the return value of GetDataGridRowValues */
                var dataValues = GetDataGridRowValues(Page3DataGrid, row);
                if (dataValues is null || dataValues.Count <= 0)
                {
                    continue;
                }
                for (var i = 0; i < dataValues.Count; i++)
                {
                    var currentColumnName = columnNames[i];
                    if (currentColumnName is null)
                    {
                        continue;
                    }

                    /* Find the currentColumnName in page5ListView */
                    foreach (CheckBox item in Page4ListView.Items)
                    {
                        if (item is null)
                        {
                            continue;
                        }

                        /* If currentColumnName is equal to the current ListView item and the item is checked */
                        if (item.Content as string == currentColumnName && item.IsChecked == true)
                        {
                            /* Add the value of the data column to filteredRow */
                            filteredRow[numCheckedColumns++] = dataValues[i];
                        }
                    }
                }

                /* Add the row */
                filteredTable.Rows.Add(filteredRow);
            }

            var saveSuccessful = Page5CreateNewExcelFileRadioButton.IsChecked == true
                ? _excel.SaveExcelFile(saveFileName, filteredTable)
                : _excel.SaveCsvFile(saveFileName, filteredTable);
            
            if (saveSuccessful)
            {
                Page5FileOpenHint.Visibility = Visibility.Hidden;
                Page6SavedFileTextBox.Text = saveFileName;
                GoToNextPage(null, null);
            }
            else
            {
                Page5FileOpenHint.Visibility = Visibility.Visible;
            }
        }
    }
}