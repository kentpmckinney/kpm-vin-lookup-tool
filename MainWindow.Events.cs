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
            if (IsFirstRun() == true)
                return;
            
            GoToNextPage(null, null);

            /* Skip the EULA page if the user has previously agreed to the terms
            but never on the first run of the application */
            if (UserHasAgreedToEula())
            {
                EulaCheckBox.IsChecked = true;
                GoToNextPage(null, null);
            }
        }


        private void EulaPageCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (EulaCheckBox.IsChecked == true)
            {
                SetUserAgreedToEula(true);
                EulaPageNext.IsEnabled = true;
            }
            else
            {
                SetUserAgreedToEula(false);
                EulaPageNext.IsEnabled = false;
            }
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

            /* Import data from the file */
            try
            {
                _excel.OpenFile(fileName);

                var sheets = _excel.GetSheetNames();
                foreach (var sheet in sheets)
                {
                    Page1SheetComboBox.Items.Add(sheet);
                }

                /* Display the open file, enable selection of a worksheet, and provide next step hint */
                Page1FileNameTextBox.Text = fileName;
                Page1SheetComboBox.IsEnabled = true;
                Page1Step1ComboHint.Visibility = Visibility.Visible;

                /* Attempt to auto-select a worksheet */
                Page1SheetComboBox.SelectedIndex = _excel.SheetLikelyToContainVins();
            }
            finally
            {
                _excel.CloseFile();
            }
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
            /* Clear highlighting in the scenario where the column of interest has changed */
            Page2DataGrid.ItemsSource = null;

            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                var worksheetIndex = Page1SheetComboBox.SelectedIndex;
                Page2DataGrid.ItemsSource = _excel.GetDataTable(worksheetIndex).DefaultView;
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


        private void Page2DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e == null)
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
                if (e.Column.Header.ToString().ToUpper() == "VIN")
                {
                    e.Column.CellStyle = new Style(typeof(DataGridCell));
                    e.Column.CellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Colors.LightYellow)));
                    e.Column.CellStyle.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Colors.Black)));
                }
            }
        }


        private void Page2DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            if (e == null)
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
            /* The intent of performing the test here is so that the user
               will see the Verify Internet Connectivity page only if they need to */
            Page3TestWebService_Click(null, null);
            GoToNextPage(null, null);
        }


        private void Page3TestWebService_Click(object sender, RoutedEventArgs e)
        {
            var uri = Page3WebServiceUri?.Text ?? string.Empty;

            if (_web.NhtsaServiceIsWorking(uri))
            {
                Page3StatusTextBlock.Text = "Success";
                Page3Hint.Visibility = Visibility.Hidden;
                Page3Next.IsEnabled = true;
            }
            else
            {
                Page3StatusTextBlock.Text = "Unable to connect to the NHTSA service. ";
                if (_web.IsConnectedToInternet())
                {
                    Page3StatusTextBlock.Text += "Please verify configuration settings and click Try Again.";
                }
                else
                {
                    Page3StatusTextBlock.Text += "Please verify Internet connectivity and click Try Again.";
                }
            }
        }


        private void Page3Next_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Page3Next?.IsEnabled == true)
            {
                GoToNextPage(null, null);
            }
        }


        private void Page3ResetDefault_Click(object sender, RoutedEventArgs e)
        {
            Page3WebServiceUri.Text = "http://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{VIN}?format=xml";
            Page3DataNodeXpath.Text = "/Response/Results/DecodedVINValues/*";
        }


        private void Page4ProgressBar_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            /* Runs only when the page is visible to the user and they will be able to see UI changes on the page */
            /* Run only once when ItemsSource is still null */
            if (Page4DataGrid.ItemsSource is null)
            {
                DownloadVinData();
            }
        }


        private async void DownloadVinData()
        {
            var uri = Page3WebServiceUri.Text;
            var xpath = Page3DataNodeXpath.Text;

            uri.ThrowIfNullOrEmpty();
            xpath.ThrowIfNullOrEmpty();

            /* Populate the list of VIN numbers */
            var vinList = Page1SourceExcelRadioButton.IsChecked == true
                ? GetVinList(Page2DataGrid, Page1ColumnComboBox.Text)
                : GetVinList(Page2DataGrid, "VIN");

            /* Add columns to the DataTable */
            var columnNames = _web.GetVinColumnHeaders(uri, xpath);
            foreach (var name in columnNames)
            {
                _vinData.Columns.Add(name);
            }

            /* Establish data binding between VinData and page4datagrid, which should happen
            after the columns have been added and before any rows are added*/
            Page4DataGrid.ItemsSource = _vinData.DefaultView;

            /* Set up the progress bar */
            Page4ProgressBar.Maximum = vinList.Count;
            Page4ProgressBar.Value = 0;
            Page4DownloadStatus.Text = "Status: Downloading...";

            /* Temporarily change the text of the Cancel button at the bottom to "Exit" for clarity */
            Page4Cancel.Content = "Exit";

            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();

            var autoCorrect = Page2AutoCorrectVinCheckBox?.IsChecked == true;
            var discardInvalid = Page2DiscardInvalidVinCheckBox?.IsChecked == true;

            var semaphoreSlim = new SemaphoreSlim(1, 4);
            
            var tasks = new List<Task>();
            foreach (var vinNumber in vinList)
            {
                await semaphoreSlim.WaitAsync(_downloadCancellationToken);
                tasks.Add(
                    /* Start a new task to download data for the current VIN number and when done pass the result to the next step */
                    Task<List<string>>.Factory
                        .StartNew(() => _web.GetVinDataRow(uri, vinNumber, xpath, autoCorrect, discardInvalid), _downloadCancellationToken)
                        .ContinueWith((task) =>
                        {
                            /* Proceed only if the user has not requested cancellation */
                            if (_downloadCancellationToken.IsCancellationRequested == false)
                            {
                                if (task.Result != null)
                                {
                                    AddVinRowToDataTable(_vinData, task.Result);
                                }
                                Page4ProgressBar.Value++;
                            }
                            semaphoreSlim.Release();
                        }, _downloadCancellationToken, TaskContinuationOptions.RunContinuationsAsynchronously, scheduler));
            }

            /* Update the user interface when all tasks have completed
            (Creates a new cancellation token so that it runs even if the user clicks cancel) */
            await Task.Factory.ContinueWhenAll(
                tasks.ToArray(),
                t =>
                {
                    UpdateUserInterfaceDownloadComplete();
                },
                new CancellationToken(), TaskContinuationOptions.PreferFairness, scheduler);
        }


        private void UpdateUserInterfaceDownloadComplete()
        {
            /* Ensure that the items in Page4DataGrid are in the correct order since
            the asynchronous and simultaneous downloading puts them out of order*/
            Page4DownloadStatus.Text = "Status: Re-ordering Results...";

            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                var vinList = GetVinList(Page2DataGrid, Page1ColumnComboBox.Text);
                OrderGridViewItems(Page4DataGrid, Page1ColumnComboBox.Text, vinList);
            }
            else
            {
                var vinList = GetVinList(Page2DataGrid, "VIN");
                OrderGridViewItems(Page4DataGrid, "VIN", vinList);
            }

            //Refresh the data view
            Page4DataGrid.Items.Refresh();

            /* Change the text of the Cancel button at the bottom back to "Cancel" */
            Page4Cancel.Content = "Cancel";

            /* Update the UI to reflect that the download is complete */
            Page4CancelDownloadButton.IsEnabled = false;
            Page4Hint.Visibility = Visibility.Hidden;
            Page4Save.IsEnabled = true;
            Page4Previous.IsEnabled = true;
            Page4Previous.IsEnabled = true;
            Page4DownloadStatus.Text = "Status: Completed";
            Page4ClipboardCopyButton.IsEnabled = true;
        }


        private void Page4CancelDownload_Click(object sender, RoutedEventArgs e) =>
            _downloadCancellationSource.Cancel(true);


        private void Page4ClipboardCopyButton_Click(object sender, RoutedEventArgs e)
        {
            var clipboardText = new StringBuilder();
            var newline = Environment.NewLine;
            const char tab = '\t';

            /* Add a header row to clipboardText */
            var columnNames = GetDataGridColumnNames(Page4DataGrid);
            for (var i = 0; i < columnNames?.Count; i++)
            {
                clipboardText.Append(columnNames[i]);
                if (i == columnNames.Count)
                {
                    clipboardText.Append(tab);
                }
            }
            clipboardText.Append(newline);
            
            /* Add all data rows to clipboardText */
            var numRows = Page4DataGrid.Items.Count;
            var numColumns = Page4DataGrid.Columns.Count;
            var lastColumn = numColumns - 1;

            for (var r = 0; r < numRows; r++)
            {
                var columnValues = GetDataGridRowValues(Page4DataGrid, r);
                for (var c = 0; c < numColumns; c++)
                {
                    if (columnValues[c] == null)
                    {
                        clipboardText.Append(tab);
                    }
                    else
                    {
                        var value = columnValues[c];
                        value = value.Replace(tab, ' ');
                        clipboardText.Append(value);
                        clipboardText.Append(c == lastColumn ? newline : tab.ToString());
                    }
                }
            }
            
            Clipboard.SetText(clipboardText.ToString());
        }


        private void Page4Previous_Click(object sender, RoutedEventArgs e)
        {
            GoToPreviousPage(null, null);
            if (Page3Next?.IsEnabled == true)
            {
                GoToPreviousPage(null, null);
            }
        }


        private void Page4Save_Click(object sender, RoutedEventArgs e)
        {
            var columnNames = GetDataGridColumnNames(Page4DataGrid);
            foreach (var name in columnNames)
            {
                var box = new CheckBox()
                {
                    Content = name,
                    IsChecked = true
                };
                Page5ListView.Items.Add(box);
            }

            GoToNextPage(null, null);
        }


        private void Page5CheckAllButton_Click(object sender, RoutedEventArgs e)
        {
            var buttonText = Page5CheckAllButton.Content as string;
            var isButtonCheckAll = buttonText == "Check All";

            foreach (CheckBox item in Page5ListView.Items)
            {
                item.IsChecked = isButtonCheckAll;
            }

            /* Toggle the text on the button to its logical opposite */
            Page5CheckAllButton.Content = isButtonCheckAll ? "Uncheck All" : "Check All";
        }


        private void Page6Browse_Click(object sender, RoutedEventArgs e)
        {
            var fileName = PromptSaveExcelFileName();

            Page6NewExcelFileTextBox.Text = string.IsNullOrWhiteSpace(fileName)
                ? string.Empty
                : fileName;

            if (string.IsNullOrWhiteSpace(Page6NewExcelFileTextBox.Text))
            {
                Page6Save.IsEnabled = false;
            }
            else
            {
                Page6Save.IsEnabled = true;
            }
        }


        private void Page6Save_Click(object sender, RoutedEventArgs e)
        {
            var saveFileName = Page6NewExcelFileTextBox.Text;

            /* Abort if the file name is an empty string */
            if (string.IsNullOrWhiteSpace(saveFileName))
            {
                return;
            }

            var filteredTable = new DataTable();
            var columnNames = GetDataGridColumnNames(Page4DataGrid);

            /* Add column headers to the DataTable */
            {
                foreach (var name in columnNames)
                {
                    foreach (CheckBox item in Page5ListView.Items)
                    {
                        if (item.Content as string == name && item.IsChecked == true)
                        {
                            filteredTable.Columns.Add(name);
                        }
                    }
                }
            }

            /* Add data rows to the worksheet for each row in the source DataGrid */
            var numRows = Page4DataGrid.Items.Count;
            for (var row = 0; row < numRows; row++)
            {
                var numCheckedColumns = 0;
                var filteredRow = filteredTable.NewRow();

                /* For each column in the source DataGrid via the return value of GetDataGridRowValues */
                var dataValues = GetDataGridRowValues(Page4DataGrid, row);
                for (var i = 0; i < dataValues?.Count; i++)
                {
                    var currentColumnName = columnNames[i];

                    /* Find the currentColumnName in page5ListView */
                    foreach (CheckBox item in Page5ListView.Items)
                    {
                        /* If currentColumnName is equal to the current ListView item and the item is checked */
                        if (item.Content as string == currentColumnName && item.IsChecked == true)
                        {
                            /* Add the value of the data column to filteredRow */
                            filteredRow[numCheckedColumns++] = dataValues[i];
                        }
                    }
                }

                filteredTable.Rows.Add(filteredRow);
            }

            var saveSuccessful = Page6CreateNewExcelFileRadioButton.IsChecked == true
                ? _excel.SaveExcelFile(saveFileName, filteredTable)
                : _excel.SaveCsvFile(saveFileName, filteredTable);
            
            if (saveSuccessful)
            {
                Page6FileOpenHint.Visibility = Visibility.Hidden;
                Page7SavedFileTextBox.Text = saveFileName;
                GoToNextPage(null, null);
            }
            else
            {
                Page6FileOpenHint.Visibility = Visibility.Visible;
            }
        }
    }
}