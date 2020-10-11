//-----------------------------------------------------------------------
// <copyright file="MainWindow.Events.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
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
    using Properties;

    /// <summary>
    /// Events for MainWindow
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>
        /// MainWindow is closing
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Window_Closing(object sender, CancelEventArgs e) =>
            Settings.Default.Save();

        /// <summary>
        /// Shut down the application
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void ShutdownApplication(object sender, RoutedEventArgs e)
        {
            /* Cancel download if in progress */
            _downloadCancellationSource.Cancel(false);

            /* Perform the shutdown */
            Application.Current.Shutdown();
        }

        /// <summary>
        /// Advance to the next page
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void GoToNextPage(object sender, RoutedEventArgs e) =>
            Dispatcher.BeginInvoke(new Action(() => TabControl.SelectedIndex++));
        
        /// <summary>
        /// Go back to the previous page
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void GoToPreviousPage(object sender, RoutedEventArgs e) =>
            Dispatcher.BeginInvoke(new Action(() => TabControl.SelectedIndex--));

        /// <summary>
        /// Opens the NavigateUri parameter of a Hyperlink element in the default browser
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Hyperlink_Click(object sender, RoutedEventArgs e) =>
            LaunchWebBrowser((e.Source as Hyperlink)?.NavigateUri.OriginalString);

        /// <summary>
        /// Displays the help file and shows the topic specified in the Tag property of the sender
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Help_Click(object sender, RoutedEventArgs e) =>
            HelpClass.ShowTopic((sender as FrameworkElement)?.Tag as string);

        /// <summary>
        /// Skip the about and EULA pages if the appropriate conditions are met
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
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

        /// <summary>
        /// Toggles the visibility of the Next button based on whether the EULA agreement checkbox is checked or not
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void EulaPageCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (EulaCheckBox.IsChecked == true)
            {
                this.SetUserAgreedToEula(true);
                EulaPageNext.IsEnabled = true;
            }
            else
            {
                this.SetUserAgreedToEula(false);
                EulaPageNext.IsEnabled = false;
            }
        }

        /// <summary>
        /// Populate worksheet names in response to the user opening an Excel file
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page1Browse_Click(object sender, RoutedEventArgs e)
        {
            /* Programmatically check the radio button to import from Excel */
            Page1SourceExcelRadioButton.IsChecked = true;

            var fileName = PromptOpenExcelFileName();
            if (!File.Exists(fileName))
                return;
            
            _excel.OpenFile(fileName);

            if (_excel.IsValidFile())
            {
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

            _excel.CloseFile();
            
        }

        /// <summary>
        /// Populate column names in response to the user selecting an Excel worksheet
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page1SheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedSheet = Page1SheetComboBox.SelectedItem.ToString();
            var columnNames = _excel.GetColumnNames(selectedSheet);
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

        /// <summary>
        /// Enable the Next button on the Enter VIN Numbers page and hide hints
        /// in response to an Excel column having been selected
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page1ColumnComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Page1Next.IsEnabled = true;
            Page1Step2ComboHint.Visibility = Visibility.Hidden;
            Page1Hint.Visibility = Visibility.Hidden;
        }

        /// <summary>
        /// Enable the Next button on the Enter VIN Numbers page and hide hints
        /// in response to the user entering text
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page1TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Page1Hint.Visibility = Visibility.Hidden;
            Page1Next.IsEnabled = true;

            /* Programmatically check the radio button indicating text source */
            Page1SourceTextRadioButton.IsChecked = true;
        }

        /// <summary>
        /// Populate the next page's DataGrid before moving on to that page
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
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
                Page2DataGrid.ItemsSource = VinTextToDataTable(Page1TextBox.Text).DefaultView;
            }

            GoToNextPage(null, null);
        }

        /// <summary>
        /// Highlight a column of the DataGrid
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page2DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (Page1SourceExcelRadioButton.IsChecked == true)
            {
                /* In this scenario the column being automatically generated matches the selected column */
                if (e.Column.Header.ToString() == Page1ColumnComboBox.SelectedItem.ToString())
                {
                    e.Column.CellStyle = new System.Windows.Style(typeof(DataGridCell));
                    e.Column.CellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Colors.LightYellow)));
                    e.Column.CellStyle.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Colors.Black)));
                }
            }
            else
            {
                /* In this scenario there is only one column and it is named VIN */
                if (e.Column.Header.ToString().ToUpper() == "VIN")
                {
                    e.Column.CellStyle = new System.Windows.Style(typeof(DataGridCell));
                    e.Column.CellStyle.Setters.Add(new Setter(BackgroundProperty, new SolidColorBrush(Colors.LightYellow)));
                    e.Column.CellStyle.Setters.Add(new Setter(ForegroundProperty, new SolidColorBrush(Colors.Black)));
                }
            }
        }

        /// <summary>
        /// Add numbered row headers
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page2DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        /// <summary>
        /// Enables or disables the Next button in response to the validation checkbox being checked or unchecked 
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
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

        /// <summary>
        /// Perform _web service test and then proceed to the next page
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page2Next_Click(object sender, RoutedEventArgs e)
        {
            /* The intent of performing the test here is so that the user
            will see the Verify Internet Connectivity page only if they need to */
            Page3TestWebService_Click(null, null);
            GoToNextPage(null, null);
        }

        /// <summary>
        /// Test the _web service
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page3TestWebService_Click(object sender, RoutedEventArgs e)
        {
            var uri = Page3WebServiceUri.Text;

            if (_web.NhtsaServiceIsWorking(uri))
            {
                Page3StatusTextBlock.Text = "Success";
                Page3Hint.Visibility = Visibility.Hidden;
                Page3Next.IsEnabled = true;
            }
            else
            {
                Page3StatusTextBlock.Text = "Unable to connect to the NHTSA service. ";
                if (_web.IsConnectedToInternet() == true)
                {
                    Page3StatusTextBlock.Text += "Please verify configuration settings and click Try Again.";
                }
                else
                {
                    Page3StatusTextBlock.Text += "Please verify Internet connectivity and click Try Again.";
                }
            }
        }

        /// <summary>
        /// Automatically advance to the next page once the Verify Internet Connectivity page's next button is enabled
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page3Next_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Page3Next?.IsEnabled == true)
            {
                GoToNextPage(null, null);
            }
        }

        /// <summary>
        /// Replace the values in the uri and xpath text boxes with default values
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page3ResetDefault_Click(object sender, RoutedEventArgs e)
        {
            Page3WebServiceUri.Text = "http://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{VIN}?format=xml";
            Page3DataNodeXpath.Text = "/Response/Results/DecodedVINValues/*";
            // test: https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/5UXWX7C5*BA?format=xml
        }

        /// <summary>
        /// Initiate the vin data download process
        /// </summary>
        /// <remarks>
        /// Runs only when the page is visible to the user and they will be able to see UI changes on the page
        /// </remarks>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page4ProgressBar_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            /* Run only once when ItemsSource is still null */
            if (Page4DataGrid.ItemsSource is null)
            {
                DownloadVinData();
            }
        }

        /// <summary>
        /// Download vin data
        /// </summary>
        private void DownloadVinData()
        {
            var uri = Page3WebServiceUri.Text;
            var xpath = Page3DataNodeXpath.Text;

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

            using (var concurrencySemaphore = new SemaphoreSlim(1, 4))
            {
                var tasks = new List<Task>();
                foreach (var vinNumber in vinList)
                {
                    // Each task added runs in a separate thread 
                    tasks.Add(
                        /* Start a new task to download data for the current VIN number and when done pass the result to the next step */
                        Task<List<string>>.Factory
                            .StartNew(() => _web.GetVinDataRow(uri, vinNumber, xpath, autoCorrect, discardInvalid), _downloadCancellationToken)
                            .ContinueWith((task) =>
                            {
                                // Proceed only if the user has not requested cancellation
                                if (_downloadCancellationToken.IsCancellationRequested == false)
                                {
                                    AddVinRowToDataTable(_vinData, task.Result);
                                    Page4ProgressBar.Value++;
                                }
                            }, _downloadCancellationToken, TaskContinuationOptions.PreferFairness, scheduler));
                }

                /* Update the user interface when all tasks have completed
                (Creates a new cancellation token so that it runs even if the user clicks cancel) */
                Task.Factory.ContinueWhenAll(
                    tasks.ToArray(),
                    t => UpdateUserInterfaceDownloadComplete(),
                    new CancellationToken(), TaskContinuationOptions.PreferFairness, scheduler);
            }
        }

        /// <summary>
        /// Updates the UI when the download of VIN data is complete
        /// </summary>
        private void UpdateUserInterfaceDownloadComplete()
        {
            /* Ensure that the items in page4datagrid are in the correct order since
            the asynchronous and simultaneous downloading puts them out of order*/
            //page4DownloadStatus.Text = "Status: Re-ordering Results...";
            //if (page1SourceExcelRadioButton.IsChecked == true)
            //{
            //    List<string> vinList = this.GetVinList(this.page2DataGrid, page1ColumnComboBox.Text);
            //    this.OrderGridViewItems(this.page4DataGrid, page1ColumnComboBox.Text, vinList);
            //}
            //else
            //{
            //    List<string> vinList = this.GetVinList(this.page2DataGrid, "VIN");
            //    this.OrderGridViewItems(this.page4DataGrid, "VIN", vinList);
            //}

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

        /// <summary>
        /// Cancel the download process
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page4CancelDownload_Click(object sender, RoutedEventArgs e) =>
            _downloadCancellationSource.Cancel(true);

        /// <summary>
        /// Copy all downloaded VIN data to the clipboard
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page4ClipboardCopyButton_Click(object sender, RoutedEventArgs e)
        {
            var clipboardText = new StringBuilder();
            var newline = Environment.NewLine;
            const char tab = '\t';

            /* Add a header row to clipboardText */
            {
                var columnNames = GetDataGridColumnNames(Page4DataGrid);
                for (var i = 0; i < columnNames.Count; i++)
                {
                    if (i == columnNames.Count)
                    {
                        clipboardText.Append(columnNames[i]);
                    }
                    else
                    {
                        clipboardText.Append(columnNames[i] + tab);
                    }
                }

                clipboardText.Append(newline);
            }

            /* Add all data rows to clipboardText */
            {
                var numRows = Page4DataGrid.Items.Count;
                var numColumns = Page4DataGrid.Columns.Count;
                var lastColumn = numColumns - 1;

                for (var r = 0; r < numRows; r++)
                {
                    var columnValues = GetDataGridRowValues(Page4DataGrid, r);

                    for (var c = 0; c < numColumns; c++)
                    {
                        if (columnValues[c] != null)
                        {
                            var value = columnValues[c];
                            value = value.Replace(tab, ' ');
                            
                            if (c == lastColumn)
                            {
                                clipboardText.Append(value + newline);
                            }
                            else
                            {
                                clipboardText.Append(value + tab);
                            }
                        }
                        else
                        {
                            clipboardText.Append(tab);
                        }
                    }
                }
            }

            Clipboard.SetText(clipboardText.ToString());
        }

        /// <summary>
        /// Populate the column list and advance to the next page
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page4Save_Click(object sender, RoutedEventArgs e)
        {
            var columnNames = GetDataGridColumnNames(Page4DataGrid);
            foreach (var name in columnNames)
            {
                var box = new CheckBox() { Content = name, IsChecked = true };
                Page5ListView.Items.Add(box);
            }

            GoToNextPage(null, null);
        }

        /// <summary>
        /// Check or uncheck all items in the column list
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
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

        /// <summary>
        /// Prompt the user for the name of the file to be saved and populate the name in a text box
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page6Browse_Click(object sender, RoutedEventArgs e)
        {
            var fileName = PromptSaveExcelFileName();

            if (!string.IsNullOrEmpty(fileName))
            {
                // The fileName is likely valid
                Page6NewExcelFileTextBox.Text = fileName;
                Page6Save.IsEnabled = true;
            }
            else
            {
                // The fileName is definitely not valid
                Page6NewExcelFileTextBox.Text = string.Empty;
                Page6Save.IsEnabled = false;
            }
        }

        /// <summary>
        /// Save data to a file
        /// </summary>
        /// <param name="sender"> The element which raised the event </param>
        /// <param name="e"> Event arguments </param>
        private void Page6Save_Click(object sender, RoutedEventArgs e)
        {
            if (Page6CreateNewExcelFileRadioButton.IsChecked == true)
            {
                var saveFileName = Page6NewExcelFileTextBox.Text;

                /* Abort if the file name is an empty string */
                if (string.IsNullOrEmpty(saveFileName))
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

                /* Add data rows to the worksheet */
                {
                    /* For each row in the source DataGrid */
                    var numRows = Page4DataGrid.Items.Count;
                    for (var row = 0; row < numRows; row++)
                    {
                        var numCheckedColumns = 0;
                        var filteredRow = filteredTable.NewRow();

                        /* For each column in the source DataGrid via the return value of GetDataGridRowValues */
                        var dataValues = GetDataGridRowValues(Page4DataGrid, row);
                        for (var i = 0; i < dataValues.Count; i++)
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
                }

                var saveSuccessful = _excel.SaveExcelFile(saveFileName, filteredTable);
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
            else if (Page6CreateNewCsvFileRadioButton.IsChecked == true)
            {
                var saveFileName = Page6NewExcelFileTextBox.Text;

                /* Abort if the file name is an empty string */
                if (string.IsNullOrEmpty(saveFileName))
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

                /* Add data rows to the worksheet */
                {
                    /* For each row in the source DataGrid */
                    var numRows = Page4DataGrid.Items.Count;
                    for (var row = 0; row < numRows; row++)
                    {
                        var numCheckedColumns = 0;
                        var filteredRow = filteredTable.NewRow();

                        /* For each column in the source DataGrid via the return value of GetDataGridRowValues */
                        var dataValues = GetDataGridRowValues(Page4DataGrid, row);
                        for (var i = 0; i < dataValues.Count; i++)
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
                }

                var saveSuccessful = _excel.SaveCsvFile(saveFileName, filteredTable);
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
}