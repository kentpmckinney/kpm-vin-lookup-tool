//-----------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="N/A">
//     Copyright (c) 2016 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

namespace VehicleInformationLookupTool
{
    using System;
    using System.Data;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;

    /// <summary>
    /// Interaction logic for MainWindow
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        /// <summary>
        /// Instance of ExcelClass which encapsulates Excel functionality
        /// </summary>
        private ExcelClass excel = new ExcelClass();

        /// <summary>
        /// Instance of WebClass which encapsulates Internet and web access
        /// </summary>
        private WebClass web = new WebClass();

        /// <summary>
        /// DataTable which stores downloaded vin data
        /// </summary>
        private DataTable vinData = new DataTable();

        /// <summary>
        /// CancellationSource which allows running tasks to be cancelled
        /// </summary>
        private CancellationTokenSource downloadCancellationSource = new CancellationTokenSource();

        /// <summary>
        /// Tasks which are provided this CancellationToken are cancellable by the CancellationSource
        /// </summary>
        private CancellationToken downloadCancellationToken;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class
        /// </summary>
        public MainWindow()
        {
            // Data initialization for MainWindow
            this.downloadCancellationToken = this.downloadCancellationSource.Token;

            // Increase the maximum number of HTTP connections (the default is 2)
            ServicePointManager.DefaultConnectionLimit = 4;
            
            this.InitializeComponent();
        }

        /// <summary>
        /// Properly dispose of data
        /// </summary>
        public void Dispose()
        {
            if (this.excel != null)
            {
                this.excel.Dispose();
                this.excel = null;
            }

            if (this.vinData != null)
            {
                this.vinData.Dispose();
                this.vinData = null;
            }

            if (this.downloadCancellationSource != null)
            {
                this.downloadCancellationSource.Dispose();
                this.downloadCancellationSource = null;
            }
        }

        /// <remarks> Please refer to MainWindow.Events.cs and MainWindow.Methods.cs for the code you might expect to see here </remarks>
    }
}
