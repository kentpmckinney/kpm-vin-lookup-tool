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
    using System.Windows;

    /// <summary>
    /// Interaction logic for MainWindow
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        /// <summary>
        /// Instance of ExcelClass which encapsulates Excel functionality
        /// </summary>
        private ExcelClass _excel = new ExcelClass();

        /// <summary>
        /// Instance of WebClass which encapsulates Internet and _web access
        /// </summary>
        private readonly WebClass _web = new WebClass();

        /// <summary>
        /// DataTable which stores downloaded vin data
        /// </summary>
        private DataTable _vinData = new DataTable();

        /// <summary>
        /// CancellationSource which allows running tasks to be cancelled
        /// </summary>
        private CancellationTokenSource _downloadCancellationSource = new CancellationTokenSource();

        /// <summary>
        /// Tasks which are provided this CancellationToken are cancellable by the CancellationSource
        /// </summary>
        private CancellationToken _downloadCancellationToken;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class
        /// </summary>
        public MainWindow()
        {
            // Data initialization for MainWindow
            this._downloadCancellationToken = this._downloadCancellationSource.Token;

            // Increase the maximum number of HTTP connections (the default is 2)
            ServicePointManager.DefaultConnectionLimit = 4;
            ServicePointManager.Expect100Continue = true;

            // Enable all commonly used HTTPS protocols
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;

            this.InitializeComponent();
        }

        /// <summary>
        /// Properly dispose of data
        /// </summary>
        public void Dispose()
        {
            if (this._excel != null)
            {
                this._excel.Dispose();
                this._excel = null;
            }

            if (this._vinData != null)
            {
                this._vinData.Dispose();
                this._vinData = null;
            }

            if (this._downloadCancellationSource != null)
            {
                this._downloadCancellationSource.Dispose();
                this._downloadCancellationSource = null;
            }
        }

    }
}
