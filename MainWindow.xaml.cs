//-----------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
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
            _downloadCancellationToken = _downloadCancellationSource.Token;

            ServicePointManager.DefaultConnectionLimit = 4;
            ServicePointManager.Expect100Continue = true;

            // Enable all commonly used HTTPS protocols
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;

            InitializeComponent();
        }

        /// <summary>
        /// Properly dispose of data
        /// </summary>
        public void Dispose()
        {
            if (_excel != null)
            {
                _excel.Dispose();
                _excel = null;
            }

            if (_vinData != null)
            {
                _vinData.Dispose();
                _vinData = null;
            }

            if (_downloadCancellationSource != null)
            {
                _downloadCancellationSource.Dispose();
                _downloadCancellationSource = null;
            }
        }

    }
}
