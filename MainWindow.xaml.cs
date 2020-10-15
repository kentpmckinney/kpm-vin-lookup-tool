//-----------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="N/A">
//     Copyright (c) 2016, 2020 Kent P. McKinney
//     Released under the terms of the MIT License
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Data;
using System.Net;
using System.Threading;
using System.Windows;

namespace VehicleInformationLookupTool
{
    public partial class MainWindow : Window, IDisposable
    {
        private ExcelClass _excel = new ExcelClass();
        private readonly WebClass _web = new WebClass();
        private DataTable _vinData = new DataTable();
        private CancellationTokenSource _downloadCancellationSource = new CancellationTokenSource();
        private CancellationToken _downloadCancellationToken;
        private Direction _navigateDirection;
        private enum Direction
        {
            Forward = 0,
            Backward = 1
        }


        public MainWindow()
        {
            /* Data initialization for MainWindow */
            _downloadCancellationToken = _downloadCancellationSource.Token;

            ServicePointManager.DefaultConnectionLimit = 4;
            ServicePointManager.Expect100Continue = true;

            /* Enable all commonly used HTTPS protocols */
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;

            InitializeComponent();
        }


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
