using SRLIMS.Data;
using SRLIMS.Services.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using System.Windows.Controls;

namespace SRLIMS
{
    public partial class MainWindow : Window
    {

        public Frame mainFrameControl => this.MainFrame;
        private ExcelReader _excelReader;

    
        public MainWindow()
        {
          

            InitializeComponent();
            MainFrame.Content = new Views.Home();

        }

    }
}
