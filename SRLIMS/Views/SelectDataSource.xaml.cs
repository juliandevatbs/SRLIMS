using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using SRLIMS.Data;
using SRLIMS.Services.Excel;
using SRLIMS.Services.Excel.ReadData;

namespace SRLIMS.Views
{
    public partial class SelectDataSource : UserControl, IDisposable
    {
        private readonly ExcelReader _excelReader;
        private readonly DbConnection _dbConnection;
        private readonly ReadReportData _chainDataGetter;

        public SelectDataSource()
        {
            InitializeComponent();
            _excelReader = new ExcelReader();
            _chainDataGetter = new ReadReportData();
            _dbConnection = new DbConnection();
        }

        private void DataSourceOption_Checked(object sender, RoutedEventArgs e)
        {
            pnlExcel.Visibility = rbExcel.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
            pnlDatabase.Visibility = rbDatabase.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        private void BrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx;*xls)|*.xlsx;*.xls| All files (*.*)|*.*",
                Title = "Select excel file"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                txtExcelPath.Text = openFileDialog.FileName;
            }
        }

        private (List<List<object>> custodyData, List<List<List<string>>> matrixData) ReadAllExcelData()
        {
            // Datos de Chain of Custody
            var custodyData = _chainDataGetter.ReadData(txtExcelPath.Text);

            var matrixData = _chainDataGetter.ReadMatrixData(txtExcelPath.Text);

            return (custodyData, matrixData);
        }

        private void QueryExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
                {
                    MessageBox.Show("Archivo no válido o no seleccionado", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var (custodyData, matrixData) = ReadAllExcelData();

                if ((custodyData == null || custodyData.Count == 0) &&
                    (matrixData == null || matrixData.All(m => m.Count == 0)))
                {
                    MessageBox.Show("No se encontraron datos en el archivo", "Información",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var excelDataView = new ExcelDataView(custodyData, matrixData);

                if (Application.Current.MainWindow is MainWindow mainWindow && mainWindow.MainFrame != null)
                {
                    if (mainWindow.MainFrame.Content is IDisposable oldView)
                        oldView.Dispose();

                    mainWindow.MainFrame.Content = excelDataView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void QueryDatabase_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtDatabaseID.Text))
                {
                    MessageBox.Show("Por favor ingrese un Lab Reporting Batch ID", "Validación",
                                  MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                var (custodyData, matrixData) = await QueryDatabaseDataAsync(txtDatabaseID.Text);

                if ((custodyData == null || custodyData.Count == 0) &&
                    (matrixData == null || matrixData.Count == 0))
                {
                    MessageBox.Show("No se encontraron datos para el ID especificado", "Información",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    Mouse.OverrideCursor = null;
                    return;
                }

                var databaseDataView = new DatabaseDataView(custodyData, matrixData);

                if (Application.Current.MainWindow is MainWindow mainWindow && mainWindow.MainFrame != null)
                {
                    if (mainWindow.MainFrame.Content is IDisposable oldView)
                        oldView.Dispose();

                    mainWindow.MainFrame.Content = databaseDataView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al consultar la base de datos: {ex.Message}", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private async Task<(List<List<object>> custodyData, List<List<List<string>>> matrixData)> QueryDatabaseDataAsync(string labReportingBatchID)
        {
            await _dbConnection.OpenAsync();

            // 1. Obtener los datos de Chain of Custody (Samples table)
            string custodyQuery = @"
                SELECT 
                    ItemID,
                    LabReportingBatchID,
                    LabSampleID,
                    ClientSampleID,
                    CollectMethod,
                    MatrixID,
                    DateCollected,
                    SamplingPersonnel,
                    CollectionAgency,
                    CustodyIntactSeal,
                    ReceiptComments,
                    LocationCode,
                    LabReceiptDate,
                    ProgramType,
                    CollectionMethod,
                    SamplingDepth,
                    ProjectNumber
                FROM Samples
                WHERE LabReportingBatchID = @batchId";

            custodyQuery = custodyQuery.Replace("@batchId", $"'{labReportingBatchID}'");
            DataTable custodyTable = await _dbConnection.ExecuteQueryAsync(custodyQuery);

            List<List<object>> custodyData = new List<List<object>>();
            foreach (DataRow row in custodyTable.Rows)
            {
                List<object> rowData = new List<object>();
                foreach (var item in row.ItemArray)
                {
                    rowData.Add(item);
                }
                custodyData.Add(rowData);
            }

            // 2. Obtener los datos de las matrices (Sample_Tests table)
            string matrixQuery = @"
                SELECT 
                    st.SampleTestsID,
                    st.ItemID,
                    st.ClientSampleID,
                    st.LabSampleID,
                    st.AnalyteName,
                    st.Result,
                    st.ResultUnits,
                    st.LabQualifiers,
                    st.ReportingLimit,
                    st.DateCollected,
                    st.MatrixID,
                    st.QCType,
                    st.ResultComments,
                    st.ReportableResult,
                    st.DateAnalyzed,
                    st.Analyst,
                    st.Notes
                FROM Sample_Tests st
                WHERE st.LabReportingBatchID = @batchId";

            matrixQuery = matrixQuery.Replace("@batchId", $"'{labReportingBatchID}'");
            DataTable allMatrixTable = await _dbConnection.ExecuteQueryAsync(matrixQuery);

            // Obtener los analitos únicos (equivalentes a las hojas de Excel)
            List<string> analytes = allMatrixTable.AsEnumerable()
                .Select(row => row["AnalyteName"].ToString())
                .Distinct()
                .ToList();

            List<List<List<string>>> matrixData = new List<List<List<string>>>();

            foreach (string analyte in analytes)
            {
                var analyteRows = allMatrixTable.AsEnumerable()
                    .Where(row => row["AnalyteName"].ToString() == analyte)
                    .ToList();

                List<List<string>> analyteData = new List<List<string>>();

                foreach (DataRow row in analyteRows)
                {
                    List<string> rowData = new List<string>
                    {
                        row["ClientSampleID"]?.ToString() ?? "",
                        row["DateCollected"]?.ToString() ?? "",
                        row["AnalyteName"]?.ToString() ?? "",
                        row["Result"]?.ToString() ?? "",
                        row["ResultUnits"]?.ToString() ?? "",
                        row["ReportingLimit"]?.ToString() ?? "",
                        row["LabQualifiers"]?.ToString() ?? "",
                        row["QCType"]?.ToString() ?? "",
                        row["ResultComments"]?.ToString() ?? "",
                        row["Notes"]?.ToString() ?? ""
                    };
                    analyteData.Add(rowData);
                }

                matrixData.Add(analyteData);
            }

            return (custodyData, matrixData);
        }

        public void Dispose()
        {
            _dbConnection?.Dispose();
        }
    }
}