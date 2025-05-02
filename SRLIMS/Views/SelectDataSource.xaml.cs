using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.EntityFrameworkCore.Query.Internal;
using Microsoft.Win32;
using SRLIMS.Services.Excel;

namespace SRLIMS.Views
{
    /// <summary>
    /// Lógica de interacción para SelectDataSource.xaml
    /// </summary>
    public partial class SelectDataSource : UserControl, IDisposable
    {
        private readonly ExcelReader _excelReader;

        public SelectDataSource()
        {
            InitializeComponent();
            _excelReader = new ExcelReader();
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
            var custodyData = _excelReader.ReadRowsAsLists(
                filePath: txtExcelPath.Text,
                startRow: 15,
                columns: new List<int> { 2, 3, 4, 5, 6, 7, 25 },
                maxRows: 20,
                sheetName: "Chain of Custody 1"
            );

            // Datos de las matrices
            List<List<List<string>>> matrixData = new List<List<List<string>>>();
            var matrixSheets = new List<string> {
        "Ammonia (7664417)",
        "Alkalinity (471341)",
        "Chlorides (16887006)"
    };

            foreach (var sheet in matrixSheets)
            {
                var excelData = _excelReader.ReadRowsAsLists(
                    filePath: txtExcelPath.Text,
                    startRow: 21,
                    columns: new List<int> { 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 },
                    maxRows: 25,
                    sheetName: sheet
                );

                List<List<string>> convertedSheet = excelData?
                    .Select(row => row.Select(cell => cell?.ToString() ?? "").ToList())
                    .ToList() ?? new List<List<string>>();

                matrixData.Add(convertedSheet);
            }

            return (custodyData, matrixData);
        }



        private void QueryExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Validaciones iniciales
                if (string.IsNullOrWhiteSpace(txtExcelPath.Text) || !File.Exists(txtExcelPath.Text))
                {
                    MessageBox.Show("Archivo no válido o no seleccionado", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Leer todos los datos
                var (custodyData, matrixData) = ReadAllExcelData();

                // Validar que hay datos
                if ((custodyData == null || custodyData.Count == 0) &&
                    (matrixData == null || matrixData.All(m => m.Count == 0)))
                {
                    MessageBox.Show("No se encontraron datos en el archivo", "Información",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Crear la vista con ambos datasets
                var excelDataView = new ExcelDataView(custodyData, matrixData);

                // Asignar al MainFrame
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





        private void QueryDatabase_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtDatabaseID.Text))
            {
                MessageBox.Show("Please enter a Lab Reporting Batch ID");
                return;
            }

            MessageBox.Show($"Samples with: {txtDatabaseID.Text}");
        }



        public void Dispose()
        {
            
        }
    }
}