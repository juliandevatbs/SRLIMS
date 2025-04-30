using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using SRLIMS.Services.Excel;

namespace SRLIMS.Views
{
    /// <summary>
    /// Lógica de interacción para SelectDataSource.xaml
    /// </summary>
    public partial class SelectDataSource : UserControl
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

        private void QueryExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Diagnóstico 1 - Verificar archivo seleccionado
                if (string.IsNullOrWhiteSpace(txtExcelPath.Text))
                {
                    MessageBox.Show("No se ha seleccionado ningún archivo", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string filePath = txtExcelPath.Text;

                // Diagnóstico 2 - Verificar existencia del archivo
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"El archivo no existe:\n{filePath}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Diagnóstico 3 - Verificar extensión
                var validExtensions = new[] { ".xls", ".xlsx", ".xlsm" };
                if (!validExtensions.Contains(System.IO.Path.GetExtension(filePath).ToLower()))
                {
                    MessageBox.Show("Extensión de archivo no válida", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Configuración conservadora
                var columnsToRead = new List<int> { 1, 2, 3 }; // Solo 3 columnas para prueba
                List<List<object>> excelData = null;

                // Bloque de diagnóstico
                try
                {
                    Debug.WriteLine("=== INICIO DIAGNÓSTICO ===");
                    Debug.WriteLine($"Leyendo archivo: {filePath}");
                    Debug.WriteLine($"Tamaño: {new FileInfo(filePath).Length} bytes");

                    excelData = _excelReader.ReadRowsAsLists(
                        filePath: filePath,
                        startRow: 1,
                        columns: columnsToRead,
                        maxRows: 10); // Solo 10 filas para prueba

                    Debug.WriteLine($"Filas leídas: {excelData?.Count ?? 0}");
                    Debug.WriteLine("=== FIN DIAGNÓSTICO ===");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ERROR CRÍTICO: {ex.ToString()}");
                    MessageBox.Show($"Error al leer el archivo:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Validación de datos
                if (excelData == null || excelData.Count == 0)
                {
                    MessageBox.Show("El archivo no contiene datos o está vacío", "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Mostrar datos
                try
                {
                    var dataView = new ExcelDataView(excelData);
                    dataView.Owner = Window.GetWindow(this);
                    dataView.Show();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al mostrar los datos:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception globalEx)
            {
                MessageBox.Show($"Error inesperado:\n{globalEx.Message}", "Error crítico", MessageBoxButton.OK, MessageBoxImage.Error);
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
    }
}
