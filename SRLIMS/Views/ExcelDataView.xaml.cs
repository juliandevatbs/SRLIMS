using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Data;
using System.Linq;

namespace SRLIMS.Views
{
    public partial class ExcelDataView : UserControl, IDisposable
    {
        private List<List<List<string>>> _originalMatrixData;
        private DataTable _fullMatrixTable;
        private DataTable _custodyDataTable;

        public ExcelDataView(List<List<object>> excelData, List<List<List<string>>> matrixData)
        {
            InitializeComponent();
            _originalMatrixData = matrixData;
            LoadData(excelData, matrixData);
        }

        private void LoadData(List<List<object>> excelData, List<List<List<string>>> matrixData)
        {
            try
            {
                if (excelData == null && matrixData == null)
                {
                    MessageBox.Show("No data received");
                    return;
                }

                if (excelData != null && excelData.Count > 0)
                {
                    _custodyDataTable = CreateCustodyTable(excelData);
                    SetupCustodyDataGrid(_custodyDataTable);
                }

                // Crear la tabla de matriz completa pero no mostrar datos inicialmente
                if (matrixData != null && matrixData.Count > 0)
                {
                    _fullMatrixTable = CreateMatrixTable(matrixData);

                    // Configurar la tabla de matriz pero con una tabla vacía que tenga la misma estructura
                    var emptyMatrixTable = _fullMatrixTable.Clone();
                    SetupMatrixDataGrid(emptyMatrixTable);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data into table: {ex.Message}");
            }
        }

        private DataTable CreateCustodyTable(List<List<object>> excelData)
        {
            var dataTable = new DataTable();

            // Headers para Chain of Custody
            string[] customHeaders = {
                "ItemID", "Sample Identification", "Sampled", "Grab or Composite",
                "Matrix", "Containers", "LabReportingBatchID"
            };

            // Crear columnas
            for (int i = 0; i < excelData[0].Count; i++)
            {
                string columnName = i < customHeaders.Length ? customHeaders[i] : $"Column {i + 1}";
                dataTable.Columns.Add(columnName);
            }

            // Agregar columna de CheckBox
            dataTable.Columns.Add("Include", typeof(bool));

            // Llenar datos
            foreach (var row in excelData)
            {
                var newRow = dataTable.NewRow();
                for (int i = 0; i < row.Count; i++)
                {
                    newRow[i] = row[i]?.ToString() ?? string.Empty;
                }
                newRow["Include"] = false;
                dataTable.Rows.Add(newRow);
            }

            return dataTable;
        }

        private DataTable CreateMatrixTable(List<List<List<string>>> matrixData)
        {
            var matrixTable = new DataTable();

            // Headers para Matrix Data
            string[] headersMatrix = {
                "Date", "Sample Id", "Sample Volume", "Ph adjustment",
                "Volume H2S04 for blank", "Volume H2S04 for sample",
                "Normality", "Result", "Notes", "Notes 2"
            };

            if (matrixData.Count > 0 && matrixData[0].Count > 0)
            {
                for (int i = 0; i < matrixData[0][0].Count; i++)
                {
                    string columnName = i < headersMatrix.Length ? headersMatrix[i] : $"Column {i + 1}";

                    // Para la columna Date, usar un tipo DateTime
                    if (i == 0 && columnName == "Date")
                    {
                        matrixTable.Columns.Add(columnName, typeof(DateTime));
                    }
                    else
                    {
                        matrixTable.Columns.Add(columnName);
                    }
                }
            }

            foreach (var sheet in matrixData)
            {
                foreach (var row in sheet)
                {
                    var newRow = matrixTable.NewRow();
                    for (int i = 0; i < row.Count; i++)
                    {
                        // Intentar convertir la fecha si es la columna Date
                        if (i == 0 && matrixTable.Columns[i].DataType == typeof(DateTime))
                        {
                            // Intentar varios formatos de fecha
                            if (DateTime.TryParse(row[i], out DateTime dateValue))
                            {
                                newRow[i] = dateValue;
                            }
                            else if (double.TryParse(row[i], out double numericDate))
                            {
                                // Si es un número, podría ser una fecha de Excel (días desde 1900-01-01)
                                try
                                {
                                    newRow[i] = DateTime.FromOADate(numericDate);
                                }
                                catch
                                {
                                    // Si falla la conversión, mantener el valor original como string
                                    newRow[i] = row[i];
                                }
                            }
                            else
                            {
                                newRow[i] = row[i];
                            }
                        }
                        else
                        {
                            newRow[i] = row[i];
                        }
                    }
                    matrixTable.Rows.Add(newRow);
                }
            }

            return matrixTable;
        }

        private void SetupCustodyDataGrid(DataTable dataTable)
        {
            CustodyDataGrid.AutoGenerateColumns = false;
            CustodyDataGrid.Columns.Clear();

            // Definir anchos de columna predeterminados para la tabla de custodia
            int[] columnWidths = { 80, 150, 100, 120, 100, 100, 150 };

            for (int i = 0; i < dataTable.Columns.Count - 1; i++)
            {
                var column = new DataGridTextColumn
                {
                    Header = dataTable.Columns[i].ColumnName,
                    Binding = new Binding($"[{i}]"),
                    Width = new DataGridLength(i < columnWidths.Length ? columnWidths[i] : 100)
                };
                CustodyDataGrid.Columns.Add(column);
            }

            var checkBoxColumn = new DataGridCheckBoxColumn
            {
                Header = "Include",
                Binding = new Binding($"[{dataTable.Columns.Count - 1}]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                Width = new DataGridLength(80)
            };
            CustodyDataGrid.Columns.Add(checkBoxColumn);

            CustodyDataGrid.ItemsSource = dataTable.DefaultView;

            // Asegurarse de que el evento SelectionChanged esté conectado
            CustodyDataGrid.SelectionChanged += CustodyDataGrid_SelectionChanged;

            // Aplicar estilo de separación
            CustodyDataGrid.ColumnHeaderHeight = 30;
            CustodyDataGrid.RowHeight = 25;
            CustodyDataGrid.HorizontalGridLinesBrush = System.Windows.Media.Brushes.LightGray;
            CustodyDataGrid.VerticalGridLinesBrush = System.Windows.Media.Brushes.LightGray;
            CustodyDataGrid.GridLinesVisibility = DataGridGridLinesVisibility.All;
        }

        private void SetupMatrixDataGrid(DataTable matrixTable)
        {
            MatrixDataGrid.AutoGenerateColumns = false;
            MatrixDataGrid.Columns.Clear();

            // Definir anchos de columna predeterminados para la tabla de matriz
            int[] columnWidths = { 120, 150, 120, 120, 150, 150, 100, 100, 150, 150 };

            for (int i = 0; i < matrixTable.Columns.Count; i++)
            {
                // Para la columna de fecha, usar un formato de fecha específico
                if (i == 0 && matrixTable.Columns[i].ColumnName == "Date")
                {
                    var dateColumn = new DataGridTextColumn
                    {
                        Header = matrixTable.Columns[i].ColumnName,
                        Width = new DataGridLength(columnWidths[i]),
                        Binding = new Binding($"[{i}]")
                        {
                            StringFormat = "yyyy-MM-dd HH:mm:ss" // Formato de fecha
                        }
                    };
                    MatrixDataGrid.Columns.Add(dateColumn);
                }
                else
                {
                    var column = new DataGridTextColumn
                    {
                        Header = matrixTable.Columns[i].ColumnName,
                        Width = new DataGridLength(i < columnWidths.Length ? columnWidths[i] : 120),
                        Binding = new Binding($"[{i}]")
                    };
                    MatrixDataGrid.Columns.Add(column);
                }
            }

            MatrixDataGrid.ItemsSource = matrixTable.DefaultView;

            // Aplicar estilo de separación
            MatrixDataGrid.ColumnHeaderHeight = 30;
            MatrixDataGrid.RowHeight = 25;
            MatrixDataGrid.HorizontalGridLinesBrush = System.Windows.Media.Brushes.LightGray;
            MatrixDataGrid.VerticalGridLinesBrush = System.Windows.Media.Brushes.LightGray;
            MatrixDataGrid.GridLinesVisibility = DataGridGridLinesVisibility.All;
        }

        public void Dispose()
        {
            CustodyDataGrid.ItemsSource = null;
            MatrixDataGrid.ItemsSource = null;
        }

        public void GetSelectedChainData_Click(object sender, RoutedEventArgs e)
        {
            var selectedData = GetSelectedChainData();
            updateMatrixTable();
        }

        // Este método actualiza la tabla de matriz - ahora solo se usa para el botón
        public void updateMatrixTable()
        {
            // Verificar si hay una fila seleccionada en la tabla de custodia
            if (CustodyDataGrid.SelectedItem is DataRowView selectedRow)
            {
                // Obtener el Sample Identification de la fila seleccionada (índice 1)
                string selectedSampleId = selectedRow.Row.ItemArray[1].ToString();

                // Mostrar los datos relacionados
                ShowRelatedMatrixData(selectedSampleId);
            }
            else
            {
                // Si no hay selección, mostrar una tabla vacía
                var emptyTable = _fullMatrixTable.Clone();
                MatrixDataGrid.ItemsSource = emptyTable.DefaultView;
            }
        }

        public List<List<object>> GetSelectedChainData()
        {
            var selectedData = new List<List<object>>();

            if (CustodyDataGrid.ItemsSource is DataView dataView)
            {
                foreach (DataRowView rowView in dataView)
                {
                    // Verificar si el checkbox está marcado
                    if (rowView.Row.Field<bool>("Include"))
                    {
                        // Obtener todos los valores de la fila excepto la columna "Include"
                        var rowData = new List<object>();
                        for (int i = 0; i < rowView.Row.ItemArray.Length - 1; i++) // -1 para excluir el checkbox
                        {
                            rowData.Add(rowView.Row.ItemArray[i]);
                        }
                        selectedData.Add(rowData);
                    }
                }
            }

            MessageBox.Show($"CANTIDAD DE MUESTRAS SELECCIONADAS: {selectedData.Count}");
            return selectedData;
        }

        private void CustodyDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Actualizar la tabla de matriz basada en la fila seleccionada
            if (CustodyDataGrid.SelectedItem is DataRowView selectedRow)
            {
                // Obtener el Sample Identification de la fila seleccionada (índice 1)
                string selectedSampleId = selectedRow.Row.ItemArray[1].ToString();

                // Actualizar la tabla de matriz para mostrar solo datos relacionados con esta muestra
                ShowRelatedMatrixData(selectedSampleId);
            }
        }

        private void MatrixDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Este evento puede utilizarse para otras funcionalidades si es necesario
        }

        // Método para mostrar los datos de matriz relacionados con un ID de muestra específico
        private void ShowRelatedMatrixData(string sampleId)
        {
            if (_fullMatrixTable == null) return;

            // Crear una tabla filtrada
            var filteredMatrixTable = _fullMatrixTable.Clone();

            // Encontrar todas las filas en la tabla matriz completa que coincidan con el sampleId
            var matchingRows = _fullMatrixTable.AsEnumerable()
                .Where(row => row["Sample Id"].ToString().Equals(sampleId))
                .ToList();

            if (matchingRows.Count > 0)
            {
                // Agregar las filas coincidentes a la tabla filtrada
                foreach (var row in matchingRows)
                {
                    var newRow = filteredMatrixTable.NewRow();
                    newRow.ItemArray = row.ItemArray;
                    filteredMatrixTable.Rows.Add(newRow);
                }
            }

            // Actualizar la tabla de matriz
            MatrixDataGrid.ItemsSource = filteredMatrixTable.DefaultView;
        }
    }
}