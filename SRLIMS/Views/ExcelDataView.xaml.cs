using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Data;

namespace SRLIMS.Views
{
    public partial class ExcelDataView : UserControl, IDisposable
    {
        public ExcelDataView(List<List<object>> excelData, List<List<List<string>>> matrixData)
        {
            InitializeComponent();
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

                // Tabla para Chain of Custody (Excel Data)
                if (excelData != null && excelData.Count > 0)
                {
                    var dataTable = CreateCustodyTable(excelData);
                    SetupCustodyDataGrid(dataTable);
                }

                // Tabla para Matrix Data
                if (matrixData != null && matrixData.Count > 0)
                {
                    var matrixTable = CreateMatrixTable(matrixData);
                    SetupMatrixDataGrid(matrixTable);
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

            // Crear columnas (usamos la primera hoja como referencia)
            if (matrixData.Count > 0 && matrixData[0].Count > 0)
            {
                for (int i = 0; i < matrixData[0][0].Count; i++)
                {
                    string columnName = i < headersMatrix.Length ? headersMatrix[i] : $"Column {i + 1}";
                    matrixTable.Columns.Add(columnName);
                }
            }

            // Llenar datos de todas las hojas
            foreach (var sheet in matrixData)
            {
                foreach (var row in sheet)
                {
                    var newRow = matrixTable.NewRow();
                    for (int i = 0; i < row.Count; i++)
                    {
                        newRow[i] = row[i];
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

            for (int i = 0; i < dataTable.Columns.Count - 1; i++)
            {
                var column = new DataGridTextColumn
                {
                    Header = dataTable.Columns[i].ColumnName,
                    Binding = new Binding($"[{i}]")
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
                }
            };
            CustodyDataGrid.Columns.Add(checkBoxColumn);

            CustodyDataGrid.ItemsSource = dataTable.DefaultView;
        }

        private void SetupMatrixDataGrid(DataTable matrixTable)
        {
            MatrixDataGrid.AutoGenerateColumns = false;
            MatrixDataGrid.Columns.Clear();

            for (int i = 0; i < matrixTable.Columns.Count; i++)
            {
                var column = new DataGridTextColumn
                {
                    Header = matrixTable.Columns[i].ColumnName,
                    Binding = new Binding($"[{i}]")
                };
                MatrixDataGrid.Columns.Add(column);
            }

            MatrixDataGrid.ItemsSource = matrixTable.DefaultView;
        }

        public void Dispose()
        {
            CustodyDataGrid.ItemsSource = null;
            MatrixDataGrid.ItemsSource = null;
        }
    }
}