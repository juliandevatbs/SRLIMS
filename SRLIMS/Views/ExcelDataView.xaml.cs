using System;
using System.Collections.Generic;
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

namespace SRLIMS.Views
{
    /// <summary>
    /// Lógica de interacción para ExcelDataView.xaml
    /// </summary>
    public partial class ExcelDataView : Window
    {
        public ExcelDataView(List<List<object>> excelData)
        {
            InitializeComponent();
            LoadData(excelData);
        }

        private void LoadData(List<List<object>> excelData)
        {

            try
            {
                if (excelData == null)
                {
                    MessageBox.Show("No se recibieron datos");
                    return;
                }


                var dataTable = new System.Data.DataTable();


                if (excelData.Count > 0)
                {
                    for (int i = 0; i < excelData[0].Count; i++)
                    {
                        dataTable.Columns.Add($"Columna {i + 1}");
                    }

                    foreach (var row in excelData)
                    {
                        dataTable.Rows.Add(row.ToArray());
                    }
                }

                DataGrid.ItemsSource = dataTable.DefaultView;
            } catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar datos en la tabla{ex.Message}");
            }
    }
    }
}


