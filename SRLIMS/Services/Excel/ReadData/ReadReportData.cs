using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SRLIMS.Services.Excel.ReadData
{
    internal class ReadReportData
    {

        ExcelReader DataReader = new ExcelReader();
        // Method to get the chain data
        public List<List<Object>> ReadData(string filepath)
        {
            List<List<object>> chainData = new List<List<object>>();
            string sheetname = "Chain of Custody 1";
            string fileRoute = filepath;
            int startRow = 15;
            List<int> columns = new List<int> { 2, 3, 4, 5, 6, 7, 8, 25};
            int maxRows = 20;

            

            try
            {
                chainData = DataReader.ReadRowsAsLists(sheetname, fileRoute, startRow, columns, maxRows);
            }
            catch (Exception ex) {

                Console.Write($"Excel reader failed: {ex.Message}");
                throw;
            
            }


            return chainData; 

        }


        public List<List<List<string>>> ReadMatrixData(string filepath)
        {


            string fileRoute = filepath;
            int startRow = 21;
            List<int> columns = new List<int> { 2, 3, 4, 5, 6, 7, 8, 25 };
            int maxRows = 25;
            List<List<List<string>>> matrixData = new List<List<List<string>>>();


            var matrixSheets = new List<string> {
                "Ammonia (7664417)",
                "Alkalinity (471341)",
                "Chlorides (16887006)"
            };

            foreach( var sheet in matrixSheets)
            {
                var sheetData = DataReader.ReadRowsAsLists(sheet, filepath, startRow, columns, maxRows);
                List<List<string>> convertedData = sheetData?
                    .Select(row => row.Select(cell => cell?.ToString() ?? "").ToList())
                    .ToList() ?? new List<List<string>>();


                matrixData.Add(convertedData);
             
            }


            return matrixData;


        }
        

       








    }
        



    }

