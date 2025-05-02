using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace SRLIMS.Services.Excel
{
    public class ExcelReader
    {
        // Habilitar licencia no comercial de EPPlus
        static ExcelReader()
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Lee datos de un archivo Excel y devuelve una lista de listas (cada lista interna es una fila)
        /// </summary>
        /// <param name="routeFile">Ruta del archivo Excel</param>
        /// <param name="startRow">Fila inicial (1-based)</param>
        /// <param name="columns">Lista de índices de columnas a leer (1-based)</param>
        /// <param name="maxRows">Número máximo de filas a leer (opcional)</param>
        /// <param name="sheetIndex">Índice de la hoja (0-based, por defecto la primera hoja)</param>
        /// <returns>Lista de listas, donde cada lista interna representa una fila con los valores de las columnas seleccionadas</returns>
        // Versión corregida del método ReadRowsAsLists
        public List<List<object>> ReadRowsAsLists(string sheetName, string filePath, int startRow, List<int> columns, int? maxRows = null)
        {
            // Validación EXTRA reforzada
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("La ruta del archivo no puede estar vacía");

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Archivo no encontrado", filePath);

            // Asegurar que las columnas sean válidas
            if (columns == null || columns.Count == 0 || columns.Any(c => c < 1))
                throw new ArgumentException("Columnas inválidas");

            var result = new List<List<object>>();

            try
            {

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook?.Worksheets == null || !package.Workbook.Worksheets.Any())
                        return result;

                    ExcelWorksheet? worksheet = !string.IsNullOrWhiteSpace(sheetName)
                        ? package.Workbook.Worksheets.FirstOrDefault(ws =>
                            ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        : package.Workbook.Worksheets[sheetName]; // Acceso a la hoja de cadena de custodia

                    if (worksheet == null)
                        throw new Exception($"Hoja no encontrada: {sheetName ?? "primera hoja"}");

                    if (worksheet.Dimension == null)
                        return result;

                    startRow = Math.Max(1, startRow);
                    int endRow = maxRows.HasValue
                        ? Math.Min(worksheet.Dimension.End.Row, startRow + maxRows.Value - 1)
                        : worksheet.Dimension.End.Row;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        var rowData = new List<object>();
                        bool hasData = false;

                        foreach (int col in columns)
                        {
                            try
                            {
                                var cell = worksheet.Cells[row, col];
                                var value = cell?.Value ?? null;
                                rowData.Add(value);
                                if (value != null) hasData = true;
                            }
                            catch
                            {
                                rowData.Add(null);
                            }
                        }

                        if (hasData || rowData.Any(x => x != null))
                            result.Add(rowData);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Fallo crítico al leer Excel. Detalles técnicos:\n" +
                                  $"• Archivo: {Path.GetFileName(filePath)}\n" +
                                  $"• Error: {ex.GetType().Name}\n" +
                                  $"• Mensaje: {ex.Message}", ex);
            }

            return result;
        }

        /// <summary>
        /// Copia un rango de celdas de una ubicación a otra
        /// </summary>
        public void CopyRange(
            string routeFile,
            int sourceStartRow,
            int sourceStartColumn,
            int sourceEndRow,
            int sourceEndColumn,
            int targetStartRow,
            int targetStartColumn,
            bool saveChanges = true)
        {
            using (var package = new ExcelPackage(new FileInfo(routeFile)))
            {
                var worksheet = package.Workbook.Worksheets.First();

                // Obtener el rango fuente
                var sourceRange = worksheet.Cells[sourceStartRow, sourceStartColumn, sourceEndRow, sourceEndColumn];

                // Calcular la posición de destino
                var targetRange = worksheet.Cells[targetStartRow, targetStartColumn,
                    targetStartRow + (sourceEndRow - sourceStartRow),
                    targetStartColumn + (sourceEndColumn - sourceStartColumn)];

                // Copiar el rango
                sourceRange.Copy(targetRange);

                // Guardar cambios si se solicita
                if (saveChanges)
                {
                    package.Save();
                }
            }
        }

        /// <summary>
        /// Pega solo valores en un rango específico
        /// </summary>
        public void PasteValues(
            string routeFile,
            int startRow,
            int startColumn,
            List<List<object>> values,
            bool saveChanges = true)
        {
            using (var package = new ExcelPackage(new FileInfo(routeFile)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Pegar cada valor en la celda correspondiente
                for (int rowOffset = 0; rowOffset < values.Count; rowOffset++)
                {
                    var rowValues = values[rowOffset];
                    for (int colOffset = 0; colOffset < rowValues.Count; colOffset++)
                    {
                        worksheet.Cells[startRow + rowOffset, startColumn + colOffset].Value = rowValues[colOffset];
                    }
                }

                if (saveChanges)
                {
                    package.Save();
                }
            }
        }

        /// <summary>
        /// Lee celdas combinadas de un<a hoja y devuelve sus rangos
        /// </summary>
        public List<ExcelAddressBase> GetMergedCells(string routeFile)
        {
            var mergedCells = new List<ExcelAddressBase>();

            using (var package = new ExcelPackage(new FileInfo(routeFile)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                foreach (var mergedCell in worksheet.MergedCells)
                {
                    var mergedRange = new ExcelAddress(mergedCell);
                    mergedCells.Add(mergedRange);
                }
            }

            return mergedCells;
        }


    }
}