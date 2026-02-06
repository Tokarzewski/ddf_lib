using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDDFAddin
{
    /// <summary>
    /// Helper methods for Excel interop operations.
    /// Provides efficient reading/writing between DataTables and Excel worksheets.
    /// </summary>
    public static class ExcelHelpers
    {
        /// <summary>
        /// Write a DataTable to an Excel worksheet starting at the specified row.
        /// Uses bulk array writing for better performance.
        /// </summary>
        public static void WriteDataTableToWorksheet(DataTable dataTable, Excel.Worksheet worksheet, int startRow = 1, int startCol = 1)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
                return;

            int rowCount = dataTable.Rows.Count;
            int colCount = dataTable.Columns.Count;

            // Create 2D array for bulk write
            object[,] data = new object[rowCount, colCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    data[i, j] = dataTable.Rows[i][j];
                }
            }

            // Get range and write all data at once (much faster than cell-by-cell)
            Excel.Range startCell = worksheet.Cells[startRow, startCol];
            Excel.Range endCell = worksheet.Cells[startRow + rowCount - 1, startCol + colCount - 1];
            Excel.Range range = worksheet.Range[startCell, endCell];

            range.Value2 = data;
        }

        /// <summary>
        /// Write column headers to an Excel worksheet.
        /// </summary>
        public static void WriteColumnHeaders(DataTable dataTable, Excel.Worksheet worksheet, int row = 1, int startCol = 1)
        {
            if (dataTable == null || dataTable.Columns.Count == 0)
                return;

            int colCount = dataTable.Columns.Count;
            object[] headers = new object[colCount];

            for (int i = 0; i < colCount; i++)
            {
                headers[i] = dataTable.Columns[i].ColumnName;
            }

            // Write headers
            Excel.Range startCell = worksheet.Cells[row, startCol];
            Excel.Range endCell = worksheet.Cells[row, startCol + colCount - 1];
            Excel.Range range = worksheet.Range[startCell, endCell];

            range.Value2 = new object[,] { headers };

            // Format headers (bold)
            range.Font.Bold = true;
        }

        /// <summary>
        /// Read an Excel worksheet into a DataTable.
        /// Assumes first row contains column headers.
        /// </summary>
        public static DataTable ReadWorksheetToDataTable(Excel.Worksheet worksheet, int headerRow = 2)
        {
            var dataTable = new DataTable(worksheet.Name);

            // Get used range
            Excel.Range usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;

            if (rowCount < headerRow)
                return dataTable;

            // Read column headers from specified row
            for (int col = 1; col <= colCount; col++)
            {
                var headerCell = (Excel.Range)usedRange.Cells[headerRow, col];
                string columnName = headerCell.Value2?.ToString() ?? $"Column{col}";
                dataTable.Columns.Add(columnName);
            }

            // Read data rows (starting after header row)
            for (int row = headerRow + 1; row <= rowCount; row++)
            {
                var dataRow = dataTable.NewRow();

                for (int col = 1; col <= colCount; col++)
                {
                    var cell = (Excel.Range)usedRange.Cells[row, col];
                    dataRow[col - 1] = cell.Value2 ?? string.Empty;
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        /// <summary>
        /// Read IDs from the first row of a worksheet.
        /// </summary>
        public static int[] ReadIdsFromWorksheet(Excel.Worksheet worksheet, int row = 1)
        {
            Excel.Range usedRange = worksheet.UsedRange;
            int colCount = usedRange.Columns.Count;

            var ids = new int[colCount];

            for (int col = 1; col <= colCount; col++)
            {
                var cell = (Excel.Range)worksheet.Cells[row, col];
                var value = cell.Value2;

                if (value != null && int.TryParse(value.ToString(), out int id))
                {
                    ids[col - 1] = id;
                }
            }

            return ids;
        }

        /// <summary>
        /// Apply formatting to a worksheet (freeze panes, auto-fit columns).
        /// </summary>
        public static void FormatWorksheet(Excel.Worksheet worksheet)
        {
            // Auto-fit all columns
            worksheet.Columns.AutoFit();

            // Freeze panes at row 3 (after IDs and column headers)
            Excel.Range freezePane = (Excel.Range)worksheet.Cells[3, 1];
            freezePane.Select();
            Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = true;
        }
    }
}
