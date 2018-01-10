using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace DatatableXlsPdfTools
{
    public class DataTableToXlsPdf
    {
        public static void ToFile(DataTable dataTable, FileInfo file)
        {
            var xlApp = new Application();
            if (xlApp == null)
            {
                throw new InvalidOperationException("Excel is not properly installed!!");
            }

            var missing = System.Reflection.Missing.Value;
            var workBook = xlApp.Workbooks.Add(missing);
            var workSheet = (Worksheet)workBook.Worksheets.Item[1];

            var rowIdx = 1;
            var colIdx = 1;
            foreach (var columnName in
                dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray())
            {
                workSheet.Cells[rowIdx, colIdx] = columnName;
                colIdx++;
            }
            rowIdx++;
            foreach (DataRow row in dataTable.Rows)
            {
                colIdx = 1;
                foreach (var item in row.ItemArray)
                {
                    workSheet.Cells[rowIdx, colIdx] = item;
                    colIdx++;
                }
                rowIdx++;
            }

            if (file.Extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                var range = workSheet.Cells;
                var border = range.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = 2d;
                workSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, file.FullName);

                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(border);
            }
            else
            {
                workBook.SaveAs(file.FullName);
            }
            
            workBook.Close(false);
            xlApp.Quit();

            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
