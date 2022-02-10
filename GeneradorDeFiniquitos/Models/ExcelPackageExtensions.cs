using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GeneradorDeFiniquitos.Models
{
    public static class ExcelPackageExtensions
    {
        public static DataTable ToDataTable(this ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
            DataTable dt = new DataTable();
            int colCount = worksheet.Dimension.End.Column;
            int rowCount = worksheet.Dimension.End.Row;

            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, colCount])
            {
                dt.Columns.Add(firstRowCell.Text);
            }
            for (int rowNumber = 2; rowNumber <= rowCount; rowNumber++)
            {
                var row = worksheet.Cells[rowNumber, 1, rowNumber, colCount];
                var newRow = dt.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(newRow);
            }
            return dt;
        }
    }
}