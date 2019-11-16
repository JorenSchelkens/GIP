using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace DefaultDomain.ExcelReading
{
    public static class ExcelRead
    {
        public static List<Row> ReadEHB(string FilePath)
        {
            FileInfo existingFile = new FileInfo(FilePath);
            List<Row> rows = new List<Row>();

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                for (int i = 2; i <= rowCount; i++)
                {
                    Row row = new Row();

                    for (int j = 3; j <= colCount; j++)
                    {
                        row.columns.Add(worksheet.Cells[i, j].Value?.ToString().Trim());
                    }

                    rows.Add(row);
                }
            }

            return rows;
        }
    }
}