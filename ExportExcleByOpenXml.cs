using ClosedXML.Excel;
using System.Data;

namespace ConsoleApp1
{

    internal class Program
    {
        static void Main(string[] args)
        {
            var data = new DataTable();
            data.Columns.Add("a");
            data.Columns.Add("b");
            data.Columns.Add("c");
            data.Columns.Add("cc");
            data.Columns.Add("cd");
            data.Rows.Add(1, 2, 3, 33, 44);
            data.Rows.Add(4, 5, 6, 55, 66);

            var headers = new List<ExcelHeaderItem>();
            headers.Add(new ExcelHeaderItem() { HeaderName = "id",ColumnSpan=2 });
            headers.Add(new ExcelHeaderItem() { HeaderName = "datas", ColumnSpan = 2 });
            headers.Add(new ExcelHeaderItem() { HeaderName = "new column", ColumnSpan = 2, RowSpan = 2 });


            ExportToExcel(data, "a.xlsx", headers);
            Console.WriteLine("done");
        }

        private static void ExportToExcel(DataTable dataTable, string fileName, List<ExcelHeaderItem> excelHeaderItems)
        {
            if (dataTable == null)
            {
                return;
            }

            int startRow = 1;
            using (XLWorkbook workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("sheet");

                // write excel header
                if (excelHeaderItems == null || excelHeaderItems.Count == 0)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = dataTable.Columns[i].ColumnName;
                        worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                }
                else
                {
                    startRow = excelHeaderItems.Max(x => x.RowSpan);
                    int columnIndex = 0;
                    foreach (var item in excelHeaderItems)
                    {
                        if (item.RowSpan > 1 || item.ColumnSpan > 1)
                        {
                            worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1 + columnIndex).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1 + columnIndex).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1+ columnIndex).Value = item.HeaderName;
                            worksheet.Range(worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1 + columnIndex), worksheet.Cell(item.RowSpan, excelHeaderItems.IndexOf(item) + columnIndex+item.ColumnSpan)).Merge();
                            columnIndex += item.ColumnSpan-1;
                        }
                        else
                        {
                            worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1).Value = item.HeaderName;
                            worksheet.Cell(1, excelHeaderItems.IndexOf(item) + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }
                        worksheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FF996515");
                        worksheet.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.FromArgb(0xFF00FF);
                    }
                }

                startRow++;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cell(startRow + i, j + 1).Value = dataTable.Rows[i][j].ToString();
                        worksheet.Cell(startRow + i, j + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                }

                workbook.SaveAs(fileName);
            }
        }


        private static void DataTableReaderDemo()
        {
            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("city");
            dt.Rows.Add("zhag", "zunyi");
            dt.Rows.Add("wang", "dd");
            var reader = dt.CreateDataReader();
            while (reader.Read())
            {
                Console.WriteLine(reader.GetString(0));
                Console.WriteLine(reader.GetString(1));
            }
        }
    }
}
