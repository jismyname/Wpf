using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using System.Data;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string file = "a.xlsx";
            var sheetdoc = SpreadsheetDocument.Create(file, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);

            var workbookpart = sheetdoc.AddWorkbookPart();
            workbookpart.Workbook = new();

            var worksheetpart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetpart.Worksheet = new Worksheet(new SheetData());


            workbookpart.AddNewPart<WorkbookStylesPart>();
            workbookpart.WorkbookStylesPart.Stylesheet = new();
            var styleSheet = workbookpart.WorkbookStylesPart.Stylesheet;
            List<UInt32> indexRef = addheaderstyle(ref styleSheet);
            styleSheet.Save();

            var sheets = sheetdoc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            var sheet = new Sheet()
            {
                Id = sheetdoc.WorkbookPart.GetIdOfPart(worksheetpart),
                SheetId = 1,
                Name = "sheet 1"
            };

            sheets.Append(sheet);
            var worksheet = worksheetpart.Worksheet;

            var sheetdata = worksheet.GetFirstChild<SheetData>();

            Merge(worksheet);

            var r = new Row();
            var cell = new Cell()
            {
                CellValue = new CellValue("j developer"),
                DataType = CellValues.String,
                CellReference="A1"
            };
            r.AppendChild(cell);
            cell.StyleIndex = indexRef[2];
           

            sheetdata.Append(r);

            ExportData(sheetdata);

            workbookpart.Workbook.Save();
            sheetdoc.Dispose();
            Console.WriteLine("done");

        }

        private static void Merge(Worksheet worksheet)
        {
            MergeCells mergeCells = new();
            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());

            var mergecell = new MergeCell() { Reference = new StringValue("A1:B1") };
            mergeCells.Append(mergecell);
            worksheet.Save();
        }

        private static void ExportData(SheetData sheetdata)
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("a");
            dt.Columns.Add("b");
            dt.Columns.Add("c");
            dt.Rows.Add(1, 2, 3);
            dt.Rows.Add(4, 5, 6);
            UInt32Value rowindex = 2;

            foreach (DataRow item in dt.Rows)
            {
                var row = new Row() { RowIndex = rowindex };
                for (int i = 0; i < item.ItemArray.Count(); i++)
                {
                    var c = new Cell()
                    {
                        CellValue = new CellValue(item[i].ToString()),
                        DataType = CellValues.String
                    };
                    row.Append(c);
                }
                sheetdata.Append(row);
                rowindex++;
            }


        }

        private static List<uint> addheaderstyle(ref Stylesheet styleSheet)
        {
            UInt32 fontid = 0, fillid = 0, cellformatid = 0;

            Font font = new Font(new FontSize() { Val = 15 },
                new Color() { Rgb = HexBinaryValue.FromString("252525") });
            styleSheet.Fonts = new();
            styleSheet.Fonts.AppendChild(font);
            styleSheet.Fonts.Count = new() { Value = 0 };
            fontid = styleSheet.Fonts.Count++;

            var patternfill = new PatternFill() { PatternType = PatternValues.Solid };
            patternfill.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString("252525") };

            styleSheet.Fills = new();
            styleSheet.Fills.Append(new Fill() { PatternFill = patternfill });
            styleSheet.Fills.Count = new() { Value = 0 };
            fillid = styleSheet.Fills.Count++;

            var cellformat = new CellFormat()
            {
                FontId = fontid,
                FillId = fillid,
                ApplyFill = true,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center }
            };

            styleSheet.CellFormats = new();
            styleSheet.CellFormats.AppendChild(cellformat);
            styleSheet.CellFormats.Count = new() { Value = 0 };
            cellformatid = styleSheet.CellFormats.Count++;

            return new List<uint> { fontid, fillid, cellformatid };
        }
    }



}
