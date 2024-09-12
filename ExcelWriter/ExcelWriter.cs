using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
namespace ExcelWriter
{
    public class ExcelWriter
    {
        public class DataSample // Any amount of columns can be used, underscores need to be used in the data model instead of spaces
        {
            public int ID { get; set; }
            public required string Name { get; set; }
            public int Age { get; set; }
            public bool Trained { get; set; }
            public DateTime Start_Date { get; set; }
            public float Application_Score { get; set; }
        }
        public static void CreateExcel<T>(List<T> data, string documentName, string path = "", string dataName = "")
        {
            if (path != "")
            {
                if (!Directory.Exists(path))
                {
                    try
                    {
                        Directory.CreateDirectory(path);
                    }
                    catch
                    {
                        path = "";
                    }
                }
                if (!path.EndsWith('\\'))
                {
                    path += "\\";
                }
            }
            string filePath = path + documentName + ".xlsx";
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = dataName
                };
                sheets.Append(sheet);
                Stylesheet stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(new Fill(new PatternFill() { PatternType = PatternValues.None })),
                    new Borders(new Border()),
                    new CellFormats(
                        new CellFormat(),
                        new CellFormat() { NumberFormatId = 49, ApplyNumberFormat = true },
                        new CellFormat() { NumberFormatId = 1, ApplyNumberFormat = true },
                        new CellFormat() { NumberFormatId = 4, ApplyNumberFormat = true },
                        new CellFormat() { NumberFormatId = 14, ApplyNumberFormat = true },
                        new CellFormat() { NumberFormatId = 4, ApplyNumberFormat = true }
                    )
                );
                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = stylesheet;
                stylesPart.Stylesheet.Save();
                Type dataModel = typeof(T);
                PropertyInfo[] properties = dataModel.GetProperties();
                Row headerRow = new Row();
                TableColumns tableColumns = new TableColumns() { Count = (UInt32)properties.Length };
                foreach (var property in properties)
                {
                    var tempName = property.Name.Replace('_', ' ');
                    tableColumns.Append(new TableColumn() { Name = tempName });
                    headerRow.AppendChild(CreateCell(tempName, CellValues.String, StyleIndex.Text));
                }
                TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
                var range = "A1:" + GetColumn(properties.Length) + data.Count.ToString();
                tableDefinitionPart.Table = new Table()
                {
                    Id = 1,
                    Name = dataName,
                    DisplayName = dataName,
                    Reference = range,
                    TotalsRowShown = false
                };
                tableDefinitionPart.Table.Append(tableColumns);
                tableDefinitionPart.Table.AppendChild(new AutoFilter() { Reference = range });
                var columns = new Columns();
                columns.Append(new Column
                {
                    Min = 1, Max = (UInt32)properties.Length, Width = 20
                });
                worksheetPart.Worksheet.InsertBefore(columns, sheetData);
                sheetData.AppendChild(headerRow);
                foreach (var dataRow in data)
                {
                    Row row = new Row();
                    foreach (var property in properties)
                    {
                        if (property.PropertyType == typeof(short) || property.PropertyType == typeof(ushort) || property.PropertyType == typeof(int) || property.PropertyType == typeof(uint) || property.PropertyType == typeof(long) || property.PropertyType == typeof(ulong))
                        {
                            row.AppendChild(CreateCell(Convert.ToString(property.GetValue(dataRow)) ?? "", CellValues.Number, StyleIndex.NumberWithoutDecimal));
                        }
                        else if (property.PropertyType == typeof(float) || property.PropertyType == typeof(double) || property.PropertyType == typeof(decimal))
                        {
                            row.AppendChild(CreateCell(Convert.ToString(property.GetValue(dataRow)) ?? "", CellValues.Number, StyleIndex.NumberWithDecimal));
                        }
                        else if (property.PropertyType == typeof(string))
                        {
                            row.AppendChild(CreateCell(Convert.ToString(property.GetValue(dataRow)) ?? "", CellValues.String, StyleIndex.Text));
                        }
                        else if (property.PropertyType == typeof(bool))
                        {
                            row.AppendChild(CreateCell(Convert.ToBoolean(property.GetValue(dataRow))));
                        }
                        else if (property.PropertyType == typeof(DateTime))
                        {
                            row.AppendChild(CreateCell(Convert.ToDateTime(property.GetValue(dataRow))));
                        }
                    }
                    sheetData.AppendChild(row);
                }
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
        }
        private static Cell CreateCell(string value, CellValues dataType, StyleIndex styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = dataType,
                StyleIndex = (UInt32)styleIndex
            };
        }
        private static Cell CreateCell(DateTime value)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = CellValues.Date,
                StyleIndex = (UInt32)StyleIndex.Date
            };
        }
        private static Cell CreateCell(bool value)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = CellValues.Boolean,
                StyleIndex = 0
            };
        }
        private enum StyleIndex
        {
            General = 0,
            Text = 1,
            NumberWithoutDecimal = 2,
            NumberWithDecimal = 3,
            Date = 4,
            Bool = 5,
        }
        private static string GetColumn(int cols)
        {
            string returnName = "";
            while (cols > 0)
            {
                int m = (cols - 1) % 26;
                returnName = Convert.ToChar(65 + m) + returnName;
                cols = (cols - m) / 26;
            }
            return returnName;
        }
    }
}
