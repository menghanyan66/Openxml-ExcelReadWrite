using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ReadExcel
{
    class UserDetails
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WriteExcelFile();

            Console.ReadKey();

            string filePath = "/Users/工作簿5.xlsx";

            // Open the spreadsheet document
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;
                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                // Get the rows in the worksheet
                var rows = worksheet.Descendants<Row>();

                // Calculate the maximum width of values for each column
                var columnWidths = CalculateColumnWidths(rows, sharedStringPart);

                // Iterate through each row
                foreach (var row in rows)
                {
                    // Get the cells in the row
                    var cells = row.Descendants<Cell>();

                    // Format and print the row values with consistent column widths
                    foreach (var cell in cells)
                    {
                        int columnIndex = cell.CellReference != null ? GetColumnIndex(cell.CellReference.Value) : -1;

                        string cellValue = GetCellValue(cell, sharedStringPart);
                        int columnWidth = columnWidths.ContainsKey(columnIndex) ? columnWidths[columnIndex] : 0;

                        // Format the console output with padding to match column widths
                        Console.Write(cellValue.PadRight(columnWidth + 2));
                    }

                    Console.WriteLine();
                }
            }

            Console.ReadLine();
        }

        static void WriteExcelFile()
        {
            List<UserDetails> persons = new List<UserDetails>()
           {
               new UserDetails() {Id=1001, Name="ABCD", City = "City1", Country="USA"},
               new UserDetails() {Id=1002, Name="PQRS", City ="City2", Country="INDIA"},
               new UserDetails() {Id=1003, Name="XYZZ", City ="City3", Country="CHINA"},
               new UserDetails() {Id=1004, Name="LMNO", City ="City4", Country="UK"},
          };

            // Lets converts our object data to Datatable for a simplified logic.
            // DataTable is the easiest way to deal with complex datatypes for easy reading and formatting.
            DataTable table = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(persons));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create("/Users/工作簿5.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (string col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }

        // Helper method to calculate the maximum width of values for each column
        static Dictionary<int, int> CalculateColumnWidths(IEnumerable<Row> rows, SharedStringTablePart sharedStringPart)
        {
            Dictionary<int, int> columnWidths = new Dictionary<int, int>();

            foreach (Row row in rows)
            {
                var cells = row.Descendants<Cell>();

                foreach (Cell cell in cells)
                {
                    string cellValue = GetCellValue(cell, sharedStringPart);

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        int columnIndex = cell.CellReference != null ? GetColumnIndex(cell.CellReference.Value) : -1;


                        if (columnWidths.ContainsKey(columnIndex))
                        {
                            int cellValueLength = cellValue.Length;
                            if (cellValueLength > columnWidths[columnIndex])
                            {
                                columnWidths[columnIndex] = cellValueLength;
                            }
                        }
                        else
                        {
                            columnWidths.Add(columnIndex, cellValue.Length);
                        }
                    }
                }
            }

            return columnWidths;
        }

        // Helper method to get the value of a cell
        static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
        {
            string cellValue = string.Empty;

            // If the cell contains a value
            if (cell.CellValue != null)
            {
                cellValue = cell.CellValue.InnerText;

                // If the cell is a shared string
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int sharedStringIndex = int.Parse(cellValue);
                    SharedStringItem sharedStringItem = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                    cellValue = sharedStringItem.InnerText;
                }
            }
            else
            {
                cellValue = " ";
            }
            return cellValue;
        }

        // Helper method to get the column index of a cell
        static int GetColumnIndex(string cellReference)
        {
            string columnName = new Regex("[A-Za-z]+").Match(cellReference).Value;
            int columnIndex = 0;
            int mul = 1;

            foreach (char c in columnName.ToUpper().ToCharArray().Reverse())
            {
                columnIndex += mul * (c - 'A' + 1);
                mul *= 26;
            }

            return columnIndex - 1;
        }
    }
}
