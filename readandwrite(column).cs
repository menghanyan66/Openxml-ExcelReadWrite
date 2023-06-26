using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Vml.Office;

namespace ReadExcel
{

    class Program
    {
        static void Main(string[] args)
        { 

            string filePath = "/Users/Desktop/工作簿1.xlsx";

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

                // Create a dictionary to store the column values
                Dictionary<int, List<string>> columnValues = new Dictionary<int, List<string>>();

                // Iterate through each row
                foreach (var row in rows)
                {
                    // Iterate through each cell in the row
                    foreach (var cell in row.Elements<Cell>())
                    {
                        int columnIndex = GetColumnIndex(cell.CellReference);

                        string cellValue = GetCellValue(cell, sharedStringPart);

                        // Check if the column index already exists in the dictionary
                        if (columnValues.ContainsKey(columnIndex))
                        {
                            // Add the cell value to the existing column's values
                            columnValues[columnIndex].Add(cellValue);
                        }
                        else
                        {
                            // Create a new list for the column values and add the cell value
                            List<string> values = new List<string>();
                            values.Add(cellValue);
                            columnValues.Add(columnIndex, values);
                        }
                    }
                }

                // Print the column values from the dictionary
                foreach (var kvp in columnValues)
                {
                    int columnIndex = kvp.Key;
                    List<string> values = kvp.Value;

                    // Print the column index
                    Console.Write($"Column {GetColumnNameFromIndex(columnIndex)}: ");

                    // Iterate through the values and print them
                    foreach (var value in values)
                    {
                        Console.Write(value + " ");
                    }

                    Console.WriteLine();
                }
            }

            Console.ReadLine();

        }

        public static string GetColumnNameFromIndex(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            if (columnIndex == 0)
            {
                return "A";
            }
            else{
                return columnName;
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
                cellValue = "blank";
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
