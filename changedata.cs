using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WriteExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "/Users/Desktop/PDI.xlsx";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;

                // Specify the data to search for
                string searchData = "ModeReason";

                // Find the cell containing the search data
                Cell targetCell = FindCellByValue(worksheet, workbookPart, searchData);
                if (targetCell != null)
                {
                    // Get the cell reference
                    string cellReference = targetCell.CellReference.Value;

                    // Get the row index and column name from the cell reference
                    uint rowIndex = GetRowIndexFromCellReference(cellReference);
                    string columnName = GetColumnNameFromCellReference(cellReference);

                    // Print the row index and column name
                    Console.WriteLine($"Data found at Row {rowIndex} and Column {columnName}");

                    // Specify the column and row where you want to update the data
                    string updateColumnName = columnName;
                    uint updateRowIndex = rowIndex;

                    // Get the cell reference for the update column and row
                    string updateCellReference = $"{updateColumnName}{updateRowIndex}";

                    // Check if the update cell already exists, or create a new one
                    Cell updateCell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference.Value == updateCellReference);
                    if (updateCell == null)
                    {
                        updateCell = new Cell() { CellReference = updateCellReference };
                        worksheetPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == updateRowIndex)?.Append(updateCell);
                    }

                    // Set the update cell value
                    updateCell.DataType = new EnumValue<CellValues>(CellValues.String);
                    updateCell.CellValue = new CellValue("changded");

                    // Save the changes to the spreadsheet document
                    worksheetPart.Worksheet.Save();
                }
                else
                {
                    Console.WriteLine("Data not found.");
                }
            }

            Console.WriteLine("Data written successfully!");
            Console.ReadLine();
        }

        // Function to find the cell containing the specified value in the worksheet
        static Cell FindCellByValue(Worksheet worksheet, WorkbookPart workbookPart, string searchValue)
        {
            foreach (Row row in worksheet.Descendants<Row>())
            {
                foreach (Cell cell in row.Descendants<Cell>())
                {
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int sharedStringId = int.Parse(cell.CellValue.Text);
                        string cellValue = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(sharedStringId).InnerText;

                        if (cellValue == searchValue)
                        {
                            return cell;
                        }
                    }
                    else if (cell.CellValue != null && cell.CellValue.Text == searchValue)
                    {
                        return cell;
                    }
                }
            }

            return null;
        }

        // Function to get the row index from the cell reference
        static uint GetRowIndexFromCellReference(string cellReference)
        {
            string rowIndex = string.Empty;
            foreach (char c in cellReference)
            {
                if (char.IsDigit(c))
                {
                    rowIndex += c;
                }
            }

            return uint.Parse(rowIndex);
        }

        // Function to get the column name from the cell reference
        static string GetColumnNameFromCellReference(string cellReference)
        {
            string columnName = string.Empty;
            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnName += c;
                }
            }

            return columnName;
        }
    }
}
