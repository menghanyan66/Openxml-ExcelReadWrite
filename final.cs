using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        string sourceFilePath = "/Users/Desktop/Book1.xlsx";
        string sourceSheetName = "L3L2PrimaryData"; // Specify the name of the sheet you want to read

        string destinationFilePath = "/Users/Desktop/1.xlsx";
        string destinationSheetName = "Sheet1"; // Specify the name of the sheet you want to write

        // Read the data from the source Excel file
        int startRow, endRow;
        string[] values = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 1);
        string[] value2 = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 2);
        string[] value3 = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 10);

        List<string> appendedValues = new List<string>();

            // Iterate over the values and check if they match any sheet names in the source Excel file
        for (int i = 0; i < values.Length; i++)
        { 

            for (int j = 0; j < int.Parse(value3[i]); j++)
            {
                string sheetNameToCompare = value2[i];

                // Check if the sheet name exists in the source Excel file
                if (SheetExists(sourceFilePath, sheetNameToCompare))
                {
                    // Get the substituted values for the specific sheet
                    string[] substitutedValues = ReadExcelData(sourceFilePath, sheetNameToCompare, out startRow, out endRow, 1);
                    string[] newvalue3 = ReadExcelData(sourceFilePath, sheetNameToCompare, out startRow, out endRow, 10);

                    for (int k = 0; k < int.Parse(newvalue3[j]); k++)
                    {
                        // Append the substituted values to the list
                        if (substitutedValues != null)
                        {
                            appendedValues.AddRange(substitutedValues);
                        }
                    }
                }
                else
                {
                    // If the sheet name does not exist, append the original value to the list
                    appendedValues.Add(values[i]);
                }
            }
        }

        // Convert the list of appended values back to an array
        values = appendedValues.ToArray();

        if (values != null)
        {
            // Write the data to the destination Excel file
            WriteExcelData(destinationFilePath, destinationSheetName, values);

            Console.WriteLine("Data copied successfully!");
        }

        Console.ReadLine();
    }

    static bool SheetExists(string filePath, string sheetName)
    {
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            return (sheet != null);
        }
    }

    static string[] ReadExcelData(string filePath, string sheetName, out int startRow, out int endRow, int x)
    {
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = FindWorksheetPart(workbookPart, sheetName);

            if (worksheetPart == null)
            {
                Console.WriteLine($"Sheet '{sheetName}' not found.");
                startRow = endRow = -1;
                return null;
            }

            Worksheet worksheet = worksheetPart.Worksheet;
            SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

            // Get all the rows in the worksheet
            var rows = worksheet.Descendants<Row>().ToList();

            // Find the start row by checking the first column for an integer value
            startRow = FindStartRow(rows);

            // Find the end row by checking the first column for an integer value
            endRow = FindEndRow(rows);

            if (startRow == -1 || endRow == -1)
            {
                Console.WriteLine("No rows with integer values found in the first column.");
                return null;
            }

            int startRowIndex = startRow + 3;
            int endRowIndex = endRow;

            rows = rows.Where(r => r.RowIndex >= (uint)startRowIndex && r.RowIndex <= (uint)endRowIndex).ToList();


            string[] values = new string[rows.Count];

            // Iterate through the rows and get the cell value from the specified column
            int rowIndex = 0;
            foreach (var row in rows)
            {
                Cell cell = row.Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == x); // Assuming the first column has index 1
                if (cell != null)
                {
                    string cellValue = GetCellValue(cell, sharedStringPart);
                    values[rowIndex] = cellValue;
                }
                rowIndex++;
            }

            return values;
        }
    }

    static int FindStartRow(List<Row> rows)
    {
        for (int i = 0; i < rows.Count; i++)
        {
            Cell cell = rows[i].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == 1); // Assuming the first column has index 1
            if (cell != null && int.TryParse(GetCellValue(cell), out int _))
            {
                return (int)rows[i].RowIndex.Value;
            }
        }
        return -1;
    }

    static int FindEndRow(List<Row> rows)
    {
        for (int i = rows.Count - 1; i >= 0; i--)
        {
            Cell cell = rows[i].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == 1); // Assuming the first column has index 1
            if (cell != null && int.TryParse(GetCellValue(cell), out int _))
            {
                return (int)rows[i].RowIndex.Value;
            }
        }
        return -1;
    }

    static string GetCellValue(Cell cell)
    {
        if (cell.CellValue != null)
        {
            return cell.CellValue.InnerText;
        }
        return string.Empty;
    }

    static void WriteExcelData(string filePath, string sheetName, string[] values)
    {
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();

            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            int rowIndex = 1; // Start from row 1

            Row headerRow = new Row() { RowIndex = (uint)rowIndex };

            for (int columnIndex = 0; columnIndex < values.Length; columnIndex++)
            {
                string columnName = GetColumnName(columnIndex);
                string cellAddress = columnName + rowIndex;

                Cell headerCell = new Cell() { CellReference = cellAddress, DataType = CellValues.String };
                headerCell.CellValue = new CellValue("Header " + (columnIndex + 1));

                headerRow.Append(headerCell);
            }

            sheetData.Append(headerRow);

            rowIndex++;

            Row dataRow = new Row() { RowIndex = (uint)rowIndex };

            for (int columnIndex = 0; columnIndex < values.Length; columnIndex++)
            {
                string columnName = GetColumnName(columnIndex);
                string cellAddress = columnName + rowIndex;

                Cell dataCell = new Cell() { CellReference = cellAddress, DataType = CellValues.String };
                dataCell.CellValue = new CellValue(values[columnIndex]);

                dataRow.Append(dataCell);
            }

            sheetData.Append(dataRow);

            worksheetPart.Worksheet.Save();
        }
    }

    static string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex + 1;
        string columnName = String.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    static Cell GetOrCreateCell(SheetData sheetData, string cellAddress)
    {
        var existingCell = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
        if (existingCell != null)
        {
            return existingCell;
        }

        string columnName = new Regex("[A-Za-z]+").Match(cellAddress).Value;
        string rowIndex = new Regex("[0-9]+").Match(cellAddress).Value;

        Row row = sheetData.Descendants<Row>().FirstOrDefault(r => r.RowIndex == uint.Parse(rowIndex));
        if (row == null)
        {
            row = new Row() { RowIndex = uint.Parse(rowIndex) };
            sheetData.Append(row);
        }

        Cell newCell = new Cell() { CellReference = cellAddress };
        row.Append(newCell);

        return newCell;
    }

    static WorksheetPart GetOrCreateWorksheetPart(SpreadsheetDocument spreadsheetDocument, string sheetName)
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        WorksheetPart worksheetPart = FindWorksheetPart(workbookPart, sheetName);

        if (worksheetPart != null)
        {
            return worksheetPart;
        }

        WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        newWorksheetPart.Worksheet = new Worksheet(new SheetData());
        newWorksheetPart.Worksheet.Save();

        Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Any())
        {
            sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
        }

        Sheet newSheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(newSheet);

        return newWorksheetPart;
    }

    static WorksheetPart FindWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

        if (sheet != null)
        {
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }

        return null;
    }

    static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
    {
        string cellValue = string.Empty;

        if (cell.CellValue != null)
        {
            cellValue = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.Number)
            {
                // Handle numeric values
                if (double.TryParse(cellValue, out double numericValue))
                {
                    cellValue = numericValue.ToString();
                }
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Handle shared string values
                int sharedStringIndex = int.Parse(cellValue);
                SharedStringItem sharedStringItem = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                cellValue = sharedStringItem.InnerText;
            }
        }

        return cellValue;
    }


    static int GetColumnIndex(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return -1;
        }

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
