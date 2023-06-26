using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        // Specify the file path of the Excel file
        string filePath = "/Users/Desktop/1.xlsx";

        // Specify the cell address where you want to enter the data
        string cellAddress = "A1";

        // Specify the data you want to enter into the cell
        string data = "Hello, World!";

        // Open the Excel file using OpenXML
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
        {
            // Get the worksheet part
            WorksheetPart worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();

            // Get the sheet data
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Get the cell reference
            Cell cell = GetOrCreateCell(sheetData, cellAddress);

            // Set the value of the cell
            cell.CellValue = new CellValue(data);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);

            // Save the changes
            worksheetPart.Worksheet.Save();
        }
    }

    static Cell GetOrCreateCell(SheetData sheetData, string cellAddress)
    {
        string columnName = GetColumnName(cellAddress);
        uint rowIndex = GetRowIndex(cellAddress);

        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellAddress };
            row.Append(cell);
        }

        return cell;
    }

    static string GetColumnName(string cellAddress)
    {
        StringBuilder columnName = new StringBuilder();
        foreach (char c in cellAddress)
        {
            if (char.IsLetter(c))
                columnName.Append(c);
            else
                break;
        }
        return columnName.ToString();
    }

    static uint GetRowIndex(string cellAddress)
    {
        StringBuilder rowIndex = new StringBuilder();
        for (int i = cellAddress.Length - 1; i >= 0; i--)
        {
            if (char.IsDigit(cellAddress[i]))
                rowIndex.Insert(0, cellAddress[i]);
            else
                break;
        }
        return uint.Parse(rowIndex.ToString());
    }
}
