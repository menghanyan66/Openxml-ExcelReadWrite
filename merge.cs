using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "/Users/Desktop/工作簿1.xlsx";
        string sheetName = "Sheet1";
        int rowNumber = 2;
        int startColumn = 8; // Column B
        int endColumn = 12;   // Column E

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
        {
            Worksheet worksheet = GetWorksheetByName(spreadsheetDocument, sheetName);
            MergeCells mergeCells = GetMergeCells(worksheet);

            string startCellReference = GetCellReference(rowNumber, startColumn);
            string endCellReference = GetCellReference(rowNumber, endColumn);

            mergeCells.Append(new MergeCell() { Reference = new StringValue($"{startCellReference}:{endCellReference}") });

            foreach (Cell cell in worksheet.Descendants<Cell>())
            {
                if (CellReferenceInRange(cell.CellReference, startCellReference, endCellReference))
                {
                    cell.CellValue = new CellValue("Merged Content");
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    cell.StyleIndex = 1;
                }
            }

            worksheet.Save();
        }
    }

    static Worksheet GetWorksheetByName(SpreadsheetDocument spreadsheetDocument, string sheetName)
    {
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

        if (sheet == null)
        {
            throw new ArgumentException($"Sheet '{sheetName}' not found in the workbook.");
        }

        WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
        return worksheetPart.Worksheet;
    }

    static MergeCells GetMergeCells(Worksheet worksheet)
    {
        return worksheet.Elements<MergeCells>().FirstOrDefault() ?? worksheet.InsertAfter(new MergeCells(), worksheet.Elements<SheetData>().FirstOrDefault());
    }

    static string GetCellReference(int rowNumber, int columnNumber)
    {
        string columnName = GetColumnName(columnNumber);
        return $"{columnName}{rowNumber}";
    }

    static string GetColumnName(int columnNumber)
    {
        int dividend = columnNumber;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            char columnLetter = (char)('A' + modulo);
            columnName = columnLetter + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    static bool CellReferenceInRange(string cellReference, string startCellReference, string endCellReference)
    {
        int startColumn = GetColumnIndexFromCellReference(startCellReference);
        int endColumn = GetColumnIndexFromCellReference(endCellReference);
        int column = GetColumnIndexFromCellReference(cellReference);

        return (column >= startColumn && column <= endColumn);
    }

    static int GetColumnIndexFromCellReference(string cellReference)
    {
        int columnIndex = 0;
        string columnName = string.Empty;

        foreach (char c in cellReference)
        {
            if (char.IsLetter(c))
            {
                columnName += c;
            }
            else
            {
                break;
            }
        }

        int power = 1;

        for (int i = columnName.Length - 1; i >= 0; i--)
        {
            columnIndex += (columnName[i] - 'A' + 1) * power;
            power *= 26;
        }

        return columnIndex;
    }
}

