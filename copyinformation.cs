using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string sourceFilePath = "/Users/Desktop/工作簿1.xlsx";
        string destinationFilePath = "/Users/Desktop/工作簿2.xlsx";

        using (SpreadsheetDocument sourceDoc = SpreadsheetDocument.Open(sourceFilePath, false))
        {
            WorkbookPart sourceWorkbookPart = sourceDoc.WorkbookPart;
            WorksheetPart sourceWorksheetPart = sourceWorkbookPart.WorksheetParts.First();

            using (SpreadsheetDocument destinationDoc = SpreadsheetDocument.Open(destinationFilePath, true))
            {
                WorkbookPart destinationWorkbookPart = destinationDoc.WorkbookPart;
                WorksheetPart destinationWorksheetPart = destinationWorkbookPart.WorksheetParts.First();

                SheetData sourceSheetData = sourceWorksheetPart.Worksheet.Elements<SheetData>().First();
                SheetData destinationSheetData = destinationWorksheetPart.Worksheet.Elements<SheetData>().First();

                foreach (Row sourceRow in sourceSheetData.Elements<Row>())
                {
                    Row destinationRow = new Row();

                    foreach (Cell sourceCell in sourceRow.Elements<Cell>())
                    {
                        Cell destinationCell = new Cell();

                        // Copy the cell value
                        if (sourceCell.CellValue != null)
                        {
                            string cellValue = sourceCell.CellValue.InnerText;
                            destinationCell.CellValue = new CellValue(cellValue);
                            destinationCell.DataType = sourceCell.DataType;
                        }

                        // Copy the cell style if applicable
                        if (sourceCell.StyleIndex != null)
                        {
                            destinationCell.StyleIndex = sourceCell.StyleIndex;
                        }

                        destinationRow.AppendChild(destinationCell);
                    }

                    destinationSheetData.AppendChild(destinationRow);
                }

                destinationWorksheetPart.Worksheet.Save();
            }
        }

        Console.WriteLine("Data copied successfully!");
    }
}
