using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;

namespace final
{
    public partial class Form1 : Form
    {
        private TextBox textBox1;
        private TextBox textBox2;
        private string sourceFilePath;
        private string destinationFilePath;
        private string sourceSheetName;

        public Form1()
        {
            textBox1 = new TextBox();
            textBox2 = new TextBox();

            textBox1.Location = new System.Drawing.Point(100, 50);
            textBox1.Size = new System.Drawing.Size(200, 20);
            Controls.Add(textBox1);

            textBox2.Location = new System.Drawing.Point(100, 100);
            textBox2.Size = new System.Drawing.Size(200, 20);
            Controls.Add(textBox2);

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
            openFileDialog.Title = "Select Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                sourceFilePath = openFileDialog.FileName;
                textBox1.Text = sourceFilePath;
            }

            listBox1.Items.Clear();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(sourceFilePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet sheet in workbookPart.Workbook.Descendants<Sheet>())
                {
                    listBox1.Items.Add(sheet.Name);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                sourceSheetName = listBox1.SelectedItem.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"error:{ex.Message}");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                MessageBox.Show("Please select the source Excel file.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
            saveFileDialog.Title = "Save Excel File";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                destinationFilePath = saveFileDialog.FileName;
                textBox2.Text = destinationFilePath;

                string destinationSheetName = "Sheet1";
                // Read data from the source Excel file
                int startRow, endRow;
                string[] values = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 2);
                string[] value2 = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 3);
                string[] value3 = ReadExcelData(sourceFilePath, sourceSheetName, out startRow, out endRow, 11);

                List<string> appendedValues = new List<string>();
                //MessageBox.Show(Convert.ToString(startRow));
                //MessageBox.Show(Convert.ToString(endRow));
                //MessageBox.Show(values[1]);

                for (int i = 0; i < (values.Length); i++)
                {
                    for (int j = 0; j < int.Parse(value3[i]); j++)
                    {
                        string sheetNameToCompare = value2[i];


                        if (SheetExists(sourceFilePath, sheetNameToCompare))
                        {
                            //MessageBox.Show(Convert.ToString(startRow));
                            //MessageBox.Show(Convert.ToString(endRow));
                            string[] substitutedValues = ReadExcelData(sourceFilePath, sheetNameToCompare, out startRow, out endRow, 2);
                            string[] newvalue3 = ReadExcelData(sourceFilePath, sheetNameToCompare, out startRow, out endRow, 11);

                            for (int k = 0; k < int.Parse(newvalue3[j]); k++)
                            {
                                if (substitutedValues != null)
                                {
                                    appendedValues.AddRange(substitutedValues);
                                }
                            }
                        }
                        else
                        {
                            appendedValues.Add(values[i]);
                        }
                    }
                }

                values = appendedValues.ToArray();

                if (values != null)
                {
                    // Write data to the destination Excel file
                    WriteExcelData(destinationFilePath, destinationSheetName, values);

                    MessageBox.Show("Data transferred successfully.");
                }
            }
        }

        private static bool SheetExists(string filePath, string sheetName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                return (sheet != null);
            }
        }

        private static string[] ReadExcelData(string filePath, string sheetName, out int startRow, out int endRow, int x)
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

        private static int FindStartRow(List<Row> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                Cell cell = rows[i].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == 2); // Assuming the first column has index 1
                if (cell != null && int.TryParse(GetCellValue(cell), out int _))
                {
                    return (int)rows[i].RowIndex.Value;
                }
            }
            return -1;
        }

        private static string GetCellValue(Cell cell, SharedStringTablePart sharedStringPart)
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

        private static int FindEndRow(List<Row> rows)
        {
            for (int i = rows.Count - 1; i >= 0; i--)
            {
                Cell cell = rows[i].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == 2); // Assuming the first column has index 1
                if (cell != null && int.TryParse(GetCellValue(cell), out int _))
                {
                    return (int)rows[i].RowIndex.Value;
                }
            }
            return -1;
        }

        private static string GetCellValue(Cell cell)
        {
            if (cell.CellValue != null)
            {
                return cell.CellValue.InnerText;
            }
            return string.Empty;
        }

        private static void WriteExcelData(string filePath, string sheetName, string[] values)
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

        private static WorksheetPart FindWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
                return null;

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private static int GetColumnIndex(string cellReference)
        {
            string columnName = Regex.Replace(cellReference, @"[\d]", string.Empty);
            int columnIndex = 0;
            int factor = 1;

            for (int i = columnName.Length - 1; i >= 0; i--)
            {
                columnIndex += (columnName[i] - 'A' + 1) * factor;
                factor *= 26;
            }

            return columnIndex;
        }
    }
}
