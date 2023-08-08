using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Random_Dates
{
    internal class ExcelHandler
    {
        private static List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };

        public DataTable ImportTableOnlyNonEmptyCellValues(DataTable dt, string FileName)
        {
            //Open excel file
            Statics.currentProgressValue = 0;
            Statics.maxProgressValue = 100;
            Statics.cancelProgress = false;
            try
            {
                bool fileExist = File.Exists(FileName);
                if (fileExist)
                {
                    using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(FileName, false))
                    {
                        var wbPart = myDoc.WorkbookPart;
                        var wsPart = wbPart.WorksheetParts.First();
                        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

                        var cellValue = GetCellValue(GetCell(sheetData, "C2"), wbPart);

                        IEnumerable<Row> rows = sheetData.Elements<Row>();
                        var rowCount = rows.Count();
                        if (rowCount > 0)
                        {
                            if (dt.Columns.Count < rows.ElementAt(0).ChildElements.Count)
                                throw new Exception($"Expects at least {dt.Columns.Count} columns.");
                            //foreach (Row r in row)
                            //{
                            Statics.maxProgressValue = rowCount;
                            for (int i = 1; i < rowCount; i++)
                            {
                                Statics.currentProgressValue += 1;
                                dt.Rows.Add();
                                for (int j = 1; j < dt.Columns.Count; j++)
                                {
                                    var cell = ((Cell)rows.ElementAt(i).ChildElements[j]);
                                    if (cell.CellValue != null)
                                        dt.Rows[i - 1][j] = cell.InnerText.Trim();
                                    else
                                        dt.Rows[i - 1][j] = "";
                                }
                                if (Statics.cancelProgress == true)
                                    break;
                            }
                            //}
                        }

                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                Statics.cancelProgress = true;
                throw new Exception("Error exporting data." +
                    Environment.NewLine + ex.Message);
            }
            return dt;
        }


        public static Cell GetCell(SheetData sheetData, string cellAddress)
        {
            uint rowIndex = uint.Parse(Regex.Match(cellAddress, @"[0-9]+").Value);
            return sheetData.Descendants<Row>().FirstOrDefault(p => p.RowIndex == rowIndex).Descendants<Cell>().FirstOrDefault(p => p.CellReference == cellAddress);
        }

        public static string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            if (cell == null) return "";
            string value = cell.InnerText;
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                        {
                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            return value;
        }

        public DataTable ImportTable(string FileName)
        {
            //Open excel file
            Statics.currentProgressValue = 0;
            Statics.maxProgressValue = 100;
            Statics.cancelProgress = false;
            DataTable dt = new DataTable();
            try
            {
                bool fileExist = File.Exists(FileName);
                if (fileExist)
                {
                    using (XLWorkbook workBook = new XLWorkbook(FileName))
                    {
                        IXLWorksheet workSheet = workBook.Worksheet(1);

                        var rowCount = workSheet.RangeUsed().RowCount();
                        if (rowCount < 5)
                            throw new Exception($"Expects at least {5} rows.");

                        //var colCount = workSheet.Row(1).CellsUsed().Count();
                        var colCount = workSheet.RangeUsed().ColumnCount();
                        if (colCount < 2)
                            throw new Exception($"Expects at least {2} columns.");

                        var rows = workSheet.Rows();

                        for (int i = 0; i < colCount; i++) {
                            dt.Columns.Add();
                            dt.Columns[dt.Columns.Count - 1].ColumnName = 
                                rows.ElementAt(2).Cell(i + 1).Value.ToString().Trim();
                        }

                        //Loop through the Worksheet rows.
                        Statics.maxProgressValue = rowCount;
                        for (int i = 0; i < rowCount - 1; i++)
                        {
                            Statics.currentProgressValue += 1;
                            dt.Rows.Add();
                            for (int j = 1; j <= dt.Columns.Count; j++)
                            {
                                var cell = (rows.ElementAt(i).Cell(j));
                                if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                    dt.Rows[i][j - 1] = cell.Value.ToString().Trim();
                                else
                                    dt.Rows[i][j - 1] = "";
                            }
                            if (Statics.cancelProgress == true)
                                break;
                        }
                        return dt;
                    }
                }
            }
            catch (Exception ex)
            {
                Statics.cancelProgress = true;
                throw new Exception("Error importing data:" +
                    Environment.NewLine + ex.Message);
            }
            return dt;
        }

        public bool ExportTable(DataTable dt, string FileName)
        {
            using (var workbook = SpreadsheetDocument.Create(FileName, 
                DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = 
                    workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(
                            s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = 
                    new DocumentFormat.OpenXml.Spreadsheet.Sheet() { 
                        Id = relationshipId, SheetId = sheetId, Name = dt.TableName 
                    };
                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in dt.Columns)
                {
                    columns.Add(column.ColumnName);

                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }


                sheetData.AppendChild(headerRow);

                foreach (System.Data.DataRow dsrow in dt.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
            }

            return true;
        }
    }
}
