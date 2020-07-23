using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IExcelDocumentsModule;
using IExcelDocumentsModule.Enums;
using IExcelDocumentsModule.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelDocumentsModule
{
    public class ExcelModule : IExcelModule
    {
        #region Methods

        public async Task WriteDocument(DataSet dataSet, string outputPath)
        {
            outputPath = await PreparePath(outputPath);

            using (var workbook = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                uint sheetId = 1;
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                foreach (DataTable table in dataSet.Tables)
                {
                    var sheetData = new SheetData();
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    UInt32 rowIdex = 0;
                    var cellIdex = 0;
                    var headerRow = new Row { RowIndex = ++rowIdex };
                    var columns = new List<string>();
                    var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };

                    sheets.Append(sheet);

                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        headerRow.AppendChild(await CreateTextCell(await GetColumnLetter(cellIdex++),
                        rowIdex, column.ColumnName ?? string.Empty));
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        cellIdex = 0;
                        var newRow = new Row { RowIndex = ++rowIdex };

                        foreach (String col in columns)
                        {
                            newRow.AppendChild(await CreateTextCell(await GetColumnLetter(cellIdex++),
                            rowIdex, dsrow[col].ToString() ?? string.Empty));
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }

        public async Task<DataSet> ReadDocument(string fileName, bool isFirstRowHeader = false)
        {
            var fileExtension = Path.GetExtension(fileName);

            if (!fileExtension.IsSupportedExcelFile())
                throw new FileFormatException("The file format is not supported!");

            var dataSet = new DataSet();

            if (fileExtension == SupportFormats.csv)
            {
                var table = new DataTable();

                using (var stremReader = new StreamReader(fileName))
                {
                    var headers = stremReader.ReadLine().Split(',');

                    foreach (string header in headers)
                        table.Columns.Add(header);

                    while (!stremReader.EndOfStream)
                    {
                        var rows = stremReader.ReadLine().Split(',');
                        var dataRow = table.NewRow();

                        for (int i = 0; i < headers.Length; i++)
                            dataRow[i] = rows[i];

                        table.Rows.Add(dataRow);
                    }
                }

                dataSet.Tables.Add(table);

                return dataSet;
            }
            else
            {
                try
                {
                    using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                        {
                            var workbookPart = doc.WorkbookPart;

                            #region Sheets processing

                            foreach (var sheet in workbookPart.Workbook.Descendants<Sheet>())
                            {
                                var table = new DataTable(sheet.Name);
                                var dataList = new List<TableRowItem>();
                                var sheetColumns = new Dictionary<string, string>();
                                var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                                dataSet.Tables.Add(table);

                                if (sheetData.Elements<Row>().Any(x => x.RowIndex == null))
                                    throw new NotSupportedException("The version of this file is not supported!");

                                #region Define header

                                if (isFirstRowHeader)
                                {
                                    foreach (var row in sheetData?.Elements<Row>()?.Where(x => x.RowIndex == 1))
                                    {
                                        if (row.Elements<Cell>().Any())
                                        {
                                            foreach (var cell in row.Elements<Cell>())
                                            {
                                                var columnAddress = await GetColumnName(cell.CellReference.Value);
                                                var columnName = await GetCellValue(cell, workbookPart);

                                                sheetColumns.TryGetValue(columnName, out string colName);

                                                if (string.IsNullOrEmpty(colName))
                                                    sheetColumns.TryAdd(columnAddress, columnName);
                                            }
                                        }
                                        else
                                        {
                                            throw new MissingFieldException("The column name in the first row was not found!");
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var row in sheetData?.Elements<Row>())
                                    {
                                        foreach (var cell in row.Elements<Cell>())
                                        {
                                            var columnAddress = await GetColumnName(cell.CellReference.Value);
                                            sheetColumns.TryGetValue(columnAddress, out string colName);

                                            if (string.IsNullOrEmpty(colName))
                                                sheetColumns.TryAdd(columnAddress, columnAddress);
                                        }
                                    }
                                }

                                #endregion

                                #region Sort columns

                                var columnsList = sheetColumns.Values.ToList();

                                if (!isFirstRowHeader)
                                    columnsList.Sort();

                                foreach (var col in columnsList)
                                    table.Columns.Add(col);

                                #endregion

                                #region Filling up dataSet

                                foreach (var row in sheetData.Elements<Row>())
                                {
                                    var dtsTableRow = table.NewRow();

                                    foreach (var cell in row.Elements<Cell>())
                                    {
                                        var columnAddress = await GetColumnName(cell.CellReference.Value);
                                        var columnIndex = await GetColumnIndex(cell.CellReference.Value);

                                        sheetColumns.TryGetValue(columnAddress, out string columnName);
                                        int.TryParse(columnIndex, out int index);

                                        var cellValue = await GetCellValue(cell, workbookPart);
                                        var firstField = isFirstRowHeader ? 1 : 0;

                                        if (index > firstField)
                                        {
                                            dataList.Add(new TableRowItem
                                            {
                                                Index = index,
                                                Name = columnName,
                                                Value = cellValue
                                            });
                                        }
                                    }
                                }

                                for (var i = isFirstRowHeader ? 2 : 1; i <= dataList.Max(x => x.Index); i++)
                                {
                                    var dataRow = table.NewRow();
                                    var tempList = dataList.Where(x => x.Index == i).ToList();

                                    foreach (var item in tempList)
                                        dataRow[item.Name] = item.Value;

                                    table.Rows.Add(dataRow);
                                }

                                #endregion
                            }

                            #endregion

                        }
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            return dataSet;
        }

        private async Task<string> PreparePath(string path)
        {
            var fileName = Path.GetFileName(path);
            var filePath = Path.GetDirectoryName(path);

            if (string.IsNullOrEmpty(filePath))
                throw new NullReferenceException("The output path cnnot be empty!");

            if (!Directory.Exists(filePath))
                throw new NullReferenceException("Directory is not exists!");

            return $@"{filePath}{Path.DirectorySeparatorChar}{fileName}";
        }

        private async Task<string> GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null)
                return null;

            var value = cell.CellFormula != null ? cell.CellValue.InnerText : cell.InnerText.Trim();

            if (cell.DataType == null)
                return value;

            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    if (stringTable != null)
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;

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

            return value;
        }

        private async Task<string> GetColumnIndex(string value) => Regex.Replace(value, @"\D", "");

        private async Task<string> GetColumnName(string value) => Regex.Replace(value, @"\d", "");

        private async Task<string> GetColumnLetter(int intCol)
        {
            var numOfFirstLetter = ((intCol) / 676) + 64;
            var numOfSecondLetter = ((intCol % 676) / 26) + 64;
            var numOfThirdLetter = (intCol % 26) + 65;

            var firstLetter = (numOfFirstLetter > 64)
                ? (char)numOfFirstLetter : ' ';
            var secondLetter = (numOfSecondLetter > 64)
                ? (char)numOfSecondLetter : ' ';
            var thirdLetter = (char)numOfThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        private async Task<Cell> CreateTextCell(string header, UInt32 index, string text)
        {
            var cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index
            };

            var inlineString = new InlineString();
            var appendText = new Text { Text = text };

            inlineString.AppendChild(appendText);
            cell.AppendChild(inlineString);
            return cell;
        }

        #endregion
    }
}
