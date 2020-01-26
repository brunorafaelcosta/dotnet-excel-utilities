using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;

namespace dotnet_excel_utilities
{
    /*
     * ToDo:
     *      - Handle other cell value types
     *      - Column auto width
     *      - Header and Column cell styling
     *      - Import exception handling
     */
    public class ExcelUtilities
    {
        const string WorksheetName = "Data";
        
        private enum RowType
        {
            Header = 1,
            Column,
            Data
        }

        #region Configuration Attributes

        [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
        public class ExportTableAttribute : Attribute
        {
            private string _header;
            private bool _hasChildren;

            public ExportTableAttribute(string header)
            {
                Header = header;
            }

            public string Header { get => _header; set { _header = value ?? throw new ArgumentNullException(nameof(value)); } }

            public bool HasChildren { get => _hasChildren; set => _hasChildren = value; }
        }

        [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
        public class ExportColumnAttribute : Attribute
        {
            public string Title { get; set; }
        }

        [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
        public class ExportColumnTableAttribute : Attribute
        {
            public ExportColumnTableAttribute()
            {
                IsCollapsed = false;
            }

            public bool IsCollapsed { get; set; }
        }

        #endregion Configuration Attributes

        public interface IData { }

        #region Import

        public ICollection<TRootData> Import<TRootData>(string fileName)
            where TRootData : IData, new()
        {
            if (!File.Exists(fileName))
                throw new FileNotFoundException();
            
            List<TRootData> result = null;

            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                    {
                        WorkbookPart wbPart = doc.WorkbookPart;

                        Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == WorksheetName).FirstOrDefault();
                        if (sheet is null)
                            throw new InvalidOperationException();
                        
                        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
                        Worksheet worksheet = wsPart.Worksheet;

                        SharedStringTablePart sharedStringTablePart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                        var sheetRows = worksheet.Descendants<Row>();
                        if (sheetRows is null)
                            throw new InvalidOperationException();

                        bool imported = ImportTable<TRootData>(sheetRows, sharedStringTable, tableStartRowIndex: 0, tableDepth: 0,
                            out int tableEndRowIndex,
                            out ICollection<TRootData> importedResult);

                        if (imported && importedResult != null)
                        {
                            result = importedResult.ToList();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public static bool ImportTable<TData>(
            IEnumerable<Row> sheetRows, SharedStringTable sheetSharedStringTable, 
            int tableStartRowIndex, int tableDepth,
            out int tableEndRowIndex,
            out ICollection<TData> importedData)
            where TData : IData, new()
        {
            Type dataType = typeof(TData);

            tableEndRowIndex = tableStartRowIndex;
            importedData = new List<TData>();

            int currentTableRowIndex = tableStartRowIndex;

            ExportTableAttribute tableConfig = (ExportTableAttribute)dataType
                .GetCustomAttributes(typeof(ExportTableAttribute), false)
                .FirstOrDefault();
            if (tableConfig is null)
                return false;

            var tableProperties = dataType
                .GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(ExportColumnAttribute)))
                .ToList();
            var tableChildrenProperties = dataType
                .GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(ExportColumnTableAttribute)))
                .ToList();

            // Table header
            // ...
            currentTableRowIndex++;

            // Table columns and data are one level deeper
            tableDepth++;

            // Table columns
            Dictionary<string, PropertyInfo> tableColumns = new Dictionary<string, PropertyInfo>();
            var columnsRow = sheetRows.ElementAt(currentTableRowIndex);
            var columnsRowCells = columnsRow.Elements<Cell>().Where(c => c.CellValue != null).ToList();
            foreach (var property in tableProperties)
            {
                ExportColumnAttribute propertyConfig = (ExportColumnAttribute)property
                    .GetCustomAttributes(typeof(ExportColumnAttribute), false).First();
                
                string columnTitle = propertyConfig.Title ?? property.Name.ToString();

                var columnRowCell = columnsRowCells.ElementAtOrDefault(tableProperties.IndexOf(property));
                if (columnRowCell is null)
                    throw new Exception($"Invalid table structure [RowIndex: {currentTableRowIndex + 1}]");

                string columnRowCellValue = GetCellValue(columnRowCell, sheetSharedStringTable);
                if (columnTitle != columnRowCellValue)
                    throw new Exception($"Invalid table structure [RowIndex: {currentTableRowIndex + 1}]");
                
                tableColumns.Add(GetColumnLetter(columnRowCell.CellReference), property);
            }
            
            currentTableRowIndex++;

            // Table data
            TData dataObj = default(TData);
            while (currentTableRowIndex < sheetRows.Count())
            {
                var dataRow = sheetRows.ElementAt(currentTableRowIndex);

                if (dataRow.OutlineLevel == tableDepth)
                {
                    dataObj = new TData();

                    var dataRowCells = dataRow.Descendants<Cell>().Where(c => c.CellValue != null).ToList();

                    foreach (var column in tableColumns)
                    {
                        var dataRowCell = dataRowCells.FirstOrDefault(c => c.CellReference == $"{column.Key}{currentTableRowIndex + 1}");
                        if (dataRowCell is null)
                            continue;

                        string dataRowCellValue = GetCellValue(dataRowCell, sheetSharedStringTable);

                        column.Value.SetValue(dataObj, dataRowCellValue, null);
                    }

                    importedData.Add(dataObj);

                    currentTableRowIndex++;
                }
                else if (dataRow.OutlineLevel == tableDepth + 1 && tableConfig.HasChildren && dataObj != null)
                {
                    foreach (var childProperty in tableChildrenProperties)
                    {
                        ExportColumnTableAttribute columnTableConfig = (ExportColumnTableAttribute)childProperty
                            .GetCustomAttributes(typeof(ExportColumnTableAttribute), false)
                            .FirstOrDefault();
                        if (columnTableConfig is null)
                            continue;

                        var childDataType = childProperty.PropertyType.GetGenericArguments()[0];

                        int childTableEndRowIndex = currentTableRowIndex;

                        object[] importTableMethodParameters = new object[]
                        {
                            sheetRows, sheetSharedStringTable,
                            currentTableRowIndex, tableDepth + 1,
                            null, null
                        };
                        object importTableMethodInvokeResult = typeof(ExcelUtilities)
                            .GetMethod(nameof(ImportTable))
                            .MakeGenericMethod(childDataType)
                            .Invoke(null, importTableMethodParameters);

                        bool importTableMethodResult = (bool)importTableMethodInvokeResult;

                        // 'tableEndRowIndex' out parameter
                        childTableEndRowIndex = (int)importTableMethodParameters[4];

                        // 'importedData' out parameter
                        if (importTableMethodResult && importTableMethodParameters[5] != null)
                        {
                            childProperty.SetValue(dataObj, importTableMethodParameters[5], null);
                        }

                        currentTableRowIndex = childTableEndRowIndex;
                    }
                }
                else
                {
                    break;
                }
            }

            tableEndRowIndex = currentTableRowIndex;

            return true;
        }

        #endregion Import

        #region Export

        public static void Export(IEnumerable<IData> data, string filename)
        {
            if (File.Exists(filename))
                File.Delete(filename);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = WorksheetName
                };
                sheets.Append(sheet);

                ExportSheetData(ref worksheetPart, ref sheetData, ref sheet, data);

                workbookpart.Workbook.Save();

                spreadsheetDocument.Close();
            }
        }

        private static void ExportSheetData(ref WorksheetPart worksheetPart, ref SheetData sheetData, ref Sheet sheet, IEnumerable<IData> data)
        {
            ExportSheetData(ref worksheetPart, ref sheetData, ref sheet, data, 0);
        }
        private static void ExportSheetData(ref WorksheetPart worksheetPart, ref SheetData sheetData, ref Sheet sheet, IEnumerable<IData> data,
            int depth, bool isCollapsed = false, bool isParentCollapsed = false)
        {
            Type dataType = data.GetType().GetGenericArguments()[0];

            ExportTableAttribute tableConfig = (ExportTableAttribute)dataType
                .GetCustomAttributes(typeof(ExportTableAttribute), false).FirstOrDefault();
            if (tableConfig is null)
                return;

            uint nextRowIndex = (uint)sheetData.Elements<Row>().Count() + 1;

            var properties = dataType
                .GetProperties()
                .Where(prop => Attribute.IsDefined(prop, typeof(ExportColumnAttribute)))
                .ToList();

            // Table header
            DocumentFormat.OpenXml.Spreadsheet.Cell headerCell;
            headerCell = ExportSheetCell(ref worksheetPart, ref sheetData, ref sheet,
                nextRowIndex, depth, RowType.Header, 0,
                tableConfig.Header,
                true, false);
            MergeCells(worksheetPart, headerCell.CellReference, IncrementColumn(headerCell.CellReference, properties.Count));

            nextRowIndex = (uint)sheetData.Elements<Row>().Count() + 1;

            // Table columns and data are one level deeper
            depth++;

            // Table columns
            foreach (var property in properties)
            {
                ExportColumnAttribute propertyConfig = (ExportColumnAttribute)property
                    .GetCustomAttributes(typeof(ExportColumnAttribute), false).First();

                string columnTitle = propertyConfig.Title ?? property.Name.ToString();
                DocumentFormat.OpenXml.Spreadsheet.Cell cell;
                cell = ExportSheetCell(ref worksheetPart, ref sheetData, ref sheet,
                    nextRowIndex, depth, RowType.Column, properties.IndexOf(property),
                    columnTitle,
                    true, false);
            }

            // Table rows
            foreach (var dataRow in data)
            {
                nextRowIndex = (uint)sheetData.Elements<Row>().Count() + 1;

                foreach (var property in properties)
                {
                    string propertyValue = property.GetValue(dataRow)?.ToString() ?? string.Empty;
                    ExportSheetCell(ref worksheetPart, ref sheetData, ref sheet,
                        nextRowIndex, depth, RowType.Data, properties.IndexOf(property),
                        propertyValue,
                        true, false);
                }

                if (tableConfig.HasChildren)
                {
                    var childrenProperties = dataType
                        .GetProperties()
                        .Where(prop => Attribute.IsDefined(prop, typeof(ExportColumnTableAttribute)))
                        .ToList();

                    foreach (var childProperty in childrenProperties)
                    {
                        ExportColumnTableAttribute columnTableConfig = (ExportColumnTableAttribute)childProperty
                            .GetCustomAttributes(typeof(ExportColumnTableAttribute), false).FirstOrDefault();
                        if (columnTableConfig is null)
                            continue;

                        ExportSheetData(ref worksheetPart, ref sheetData, ref sheet, 
                            (IEnumerable<IData>)childProperty.GetValue(dataRow), depth + 1, columnTableConfig.IsCollapsed, isCollapsed);
                    }
                }
            }
        }

        private static DocumentFormat.OpenXml.Spreadsheet.Cell ExportSheetCell(ref WorksheetPart worksheetPart, ref SheetData sheetData, ref Sheet sheet,
            uint rowIndex, int rowDepth, RowType rowType, int columnIndex,
            string cellValue,
            bool isVisible, bool isCollapsed
            )
        {
            int displayColumnIndex = columnIndex + rowDepth + 1;
            string columnName = GetColumnName(displayColumnIndex);

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() > 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row()
                {
                    RowIndex = rowIndex,
                    DyDescent = 0.25D,
                    Collapsed = isCollapsed,
                    OutlineLevel = (Byte)rowDepth,
                    Hidden = !isVisible
                };
                sheetData.Append(row);
            }

            if (row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                return row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).First();

            // Cell reference
            string cellReferenceStr = columnName + rowIndex;
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if (cell.CellReference.Value.Length == cellReferenceStr.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReferenceStr, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell()
            {
                CellReference = cellReferenceStr,
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };

            row.InsertBefore(newCell, refCell);

            worksheetPart.Worksheet.Save();

            return newCell;
        }

        #endregion Export

        #region Helpers

        private static void MergeCells(WorksheetPart worksheetPart, string fromCellReference, string toCellReference)
        {
            MergeCells mergeCells;

            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();

                if (worksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<CustomSheetView>().First());
                else
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            }

            MergeCell mergeCell = new MergeCell()
            {
                Reference = new StringValue(fromCellReference + ":" + toCellReference)
            };

            mergeCells.Append(mergeCell);
        }

        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable = null)
        {
            string value = default(string);

            try
            {
                string valueStr = cell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (sharedStringTable != null)
                            {
                                value = sharedStringTable.ElementAt(int.Parse(valueStr)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (valueStr)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                        
                        case CellValues.String:
                        case CellValues.InlineString:
                        default:
                            value = (string)valueStr;
                            break;
                    }
                }
                else
                {
                    value = (string)valueStr;
                }
            }
            catch (System.Exception)
            {
                value = default(string);
            }

            return value;
        }

        private static string IncrementRow(string cell)
        {
            string reg = @"^([A-Za-z]+)(\d+)$";
            Match m = Regex.Match(cell, reg);

            if (!m.Success)
            {
                throw new ArgumentException(cell + " is not a valid cell reference");
            }

            int rowNumber = int.Parse(m.Groups[2].Value);
            return m.Groups[1].Value.ToUpper() + (++rowNumber);
        }

        private static string IncrementColumn(string cell, int delta = 1)
        {
            string reg = @"^([A-Za-z]+)(\d+)$";
            Match m = Regex.Match(cell, reg);

            if (!m.Success)
            {
                throw new ArgumentException(cell + " is not a valid cell reference");
            }

            string colLetters = IncrementColumnName(m.Groups[1].Value.ToUpper(), delta);
            
            return colLetters + m.Groups[2].Value;
        }

        private static string IncrementColumnName(string startColumnName, int delta = 1)
        {
            if (string.IsNullOrEmpty(startColumnName) || !Regex.IsMatch(startColumnName, @"^[a-zA-Z]+$"))
            {
                throw new ArgumentException(startColumnName + " is not a valid column name");
            }

            string colLetters = startColumnName.Trim().ToUpper();
            for (int d = 0; d < delta; d++)
            {
                int len = colLetters.Length;
                char lastLetter = colLetters[len - 1];

                if (lastLetter < 'Z')
                {
                    colLetters = colLetters.Substring(0, len - 1) + (++lastLetter);
                }
                else if (Regex.IsMatch(colLetters, "^Z+$"))
                {
                    colLetters = new string('A', len + 1);
                }
                else
                {
                    int base26 = 0;
                    int multiplier = 1;

                    for (int i = len - 1; i >= 0; --i)
                    {
                        base26 += multiplier * (colLetters[i] - 65);
                        multiplier *= 26;
                    }

                    base26++;
                    string temp = "";

                    while (base26 > 0)
                    {
                        temp = (char)(base26 % 26 + 65) + temp;
                        base26 /= 26;
                    }

                    colLetters = temp;
                }
            }

            return colLetters;
        }

        private static string GetColumnName(int Index)
        {
            return IncrementColumnName("A", Index - 1).ToString();
        }

        private static string GetColumnLetter(string cellReference)
        {
            string reg = @"^([A-Za-z]+)(\d+)$";
            Match m = Regex.Match(cellReference, reg);

            if (!m.Success)
            {
                throw new ArgumentException(cellReference + " is not a valid cell reference");
            }
            
            return m.Groups[1].Value.ToUpper();
        }

        #endregion Helpers
    }
}
