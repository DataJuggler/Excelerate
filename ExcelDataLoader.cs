

#region using statements

using DataJuggler.UltimateHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI;
using NPOI.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

#endregion

namespace DataJuggler.Excelerate
{

    #region class ExcelDataLoader
    /// <summary>
    /// This class is used to load Excel Workbooks and Worksheets
    /// </summary>
    public class ExcelDataLoader
    {
        
        #region Methods

           #region GetCellText(ISheet sheet, int row, int col)
            /// <summary>
            /// This method returns the Cell Text (what you see, including formatting)
            /// </summary>
            public static string GetCellText(ISheet sheet, int row, int col)
            {
                // initial value
                string cellText = "";

                try
                {
                    // if the row is in range
                    if (row < sheet.LastRowNum)
                    {  
                        // Get the row
                        IRow currentRow = sheet.GetRow(row);

                        // if the cell is in range
                        if (col < currentRow.LastCellNum)
                        {
                            // Get the cell
                            ICell cell = currentRow.GetCell(col);

                            // if the cell exists
                            if ((NullHelper.Exists(cell)) && (TextHelper.Exists(cell.ToString())))
                            {  
                                // Setg retue
                                cellText = cell.ToString();
                            }
                        }
                    }
                }
                catch (Exception error)
                {   
                    // Return Exception
                    cellText = "Exception: " + error.ToString();
                }

                // return value
                return cellText;
            }           
            #endregion

            #region GetCellValue()
            /// <summary>
            /// This method returns the Cell Value
            /// </summary>
            public static object GetCellValue(ISheet sheet, int row, int col)
            {
                // initial value
                object cellValue = "";

                try
                {
                    // if the row is in range
                    if (row < sheet.LastRowNum)
                    {  
                        // Get the row
                        IRow currentRow = sheet.GetRow(row);

                        // if the cell is in range
                        if (col < currentRow.LastCellNum)
                        {
                            // Get the cell
                            ICell cell = currentRow.GetCell(col);

                            // if the cell exists
                            if (NullHelper.Exists(cell))
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Boolean:
                                        cellValue = cell.BooleanCellValue;
                                        break;

                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell))
                                            cellValue = cell.DateCellValue;
                                        else
                                            cellValue = cell.NumericCellValue;
                                        break;

                                    case CellType.String:
                                        cellValue = cell.StringCellValue;
                                        break;

                                    case CellType.Blank:
                                        cellValue = null;
                                        break;

                                    case CellType.Formula:
                                        cellValue = GetFormulaCellValue(cell);
                                        break;

                                    default:
                                        cellValue = cell.ToString();
                                        break;
                                }
                            }
                        }
                    }
                }
                catch (Exception error)
                {   
                    // Return Exception
                    cellValue = "Exception: " + error.ToString();
                }

                // return value
                return cellValue;
            }           
            #endregion
            
            #region GetColumnIndex(ISheet excelWorksheet, string columnName)
            /// <summary>
            /// This method returns the Column Index
            /// </summary>
            public static int GetColumnIndex(ISheet excelWorksheet, string columnName)
            {
                // initial value
                int columnIndex = -1;

                // If the columnName string exists
                if ((TextHelper.Exists(columnName)) && (NullHelper.Exists(excelWorksheet)))
                {
                    // Iterate through the first row (header row) to find the column name
                    IRow headerRow = excelWorksheet.GetRow(0);

                    for (int i = 0; i < headerRow.LastCellNum; i++)
                    {
                        ICell cell = headerRow.GetCell(i);
                        if (cell != null && cell.CellType == CellType.String && cell.StringCellValue.Trim() == columnName)
                        {
                            columnIndex = i;
                            break;
                        }
                    }
                }
                
                // return value
                return columnIndex;
            }
            #endregion
            
            #region GetFormulaCellValue(ICell cell)
            /// <summary>
            /// method returns the Formula Cell Value
            /// </summary>
            public static object GetFormulaCellValue(ICell cell)
            {
                // initial value
                object cellValue;

                switch (cell.CachedFormulaResultType)
                {
                    case CellType.Boolean:
                        cellValue = cell.BooleanCellValue;

                        // required                        
                        break;
                    
                    case CellType.Numeric:
                    
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            // set the return value
                            cellValue = cell.DateCellValue;
                        }
                        else
                        {
                            // set the return value
                            cellValue = cell.NumericCellValue;
                        }

                        // required
                        break;
                    
                    case CellType.String:

                        // set the return value
                        cellValue = cell.StringCellValue;

                    break;
                    
                    case CellType.Blank:

                        // null
                        cellValue = null;

                        // required
                        break;
                    
                    default:

                        // Return as a string
                        cellValue = cell.ToString();

                        // required
                        break;
                }
                return cellValue;
            }
            #endregion
            
            #region GetSheetNames(string path)
            /// <summary>
            /// method returns the Sheet Names for the workbook given
            /// </summary>
            public static List<string> GetSheetNames(string path)
            {
                // initial value
                List<string> sheetNames = new List<string>();

                // verify the path exists
                if ((TextHelper.Exists(path)) && (FileHelper.Exists(path)))
                {
                    // Open the Excel file
                    using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        // Create the workbook
                        IWorkbook workbook = WorkbookFactory.Create(fileStream);

                        // Get the sheetNames from the override
                        sheetNames = GetSheetNames(workbook);                    
                    }
                }

                // return value
                return sheetNames;
            }
            #endregion

            #region GetSheetNames(IWorkbook excelWorkbook)
            /// <summary>
            /// method returns the Sheet Names for the workbook given
            /// </summary>
            public static List<string> GetSheetNames(IWorkbook workbook)
            {
                // initial value
                List<string> sheetNames = new List<string>();

                // If the workbook object exists
                if (NullHelper.Exists(workbook))
                {
                    // Iterate through the sheets
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        // Add the sheet name to the list
                        sheetNames.Add(workbook.GetSheetName(i));
                    }
                }

                // return value
                return sheetNames;
            }
            #endregion
            
            #region LoadAllData(string path)
            /// <summary>
            /// method returns the All Data
            /// </summary>
            public static Workbook LoadAllData(string path)
            {
                // initial value
                Workbook workbook = null;

                // load the workbook
                XSSFWorkbook excelWorkbook = LoadExcelWorkbook(path);
                    
                // If the excelWorkbook object exists
                if (NullHelper.Exists(excelWorkbook))
                {
                    // Create a new instance of a 'Workbook' object.
                    workbook = new Workbook();

                    // Get the sheetNames
                    List<string> sheetNames = GetSheetNames(excelWorkbook);

                    // If the sheetNames collection exists and has one or more items
                    if (ListHelper.HasOneOrMoreItems(sheetNames))
                    {
                        // Iterate the collection of string objects
                        foreach (string sheetName in sheetNames)
                        {
                            // Create a new instance of a 'LoadWorksheetInfo' object.
                            WorksheetInfo loadWorksheetInfo = new WorksheetInfo();

                            // Set the sheetName
                            loadWorksheetInfo.SheetName = sheetName;

                            // Set the LoadColumnOption
                            loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;

                            // Load this worksheet
                            Worksheet worksheet = LoadWorksheet(excelWorkbook, loadWorksheetInfo);

                            // if the workbook exists
                            if (NullHelper.Exists(worksheet))
                            {
                                // Add this item
                                workbook.Worksheets.Add(worksheet);
                            }
                        }
                    }
                }

                // return value
                return workbook;
            }
            #endregion
            
            #region LoadExcelWorkbook(string path)
            /// <summary>
            /// This method returns the Excel Workbook
            /// </summary>
            public static XSSFWorkbook LoadExcelWorkbook(string path)
            {
                // initial value
                XSSFWorkbook excelWorkbook = null;

                // If the path string exists
                if (TextHelper.Exists(path))
                {
                    // load the Workbook
                    using (var stream = new FileStream(path, FileMode.Open))
                    {
                        stream.Position = 0;
                        excelWorkbook = new XSSFWorkbook(stream);
                    }
                }
                
                // return value
                return excelWorkbook;
            }
            #endregion

            #region LoadWorkbook(string path, List<LoadWorksheetInfo> sheetsToLoad)
            /// <summary>
            /// This method loads a Workbook for the path given
            /// </summary>
            /// <param name="path"></param>
            /// <returns></returns>
            public static Workbook LoadWorkbook(string path, List<WorksheetInfo> sheetsToLoad)
            {
                // initial value
                Workbook workbook = new Workbook();

                // If the path string exists and there is one or more sheetsToLoad
                if (TextHelper.Exists(path) && (ListHelper.HasOneOrMoreItems(sheetsToLoad)))
                {
                    // Load the workbook
                    XSSFWorkbook excelWorkbook = LoadExcelWorkbook(path);

                    // If the excelWorkbook object exists
                    if (NullHelper.Exists(excelWorkbook))
                    {
                        // Iterate the collection of LoadWorksheetInfo objects
                        foreach (WorksheetInfo loadWorksheetInfo in sheetsToLoad)
                        {
                            // Create a workSheet object
                            Worksheet worksheet = LoadWorksheet(excelWorkbook, loadWorksheetInfo);

                           // If the worksheet object exists
                           if (NullHelper.Exists(worksheet))
                           {
                                // Add this worksheet
                                workbook.Worksheets.Add(worksheet);
                           }
                        }
                    }
                }

                // return value
                return workbook;
            }
            #endregion

            #region LoadWorkbook(string path, LoadWorksheetInfo loadWorksheetInfo)
            /// <summary>
            /// This method loads a Workbook and only one shheet for the path given
            /// </summary>
            /// <param name="path"></param>
            /// <returns></returns>
            public static Workbook LoadWorkbook(string path, WorksheetInfo loadWorksheetInfo)
            {
                // initial value
                Workbook workbook = new Workbook();

                // If the path string exists and the sheetToLoad exists
                if (TextHelper.Exists(path) && (NullHelper.Exists(loadWorksheetInfo)))
                {
                    // Create a new instance of a 'FileInfo' object.
                    FileInfo fileInfo = new FileInfo(path);

                    // Load the workbook
                    XSSFWorkbook excelWorkbook = LoadExcelWorkbook(path);

                    // If the excelWorkbook object exists
                    if (NullHelper.Exists(excelWorkbook))
                    {  
                        // Create a workSheet object
                        Worksheet workSheet = LoadWorksheet(excelWorkbook, loadWorksheetInfo);

                        // If the workSheet object exists
                        if (NullHelper.Exists(workSheet))
                        {
                            // Add this worksheet
                            workbook.Worksheets.Add(workSheet);
                        }
                    }
                }

                // return value
                return workbook;
            }
            #endregion

            #region LoadWorksheet(XSSFWorkbook excelWorkbook, LoadWorksheetInfo loadWorksheetInfo)
            /// <summary>
            /// This method returns the Worksheet
            /// </summary>
            public static Worksheet LoadWorksheet(XSSFWorkbook excelWorkbook, WorksheetInfo loadWorksheetInfo)
            {
                // initial value
                Worksheet worksheet = null;

                // locals
                int rowNumber = 0;
                int colNumber = 1;                
                Column column = null;
                int columnIndex = -1;
                bool skipColumn = false;
                int tempIndex = -1;
                
                // verify both objects exist
                if (NullHelper.Exists(excelWorkbook, loadWorksheetInfo))
                {
                    try
                    {
                        //reading Project Information
                        ISheet excelWorksheet = excelWorkbook.GetSheet(loadWorksheetInfo.SheetName);

                        // If the excelWorksheet object exists
                        if (NullHelper.Exists(excelWorksheet))
                        {
                            // set the rowCount and colCount
                            int rowCount = excelWorksheet.LastRowNum;
                            IRow headerRow = excelWorksheet.GetRow(0);
                            int colCount = headerRow.LastCellNum;

                            // if the MawRowsToLoad is set and the MaxRowsToLoad is less than RowCount
                            if ((loadWorksheetInfo.MaxRowsToLoad > 0) && (loadWorksheetInfo.MaxRowsToLoad < rowCount))
                            {
                                // Only load this many rows
                                rowCount = loadWorksheetInfo.MaxRowsToLoad;
                            }

                            // if only a specified number of columns should be loaded
                            if (loadWorksheetInfo.LoadColumnOptions == LoadColumnOptionsEnum.LoadFirstXColumns)
                            {
                                // if the ColumnsToLoad is set and the ColumsnToLoad is less than the number of columns
                                if ((loadWorksheetInfo.ColumnsToLoad > 0) && (loadWorksheetInfo.ColumnsToLoad < colCount))
                                {
                                    // Set the colCount
                                    colCount = loadWorksheetInfo.ColumnsToLoad;
                                }
                            }

                            // verify there are rows and columns
                            if ((rowCount > 0) && (colCount > 0))
                            {
                                // Create a new instance of a 'Worksheet' object.
                                worksheet = new Worksheet();

                                // Store the loadWorksheetInfo, so saving is easier
                                worksheet.WorksheetInfo = loadWorksheetInfo;

                                // Set the sheetName
                                worksheet.Name = loadWorksheetInfo.SheetName;

                                do
                                {   
                                    // Increment the value for rowNumber
                                    rowNumber++;

                                    // Create a new instance of a 'Row' object.
                                    Row row = new Row();

                                    // Set the rowNumber
                                    row.Number = rowNumber;

                                    // Set IsHeaderRow to true, since the header row has to be in the top row
                                    row.IsHeaderRow = (rowNumber == 1);

                                    // Set the rowId
                                    row.Id = Guid.NewGuid();

                                    // now load the columns for this row

                                    // if load specified columns is true and there are one or more columns specified
                                    if ((loadWorksheetInfo.LoadColumnOptions == LoadColumnOptionsEnum.LoadSpecifiedColumns) && (ListHelper.HasOneOrMoreItems(loadWorksheetInfo.SpecifiedColumnNames)))
                                    {
                                        // load the specified columns
                                        foreach (SpecifiedColumnName columnName in loadWorksheetInfo.SpecifiedColumnNames)
                                        {
                                            // if the Index is greater than zero
                                            if (columnName.HasIndex)
                                            {
                                                // Set the index that was already looked up
                                                columnIndex = columnName.Index;
                                            }
                                            else if (!columnName.NotFound)
                                            {
                                                // find the columnIndex
                                                columnIndex = GetColumnIndex(excelWorksheet, columnName.Name);

                                                // if the columnIndex was not found
                                                if (columnIndex < 1)
                                                {
                                                    // not found
                                                    columnName.NotFound = true;
                                                }
                                            }

                                            // if the index was found
                                            if (columnIndex > 0)
                                            {
                                                // increment the value
                                                colNumber++;

                                                // Create a new instance of a 'Column' object.
                                                column = new Column();

                                                // Set the index
                                                column.Index = columnIndex;

                                                // Set the value
                                                column.RowNumber = rowNumber;
                                                column.ColumnNumber = columnIndex; 

                                                // Get the ColumnValue
                                                column.ColumnValue = GetCellValue(excelWorksheet, rowNumber, colNumber);

                                                // Get the CellText
                                                column.ColumnText = GetCellText(excelWorksheet, rowNumber, colNumber);

                                                // Add this column
                                                row.Columns.Add(column);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // iterate the columns up to colCount
                                        for (int x = 1; x <= colCount; x++)
                                        {
                                            // reset
                                            skipColumn = false;
                                            
                                            if (ListHelper.HasOneOrMoreItems(loadWorksheetInfo.ExcludedColumnIndexes))
                                            {
                                                // attempt to find this index
                                                tempIndex = loadWorksheetInfo.ExcludedColumnIndexes.IndexOf(x);

                                                // if this column index was found
                                                skipColumn = (tempIndex >= 0);
                                            }

                                            // if the value for skipColumn is false
                                            if (!skipColumn)
                                            {
                                                // Create a new instance of a 'Column' object.
                                                column = new Column();

                                                // Set the value
                                                column.RowNumber = rowNumber;
                                                column.ColumnNumber = x;

                                                // Get the ColumnValue
                                                column.ColumnValue = GetCellValue(excelWorksheet, rowNumber, x);

                                                // Get the CellText
                                                column.ColumnText = GetCellText(excelWorksheet, rowNumber, x);

                                                // Set the index
                                                column.Index = x;

                                                // Add this column
                                                row.Columns.Add(column);
                                            }
                                        }
                                    }

                                    // Add this row
                                    worksheet.Rows.Add(row);

                                } while (rowNumber < rowCount);
                            }
                        }
                    }
                    catch (Exception error)
                    {
                        // Use this to attach logging or other centralized error handling
                        DebugHelper.WriteDebugError("LoadWorksheet", "ExcelDataLoader", error);
                    }
                }
                
                // return value
                return worksheet;
            }
            #endregion

            #region LoadWorksheet(string path, LoadWorksheetInfo loadWorksheetInfo)
            /// <summary>
            /// This method returns a single Worksheet
            /// </summary>
            public static Worksheet LoadWorksheet(string path, WorksheetInfo loadWorksheetInfo)
            {
                // initial value
                Worksheet worksheet = null;

                // load the workbook
                Workbook workbook = LoadWorkbook(path, loadWorksheetInfo);

                // if the workbook
                if ((NullHelper.Exists(workbook)) && (ListHelper.HasOneOrMoreItems(workbook.Worksheets)))
                {
                    // set the return value
                    worksheet = workbook.Worksheets[0];
                }
                
                // return value
                return worksheet;
            }
            #endregion

            #region LoadWorksheet(string path, LoadWorksheetInfo loadWorksheetInfo)
            /// <summary>
            /// This method returns a single Worksheet. For this override worksheetInfo.Path must be set.
            /// </summary>
            public static Worksheet LoadWorksheet(WorksheetInfo worksheetInfo)
            {
                // initial value
                Worksheet worksheet = null;

                // If the worksheetInfo object exists
                if (NullHelper.Exists(worksheetInfo))
                {
                    // call the override
                    worksheet = LoadWorksheet(worksheetInfo.Path, worksheetInfo);
                }
                
                // return value
                return worksheet;
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
