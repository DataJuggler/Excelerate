

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using DataJuggler.UltimateHelper;

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

            #region GetCellValue()
            /// <summary>
            /// This method returns the Cell Value
            /// </summary>
            public static object GetCellValue(ExcelWorksheet sheet, int row, int col)
            {
                // initial value
                object cellValue = "";

                // If the sheet object exists
                if (NullHelper.Exists(sheet))                
                {
                    // if the value of this cell exists
                    if (NullHelper.Exists(sheet.Cells[row, col].Value))
                    {
                        // Setg retue
                        cellValue = sheet.Cells[row, col].Value;
                    }
                }

                // return value
                return cellValue;
            }           
            #endregion
            
            #region GetColumnIndex(ExcelWorksheet excelWorksheet, string columnName)
            /// <summary>
            /// This method returns the Column Index
            /// </summary>
            public static int GetColumnIndex(ExcelWorksheet excelWorksheet, string columnName)
            {
                // initial value
                int columnIndex = -1;

                // If the columnName string exists
                if ((TextHelper.Exists(columnName)) && (NullHelper.Exists(excelWorksheet)))
                {
                    // Set the return value
                    columnIndex = excelWorksheet.Cells["1:1"].First(c => c.Value.ToString() == columnName).Start.Column;
                }
                
                // return value
                return columnIndex;
            }
            #endregion
            
            #region LoadExcelWorkbook(string path)
            /// <summary>
            /// This method returns the Excel Workbook
            /// </summary>
            public static ExcelWorkbook LoadExcelWorkbook(string path)
            {
                // initial value
                ExcelWorkbook excelWorkbook = null;

                // If the path string exists and there is one or more sheetsToLoad
                if (TextHelper.Exists(path))
                {
                    // Create a new instance of a 'FileInfo' object.
                    FileInfo fileInfo = new FileInfo(path);

                    // Create the package
                    var package = new ExcelPackage(fileInfo);

                    // get the workbook
                    excelWorkbook = package.Workbook;
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
            public static Workbook LoadWorkbook(string path, List<LoadWorksheetInfo> sheetsToLoad)
            {
                // initial value
                Workbook workbook = new Workbook();

                // If the path string exists and there is one or more sheetsToLoad
                if (TextHelper.Exists(path) && (ListHelper.HasOneOrMoreItems(sheetsToLoad)))
                {
                    // Create a new instance of a 'FileInfo' object.
                    FileInfo fileInfo = new FileInfo(path);

                    // Create the package
                    var package = new ExcelPackage(fileInfo);

                    // get the workbook
                    ExcelWorkbook excelWorkbook = package.Workbook;
                
                    // If the excelWorkbook object exists
                    if (NullHelper.Exists(excelWorkbook))
                    {
                        // Iterate the collection of LoadWorksheetInfo objects
                        foreach (LoadWorksheetInfo loadWorksheetInfo in sheetsToLoad)
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
            public static Workbook LoadWorkbook(string path, LoadWorksheetInfo loadWorksheetInfo)
            {
                // initial value
                Workbook workbook = new Workbook();

                // If the path string exists and the sheetToLoad exists
                if (TextHelper.Exists(path) && (NullHelper.Exists(loadWorksheetInfo)))
                {
                    // Create a new instance of a 'FileInfo' object.
                    FileInfo fileInfo = new FileInfo(path);

                    // Create the package
                    var package = new ExcelPackage(fileInfo);

                    // get the workbook
                    ExcelWorkbook excelWorkbook = package.Workbook;
                
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

            #region LoadWorksheet(ExcelWorkbook excelWorkbook, LoadWorksheetInfo loadWorksheetInfo)
            /// <summary>
            /// This method returns the Worksheet
            /// </summary>
            public static Worksheet LoadWorksheet(ExcelWorkbook excelWorkbook, LoadWorksheetInfo loadWorksheetInfo)
            {
                // initial value
                Worksheet worksheet = null;

                // locals
                int rowNumber = 0;
                int colNumber = 1;                
                Column column = null;
                int columnIndex = -1;

                // verify both objects exist
                if (NullHelper.Exists(excelWorkbook, loadWorksheetInfo))
                {
                    try
                    {
                        //reading Project Information
                        ExcelWorksheet excelWorksheet = excelWorkbook.Worksheets[loadWorksheetInfo.SheetName];

                        // If the excelWorksheet object exists
                        if (NullHelper.Exists(excelWorksheet))
                        {
                            // set the rowCount and colCount
                            int rowCount = excelWorksheet.Dimension.Rows;
                            int colCount = excelWorksheet.Dimension.Columns;

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

                                // Load th
                                worksheet.Name = loadWorksheetInfo.SheetName;

                                do
                                {   
                                    // Increment the value for rowNumber
                                    rowNumber++;

                                    // Create a new instance of a 'Row' object.
                                    Row row = new Row();

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

                                                // Set the value
                                                column.RowNumber = rowNumber;
                                                column.ColumnNumber = columnIndex; 

                                                // Get the ColumnValue
                                                column.ColumnValue = GetCellValue(excelWorksheet, rowNumber, colNumber);

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
                                            // Create a new instance of a 'Column' object.
                                            column = new Column();

                                            // Set the value
                                            column.RowNumber = rowNumber;
                                            column.ColumnNumber = x;

                                            // Get the ColumnValue
                                            column.ColumnValue = GetCellValue(excelWorksheet, rowNumber, x);

                                            // Add this column
                                            row.Columns.Add(column);
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
            
        #endregion
        
    }
    #endregion

}
