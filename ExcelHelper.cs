

#region using statements

using DataJuggler.Net7;
using DataJuggler.Net7.Delegates;
using DataJuggler.UltimateHelper;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;


#endregion

namespace DataJuggler.Excelerate
{

    #region class ExcelHelper
    /// <summary>
    /// This class is used to help with common Excel tasks, with the main one being Saving Data to Excel.
    /// </summary>
    public class ExcelHelper
    {

        #region Methods

            #region CreateWorkbook(FileInfo worksheetInfo, List<LoadWorksheetInfo> worksheets, ProgressStatusCallback callback = null, string fontName = "Verdana", double fontSize = 11)
            /// <summary>
            /// Creates an Excel Workbook. When called by SQLSnapshot, the datarows and fields can be loaded and written out here.
            /// </summary>
            public static void CreateWorkbook(FileInfo worksheetInfo, List<LoadWorksheetInfo> worksheets, ProgressStatusCallback callback = null, string fontName = "Verdana", double fontSize = 11)
            {
                // Create a new instance of an 'ExcelPackage' object.
                ExcelPackage excel = new ExcelPackage();

                // locals
                int index = 0;
                int rowNumber = 1;
                int startRowNumber = 1;
                int progress = 0;
                int subProgress = 0;

                // If the worksheets collection exists and has one or more items
                if (ListHelper.HasOneOrMoreItems(worksheets))
                {
                    // If the progress object exists
                    if (NullHelper.Exists(callback))
                    {
                        // Notify the caller
                        callback(worksheets.Count * 2, worksheets.Count + progress, "Begin exporting data rows", 0, 0, "Starting");
                    }

                    // Iterate the collection of LoadWorksheetInfo objects
                    foreach (LoadWorksheetInfo sheet in worksheets)
                    {
                        // reset
                        index = 0;
                        rowNumber = 1;

                        // If the callback object exists
                        if (NullHelper.Exists(callback, sheet.Rows))
                        {
                            // Notify the caller
                            callback(worksheets.Count * 2, worksheets.Count + progress, "Exporting data rows", sheet.Rows.Count, 0, "Exporting sheet " + sheet.SheetName);
                        }

                        // Create the sheet
                        ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(sheet.SheetName);

                        // Beging writing header row

                        // if the Fields collection exists
                        if (sheet.HasFields)
                        {
                            // Write out the HeaderRow

                            // iterate the fields
                            foreach (DataField field in sheet.Fields)
                            {
                                // increment the value for index
                                index++;

                                // Set the fieldName
                                worksheet.Cells[rowNumber, index].Value = field.FieldName;    
                            }

                            // Set the header to bold
                            worksheet.Cells[rowNumber, 1, rowNumber, index].Style.Font.Name = fontName;
                            worksheet.Cells[rowNumber, 1, rowNumber, index].Style.Font.Size = (float) fontSize;
                            worksheet.Cells[rowNumber, 1, rowNumber, index].Style.Font.Bold = true;
                            worksheet.Cells[rowNumber, 1, rowNumber, index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[rowNumber, 1, rowNumber, index].AutoFitColumns();

                            // Increment the value for rowNumber
                            rowNumber++;

                            // needed when formatting at the end of this method
                            startRowNumber = rowNumber;
                        }

                        // Beging writing data rows

                        // write out the rows collection
                        if (sheet.HasRows)
                        {
                            // If the callback object exists
                            if (NullHelper.Exists(callback))
                            {
                                // Notify the caller
                                callback(worksheets.Count * 2, worksheets.Count + progress, "Exporting data rows", sheet.Rows.Count, 0, "Exporting sheet " + sheet.SheetName);
                            }

                            // iterate the rows
                            foreach (DataRow row in sheet.Rows)
                            {                                
                                // reset
                                index = 0;
                                subProgress = 0;
                                
                                // if there are one or more fields
                                if (ListHelper.HasOneOrMoreItems(row.Fields))
                                {
                                    // iterate row.Fields
                                    foreach (DataField field in row.Fields)
                                    {
                                        // increment the value for index
                                        index++;

                                        // Set the fieldName
                                        worksheet.Cells[rowNumber, index].Value = field.FieldValue;

                                        // if the first row
                                        if (rowNumber == 2)
                                        {
                                            // if this is a date
                                            if (field.DataType == DataManager.DataTypeEnum.DateTime)
                                            {
                                                // Format the column as a date (testing this now)
                                                worksheet.Column(index).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                                            }
                                        }
                                    }
                                }

                                // Increment the value for rowNumber
                                rowNumber++;

                                // Increment the value for subProgress
                                subProgress++;

                                // if every 100th row
                                if (subProgress % 100 == 0)
                                {
                                    // If the callback object exists
                                    if (NullHelper.Exists(callback))
                                    {
                                        // Notify the caller
                                        callback(worksheets.Count * 2, worksheets.Count + progress, "Exporting data rows", sheet.Rows.Count, subProgress, "Exporting sheet " + sheet.SheetName);
                                    }
                                }
                            }

                            // Format the data rows
                            worksheet.Cells[startRowNumber, 1, rowNumber, index].Style.Font.Name = fontName;
                            worksheet.Cells[startRowNumber, 1, rowNumber, index].Style.Font.Size = (float) fontSize;                
                            worksheet.Cells[startRowNumber, 1, rowNumber, index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;                            
                        }

                        // Increment the value for progress
                        progress++;

                        // If the callback object exists
                        if (NullHelper.Exists(callback))
                        {
                            // Notify the caller
                            callback(worksheets.Count * 2, worksheets.Count + progress, "Exporting data rows", sheet.Rows.Count, subProgress, "Exporting sheet " + sheet.SheetName + " complete.");
                        }
                    }
                }

                // Save the file
                excel.SaveAs(worksheetInfo);
            }
            #endregion
            
            #region GetColumnLetter(int column)
            /// <summary>
            /// returns the Column Letter for the column index (1 = A, 2 = B, 27 = AA, 78 = "ZZZ" I think)
            /// </summary>
            public static string GetColumnLetter(int column)
            {
                // initial value
                string columnLetter = ExcelCellAddress.GetColumnLetter(column);
                
                // return value
                return columnLetter;
            }
            #endregion
            
            #region SaveBatch(string path, Batch batch, onlyColumnsWithChanges = false)
            /// <summary>
            /// returns the Batch
            /// </summary>
            public static bool SaveBatch(string path, Batch batch, bool onlyColumnsWithChanges = false)
            {
                // initial value
                bool saved = false;

                try
                {   
                    // if the batch exists and the batch has BatchItems (represents a Worksheet) and the path to the Excel file exists on disk
                    if ((NullHelper.Exists(batch)) && (batch.HasBatchItems) && (FileHelper.Exists(path)))
                    {
                        // load the package
                        ExcelPackage package = ExcelDataLoader.LoadExcelPackage(path);
                        
                        // iterate the batchItems
                        foreach (BatchItem batchItem in batch.BatchItems)
                        {
                            // if the batchItem has Updates and a WorksheetInfo
                            if ((batchItem.HasUpdates) && (batchItem.HasWorksheetInfo))
                            {
                                // Get the sheet
                                ExcelWorksheet excelworksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == batchItem.WorksheetInfo.SheetName);

                                // If the excelworksheet object exists
                                if (NullHelper.Exists(excelworksheet))
                                {
                                    // iterate the rows to update
                                    foreach (Row row in batchItem.Updates)
                                    {  
                                        // If the value for the property rowNumber.HasColumns is true
                                        if (row.HasColumns)
                                        {
                                            // iterate the rows
                                            foreach (Column column in row.Columns)
                                            {
                                                if ((onlyColumnsWithChanges) && (column.HasChanges))
                                                {
                                                    // Set the value
                                                    excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                                }
                                                else
                                                {
                                                    // Set the value
                                                    excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Save the package
                        package.Save();

                        // all is good
                        saved = true;
                    }
                }
                catch (Exception error)
                {
                    // for debugging only for now
                    DebugHelper.WriteDebugError("SaveBatch", "ExcelHelper.cs", error);
                }
                
                // return value
                return saved;
            }
            #endregion
            
            #region SaveBatchItem(string path, BatchItem batchItem, bool onlyColumnsWithChanges = false)
            /// <summary>
            /// Save and then returns the batchItem
            /// </summary>
            public static bool SaveBatchItem(string path, BatchItem batchItem, bool onlyColumnsWithChanges = false)
            {
                // initial value
                bool saved = false;

                try
                {   
                    // if the batchItem exists and the batchItem has Updates (rows to update) and the path to the Excel file exists on disk
                    if ((NullHelper.Exists(batchItem)) && (batchItem.HasUpdates) && (batchItem.HasWorksheetInfo) && (FileHelper.Exists(path)))
                    {
                        // load the package
                        ExcelPackage package = ExcelDataLoader.LoadExcelPackage(path);

                        // Get the sheet
                        ExcelWorksheet excelworksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == batchItem.WorksheetInfo.SheetName);

                        // If the excelworksheet object exists
                        if (NullHelper.Exists(excelworksheet))
                        {
                            // iterate the rows to update
                            foreach (Row row in batchItem.Updates)
                            {
                                // If the value for the property rowNumber.HasColumns is true
                                if (row.HasColumns)
                                {
                                    // iterate the rows
                                    foreach (Column column in row.Columns)
                                    {
                                        if ((onlyColumnsWithChanges) && (column.HasChanges))
                                        {
                                            // Set the value
                                            excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                        }
                                        else
                                        {
                                            // Set the value
                                            excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                        }
                                    }
                                }
                            }
                        }

                        // Save the package
                        package.Save();

                        // all is good
                        saved = true;
                    }
                }
                catch (Exception error)
                {
                    // for debugging only for now
                    DebugHelper.WriteDebugError("SaveBatchItem", "ExcelHelper.cs", error);
                }
                
                // return value
                return saved;
            }
            #endregion
            
            #region SaveRow(string path, Row row, Worksheet worksheet, bool onlyColumnsWithChanges = false)
            /// <summary>
            /// returns the Row
            /// </summary>
            public static bool SaveRow(string path, Row row, Worksheet worksheet, bool onlyColumnsWithChanges = false)
            {
                // initial value
                bool saved = false;

                try
                {   
                    // if the worksheet exists and the worksheet.WorksheetInfo exists and the path to the worksheet exists
                    if ((NullHelper.Exists(worksheet, row)) && (worksheet.HasWorksheetInfo) && (FileHelper.Exists(path)) && (row.HasColumns))
                    {
                        // load the package
                        ExcelPackage package = ExcelDataLoader.LoadExcelPackage(path);

                        // Get the sheet
                        ExcelWorksheet excelworksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == worksheet.Name);

                        // If the excelworksheet object exists
                        if (NullHelper.Exists(excelworksheet))
                        {
                            // iterate the rows
                            foreach (Column column in row.Columns)
                            {
                                if ((onlyColumnsWithChanges) && (column.HasChanges))
                                {
                                    // Set the value
                                    excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                }
                                else
                                {
                                    // Set the value
                                    excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
                                }
                            }
                        }

                        // Save the package
                        package.Save();

                        // all is good
                        saved = true;
                    }
                }
                catch (Exception error)
                {
                    // for debugging only for now
                    DebugHelper.WriteDebugError("SaveRow", "ExcelHelper.cs", error);
                }
                
                // return value
                return saved;
            }
            #endregion

        #endregion

    }
    #endregion

}
