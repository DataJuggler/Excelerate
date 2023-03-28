

#region using statements

using DataJuggler.Net7;
using DataJuggler.UltimateHelper;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
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

            #region CreateWorkbook(FileInfo worksheetInfo, List<LoadWorksheetInfo> worksheets)
            /// <summary>
            /// Create Workbook
            /// </summary>
            public static void CreateWorkbook(FileInfo worksheetInfo, List<LoadWorksheetInfo> worksheets)
            {
                // Create a new instance of an 'ExcelPackage' object.
                ExcelPackage excel = new ExcelPackage();

                // local
                int index = 0;

                // If the worksheets collection exists and has one or more items
                if (ListHelper.HasOneOrMoreItems(worksheets))
                {
                    // Iterate the collection of LoadWorksheetInfo objects
                    foreach (LoadWorksheetInfo sheet in worksheets)
                    {
                        // reset
                        index = 0;

                        // name of the sheet
                        ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(sheet.SheetName);

                        // if the Fields collection exists
                        if (sheet.HasFields)
                        {
                            // order the fields by FieldOrdinal
                            sheet.Fields = sheet.Fields.OrderBy(x => x.FieldOrdinal).ToList();

                            // iterate the fields
                            foreach (DataField field in sheet.Fields)
                            {
                                // increment the value for index
                                index++;

                                // Set the fieldName
                                worksheet.Cells[1, index].Value = field.FieldName;    
                            }

                            // Set the header to bold
                            worksheet.Cells[1, 1, 1, index].Style.Font.Name = "Verdana";
                            worksheet.Cells[1, 1, 1, index].Style.Font.Size = 12;
                            worksheet.Cells[1, 1, 1, index].Style.Font.Bold = true;
                            worksheet.Cells[1, 1, 1, index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[1, 1, 1, index].AutoFitColumns();
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
            
            #region SaveBatch(string path, Batch batch)
            /// <summary>
            /// returns the Batch
            /// </summary>
            public static bool SaveBatch(string path, Batch batch)
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
                                        // If the value for the property row.HasColumns is true
                                        if (row.HasColumns)
                                        {
                                            // iterate the rows
                                            foreach (Column column in row.Columns)
                                            {
                                                // Set the value
                                                excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
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
            
            #region SaveBatchItem(string path, BatchItem batchItem)
            /// <summary>
            /// Save and then returns the batchItem
            /// </summary>
            public static bool SaveBatchItem(string path, BatchItem batchItem)
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
                                // If the value for the property row.HasColumns is true
                                if (row.HasColumns)
                                {
                                    // iterate the rows
                                    foreach (Column column in row.Columns)
                                    {
                                        // Set the value
                                        excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
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
            
            #region SaveRow(string path, Row row, Worksheet worksheet)
            /// <summary>
            /// returns the Row
            /// </summary>
            public static bool SaveRow(string path, Row row, Worksheet worksheet)
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
                                // Set the value
                                excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
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
