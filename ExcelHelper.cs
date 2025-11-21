

#region using statements

using DataJuggler.Excelerate.Delegates;
using DataJuggler.Excelerate.Interfaces;
using DataJuggler.NET.Data;
using DataJuggler.NET.Data.Delegates;
using DataJuggler.UltimateHelper;
using DataJuggler.UltimateHelper.Objects;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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

            #region CreateWorkbook(FileInfo workbookFileInfo, List<LoadWorksheetInfo> worksheets, ProgressStatusCallback callback = null, string fontName = "Verdana", double fontSize = 11)
            /// <summary>
            /// Creates an Excel Workbook. When called by SQLSnapshot, the datarows and fields can be loaded and written out here.
            /// </summary>
            public static void CreateWorkbook(FileInfo workbookFileInfo, List<WorksheetInfo> worksheets, ProgressStatusCallback callback = null, string fontName = "Verdana", double fontSize = 11)
            {
                // locals
                int index = 0;
                int rowNumber = 0;
                int startRowNumber = 1;
                int progress = 0;
                int subProgress = 0;

                // Create a new workbook
                XSSFWorkbook workbook = new XSSFWorkbook();

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
                    foreach (WorksheetInfo sheet in worksheets)
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

                        // Styles needed for formatting

                        // Create a cell style for date formatting
                        ICellStyle dateCellStyle = workbook.CreateCellStyle();
                        short dateFormat = workbook.CreateDataFormat().GetFormat(DateTimeFormatInfo.CurrentInfo.ShortDatePattern);
                        dateCellStyle.Alignment = HorizontalAlignment.Left;
                        dateCellStyle.DataFormat = dateFormat;

                        // Define the font and cell style
                        IFont boldFont = workbook.CreateFont();
                        boldFont.FontName = fontName;
                        boldFont.FontHeightInPoints = fontSize;
                        boldFont.IsBold = true;
                        ICellStyle boldCellStyle = workbook.CreateCellStyle();
                        // Center the Header Row
                        boldCellStyle.Alignment = HorizontalAlignment.Center;
                        boldCellStyle.SetFont(boldFont);

                        // Create a cell style for font and alignment
                        ICellStyle cellStyle = workbook.CreateCellStyle();

                        // Set font settings
                        IFont font = workbook.CreateFont();
                        font.FontName = fontName;
                        font.FontHeightInPoints = fontSize;
                        font.IsBold = false;
                        cellStyle.Alignment = HorizontalAlignment.Left;
                        cellStyle.SetFont(font);

                        // Create a new sheet in the workbook
                        ISheet worksheet = workbook.CreateSheet(sheet.SheetName);

                        // Beging writing header row

                        // Create a row in the sheet
                        IRow headerRow = worksheet.CreateRow(0);

                        // if the Fields collection exists
                        if (sheet.HasFields)
                        {
                            // Write out the HeaderRow

                            // iterate the fields
                            foreach (DataField dataField in sheet.Fields)
                            {
                                // increment the value for index
                                index++;

                                // Set the fieldName
                                ICell cell = headerRow.CreateCell(index);

                                // Set the fieldName
                                cell.SetCellValue(dataField.FieldName);

                                // Set to bold
                                cell.CellStyle = boldCellStyle;
                            }

                            // Auto-fit the columns
                            for (int i = 0; i <= index; i++)
                            {
                                // Auto fit the columns
                                worksheet.AutoSizeColumn(i);
                            }

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

                                // Create a row in the sheet
                                IRow currentRow = worksheet.CreateRow(rowNumber);
                                
                                // if there are one or more fields
                                if (ListHelper.HasOneOrMoreItems(row.Fields))
                                {
                                    // iterate row.Fields
                                    foreach (DataField field in row.Fields)
                                    {
                                        // increment the value for index
                                        index++;

                                        // Set the fieldName
                                        ICell cell = headerRow.CreateCell(index);

                                        // Set the fieldName
                                        cell.SetCellValue(field.FieldName);

                                        // if this is a date
                                        if (field.DataType == DataManager.DataTypeEnum.DateTime)
                                        {
                                            // Setup as a date
                                            cell.CellStyle = dateCellStyle;
                                        }
                                        else
                                        {
                                            // Set up default font and size plus left alignment
                                            cell.CellStyle = cellStyle;
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

                // Now Save the workbook
                using (var fileStream = new FileStream(workbookFileInfo.FullName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }
            }
            #endregion
            
            #region ParseChangedColumnIndexes(string changedColumns)
            /// <summary>
            /// returns a list of Changed Column Indexes
            /// </summary>
            public static List<int> ParseChangedColumnIndexes(string changedColumns)
            {
                // initial value
                List<int> changedColumnIndexes = null;

                // If the changedColumns string exists
                if (TextHelper.Exists(changedColumns))
                {
                    // Create a new collection of 'int' objects.
                    changedColumnIndexes = new List<int>();

                    // Create a delimiter that is only a comma
                    char[] delimiters = { ',' };

                    // get the words
                    List<Word> words = TextHelper.GetWords(changedColumns, delimiters);

                    // If the words collection exists and has one or more items
                    if (ListHelper.HasOneOrMoreItems(words))
                    {
                        // Iterate the collection of Word objects
                        foreach (Word word in words)
                        {
                            // parse this column index
                            int columnIndex = NumericHelper.ParseInteger(word.Text, -1, -2);

                            // if this columnIndex was found
                            if (columnIndex >= 0)
                            {
                                // add this item
                                changedColumnIndexes.Add(columnIndex);
                            }
                        }
                    }
                }
                
                // return value
                return changedColumnIndexes;
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

                 // load the package
                XSSFWorkbook workbook = null;

                try
                {   
                    // if the batch exists and the batch has BatchItems (represents a Worksheet) and the path to the Excel file exists on disk
                    if ((NullHelper.Exists(batch)) && (batch.HasBatchItems) && (FileHelper.Exists(path)))
                    {
                       workbook = ExcelDataLoader.LoadExcelWorkbook(path);
                        
                        // iterate the batchItems
                        foreach (BatchItem batchItem in batch.BatchItems)
                        {
                            // if the batchItem has Updates and a WorksheetInfo
                            if ((batchItem.HasUpdates) && (batchItem.HasWorksheetInfo))
                            {
                                // Get the sheet
                                ISheet excelworksheet = workbook.GetSheet(batchItem.WorksheetInfo.SheetName);

                                // If the excelworksheet object exists
                                if (NullHelper.Exists(excelworksheet))
                                {
                                    // iterate the rows to update
                                    foreach (Row row in batchItem.Updates)
                                    {
                                        // Get the row
                                        IRow currentRow = excelworksheet.GetRow(row.Number);

                                        // If the value for the property rowNumber.HasColumns is true
                                        if (row.HasColumns)
                                        {
                                            // iterate the rows
                                            foreach (Column column in row.Columns)
                                            {
                                                // Get the cell
                                                ICell cell = currentRow.GetCell(column.ColumnNumber);

                                                if (NullHelper.Exists(cell))
                                                {
                                                    if ((onlyColumnsWithChanges) && (column.HasChanges))
                                                    {
                                                        // Set the value
                                                        cell.SetCellValue(column.ColumnValue.ToString());
                                                    }
                                                    else
                                                    {
                                                        // Set the value
                                                        cell.SetCellValue(column.ColumnValue.ToString());
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Save the workbook to a file
                        using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            // write changes
                            workbook.Write(fileStream);
                        }

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
                        // Load the workbook
                        XSSFWorkbook workbook = ExcelDataLoader.LoadExcelWorkbook(path);

                        // If the workbook object exists
                        if (NullHelper.Exists(workbook))
                        {
                            // Get the sheet
                            ISheet excelworksheet = workbook.GetSheet(batchItem.WorksheetInfo.SheetName);

                            // If the excelworksheet object exists
                            if (NullHelper.Exists(excelworksheet))
                            {
                                // iterate the rows to update
                                foreach (Row row in batchItem.Updates)
                                {
                                    // If the value for the property rowNumber.HasColumns is true
                                    if (row.HasColumns)
                                    {
                                        // Find this row
                                        IRow currentRow = excelworksheet.GetRow(row.Number);

                                        // iterate the rows
                                        foreach (Column column in row.Columns)
                                        {
                                            // Get the cell
                                            ICell cell = currentRow.GetCell(column.ColumnNumber);

                                            // If the cell object exists
                                            if (NullHelper.Exists(cell))
                                            {
                                                if ((onlyColumnsWithChanges) && (column.HasChanges))
                                                {
                                                    // Set the value
                                                    SetOptimizedCellValue(cell, column.ColumnValue);
                                                }
                                                else
                                                {
                                                    // Set the value
                                                    SetOptimizedCellValue(cell, column.ColumnValue);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Save the workbook to a file
                        using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            // write changes
                            workbook.Write(fileStream);
                        }

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

                // Load the excel workbook
                XSSFWorkbook excelWorkbook = null;

                try
                {   
                    // if the worksheet exists and the worksheet.WorksheetInfo exists and the path to the worksheet exists
                    if ((NullHelper.Exists(worksheet, row)) && (worksheet.HasWorksheetInfo) && (FileHelper.Exists(path)) && (row.HasColumns))
                    {
                        // Load the workbook
                        excelWorkbook = ExcelDataLoader.LoadExcelWorkbook(path);

                        // if exists
                        if (NullHelper.Exists(excelWorkbook))
                        {
                            // Get the sheet
                            ISheet excelworksheet = excelWorkbook.GetSheet(worksheet.Name);

                            // Find the currentRow
                            IRow currentRow = excelworksheet.GetRow(row.Number);

                            // If the excelworksheet object exists
                            if (NullHelper.Exists(excelworksheet))
                            {
                                // iterate the rows
                                foreach (Column column in row.Columns)
                                {
                                    // Find the Cell
                                    ICell cell = currentRow.GetCell(column.ColumnNumber);

                                    // If the cell object exists
                                    if (NullHelper.Exists(cell))
                                    {
                                        // if only columns with changes and this column has changes
                                        if ((onlyColumnsWithChanges) && (column.HasChanges))
                                        {
                                            // Set the value
                                            SetOptimizedCellValue(cell, column.ColumnValue);
                                        }
                                        else
                                        {
                                            // Set the value
                                            SetOptimizedCellValue(cell, column.ColumnValue);
                                        }
                                    }
                                }
                            }
                        }

                        // Save the workbook to a file
                        using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            // write changes
                            excelWorkbook.Write(fileStream);
                        }

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

            #region SaveWorksheet(List<IExcelerateObject> excelerateObjects,Worksheet worksheet, WorksheetInfo worksheetInfo, SaveInProgressCallback callback, int saveBatchInterval = 100)
            /// <summary>
            /// returns the Worksheet
            /// </summary>
            public static SaveWorksheetResponse SaveWorksheet(List<IExcelerateObject> excelerateObjects, Worksheet worksheet, WorksheetInfo worksheetInfo, SaveInProgressCallback callback, int saveBatchInterval = 100)
            {
                // initial value
                SaveWorksheetResponse response = new SaveWorksheetResponse();

                // locals
                Batch batch = new Batch();                    
                BatchItem batchItem = new BatchItem();
                batchItem.WorksheetInfo = worksheetInfo;
                batch.BatchItems.Add(batchItem);

                // If the excelerateObjects collection exists and has one or more items
                if (ListHelper.HasOneOrMoreItems(excelerateObjects))
                {
                    // Setup the graph before we start
                    response.CurrentRowNumber = 0;
                    response.TotalRows = excelerateObjects.Count;

                    // If the callback object exists
                    if (NullHelper.Exists(callback))
                    {
                        // call back to the client
                        callback(response);
                    }

                    // Iterate the collection of IExcelerateObject objects
                    foreach (IExcelerateObject excelerateObject in excelerateObjects)
                    {
                        // Increment the value for CurrentRowNumber
                        response.CurrentRowNumber++;

                        // find the row                        
                        Row row = worksheet.Rows.FirstOrDefault(x => x.Id == excelerateObject.RowId);

                        // If the row object exists
                        if (NullHelper.Exists(row))
                        {
                            // Save the property values columns in the row
                            excelerateObject.Save(row);

                            // Add this row
                            batchItem.Updates.Add(row);

                            // if the row exists
                            if (batch.BatchItems[0].Updates.Count == saveBatchInterval)
                            {
                                // perform the save
                                bool saved = ExcelHelper.SaveBatch(worksheetInfo.Path, batch, true);

                                 // if the value for saved is true
                                if ((saved) && (NullHelper.Exists(callback)))
                                {
                                    // update the rows saved
                                    response.RowsSaved += saveBatchInterval;

                                    // notify the client
                                    callback(response);

                                    // recreate the objects
                                    batch = new Batch();                    
                                    batchItem = new BatchItem();
                                    batchItem.WorksheetInfo = worksheetInfo;
                                    batch.BatchItems.Add(batchItem);
                                }
                            }
                        }
                    }
                }
                
                // return value
                return response;
            }
            #endregion
            
            #region SetOptimizedCellValue(ICell cell, object value)
            /// <summary>
            /// method returns the Optimized Cell Value
            /// </summary>
            public static void SetOptimizedCellValue(ICell cell, object value)
            {
                if (value is bool boolValue)
                {
                    cell.SetCellValue(boolValue);
                }
                else if (value is double doubleValue)
                {
                    cell.SetCellValue(doubleValue);
                }
                else if (value is int intValue)
                {
                    cell.SetCellValue(intValue);
                }
                else if (value is string stringValue)
                {
                    cell.SetCellValue(stringValue);
                }
                // Add other types as needed
                else
                {
                    cell.SetCellValue(value.ToString());
                }
            }
            #endregion
            
        #endregion

    }
    #endregion

}
