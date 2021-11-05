

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using DataJuggler.UltimateHelper;

#endregion

namespace DataJuggler.Excelerate
{

    #region class ExcelHelper
    /// <summary>
    /// This class is used to help with common Excel tasks, with the main one being Saving data to Excel.
    /// </summary>
    public class ExcelHelper
    {

        #region Methods

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

                        // iterate the rows
                        foreach (Column column in row.Columns)
                        {
                            // Set the value
                            excelworksheet.SetValue(column.RowNumber, column.ColumnNumber, column.ColumnValue);
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
