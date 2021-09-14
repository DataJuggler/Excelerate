

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataJuggler.Excelerate;

#endregion

namespace DataJuggler.Excelerate
{

    #region class LoadWorksheetInfo
    /// <summary>
    /// This class is here so each Worksheet can have its own Load options
    /// </summary>
    public class LoadWorksheetInfo    
    {

        #region Private Variables
        private LoadColumnOptionsEnum loadColumnOptions;
        private string sheetName;
        private List<SpecifiedColumnName> specifiedColumnNames;
        private int maxRowsToLoad;
        private bool hasHeaderRow;
        private int columnsToLoad;
        #endregion

        #region Properties

            #region ColumnsToLoad
            /// <summary>
            /// This property gets or sets the value for 'ColumnsToLoad'.
            /// This property is only used when LoadColumnOptionsEnum.LoadFirstXColumns
            /// is set to true and this value is set to a value less than the column count of the worksheet.
            /// This is useful if you want to load the first X number of columns.
            /// Example: if ColumnsToLoad = 6, this would be columns A - F.
            /// </summary>
            public int ColumnsToLoad
            {
                get { return columnsToLoad; }
                set { columnsToLoad = value; }
            }
            #endregion
            
            #region HasHeaderRow
            /// <summary>
            /// This property gets or sets the value for 'HasHeaderRow'.
            /// </summary>
            public bool HasHeaderRow
            {
                get { return hasHeaderRow; }
                set { hasHeaderRow = value; }
            }
            #endregion
            
            #region LoadColumnOptions
            /// <summary>
            /// This property gets or sets the value for 'LoadColumnOptions'.
            /// </summary>
            public LoadColumnOptionsEnum LoadColumnOptions
            {
                get { return loadColumnOptions; }
                set { loadColumnOptions = value; }
            }
            #endregion
            
            #region MaxRowsToLoad
            /// <summary>
            /// This property gets or sets the value for 'MaxRowsToLoad'.
            /// </summary>
            public int MaxRowsToLoad
            {
                get { return maxRowsToLoad; }
                set { maxRowsToLoad = value; }
            }
            #endregion
            
            #region SheetName
            /// <summary>
            /// This property gets or sets the value for 'SheetName'.
            /// </summary>
            public string SheetName
            {
                get { return sheetName; }
                set { sheetName = value; }
            }
            #endregion
            
            #region SpecifiedColumnNames
            /// <summary>
            /// This property gets or sets the value for 'SpecifiedColumnNames'.
            /// </summary>
            public List<SpecifiedColumnName> SpecifiedColumnNames
            {
                get { return specifiedColumnNames; }
                set { specifiedColumnNames = value; }
            }
            #endregion
            
        #endregion

    }
    #endregion

}
