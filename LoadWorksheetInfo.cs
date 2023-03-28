

#region using statements

using System.Collections.Generic;
using DataJuggler.Net7;

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
        private List<int> excludedColumnIndexes;
        private List<DataField> fields;
        private List<DataRow> rows;
        #endregion

        #region Constructor
        /// <summary>
        /// Create a new instancce of a LoadWorkSheetInfo object
        /// </summary>
        public LoadWorksheetInfo()
        {
            // Create the lists
            SpecifiedColumnNames = new List<SpecifiedColumnName>();
            ExcludedColumnIndexes = new List<int>();
        }
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
            
            #region ExcludedColumnIndexes
            /// <summary>
            /// This property gets or sets the value for 'ExcludedColumnIndexes'.
            /// </summary>
            public List<int> ExcludedColumnIndexes
            {
                get { return excludedColumnIndexes; }
                set { excludedColumnIndexes = value; }
            }
            #endregion
            
            #region Fields
            /// <summary>
            /// This property gets or sets the value for 'Fields'.
            /// </summary>
            public List<DataField> Fields
            {
                get { return fields; }
                set { fields = value; }
            }
            #endregion
            
            #region HasFields
            /// <summary>
            /// This property returns true if this object has a 'Fields'.
            /// </summary>
            public bool HasFields
            {
                get
                {
                    // initial value
                    bool hasFields = (this.Fields != null);
                    
                    // return value
                    return hasFields;
                }
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
            
            #region HasRows
            /// <summary>
            /// This property returns true if this object has a 'Rows'.
            /// </summary>
            public bool HasRows
            {
                get
                {
                    // initial value
                    bool hasRows = (this.Rows != null);
                    
                    // return value
                    return hasRows;
                }
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
            
            #region Rows
            /// <summary>
            /// This property gets or sets the value for 'Rows'.
            /// </summary>
            public List<DataRow> Rows
            {
                get { return rows; }
                set { rows = value; }
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
