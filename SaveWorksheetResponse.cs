

#region using statements

using System;
using System.Collections.Generic;

#endregion

namespace DataJuggler.Excelerate
{

    #region class SaveWorksheetResponse
    /// <summary>
    /// This class is used to return information from SaveWorksheet or
    /// SaveWorksheets (coming soon).
    /// </summary>
    public class SaveWorksheetResponse
    {
        
        #region Private Variables
        private int rowsSaved;
        private int totalRows;
        private int currentRowNumber;
        private List<Exception> exceptions;
        #endregion

        #region Properties

            #region CurrentRowNumber
            /// <summary>
            /// This property gets or sets the value for 'CurrentRowNumber'.
            /// </summary>
            public int CurrentRowNumber
            {
                get { return currentRowNumber; }
                set { currentRowNumber = value; }
            }
            #endregion
            
            #region Exceptions
            /// <summary>
            /// This property gets or sets the value for 'Exceptions'.
            /// </summary>
            public List<Exception> Exceptions
            {
                get { return exceptions; }
                set { exceptions = value; }
            }
            #endregion
            
            #region RowsSaved
            /// <summary>
            /// This property gets or sets the value for 'RowsSaved'.
            /// </summary>
            public int RowsSaved
            {
                get { return rowsSaved; }
                set { rowsSaved = value; }
            }
            #endregion
            
            #region TotalRows
            /// <summary>
            /// This property gets or sets the value for 'TotalRows'.
            /// </summary>
            public int TotalRows
            {
                get { return totalRows; }
                set { totalRows = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
