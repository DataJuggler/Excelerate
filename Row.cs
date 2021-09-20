

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class Row
    /// <summary>
    /// This class represents the Columns make up the Data for an excel sheet.
    /// </summary>
    public class Row
    {
        
        #region Private Variables
        private List<Column> columns;
        private bool isHeaderRow;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'Row' object.
        /// </summary>
        public Row()
        {
            // Create a new collection of 'Column' objects.
            Columns = new List<Column>();
        }
        #endregion
        
        #region Properties
            
            #region Columns
            /// <summary>
            /// This property gets or sets the value for 'Columns'.
            /// </summary>
            public List<Column> Columns
            {
                get { return columns; }
                set { columns = value; }
            }
            #endregion
            
            #region HasColumns
            /// <summary>
            /// This property returns true if this object has a 'Columns'.
            /// </summary>
            public bool HasColumns
            {
                get
                {
                    // initial value
                    bool hasColumns = (this.Columns != null);
                    
                    // return value
                    return hasColumns;
                }
            }
            #endregion
            
            #region IsHeaderRow
            /// <summary>
            /// This property gets or sets the value for 'IsHeaderRow'.
            /// </summary>
            public bool IsHeaderRow
            {
                get { return isHeaderRow; }
                set { isHeaderRow = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
