

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class Worksheet
    /// <summary>
    /// This represents a Worksheet in an Excel Workbook
    /// </summary>
    public class Worksheet
    {
        
        #region Private Variables
        private List<Column> columns;
        private List<Row> rows;
        private string name;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'Worksheet' object.
        /// </summary>
        public Worksheet()
        {
            // Create both lists
            Columns = new List<Column>();
            Rows = new List<Row>();
        }
        #endregion
        
        #region Properties
            
            #region Columns
            /// <summary>
            /// This property gets or sets the value for 'Columns'.
            /// Worksheet.Rows has the actual Data. The Columns collection
            /// here is for a 'HeaderRow' and contains information about
            /// column data types.
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
            
            #region Name
            /// <summary>
            /// This property gets or sets the value for 'Name'.
            /// </summary>
            public string Name
            {
                get { return name; }
                set { name = value; }
            }
            #endregion
            
            #region Rows
            /// <summary>
            /// This property gets or sets the value for 'Rows'.
            /// </summary>
            public List<Row> Rows
            {
                get { return rows; }
                set { rows = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
