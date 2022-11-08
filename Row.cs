

#region using statements

using System;
using System.Collections.Generic;
using DataJuggler.UltimateHelper;

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
        private int number;
        private Guid id;
        private bool isHeaderRow;
        private string className;
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
        
        #region Methods

            #region FindColumn(string name)
            /// <summary>
            /// returns the Column
            /// </summary>
            public Column FindColumn(string name)
            {
                // initial value
                Column column = null;

                // if the value for HasColumns is true
                if (HasColumns)
                {
                    // Iterate the collection of Column objects
                    foreach (Column tempColumn in Columns)
                    {
                        if (TextHelper.IsEqual(tempColumn.ColumnName, name))
                        {
                            // set the return value
                            column = tempColumn;

                            // break out of the loop
                            break;
                        }
                    }
                }
                
                // return value
                return column;
            }
            #endregion
            
        #endregion
        
        #region Properties
            
            #region ClassName
            /// <summary>
            /// This property gets or sets the value for 'ClassName'.
            /// </summary>
            public string ClassName
            {
                get { return className; }
                set { className = value; }
            }
            #endregion
            
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
            
            #region Id
            /// <summary>
            /// This property gets or sets the value for 'Id'.
            /// </summary>
            public Guid Id
            {
                get { return id; }
                set { id = value; }
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
            
            #region Number
            /// <summary>
            /// This property gets or sets the value for 'Number'.
            /// </summary>
            public int Number
            {
                get { return number; }
                set { number = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
