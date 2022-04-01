

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataJuggler.Net6;
using DataJuggler.UltimateHelper;

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
        private List<Row> rows;
        private string name;
        private LoadWorksheetInfo worksheetInfo;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'Worksheet' object.
        /// </summary>
        public Worksheet()
        {            
            Rows = new List<Row>();
        }
        #endregion

        #region Properties
            
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
            
            #region HasWorksheetInfo
            /// <summary>
            /// This property returns true if this object has a 'WorksheetInfo'.
            /// </summary>
            public bool HasWorksheetInfo
            {
                get
                {
                    // initial value
                    bool hasWorksheetInfo = (this.WorksheetInfo != null);
                    
                    // return value
                    return hasWorksheetInfo;
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
            
            #region WorksheetInfo
            /// <summary>
            /// This property gets or sets the value for 'WorksheetInfo'.
            /// </summary>
            public LoadWorksheetInfo WorksheetInfo
            {
                get { return worksheetInfo; }
                set { worksheetInfo = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
