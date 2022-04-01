

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class BatchItem
    /// <summary>
    /// This class represents all the changes for a Worksheet.
    /// </summary>
    public class BatchItem
    {
        
        #region Private Variables
        private LoadWorksheetInfo worksheetInfo;
        private List<Row> updates;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'BatchItem' object.
        /// </summary>
        public BatchItem()
        {
            // Create a new collection of 'Row' objects.
            this.Updates = new List<Row>();

            // Create the WorksheetInfo
            this.WorksheetInfo = new LoadWorksheetInfo();
        }
        #endregion
        
        #region Properties
            
            #region HasUpdates
            /// <summary>
            /// This property returns true if this object has an 'Updates'.
            /// </summary>
            public bool HasUpdates
            {
                get
                {
                    // initial value
                    bool hasUpdates = (this.Updates != null);
                    
                    // return value
                    return hasUpdates;
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
            
            #region Updates
            /// <summary>
            /// This property gets or sets the value for 'Updates'.
            /// </summary>
            public List<Row> Updates
            {
                get { return updates; }
                set { updates = value; }
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
