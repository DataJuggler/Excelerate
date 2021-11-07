

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class Batch
    /// <summary>
    /// This class is used to update one or more BatchItems.
    /// </summary>
    public class Batch
    {
        
        #region Private Variables
        private List<BatchItem> batchItems;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'Batch' object.
        /// </summary>
        public Batch()
        {
            // Create a new instance of a 'batchItems' object.
            BatchItems = new List<BatchItem>();
        }
        #endregion
        
        #region Properties
            
            #region BatchItems
            /// <summary>
            /// This property gets or sets the value for 'BatchItems'.
            /// </summary>
            public List<BatchItem> BatchItems
            {
                get { return batchItems; }
                set { batchItems = value; }
            }
            #endregion
            
            #region HasBatchItems
            /// <summary>
            /// This property returns true if this object has a 'BatchItems'.
            /// </summary>
            public bool HasBatchItems
            {
                get
                {
                    // initial value
                    bool hasBatchItems = (this.BatchItems != null);
                    
                    // return value
                    return hasBatchItems;
                }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
