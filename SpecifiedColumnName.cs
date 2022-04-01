

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class SpecifiedColumnName
    /// <summary>
    /// This class is used to specify column names.
    /// The purpose of the class, is so a specified column name
    /// can hold a column index value, so it can be retried on 
    /// every row without going back to lookup the index.
    /// </summary>
    public class SpecifiedColumnName
    {
        
        #region Private Variables
        private string name;
        private int index;
        private bool notFound;
        #endregion

        #region Properties

            #region HasIndex
            /// <summary>
            /// This property returns true if the 'Index' is set.
            /// </summary>
            public bool HasIndex
            {
                get
                {
                    // initial value
                    bool hasIndex = (this.Index > 0);
                    
                    // return value
                    return hasIndex;
                }
            }
            #endregion
            
            #region Index
            /// <summary>
            /// This property gets or sets the value for 'Index'.
            /// </summary>
            public int Index
            {
                get { return index; }
                set { index = value; }
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
            
            #region NotFound
            /// <summary>
            /// This property gets or sets the value for 'NotFound'.
            /// </summary>
            public bool NotFound
            {
                get { return notFound; }
                set { notFound = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
