

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace DataJuggler.Excelerate
{

    #region class CodeGenerationResponse
    /// <summary>
    /// This class is used to return a response with the path to the classes generated so they can be downloaded.
    /// </summary>
    public class CodeGenerationResponse
    {
        
        #region Private Variables
        private bool success;
        private string fullPath;
        private string fileName;
        #endregion
        
        #region Properties
            
            #region FileName
            /// <summary>
            /// This property gets or sets the value for 'FileName'.
            /// </summary>
            public string FileName
            {
                get { return fileName; }
                set { fileName = value; }
            }
            #endregion
            
            #region FullPath
            /// <summary>
            /// This property gets or sets the value for 'FullPath'.
            /// </summary>
            public string FullPath
            {
                get { return fullPath; }
                set { fullPath = value; }
            }
            #endregion
            
            #region HasFileName
            /// <summary>
            /// This property returns true if the 'FileName' exists.
            /// </summary>
            public bool HasFileName
            {
                get
                {
                    // initial value
                    bool hasFileName = (!String.IsNullOrEmpty(this.FileName));
                    
                    // return value
                    return hasFileName;
                }
            }
            #endregion
            
            #region HasFullPath
            /// <summary>
            /// This property returns true if the 'FullPath' exists.
            /// </summary>
            public bool HasFullPath
            {
                get
                {
                    // initial value
                    bool hasFullPath = (!String.IsNullOrEmpty(this.FullPath));
                    
                    // return value
                    return hasFullPath;
                }
            }
            #endregion
            
            #region Success
            /// <summary>
            /// This property gets or sets the value for 'Success'.
            /// </summary>
            public bool Success
            {
                get { return success; }
                set { success = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
