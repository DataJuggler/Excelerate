

#region using statements

using System;
using System.Collections.Generic;
using DataJuggler.UltimateHelper;

#endregion

namespace DataJuggler.Excelerate
{

    #region class Workbook
    /// <summary>
    /// This class represents an Excel Workbook.
    /// </summary>
    public class Workbook
    {
        
        #region Private Variables
        private List<Worksheet> worksheets;
        private string response;
        private List<Exception> errors;
        #endregion

        #region Constructor
        /// <summary>
        /// Create a new instance of a Workbook object.
        /// </summary>
        public Workbook()
        {
            // Create a new collection of 'Worksheet' objects.
            Worksheets = new List<Worksheet>();
        }
        #endregion

        #region Methods

            #region GetWorksheetIndex(string worksheetName)
            /// <summary>
            /// This method returns the Worksheet Index
            /// </summary>
            public int GetWorksheetIndex(string worksheetName)
            {
                // initial value
                int worksheetIndex = -1;

                // local
                int index = -1;

                // if the value for HasWorksheets is true
                if (HasWorksheets)
                {
                    // Iterate the collection of Worksheet objects
                    foreach (Worksheet worksheet in worksheets)
                    {
                        // Increment the value for index
                        index++;

                        // if this is the sheet intended
                        if (TextHelper.IsEqual(worksheet.Name, worksheetName))
                        {
                            // set the return value                            
                            worksheetIndex = index;

                            // break out of loop
                            break;
                        }
                    }
                }
                
                // return value
                return worksheetIndex;
            }
            #endregion

            #region GetWorksheet(string worksheetName)
            /// <summary>
            /// This method returns the Worksheet
            /// </summary>
            public Worksheet GetWorksheet(string worksheetName)
            {
                // initial value
                Worksheet worksheet = null;

                // local
                int worksheetIndex = GetWorksheetIndex(worksheetName);

                // if the worksheet index was found
                if (worksheetIndex >= 0)
                {
                    // set the return value
                    worksheet = Worksheets[worksheetIndex];
                }

                // return value
                return worksheet;
            }
            #endregion
            
        #endregion

        #region Properties

            #region Errors
            /// <summary>
            /// This property gets or sets the value for 'Errors'.
            /// </summary>
            public List<Exception> Errors
            {
                get { return errors; }
                set { errors = value; }
            }
            #endregion
            
            #region HasErrors
            /// <summary>
            /// This property returns true if this object has an 'Errors'.
            /// </summary>
            public bool HasErrors
            {
                get
                {
                    // initial value
                    bool hasErrors = (this.Errors != null);
                    
                    // return value
                    return hasErrors;
                }
            }
            #endregion
            
            #region HasResponse
            /// <summary>
            /// This property returns true if the 'Response' exists.
            /// </summary>
            public bool HasResponse
            {
                get
                {
                    // initial value
                    bool hasResponse = (!String.IsNullOrEmpty(this.Response));
                    
                    // return value
                    return hasResponse;
                }
            }
            #endregion
            
            #region HasWorksheets
            /// <summary>
            /// This property returns true if this object has a 'Worksheets'.
            /// </summary>
            public bool HasWorksheets
            {
                get
                {
                    // initial value
                    bool hasWorksheets = (this.Worksheets != null);
                    
                    // return value
                    return hasWorksheets;
                }
            }
            #endregion
            
            #region Response
            /// <summary>
            /// This property gets or sets the value for 'Response'.
            /// </summary>
            public string Response
            {
                get { return response; }
                set { response = value; }
            }
            #endregion
            
            #region Worksheets
            /// <summary>
            /// This property gets or sets the value for 'Worksheets'.
            /// </summary>
            public List<Worksheet> Worksheets
            {
                get { return worksheets; }
                set { worksheets = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}