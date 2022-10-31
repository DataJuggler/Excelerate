

#region using statements

using DataJuggler.Net7;
using DataJuggler.Net7.Enumerations;

#endregion

namespace DataJuggler.Excelerate
{

    #region class DataTypeScorer
    /// <summary>
    /// This class is used to make the Determine Datatype better.
    /// The idea is to add up matches, and return the best match
    /// instead of my first attempt kind of guessed at the first
    /// correct match.
    /// </summary>
    public class DataTypeScorer
    {
        
        #region Private Variables
        private int guidCount;
        private int dateCount;
        private int decimalCount;
        private int intCount;
        private int boolCount;
        private int stringCount;
        #endregion

        #region Properties

            #region BoolCount
            /// <summary>
            /// This property gets or sets the value for 'BoolCount'.
            /// </summary>
            public int BoolCount
            {
                get { return boolCount; }
                set { boolCount = value; }
            }
            #endregion
            
            #region DateCount
            /// <summary>
            /// This property gets or sets the value for 'DateCount'.
            /// </summary>
            public int DateCount
            {
                get { return dateCount; }
                set { dateCount = value; }
            }
            #endregion
            
            #region DecimalCount
            /// <summary>
            /// This property gets or sets the value for 'DecimalCount'.
            /// </summary>
            public int DecimalCount
            {
                get { return decimalCount; }
                set { decimalCount = value; }
            }
            #endregion
            
            #region GuidCount
            /// <summary>
            /// This property gets or sets the value for 'GuidCount'.
            /// </summary>
            public int GuidCount
            {
                get { return guidCount; }
                set { guidCount = value; }
            }
            #endregion
            
            #region IntCount
            /// <summary>
            /// This property gets or sets the value for 'IntCount'.
            /// </summary>
            public int IntCount
            {
                get { return intCount; }
                set { intCount = value; }
            }
            #endregion
            
            #region StringCount
            /// <summary>
            /// This property gets or sets the value for 'StringCount'.
            /// </summary>
            public int StringCount
            {
                get { return stringCount; }
                set { stringCount = value; }
            }
            #endregion
            
            #region TopDataType
            /// <summary>
            /// This read only property returns the value of TopDataType from the object DataType.
            /// </summary>
            public DataManager.DataTypeEnum TopDataType
            {
                
                get
                {
                    // initial value
                    DataManager.DataTypeEnum topDataType = DataManager.DataTypeEnum.String;

                    // local
                    int highest = StringCount;
                    
                    // if higher than the current value
                    if (BoolCount > StringCount)
                    {
                        // Set the return value
                        topDataType = DataManager.DataTypeEnum.Boolean;

                        // Set the new highest
                        highest = BoolCount;
                    }

                    // now use highest as the comparison

                    // if a double
                    if (DecimalCount > highest)
                    {
                        // Set the return value
                        topDataType = DataManager.DataTypeEnum.Double;

                        // set the new highest
                        highest = DecimalCount;
                    }

                    // if a double
                    if (DateCount > highest)
                    {
                        // Set the return value
                        topDataType = DataManager.DataTypeEnum.DateTime;

                        // set the new highest
                        highest = DateCount;
                    }

                    // if a guid
                    if (GuidCount > highest)
                    {
                         // Set the return value
                        topDataType = DataManager.DataTypeEnum.Guid;

                        // set the new highest
                        highest = GuidCount;
                    }

                    // if an int
                    if (IntCount > highest)
                    {
                         // Set the return value
                        topDataType = DataManager.DataTypeEnum.Integer;

                        // set the new highest
                        highest = IntCount;
                    }
                    
                    // return value
                    return topDataType;
                }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
