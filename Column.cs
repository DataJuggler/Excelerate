

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataJuggler.UltimateHelper;
using DataJuggler.Net5;

#endregion

namespace DataJuggler.Excelerate
{

    #region class Column
    /// <summary>
    /// This class represents one column in a Worksheet
    /// </summary>
    public class Column
    {
        
        #region Private Variables
        private int rowNumber;
        private int columnNumber;
        private bool columnContainsData;
        private object columnValue;
        private string columnName;
        private string originalName;
        private string columnText;
        private bool exportBooleanAsOneOrZero;
        private DataManager.DataTypeEnum dataType;
        #endregion

        #region Properties

            #region BoolValue
            /// <summary>
            /// This property gets or sets the value for 'BoolValue'.
            /// </summary>
            public bool BoolValue
            {
                get 
                {
                    // initial value
                    bool boolValue = false;

                    // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                        try
                        {
                            // attempt to cast as a bool
                            boolValue = BooleanHelper.ParseBoolean(StringValue, false, false);
                        }
                        catch (Exception error)
                        {
                            // for debugging only
                            DebugHelper.WriteDebugError("BoolValue", "ExcelImort.Column", error);
                        }
                    }

                    // initial value
                    return boolValue;
                }
            }
            #endregion
            
            #region ColumnContainsData
            /// <summary>
            /// This property gets or sets the value for 'ColumnContainsData'.
            /// This property is referring to a complete list of Rows, and only
            /// columns that one or more rows have data will be included.
            /// </summary>
            public bool ColumnContainsData
            {
                get { return columnContainsData; }
                set { columnContainsData = value; }
            }
            #endregion
            
            #region ColumnName
            /// <summary>
            /// This property gets or sets the value for 'ColumnName'.
            /// </summary>
            public string ColumnName
            {
                get { return columnName; }
                set { columnName = value; }
            }
            #endregion
            
            #region ColumnNumber
            /// <summary>
            /// This property gets or sets the value for 'ColumnNumber'.
            /// </summary>
            public int ColumnNumber
            {
                get { return columnNumber; }
                set { columnNumber = value; }
            }
            #endregion
            
            #region ColumnText
            /// <summary>
            /// This property gets or sets the value for 'ColumnText'.
            /// </summary>
            public string ColumnText
            {
                get { return columnText; }
                set { columnText = value; }
            }
            #endregion
            
            #region ColumnValue
            /// <summary>
            /// This property gets or sets the value for 'ColumnValue'.
            /// </summary>
            public object ColumnValue
            {
                get 
                {
                    // set the initial value
                    object returnValue = columnValue;

                    // formatting changes here. Something more advanced might be needed.
                    if ((DataType == DataManager.DataTypeEnum.Boolean) && (ExportBooleanAsOneOrZero))
                    {
                        // if the value for BoolValue is true
                        if (BoolValue)
                        {
                            // set the returnValue to 1 for true
                            returnValue = "1";
                        }
                        else
                        {
                            // set the returnValue to 0 for false
                            returnValue = "0";
                        }
                    }

                    // return value
                    return returnValue;
                }
                set { columnValue = value; }
            }
            #endregion
            
            #region DataType
            /// <summary>
            /// This property gets or sets the value for 'DataType'.
            /// </summary>
            public DataManager.DataTypeEnum DataType
            {
                get { return dataType; }
                set { dataType = value; }
            }
            #endregion
            
            #region DateValue
            /// <summary>
            /// This property gets or sets the value for 'DateValue'.
            /// </summary>
            public DateTime? DateValue
            {
                get 
                {
                    // initial value
                    DateTime? dateValue = null;

                    // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                         try
                        {
                            // attempt to cast as a DateTime
                            dateValue = (DateTime) ColumnValue;
                        }
                        catch (Exception error)
                        {
                            // for debugging only
                            DebugHelper.WriteDebugError("DateValue", "ExcelImort.Column", error);
                        }
                    }
                    
                    // initial value
                    return dateValue;
                }
            }
            #endregion
            
            #region DecimalValue
            /// <summary>
            /// This property gets or sets the value for 'DecimalValue'.
            /// </summary>
            public Decimal DecimalValue
            {
                get
                {
                    // initial value
                    Decimal decimalValue = 0;

                    // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                         try
                        {
                            // attempt to cast as a Decimal
                            decimalValue = NumericHelper.ParseDecimal(StringValue, decimalValue, decimalValue);
                        }
                        catch (Exception error)
                        {
                            // for debugging only
                            DebugHelper.WriteDebugError("DecimalValue", "ExcelImort.Column", error);
                        }
                    }

                    // return value
                    return decimalValue;
                }
            }
            #endregion

            #region ExportBooleanAsOneOrZero
            /// <summary>
            /// This property gets or sets the value for 'ExportBooleanAsOneOrZero'.
            /// </summary>
            public bool ExportBooleanAsOneOrZero
            {
                get { return exportBooleanAsOneOrZero; }
                set { exportBooleanAsOneOrZero = value; }
            }
            #endregion
            
            #region GuidValue
            /// <summary>
            /// This property returns the value for 'GuidValue'.
            /// </summary>
            public Guid GuidValue
            {
                get
                {
                    // initial value
                    Guid guidValue = Guid.Empty;

                    // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                         try
                        {
                            // attempt to cast as a Decimal
                            guidValue = (Guid) this.ColumnValue;
                        }
                        catch (Exception error)
                        {
                            // for debugging only
                            DebugHelper.WriteDebugError("GuidValue", "Column", error);
                        }
                    }

                    // return value
                    return guidValue;
                }
            }
            #endregion
            
            #region HasColumnName
            /// <summary>
            /// This property returns true if the 'ColumnName' exists.
            /// </summary>
            public bool HasColumnName
            {
                get
                {
                    // initial value
                    bool hasColumnName = (!String.IsNullOrEmpty(this.ColumnName));
                    
                    // return value
                    return hasColumnName;
                }
            }
            #endregion
            
            #region HasColumnValue
            /// <summary>
            /// This property returns true if this object has a 'ColumnValue'.
            /// </summary>
            public bool HasColumnValue
            {
                get
                {
                    // initial value
                    bool hasColumnValue = (this.ColumnValue != null);
                    
                    // return value
                    return hasColumnValue;
                }
            }
            #endregion
            
            #region HasDateValue
            /// <summary>
            /// This property returns true if this object has a 'DateValue'.
            /// </summary>
            public bool HasDateValue
            {
                get
                {
                    // initial value
                    bool hasDateValue = ((this.DateValue != null) && (DateValue.HasValue) && (DateValue.Value.Year > 1900));
                    
                    // return value
                    return hasDateValue;
                }
            }
            #endregion
            
            #region HasDecimalValue
            /// <summary>
            /// This property returns true if this object has a 'DecimalValue'.
            /// </summary>
            public bool HasDecimalValue
            {
                get
                {
                    // initial value
                    bool hasDecimalValue = (this.DecimalValue > 0);
                    
                    // return value
                    return hasDecimalValue;
                }
            }
            #endregion
            
            #region HasIntValue
            /// <summary>
            /// This property returns true if the 'IntValue' is set.
            /// </summary>
            public bool HasIntValue
            {
                get
                {
                    // initial value
                    bool hasIntValue = (this.IntValue > 0);
                    
                    // return value
                    return hasIntValue;
                }
            }
            #endregion
            
            #region HasStringValue
            /// <summary>
            /// This property returns true if the 'StringValue' exists.
            /// </summary>
            public bool HasStringValue
            {
                get
                {
                    // initial value
                    bool hasStringValue = (!String.IsNullOrEmpty(this.StringValue));
                    
                    // return value
                    return hasStringValue;
                }
            }
            #endregion
            
            #region IntValue
            /// <summary>
            /// This property returns the value for 'IntValue'.
            /// </summary>
            public int IntValue
            {
                get 
                { 
                    // initial value
                    int intValue = 0;

                    // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                         try
                        {
                            // attempt to cast as an int
                            intValue = NumericHelper.ParseInteger(StringValue, 0, 0);
                        }
                        catch (Exception error)
                        {
                            // for debugging only
                            DebugHelper.WriteDebugError("IntValue", "ExcelImort.Column", error);
                        }
                    }

                    // return value
                    return intValue;
                }
            }
            #endregion
            
            #region OriginalName
            /// <summary>
            /// This property gets or sets the value for 'OriginalName'.
            /// The original name is the name as it is in the source Excel.
            /// The purpose of this property is after ReplaceSpecialCharacters is called,
            /// the ColumnName is set and I needed a way too display the name the way
            /// it is in the source excel.
            /// </summary>
            public string OriginalName
            {
                get { return originalName; }
                set { originalName = value; }
            }
            #endregion
            
            #region RowNumber
            /// <summary>
            /// This property gets or sets the value for 'RowNumber'.
            /// </summary>
            public int RowNumber
            {
                get { return rowNumber; }
                set { rowNumber = value; }
            }
            #endregion
            
            #region StringValue
            /// <summary>
            /// This property gets or sets the value for 'StringValue'.
            /// </summary>
            public string StringValue
            {
                get
                {
                    // initial value
                    string stringValue = "";

                     // if the value for HasColumnValue is true
                    if (HasColumnValue)
                    {
                        try
                        {
                            // attempt to cast as a string
                            stringValue = ColumnValue.ToString();
                        }
                        catch (Exception error)
                        {
                            // set to an emptyString versus null
                            stringValue = "";

                            // for debugging only
                            DebugHelper.WriteDebugError("StringValue", "Column", error);
                        }
                    }

                    // returnValue
                    return stringValue;
                }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
