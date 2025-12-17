

#region using statements

using DataJuggler.Excelerate.Enumerations;
using DataJuggler.NET.Data;
using DataJuggler.UltimateHelper;
using System;
using System.Collections.Generic;
using System.Reflection;

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
        private Guid id;
        private int rowNumber;
        private int columnNumber;
        private bool columnContainsData;
        private object columnValue;        
        private string columnName;
        private string originalName;
        private string columnText;
        private string editText;
        private bool exportBooleanAsOneOrZero;
        private int index;
        private Row row;
        private DataManager.DataTypeEnum dataType;
        private bool hasChanges;

        // Added for DataJuggler.Blazor.Components.Grid
        private double borderWidth;
        private double imageHeight;
        private double imageWidth;
        private string imageUrl;
        private double width;
        private double height;
        private string unit;
        private string caption;
        private string className;
        private EditorTypeEnum editorType;
        private string editorClassName;
        private int zIndex;
        private bool hidden;
        private bool setFocusOnFirstRender;
        private bool primaryKey;
        private bool editMode;
        private bool isImageButton;
        private bool isImage;
        private string buttonClassName;
        private string buttonUrl;
        private int buttonNumber;
        private string format;
        #endregion

        #region Constructors

            #region Default Constructor
            /// <summary>
            /// Create a new instance of a Column object
            /// </summary>
            public Column()
            {
                // Create
                Id = Guid.NewGuid();
            }
            #endregion

            #region Parametized Constuctor Column(string columnName, int rowNumber, int colNumber, DataManager.DataTypeEnum dataType)
            /// <summary>
            /// Create a new instance of a Column object and set it's Data types
            /// </summary>
            /// <param name="columnName"></param>
            /// <param name="rowNumber"></param>
            /// <param name="columnNumber"></param>
            /// <param name="dataType"></param>
            public Column(string columnName, int rowNumber, int columnNumber, DataManager.DataTypeEnum dataType)
            {
                // Create
                Id = Guid.NewGuid();

                // store the args
                ColumnName = columnName;
                RowNumber = rowNumber;
                ColumnNumber = columnNumber;
                DataType = dataType;

                // Editors need to be in front
                ZIndex = 100;
            }
        #endregion

        #endregion

        #region Methods

            #region Clone()
            /// <summary>
            /// Creates a deep copy of this Column (except Row and ColumnValue, which
            /// should be set by the grid when creating rows).
            /// </summary>
            public Column Clone()
            {
                Column clone = new Column();

                // core identity
                clone.ColumnName = this.ColumnName;
                clone.OriginalName = this.OriginalName;
                clone.ColumnNumber = this.ColumnNumber;
                clone.DataType = this.DataType;
                clone.Width = this.Width;
                clone.Height = this.Height;
                clone.Unit = this.Unit;

                // display / style
                clone.Caption = this.Caption;
                clone.ClassName = this.ClassName;
                clone.EditorType = this.EditorType;
                clone.EditorClassName = this.EditorClassName;
                clone.ZIndex = this.ZIndex;
                clone.Hidden = this.Hidden;
                clone.SetFocusOnFirstRender = this.SetFocusOnFirstRender;
                clone.EditMode = this.EditMode;
                clone.IsImageButton = this.IsImageButton;
                clone.ButtonClassName = this.ButtonClassName;
                clone.ButtonUrl = this.ButtonUrl;
                clone.ButtonNumber = this.ButtonNumber;

                // Row and ColumnValue intentionally left null
                return clone;
            }
            #endregion

            #region ToString()
            /// <summary>
            /// method returns the String
            /// </summary>
            public override string ToString()
            {
                // Return the column name when ToString is called
                return ColumnName;
            }
            #endregion
            
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

                    // if the columnValue exists (can't use HasColumnValue here or capital ColumnValue)
                    if (NullHelper.Exists(columnValue))
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
            
            #region BorderWidth
            /// <summary>
            /// This property gets or sets the value for 'BorderWidth'.
            /// </summary>
            public double BorderWidth
            {
                get { return borderWidth; }
                set { borderWidth = value; }
            }
            #endregion
            
            #region BorderWidthStyle
            /// <summary>
            /// This read only property returns the value of BorderWidth + "px";
            /// </summary>
            public string BorderWidthStyle
            {

                get
                {
                    // initial value
                    string borderWidthStyle = BorderWidth + "px";
                    
                    // return value
                    return borderWidthStyle;
                }
            }
            #endregion

            #region ButtonClassName
            /// <summary>
            /// This property gets or sets the value for 'ButtonClassName'.
            /// </summary>
            public string ButtonClassName
            {
                get { return buttonClassName; }
                set { buttonClassName = value; }
            }
            #endregion
            
            #region ButtonNumber
            /// <summary>
            /// This property gets or sets the value for 'ButtonNumber'.
            /// </summary>
            public int ButtonNumber
            {
                get { return buttonNumber; }
                set { buttonNumber = value; }
            }
            #endregion
            
            #region ButtonUrl
            /// <summary>
            /// This property gets or sets the value for 'ButtonUrl'.
            /// </summary>
            public string ButtonUrl
            {
                get { return buttonUrl; }
                set { buttonUrl = value; }
            }
            #endregion
            
            #region Caption
            /// <summary>
            /// This property gets or sets the value for 'Caption'.
            /// </summary>
            public string Caption
            {
                get { return caption; }
                set { caption = value; }
            }
            #endregion
            
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
            
            #region ColumnContainsData
            /// <summary>
            /// This property gets or sets the value for 'ColumnContainsData'.
            /// This property is referring to a complete list of Rows, and only
            /// columns that one or more rows have Data will be included.
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
                            returnValue = 1;
                        }
                        else
                        {
                            // set the returnValue to 0 for false
                            returnValue = 0;
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

                     // if the columnValue exists (can't use HasColumnValue or capital ColumnValue here)
                    if (NullHelper.Exists(columnValue))
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

                    // if the columnValue exists (can't use HasColumnValue or the Capital ColumnValue here)
                    if (NullHelper.Exists(columnValue))
                    {
                         try
                        {
                            // attempt to cast as a Decimal
                            decimalValue = NumericHelper.ParseDecimal(StringValue.Replace("$", ""), decimalValue, decimalValue);
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

            #region EditMode
            /// <summary>
            /// This property gets or sets the value for 'EditMode'.
            /// </summary>
            public bool EditMode
            {
                get { return editMode; }
                set { editMode = value; }
            }
            #endregion
            
            #region EditorClassName
            /// <summary>
            /// This property gets or sets the value for 'EditorClassName'.
            /// </summary>            
            public string EditorClassName
            {
                get { return editorClassName; }
                set { editorClassName = value; }
            }
            #endregion
            
            #region EditorText
            /// <summary>
            /// This read only property returns the value of EditorText from the object EditText.
            /// </summary>
            public string EditorText
            {
                
                get
                {
                    // initial value
                    string editorText = ColumnText;
                    
                    // if EditText exists
                    if (TextHelper.Exists(EditText))
                    {
                        // set the return value
                        editorText = EditText;
                    }
                    
                    // return value
                    return editorText;
                }
            }
            #endregion
            
            #region EditorType
            /// <summary>
            /// This property gets or sets the value for 'EditorType'.
            /// </summary>
            public EditorTypeEnum EditorType
            {
                get { return editorType; }
                set { editorType = value; }
            }
            #endregion
            
            #region EditText
            /// <summary>
            /// This property gets or sets the value for 'EditText'.
            /// </summary>
            public string EditText
            {
                get { return editText; }
                set { editText = value; }
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
            
            #region Format
            /// <summary>
            /// This property gets or sets the value for 'Format'.
            /// </summary>
            public string Format
            {
                get { return format; }
                set { format = value; }
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

                    // if the columnValue exists (can't use HasColumnValue or the capital HasColumnvalue here)
                    if (NullHelper.Exists(columnValue))
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
            
            #region HasChanges
            /// <summary>
            /// This property gets or sets the value for 'HasChanges'.
            /// </summary>
            public bool HasChanges
            {
                get { return hasChanges; }
                set { hasChanges = value; }
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
            
            #region HasId
            /// <summary>
            /// This property returns true if this object has an 'Id'.
            /// </summary>
            public bool HasId
            {
                get
                {
                    // initial value
                    bool hasId = (this.Id != Guid.Empty);
                    
                    // return value
                    return hasId;
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
            
            #region HasRow
            /// <summary>
            /// This property returns true if this object has a 'Row'.
            /// </summary>
            public bool HasRow
            {
                get
                {
                    // initial value
                    bool hasRow = (this.Row != null);
                    
                    // return value
                    return hasRow;
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
            
            #region Height
            /// <summary>
            /// This property gets or sets the value for 'Height'.
            /// </summary>
            public double Height
            {
                get { return height; }
                set { height = value; }
            }
            #endregion
            
            #region HeightPlusUnit
            /// <summary>
            /// This read only property returns the value of Height + the Unit.
            /// </summary>
            public string HeightPlusUnit
            {
                
                get
                {
                    // set the return value
                    string heightPlusUnit = Height + unit;
                    
                    // return value
                    return heightPlusUnit;
                }
            }
            #endregion
            
            #region Hidden
            /// <summary>
            /// This property gets or sets the value for 'Hidden'.
            /// </summary>            
            public bool Hidden
            {
                get { return hidden; }
                set { hidden = value; }
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
            
            #region ImageHeight
            /// <summary>
            /// This property gets or sets the value for 'ImageHeight'.
            /// </summary>
            public double ImageHeight
            {
                get { return imageHeight; }
                set { imageHeight = value; }
            }
            #endregion
            
            #region ImageUrl
            /// <summary>
            /// This property gets or sets the value for 'ImageUrl'.
            /// </summary>
            public string ImageUrl
            {
                get { return imageUrl; }
                set { imageUrl = value; }
            }
            #endregion
            
            #region ImageWidth
            /// <summary>
            /// This property gets or sets the value for 'ImageWidth'.
            /// </summary>
            public double ImageWidth
            {
                get { return imageWidth; }
                set { imageWidth = value; }
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

                    // if the columnValue exists (can't use HasColumnValue or the capital ColumnValue here)
                    if (NullHelper.Exists(columnValue))
                    {
                         try
                        {
                            // attempt to cast as an int and remove dollar signs in case currency is present
                            intValue = NumericHelper.ParseInteger(StringValue.Replace("$", ""), 0, 0);
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
            
            #region IsEditable
            /// <summary>
            /// This read only property returns true if this object is not read only.
            /// </summary>
            public bool IsEditable
            {
                
                get
                {
                    // initial value
                    bool isEditable = false;
                    
                    // if EditorType exists
                    isEditable = (EditorType != EditorTypeEnum.ReadOnly);
                    
                    // return value
                    return isEditable;
                }
            }
            #endregion
            
            #region IsImage
            /// <summary>
            /// This property gets or sets the value for 'IsImage'.
            /// </summary>
            public bool IsImage
            {
                get { return isImage; }
                set { isImage = value; }
            }
            #endregion
        
            #region IsImageButton
            /// <summary>
            /// This property gets or sets the value for 'IsImageButton'.
            /// </summary>
            public bool IsImageButton
            {
                get { return isImageButton; }
                set { isImageButton = value; }
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
            
            #region PrimaryKey
            /// <summary>
            /// This property gets or sets the value for 'PrimaryKey'.
            /// </summary>
            public bool PrimaryKey
            {
                get { return primaryKey; }
                set { primaryKey = value; }
            }
            #endregion
            
            #region ReadOnly
            /// <summary>
            /// This read only property returns true if EditType = EditTypeEnum.ReadOnly
            /// </summary>
            public bool ReadOnly
            {
                
                get
                {
                    // initial value
                    bool readOnly = false;
                    
                    // set the return value
                    readOnly = (EditorType == EditorTypeEnum.ReadOnly);
                    
                    // return value
                    return readOnly;
                }
            }
            #endregion
            
            #region Row
            /// <summary>
            /// This property gets or sets the value for 'Row'.
            /// </summary>
            public Row Row
            {
                get { return row; }
                set { row = value; }
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
           
            #region SetFocusOnFirstRender
            /// <summary>
            /// This property gets or sets the value for 'SetFocusOnFirstRender'.
            /// </summary>
            public bool SetFocusOnFirstRender
            {
                get { return setFocusOnFirstRender; }
                set { setFocusOnFirstRender = value; }
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

                    // if the columnValue exists (can't use HasColumnValue or the capital here)
                    if (NullHelper.Exists(columnValue))
                    {
                        try
                        {
                            // attempt to cast as a string
                            stringValue = columnValue.ToString();
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
            
            #region Unit
            /// <summary>
            /// This property gets or sets the value for 'Unit'.
            /// </summary>
            public string Unit
            {
                get { return unit; }
                set { unit = value; }
            }
            #endregion
            
            #region Width
            /// <summary>
            /// This property gets or sets the value for 'Width'.
            /// </summary>
            public double Width
            {
                get { return width; }
                set { width = value; }
            }
            #endregion
            
            #region WidthPlusUnit
            /// <summary>
            /// This read only property returns the value of Width' + the Unit. 
            /// </summary>
            public string WidthPlusUnit
            {
                
                get
                {
                    // set the return value
                    string widthPlusUnit = Width + Unit;
                    
                    // return value
                    return widthPlusUnit;
                }
            }
            #endregion
            
            #region ZIndex
            /// <summary>
            /// This property gets or sets the value for 'ZIndex'.
            /// </summary>
            public int ZIndex
            {
                get { return zIndex; }
                set { zIndex = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}