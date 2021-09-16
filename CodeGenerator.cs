

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataJuggler.Net5;
using DataJuggler.UltimateHelper;

#endregion

namespace DataJuggler.Excelerate
{

    #region class CodeGenerator
    /// <summary>
    /// method [Enter Method Description]
    /// </summary>
    public class CodeGenerator : CSharpClassWriter
    {
        
        #region Private Variables
        private Worksheet worksheet;
        private string outputFolder;
        private string className;        
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a CodeGenerator object
        /// </summary>
        public CodeGenerator(Worksheet worksheet, string outputFolder, string className): base(false)
        {
            // Store the args
            Worksheet = worksheet;
            OutputFolder = outputFolder;
            ClassName = className;

            // Added so that the name comes out [className].cs
            DoNotPrependExtension = true;

            // At this point not using partial classes, not ruling it out for later
            DoNotUsePartialClass = true;
        }
        #endregion

        #region Methods

            #region AttemptToDetermineDataType(int columnIndex)
            /// <summary>
            /// This method returns the To Determine Data Type
            /// </summary>
            /// <param name="fieldName">The fieldName is only needed here for debugging.</param>
            /// <param name="defaultDataType"></param>
            /// <param name="columnIndex">This is the 0 based index of this column in a DataJuggler.Excelerate.Row.Columns.
            /// This is not to be confused with a 1 based Excel Column. Be aware this index is 1 less than the Column Number
            /// field using EPPPlus / Excel.
            /// </param>
            public DataManager.DataTypeEnum AttemptToDetermineDataType(string fieldName, int columnIndex, DataManager.DataTypeEnum defaultDataType = DataManager.DataTypeEnum.Object)
            {
                // initial value (use object or string is probably best)
                DataManager.DataTypeEnum dataType = defaultDataType;

                // locals
                string temp = "";
                int tempInt = 0;
                Decimal tempDecimal = 0;
                DateTime? tempDate = new DateTime();
                int maxToLookAt = 50;
                int lookedAt = 0;
                
                // If the worksheet object exists
                if ((NullHelper.Exists(worksheet)) && (ListHelper.HasOneOrMoreItems(worksheet.Rows)))
                {
                    for (int x = 1; x < worksheet.Rows.Count; x++)
                    {
                        // get the value in this position
                        temp = worksheet.Rows[x].Columns[columnIndex].ColumnText;

                        // If the temp string exists, and this is not a 0, 0's are hard to tell anything from
                        if ((TextHelper.Exists(temp)) && (temp != "0") && (temp != "0.00"))
                        {
                            // get the values
                            tempInt = worksheet.Rows[x].Columns[columnIndex].IntValue;
                            tempDecimal = worksheet.Rows[x].Columns[columnIndex].DecimalValue;
                            tempDate = worksheet.Rows[x].Columns[columnIndex].DateValue;

                            // if a column starts with preceding zeros, and it is a number I am counting this as string
                            if ((temp.StartsWith("0")) && ((tempInt != 0) || (tempDecimal != 0)))
                            {
                                // Set to string
                                dataType = DataManager.DataTypeEnum.String;

                                // break out of the loop
                                break;
                            }

                            // if this is a number
                            if ((tempInt != 0) || (tempDecimal != 0))
                            {
                                // if the string contains a decimal point
                                if (temp.Contains("."))
                                {
                                    // Use Decimal
                                    dataType = DataManager.DataTypeEnum.Decimal;

                                    // break out
                                    break;
                                }
                                else
                                {
                                    // Use Integer, but keep looking
                                    dataType = DataManager.DataTypeEnum.Integer;
                                }
                            }
                            else if ((tempDate.HasValue) && (tempDate.Value.Year > 1900))
                            {
                                // Use Date
                                dataType = DataManager.DataTypeEnum.DateTime;

                                // once we determine we used DateTime, break out of loop
                                break;
                            }
                            else
                            {  
                                // Use String
                                dataType = DataManager.DataTypeEnum.String;
                                
                                // Use for scoring
                                break;
                            }
                            
                            // Increment the value for lookedAt
                            lookedAt++;

                            // if we have looked at enough
                            if (lookedAt > maxToLookAt)
                            {
                                // break out of the loop
                                break;
                            }
                        }
                    }
                }
                
                // return value
                return dataType;
            }
            #endregion
            
            #region GenerateClassFromWorksheet(string namespaceName)
            /// <summary>
            /// This method returns a Class From the Worksheet supplied in the Constructor.
            /// The Worksheet must have a HeaderRow for the top row.
            /// </summary>
            public bool GenerateClassFromWorksheet(string namespaceName)
            {
                // initial value
                bool success  = false;

                // local
                int columnIndex = -1;

                // if the value for IsValid is true (means there is a worksheet and it has at least one row)
                if (IsValid)
                {
                    // Get the first row
                    Row row = worksheet.Rows[0];

                    // if the row exists
                    if (ListHelper.HasOneOrMoreItems(row.Columns))
                    {
                        // Create a DataTable to create the 
                        DataManager dataManager = new DataManager(outputFolder, this.ClassName, DataManager.ClassOutputLanguage.CSharp);

                        // add the Namespace
                        dataManager.NamespaceName = namespaceName;

                        // Create a database to hold the tables
                        Database database = new Database();

                        // Create a new instance of a 'DataTable' object.
                        DataTable dataTable = new DataTable(database);

                        // Set the Name
                        dataTable.Name = ClassName;

                        // Create DataFields for each column
                        foreach (Column column in row.Columns)
                        {
                            // if the ColumNValue exists
                            if (column.HasColumnValue)
                            {
                                // Increment the value for columnIndex
                                columnIndex++;

                                // Create a field for this column
                                DataField field = new DataField();

                                // Set the name, but replace out things that make it an illegal field name like spaces or dashes
                                field.FieldName = TextHelper.CapitalizeFirstChar(ReplaceInvalidCharacters(column.StringValue));

                                // Set the FieldOrdinal
                                field.FieldOrdinal = columnIndex;

                                // DetermineDataType
                                field.DataType = AttemptToDetermineDataType(field.FieldName, columnIndex);

                                // Add this field
                                dataTable.Fields.Add(field);
                            }
                        }

                        // Sort the fields
                        List<DataField> sortedFields = dataTable.Fields.OrderBy(x => x.FieldName).ToList();

                        // Sort the fields so the Properties will be in order
                        dataTable.Fields = sortedFields;

                        // Create a dataManager
                        database.Tables.Add(dataTable);

                        // Add this database
                        dataManager.Databases.Add(database);

                        // Write out the class
                        success = WriteDataClasses(dataManager);
                    }
                }
                
                // return value
                return success;
            }
            #endregion
            
            #region ReplaceInvalidCharacters(string fieldName)
            /// <summary>
            /// This method returns the Invalid Characters
            /// </summary>
            public string ReplaceInvalidCharacters(string fieldName)
            {
                // If the fieldName string exists
                if (TextHelper.Exists(fieldName))
                {
                    // Remove any weird characters. This list may grow
                    fieldName = fieldName.Replace("-", "").Replace(" ", "").Replace(".", "").Replace("#", "").Replace("$", "").Replace("_", "").Replace("+", "").Replace("(", "").Replace(")", "").Replace("[", "").Replace("]", "");

                    // Percents are handled differently
                    fieldName = fieldName.Replace("%", "Percent");
                }
                
                // return value
                return fieldName;
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
            
            #region HasClassName
            /// <summary>
            /// This property returns true if the 'ClassName' exists.
            /// </summary>
            public bool HasClassName
            {
                get
                {
                    // initial value
                    bool hasClassName = (!String.IsNullOrEmpty(this.ClassName));
                    
                    // return value
                    return hasClassName;
                }
            }
            #endregion
            
            #region HasOutputFolder
            /// <summary>
            /// This property returns true if the 'OutputFolder' exists.
            /// </summary>
            public bool HasOutputFolder
            {
                get
                {
                    // initial value
                    bool hasOutputFolder = (!String.IsNullOrEmpty(this.OutputFolder));
                    
                    // return value
                    return hasOutputFolder;
                }
            }
            #endregion
            
            #region HasWorksheet
            /// <summary>
            /// This property returns true if this object has a 'Worksheet'.
            /// </summary>
            public bool HasWorksheet
            {
                get
                {
                    // initial value
                    bool hasWorksheet = (this.Worksheet != null);
                    
                    // return value
                    return hasWorksheet;
                }
            }
            #endregion
            
            #region IsValid
            /// <summary>
            /// This read only property returns the value for 'IsValid'.
            /// </summary>
            public bool IsValid
            {
                get
                {
                    // initial value
                    bool isValid = false;

                    // valid if all exist
                    bool hasAllRequiredProperties = (HasClassName) && (HasOutputFolder) && (HasWorksheet);

                    // if the value for hasAllRequiredProperties is true
                    if (hasAllRequiredProperties)
                    {
                        // If the value for the property worksheet.HasRows is true
                        if (worksheet.HasRows)
                        {
                            // valid if there are one or more rows (the top row is used for the header)                            
                            isValid = ListHelper.HasOneOrMoreItems(worksheet.Rows);
                        }
                    }

                    // return value
                    return isValid;
                }
            }
            #endregion
            
            #region OutputFolder
            /// <summary>
            /// This property gets or sets the value for 'OutputFolder'.
            /// </summary>
            public string OutputFolder
            {
                get { return outputFolder; }
                set { outputFolder = value; }
            }
            #endregion
            
            #region Worksheet
            /// <summary>
            /// This property gets or sets the value for 'Worksheet'.
            /// </summary>
            public Worksheet Worksheet
            {
                get { return worksheet; }
                set { worksheet = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
