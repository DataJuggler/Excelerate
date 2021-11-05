

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataJuggler.Net5;
using DataJuggler.UltimateHelper;
using DataJuggler.UltimateHelper.Objects;
using System.IO;

#endregion

namespace DataJuggler.Excelerate
{

    #region class CodeGenerator
    /// <summary>
    /// This class is used to code generate classes based on the columns from an Excel worksheet
    /// </summary>
    public class CodeGenerator : CSharpClassWriter
    {
        
        #region Private Variables
        private Worksheet worksheet;
        private string outputFolder;
        private string className;
        private const string RowId = "RowId";
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

            #region AddLoadMethod(Row row, ref StringBuilder sb)
            /// <summary>
            /// This method Add Load Method
            /// </summary>            
            public void AddLoadMethod(Row row, ref StringBuilder sb)
            {
                 // locals
                int columnIndex = -1;
                string indent = "            ";
                string indent2 = "                ";
                string indent3 = "                    ";
                columnIndex = -1;
                
                 // Add a blank line
                sb.Append(Environment.NewLine);

                // Add a region
                sb.Append(indent);
                sb.Append("#region Load(Row row)");
                sb.Append(Environment.NewLine);

                // Now add the method summary

                // Summary Line 1
                sb.Append(indent);
                sb.Append("/// <summary>");
                sb.Append(Environment.NewLine);

                // Summary Line 2
                sb.Append(indent);
                sb.Append("/// This method loads a " + ClassName + " object from a Row.");
                sb.Append(Environment.NewLine);

                // Summary Line 3
                sb.Append(indent);
                sb.Append("/// </Summary>");
                sb.Append(Environment.NewLine);

                // Add the param comment
                sb.Append(indent);
                sb.Append("/// <param name=\"row\">The row which the row.Columns[x].ColumnValue will be used to load this object.</param>");
                sb.Append(Environment.NewLine);

                // Now add the indent
                sb.Append(indent);

                // Set the methodDeclarationLine
                string methodDeclarationLine = "public void Load(Row row)";

                // Add this line
                sb.Append(methodDeclarationLine);

                // Add a new line
                sb.Append(Environment.NewLine);

                // Now add the method
                sb.Append(indent);

                // Add an open bracket
                sb.Append('{');

                // Add a new line
                sb.Append(Environment.NewLine);

                // Add a comment
                sb.Append(indent2);

                // Add this
                sb.Append("// If the row exists and the row's column collection exists");
                sb.Append(Environment.NewLine);

                // Add a check for the column
                sb.Append(indent2);

                // create the ifLine
                sb.Append("if ((NullHelper.Exists(row)) && (row.HasColumns))");

                // Add a new line here before the paren
                sb.Append(Environment.NewLine);

                // Add an open paren
                sb.Append(indent2);
                sb.Append('{');

                // Add a new line
                sb.Append(Environment.NewLine);
               
                // Create DataFields for each column
                foreach (Column column in row.Columns)
                {
                    // if the ColumnName Exists
                    if ((column.HasColumnName) && (column.ColumnName != RowId))
                    {
                        // Increment the value for columnIndex
                        columnIndex++;

                        // Now add the indent3 (8 spaces extra)
                        sb.Append(indent3);

                        // Set the Column Name (Property Name)
                        sb.Append(column.ColumnName);

                        // Set Equals
                        sb.Append(" = ");

                        // if Decimal, must cast as a double
                        if (column.DataType == DataManager.DataTypeEnum.Decimal)
                        {
                            // Cast as a double
                            sb.Append("(double) ");
                        }

                        // add the start of this column
                        sb.Append("row.Columns[");
                                                        
                        // add the index
                        sb.Append(columnIndex);

                        // determine the action by the DataType
                        switch (column.DataType)
                        {
                            case DataManager.DataTypeEnum.Integer:

                                // Set the value
                                sb.Append("].IntValue;");

                                // required
                                break;

                            case DataManager.DataTypeEnum.Decimal:

                                // Set the value
                                sb.Append("].DecimalValue;");

                                // required
                                break;

                            case DataManager.DataTypeEnum.DateTime:

                                // Set the value
                                sb.Append("].DateValue;");

                                // required
                                break;

                            case DataManager.DataTypeEnum.String:

                                // Set the value
                                sb.Append("].StringValue;");

                                // required
                                break;

                            case DataManager.DataTypeEnum.Boolean:

                                // Set the value
                                sb.Append("].BoolValue;");

                                // required
                                break;

                            case DataManager.DataTypeEnum.Guid:

                                // Set the value
                                sb.Append("].GuidValue;");

                                // required
                                break;

                            default:

                                // Set the value
                                sb.Append("].ColumnValue;");

                                // required
                                break;
                        }

                        // Add a new line
                        sb.Append(Environment.NewLine);
                    }
                }

                // Add a closing bracket
                sb.Append(indent2);
                sb.Append('}');
                sb.Append(Environment.NewLine);

                // Add an extra blank line
                sb.Append(Environment.NewLine);

                // Add a comment for the RowId
                sb.Append(indent2);
                sb.Append("// Set RowId");
                sb.Append(Environment.NewLine);

                // add RowId
                sb.Append(indent2);
                sb.Append("RowId = row.Id;");
                sb.Append(Environment.NewLine);

                // Add indent
                sb.Append(indent);

                // Add a closing bracket
                sb.Append('}');

                // Add a new line
                sb.Append(Environment.NewLine);

                // Add the endregion
                sb.Append(indent);
                sb.Append("#endregion");
                sb.Append(Environment.NewLine);
            }
            #endregion

            #region AddSaveMethod(Row row, ref StringBuilder sb)
            /// <summary>
            /// This method Adds a Save Method
            /// </summary>            
            public void AddSaveMethod(Row row, ref StringBuilder sb)
            {
                 // locals
                int columnIndex = -1;
                string indent = "            ";
                string indent2 = "                ";
                string indent3 = "                    ";                 
                columnIndex = -1;
                
                 // Add a blank line
                sb.Append(Environment.NewLine);

                // Add a region
                sb.Append(indent);
                sb.Append("#region Save(Row row)");
                sb.Append(Environment.NewLine);

                // Now add the method summary

                // Summary Line 1
                sb.Append(indent);
                sb.Append("/// <summary>");
                sb.Append(Environment.NewLine);

                // Summary Line 2
                sb.Append(indent);
                sb.Append("/// This method saves a " + ClassName + " object back to a Row.");
                sb.Append(Environment.NewLine);

                // Summary Line 3
                sb.Append(indent);
                sb.Append("/// </Summary>");
                sb.Append(Environment.NewLine);

                // Add the param comment
                sb.Append(indent);
                sb.Append("/// <param name=\"row\">The row which the row.Columns[x].ColumnValue will be set to Save back to Excel.</param>");
                sb.Append(Environment.NewLine);

                // Now add the indent
                sb.Append(indent);

                // Set the methodDeclarationLine
                string methodDeclarationLine = "public Row Save(Row row)";

                // Add this line
                sb.Append(methodDeclarationLine);

                // Add a new line
                sb.Append(Environment.NewLine);

                // Now add the method
                sb.Append(indent);

                // Add an open bracket
                sb.Append('{');

                // Add a new line
                sb.Append(Environment.NewLine);

                // Add a comment
                sb.Append(indent2);

                // Add this
                sb.Append("// If the row exists and the row's column collection exists");
                sb.Append(Environment.NewLine);

                // Add a check for the column
                sb.Append(indent2);

                // create the ifLine
                sb.Append("if ((NullHelper.Exists(row)) && (row.HasColumns))");

                // Add a new line here before the paren
                sb.Append(Environment.NewLine);

                // Add an open paren
                sb.Append(indent2);
                sb.Append('{');

                // Add a new line
                sb.Append(Environment.NewLine);

                // Create DataFields for each column
                foreach (Column column in row.Columns)
                {
                    // if the ColumnName Exists
                    if ((column.HasColumnName) && (column.ColumnName != RowId))
                    {
                        // Increment the value for columnIndex
                        columnIndex++;

                        // Now add the indent3 (8 spaces extra)
                        sb.Append(indent3);

                        // Set the columnValue
                        sb.Append("row.Columns[");
                        
                        // add the columnIndex
                        sb.Append(columnIndex);

                        // Append '].ColumnValue = '
                        sb.Append("].ColumnValue = ");

                        // Set the Column Name (Property Name)
                        sb.Append(column.ColumnName);

                        // Append closing semicolon
                        sb.Append(";");

                        // Add a new line
                        sb.Append(Environment.NewLine);
                    }
                }

                // Add a closing bracket
                sb.Append(indent2);
                sb.Append('}');
                sb.Append(Environment.NewLine);

                // Add a new line
                sb.Append(Environment.NewLine);

                // write a comment
                sb.Append(indent2);
                sb.Append("// return value");
                sb.Append(Environment.NewLine);

                // write the return value row
                sb.Append(indent2);
                sb.Append("return row;");
                sb.Append(Environment.NewLine);

                // Add indent
                sb.Append(indent);

                // Add a closing bracket
                sb.Append('}');

                // Add a new line
                sb.Append(Environment.NewLine);

                // Add the endregion
                sb.Append(indent);
                sb.Append("#endregion");
                sb.Append(Environment.NewLine);
            }
            #endregion
            
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
                if ((NullHelper.Exists(worksheet)) && (ListHelper.HasOneOrMoreItems(worksheet.Rows)) && (TextHelper.Exists(fieldName)))
                {
                    // get a lowercase version of the fieldName
                    fieldName = fieldName.ToLower();

                    for (int x = 1; x < worksheet.Rows.Count; x++)
                    {
                        // get the value in this position
                        temp = worksheet.Rows[x].Columns[columnIndex].ColumnText;

                        // If the temp string exists, and this is not a 0, 0's are hard to tell anything from
                        if ((temp != "0") && (temp != "0.00"))
                        {
                            // get the values
                            tempInt = worksheet.Rows[x].Columns[columnIndex].IntValue;
                            tempDecimal = worksheet.Rows[x].Columns[columnIndex].DecimalValue;
                            tempDate = worksheet.Rows[x].Columns[columnIndex].DateValue;
                            
                            // if true or false
                            if ((temp.ToLower() == "true") || (temp.ToLower() == "false"))
                            {
                                // this is a boolean
                                dataType = DataManager.DataTypeEnum.Boolean;

                                // break
                                break;
                            }
                            else if (fieldName == "active")
                            {
                                // hard coding Active as boolean, because I need it for the Demo and Active usually is a boolean

                                // this is a boolean
                                dataType = DataManager.DataTypeEnum.Boolean;

                                // break
                                break;
                            }
                            else if ((fieldName == "zip") || (fieldName == "zipcode") || (fieldName == "postal") || (fieldName == "postalcode"))
                            {
                                // this is a string, not an int
                                dataType = DataManager.DataTypeEnum.String;

                                // break out
                                break;
                            }
                            else
                            {
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

                                    // exit loop
                                    break;
                                }
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
            
            #region GenerateClassFromWorksheet(string namespaceName, bool appendPartialGuid = true)
            /// <summary>
            /// This method returns a Class From the Worksheet supplied in the Constructor.
            /// The Worksheet must have a HeaderRow for the top row.
            /// </summary>
            public CodeGenerationResponse GenerateClassFromWorksheet(string namespaceName, bool appendPartialGuidToFileNameForUniquenessInFolder = true)
            {
                // initial value
                CodeGenerationResponse response = new CodeGenerationResponse();

                // locals
                int columnIndex = -1;
                int lineNumber = 0;
                
                // if the value for IsValid is true (means there is a worksheet and it has at least one row)
                if (IsValid)
                {
                    // Get the first row
                    Row row = worksheet.Rows[0];

                    // if the row exists
                    if (ListHelper.HasOneOrMoreItems(row.Columns))
                    {
                        // Update 10.31.2021: Creating a column to hold the RowId
                        Column rowIdColumn = new Column();

                        // Set the rowId
                        rowIdColumn.ColumnName = RowId;

                        // Set the DataType
                        rowIdColumn.DataType = DataManager.DataTypeEnum.Guid;

                        // Extra column
                        rowIdColumn.ColumnNumber = row.Columns.Count + 1;

                        // Add this rowIdColumn
                        row.Columns.Add(rowIdColumn);

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
                            // if the ColumnValue exists or if this is the RowId field
                            if ((column.HasColumnValue) || (TextHelper.IsEqual(column.ColumnName, RowId)))
                            {
                                // Create a field for this column
                                DataField field = new DataField();

                                // if this is not the RowId
                                if (!TextHelper.IsEqual(column.ColumnName, RowId))
                                {
                                    // Increment the value for columnIndex
                                    columnIndex++;

                                    // Store the orininalName so it can be used during Export.
                                    column.OriginalName = column.StringValue;

                                    // Set the name, but replace out things that make it an illegal field name like spaces or dashes
                                    field.FieldName = TextHelper.CapitalizeFirstChar(ReplaceInvalidCharacters(column.StringValue));

                                    // Set the ColumnName in the Column
                                    column.ColumnName = field.FieldName;
                                }
                                else
                                {
                                    // Store
                                    field.FieldName = column.ColumnName;

                                    // Set the OriginalName (not sure if this is needed, more for if needed)
                                    column.OriginalName = field.FieldName;
                                }

                                // if this is the RowId
                                if (TextHelper.IsEqual(field.FieldName, RowId))
                                {
                                    // If this is RowId
                                    field.DataType = DataManager.DataTypeEnum.Guid;
                                }
                                else
                                {
                                    // DetermineDataType
                                    field.DataType = AttemptToDetermineDataType(field.FieldName, columnIndex);

                                    // Store the DataType in the column, so the Loader knows how to handle this column
                                    column.DataType = field.DataType;
                                }

                                // Set the FieldOrdinal
                                field.FieldOrdinal = columnIndex;

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

                        // Create a referencesSet
                        ReferencesSet referencesSet = new ReferencesSet("References");

                        // Create a couple references
                        Reference reference = new Reference("DataJuggler.Excelerate", 1);
                        Reference reference2 = new Reference("DataJuggler.UltimateHelper", 2);
                        Reference reference3 = new Reference("System", 3);

                        // Add the references to the ReferencesSet
                        referencesSet.Add(reference);
                        referencesSet.Add(reference2);
                        referencesSet.Add(reference3);

                        // Set the references
                        dataManager.References = referencesSet;

                        // Write out the class
                        response.Success = WriteDataClasses(dataManager);

                        // if the success = true, and one or more files were created during the build
                        if ((response.Success) && (ListHelper.HasOneOrMoreItems(CreatedFilePaths)))
                        {
                            // if exactly one file was created (should be, since the build here is done one at a time for now)
                            if (CreatedFilePaths.Count == 1)
                            {
                                // Get the filePath of the class just created
                                string filePath = CreatedFilePaths[0];

                                // Get the text of the file
                                string content = File.ReadAllText(filePath);

                                // create a stringbuilder to rebuild this file
                                StringBuilder sb = new StringBuilder();

                                // If the content string exists
                                if (TextHelper.Exists(content))
                                {
                                    // parse the fileText
                                    List<TextLine> lines = TextHelper.GetTextLines(content);

                                    // if the lines exist
                                    if (ListHelper.HasOneOrMoreItems(lines))
                                    {
                                        // need to skip one line after the method is created
                                        bool skipNextLine = false;
                                        bool skipNextLineIfBlank = false;
    
                                        // Iterate the collection of TextLine objects
                                        foreach (TextLine line in lines)
                                        {
                                            // increment lineNumber
                                            lineNumber++;

                                            // if skipNextLineIfBlank
                                            if ((skipNextLineIfBlank) && (!TextHelper.Exists(line.Text)) && (lineNumber > 2))
                                            {
                                                // skip this line
                                                skipNextLine = true;
                                            }

                                            // if the value for skipNextLine is false
                                            if (!skipNextLine)
                                            {
                                                // Add this line
                                                sb.Append(line.Text);
                                                sb.Append(Environment.NewLine);

                                                // reset
                                                skipNextLineIfBlank = false;
                                            }
                                            else if (!TextHelper.Exists(line.Text))
                                            {
                                                // if currently true
                                                if (skipNextLine)
                                                {
                                                    // skip thius line
                                                    skipNextLine = false;
                                                }
                                                else
                                                {
                                                    // Set to true
                                                    skipNextLineIfBlank = true;
                                                }
                                            }

                                            // if this is the Methods line
                                            if (TextHelper.IsEqual(line.Text.Trim(), "#region Methods"))
                                            {
                                                // Pass in the string builder here, saves a bunch of code in this method
                                               AddLoadMethod(row, ref sb);

                                               // Add the Save method
                                               AddSaveMethod(row, ref sb);
                                            }

                                            // if this line is blank
                                            if (!TextHelper.Exists(line.Text))
                                            {
                                                // Skip this line
                                                skipNextLineIfBlank = true;
                                            }
                                        }

                                        // Now the fileContent has been rebuilt in the string builder
                                        string newFileText = sb.ToString().TrimEnd();

                                        // if true
                                        if (appendPartialGuidToFileNameForUniquenessInFolder)
                                        {
                                            // Create a PartialGuid
                                            filePath = FileHelper.CreateFileNameWithPartialGuid(filePath, 12);
                                        }

                                        // Now write the new file text
                                        File.WriteAllText(filePath, newFileText);

                                        // Create a FileInfo
                                        FileInfo fileInfo = new FileInfo(filePath);

                                        // Set the FileName                                        
                                        response.FileName = fileInfo.Name;

                                        // Set the fullPath
                                        response.FullPath = filePath;
                                    }
                                }
                            }
                        }
                    }
                }
                
                // return value
                return response;
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
