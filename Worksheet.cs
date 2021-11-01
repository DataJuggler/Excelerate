

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

    #region class Worksheet
    /// <summary>
    /// This represents a Worksheet in an Excel Workbook
    /// </summary>
    public class Worksheet
    {
        
        #region Private Variables
        private List<Column> columns;
        private List<Row> rows;
        private string name;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'Worksheet' object.
        /// </summary>
        public Worksheet()
        {
            // Create both lists
            Columns = new List<Column>();
            Rows = new List<Row>();
        }
        #endregion

        #region Methods

            #region ClearColumnValues()
            /// <summary>
            /// Clear Column Values
            /// </summary>
            public void ClearColumnValues()
            {
                // if the value for HasColumns is true
                if (HasColumns)
                {
                    // Iterate the collection of Column objects
                    foreach (Column column in columns)
                    {
                        switch (column.DataType)
                        {
                            case DataManager.DataTypeEnum.Integer:
                            case DataManager.DataTypeEnum.Double:
                            case DataManager.DataTypeEnum.Decimal:
                            case DataManager.DataTypeEnum.Percentage:
                            
                            // Not really used here, but doesn't hurt for future
                            case DataManager.DataTypeEnum.Autonumber:

                                // erase
                                column.ColumnValue = 0;

                                // required
                                break;

                            case DataManager.DataTypeEnum.String:

                                // erase
                                column.ColumnValue = "";

                                // required
                                break;

                            case DataManager.DataTypeEnum.Guid:

                                // erase
                                column.ColumnValue = Guid.Empty;

                                // required
                                break;

                            case DataManager.DataTypeEnum.DateTime:

                                // erase (sort of)
                                column.ColumnValue = new DateTime(1, 1, 1900);

                                // required
                                break;

                            case DataManager.DataTypeEnum.Boolean:
                            case DataManager.DataTypeEnum.YesNo:

                                // set to false for erase
                                column.ColumnValue = false;
                                
                                // required
                                break;
                        }
                    }
                }
            }
            #endregion
            
            #region CreateNewRowColumns(int rowNumber)
            /// <summary>
            /// returns a list of New Row Columns
            /// </summary>
            public List<Column> CreateNewRowColumns(int rowNumber)
            {
                // initial value
                List<Column> columns = new List<Column>();

                // if there are one or more columns
                if (ListHelper.HasOneOrMoreItems(this.Columns))
                {
                    // iterate the columns
                    foreach (Column column in this.Columns)
                    {
                        // Create
                        Column newColumn = new Column();

                        // Set the DataType
                        newColumn.DataType = column.DataType;

                        // Set the Name & ColumnNumber & RowNumber
                        newColumn.ColumnName = column.ColumnName;
                        newColumn.ColumnNumber = column.ColumnNumber;
                        newColumn.RowNumber = rowNumber;

                        // Add this column
                        columns.Add(newColumn);
                    }
                }
                
                // return value
                return columns;
            }
            #endregion
            
            #region NewRow()
            /// <summary>
            /// This method creates a new row, and sets the columns to this objects columns
            /// </summary>
            public Row NewRow()
            {
                // Create a new instance of a 'Row' object.
                Row row = new Row();

                // Create a new Guid
                row.Id = Guid.NewGuid();

                // Set the number
                row.Number = Rows.Count + 1;

                // if the Columns exist
                if (HasColumns)
                {
                    // Use these columns
                    row.Columns = CreateNewRowColumns(row.Number);
                }
                else
                {   
                    // Create a new collection of 'Column' objects.
                    row.Columns = new List<Column>();
                }

                // return value
                return row;
            }
            #endregion
            
        #endregion

        #region Properties

        #region Columns
        /// <summary>
        /// This property gets or sets the value for 'Columns'.
        /// Worksheet.Rows has the actual Data. The Columns collection
        /// here is for a 'HeaderRow' and contains information about
        /// column data types.
        /// </summary>
        public List<Column> Columns
            {
                get { return columns; }
                set { columns = value; }
            }
            #endregion
            
            #region HasColumns
            /// <summary>
            /// This property returns true if this object has a 'Columns'.
            /// </summary>
            public bool HasColumns
            {
                get
                {
                    // initial value
                    bool hasColumns = (this.Columns != null);
                    
                    // return value
                    return hasColumns;
                }
            }
            #endregion
            
            #region HasRows
            /// <summary>
            /// This property returns true if this object has a 'Rows'.
            /// </summary>
            public bool HasRows
            {
                get
                {
                    // initial value
                    bool hasRows = (this.Rows != null);
                    
                    // return value
                    return hasRows;
                }
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
            
            #region Rows
            /// <summary>
            /// This property gets or sets the value for 'Rows'.
            /// </summary>
            public List<Row> Rows
            {
                get { return rows; }
                set { rows = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
