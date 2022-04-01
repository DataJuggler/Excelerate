

#region using statements

using DataJuggler.UltimateHelper;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using DataJuggler.Win.Controls.Interfaces;
using OfficeOpenXml;
using System.IO;

#endregion

namespace DataJuggler.Excelerate.Sample
{

    #region class MainForm
    /// <summary>
    /// This is the MainForm for this app.
    /// </summary>
    public partial class MainForm : Form, ITextChanged
    {
        
        #region Private Variables
        private Worksheet worksheet;
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of a 'MainForm' object.
        /// </summary>
        public MainForm()
        {
            // Create Controls
            InitializeComponent();

            // Perform initializations for this object
            Init();
        }
        #endregion

        #region Events

            #region CodeGenerateButton_Click(object sender, EventArgs e)
            /// <summary>
            /// event is fired when the 'CodeGenerateButton' is clicked.
            /// </summary>
            private void CodeGenerateButton_Click(object sender, EventArgs e)
            {
                // Remove focus from the button just clicked
                OffScreenButton.Focus();

                // if the value for HasWorksheet is true
                if ((HasWorksheet) && (ListHelper.HasOneOrMoreItems(Worksheet.Rows)))
                {
                    // The file I am using to test has 3 rows at the top. Take this out if I accidently check this in
                    // worksheet.Rows.RemoveRange(0, 3);

                    // Set the outputFolder
                    string outputFolder = OutputFolderControl.Text;

                    // Set the className (as a test)
                    string className = "SalesTaxEntry";

                    // Create a new instance of a CodeGenerator
                    CodeGenerator codeGenerator = new CodeGenerator(worksheet, outputFolder, className);

                    // Generate a class and set the Namespace
                    CodeGenerationResponse response = codeGenerator.GenerateClassFromWorksheet("STATS.Objects");

                    // Show the results
                    MessageBox.Show("Success: " + response.Success);
                }
            }
            #endregion

            #region LoadWorksheetButton_Click(object sender, EventArgs e)
            /// <summary>
            /// event is fired when the 'TestButton' is clicked.
            /// </summary>
            private void LoadWorksheetButton_Click(object sender, EventArgs e)
            {
                // Remove focus from the button just clicked
                OffScreenButton.Focus();

                // Set the text
                string path = WorksheetControl.Text;

                // Create a new instance of a 'LoadWorksheetInfo' object.
                LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                // Set the SheetName
                loadWorksheetInfo.SheetName = SheetnameControl.SelectedObject.ToString();

                // Only load the first 12 columns for this test
                loadWorksheetInfo.ColumnsToLoad = 12;

                // Set the LoadColumnOptions
                loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadFirstXColumns;

                // other options. ExcludedColumns is a coming soon feature.
                // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
                // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadSpecifiedColumns;

                // load the worksheet
                Worksheet worksheet = ExcelDataLoader.LoadWorksheet(path, loadWorksheetInfo);

                // if the workbook exists
                if ((NullHelper.Exists(worksheet)) && (SheetnameControl.HasSelectedObject))
                {
                    // Store the property on this form
                    this.Worksheet = worksheet;

                    // If the Worksheet exists, the Code Generate Button exists
                    CodeGenerateButton.Enabled = (HasWorksheet && (OutputFolderControl.HasText));

                    // if the rows collection was found
                    if (worksheet.HasRows)
                    {
                        // Show a message as a test
                        MessageBox.Show("Worksheet Loaded", "Finished");

                        //// test only
                        //int rows = worksheet.Rows.Count;

                        //// Show a message as a test
                        //MessageBox.Show("There were " + String.Format("{0:n0}",  rows) + " rows found in the worksheet");

                        //int cols = worksheet.Rows[1124].Columns.Count;

                        //// Show a message as a test
                        //MessageBox.Show("There were " + String.Format("{0:n0}",  cols) + " columns found in the row index 1125.");

                        // // Get a nullable date
                        // string columnValue = worksheet.Rows[1124].Columns[3].DateValue;

                        //// Show a message with the columnValue
                        //MessageBox.Show("Column Value: " + columnValue);
                    }
                }
            }
            #endregion
            
            #region OnTextChanged(Control sender, string text)
            /// <summary>
            /// event is fired when On Text Changed
            /// </summary>
            public void OnTextChanged(Control sender, string text)
            {
                // here we must lookup the first sheet name, so I don't put my clients name
                // in this Git Hub repo.

                LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                // Load all columns
                loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;

                // local
                string firstSheetName = "";

                // create a workbook from the path to look up the first sheet name
                using (ExcelWorkbook workbook = ExcelDataLoader.LoadExcelWorkbook(text))
                {
                    // if the workbook exists
                    if ((NullHelper.Exists(workbook)) && (workbook.Worksheets.Count > 0))
                    {
                        // Create a new collection of 'string' objects.
                        List<string> worksheetNames = new List<string>();

                        // Set the firstSheetName
                        firstSheetName = workbook.Worksheets[0].Name;

                        // iterate worksheets
                        for (int x = 0; x < workbook.Worksheets.Count; x++)
                        {
                            // Add this string
                            worksheetNames.Add(workbook.Worksheets[x].Name);
                        }

                        // Load the list
                        SheetnameControl.LoadItems(worksheetNames);

                        // Select the first item
                        SheetnameControl.SelectedIndex = SheetnameControl.FindItemIndexByValue(firstSheetName);
                    }
                }
            }
            #endregion

        #endregion

        #region Methods

            #region Init()
            /// <summary>
            /// This method performs initializations for this object.
            /// </summary>
            public void Init()
            {
                // Setup the listener
                this.WorksheetControl.OnTextChangedListener = this;
            }
        #endregion

        #endregion

        #region Properties
            
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
