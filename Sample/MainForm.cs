

#region using statements

using DataJuggler.UltimateHelper;
using System;
using System.Windows.Forms;
using DataJuggler.Win.Controls.Interfaces;
using OfficeOpenXml;

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

                // create a workbook from the path to look up the first sheet name
                using (ExcelWorkbook workbook = ExcelDataLoader.LoadExcelWorkbook(text))
                {
                    // if the workbook exists
                    if ((NullHelper.Exists(workbook)) && (workbook.Worksheets.Count > 0))
                    {
                        // Get the worksheetName
                        string workSheetName = workbook.Worksheets[0].Name;

                        // Display the name
                        SheetNameControl.Text = workSheetName;
                    }
                }
            }
            #endregion
            
            #region TestButton_Click(object sender, EventArgs e)
            /// <summary>
            /// event is fired when the 'TestButton' is clicked.
            /// </summary>
            private void TestButton_Click(object sender, EventArgs e)
            {
                // Set the text
                string path = WorksheetControl.Text;

                // Create a new instance of a 'LoadWorksheetInfo' object.
                LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                // Set the SheetName
                loadWorksheetInfo.SheetName = SheetNameControl.Text;

                // Only load the first 12 columns for this test
                loadWorksheetInfo.ColumnsToLoad = 12;

                // Set the LoadColumnOptions
                loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadFirstXColumns;
                // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
                // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadSpecifiedColumns;

                // load the workbook
                Workbook workbook = ExcelDataLoader.LoadWorkbook(path, loadWorksheetInfo);

                // if the workbook exists
                if ((NullHelper.Exists(workbook)) && (ListHelper.HasOneOrMoreItems(workbook.Worksheets)))
                {
                    // get the index
                    int index = workbook.GetWorksheetIndex(SheetNameControl.Text);

                    // if the index was found
                    if (index >= 0)
                    {
                        // set the worksheet
                        Worksheet worksheet = workbook.Worksheets[index];

                        // if the rows collection was found
                        if (worksheet.HasRows)
                        {
                            // test only
                            int rows = worksheet.Rows.Count;

                            // Show a message if it works
                            MessageBox.Show("There were " + String.Format("{0:n0}",  rows) + " rows found in the worksheet");

                            int cols = worksheet.Rows[1124].Columns.Count;

                            // Show a message if it works
                            MessageBox.Show("There were " + String.Format("{0:n0}",  cols) + " columns found in the row index 1125.");

                            // Get a nullable date
                            string columnValue = worksheet.Rows[1124].Columns[3].StringValue;

                            // Show a message of the columnValue
                            MessageBox.Show("Column Value: " + columnValue);
                        }
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

    }
    #endregion

}
