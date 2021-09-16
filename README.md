# DataJuggler.Excelerate
Excelerate uses EPPPlus version 4.5.3.3 (last free version), and it makes it easy to load Workbooks or Worksheets.

Nuget package version 1.1.0 was just published: DataJuggler.Excelerate

A class named CodeGenerator was just assed, and now using the same CSharpClassWriter that code generates for DataTier.Net, I code generate
classes based on your header row.

I have a couple of clients that I build programs that automate combining columns from multiple Worksheets to form reports.

Rather than continue to write custom loaders, I really only need custom Exporters in most cases.

This short code snippet will load all the rows from a worksheet:

Snippet is from a Windows Form .Net 5 project, located in the Sample folder of this project. Very simple for now:

    using DataJuggler.UltimateHelper;
    using DataJuggler.Excelerate;
    using System;
    using System.Windows.Forms;

    // get the path
    string path = WorksheetControl.Text;

    // Create a new instance of a 'LoadWorksheetInfo' object.
    LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();
    
    // Set the SheetName
    loadWorksheetInfo.SheetName = SheetNameControl.Text;

    // Only load the first 12 columns for this test
    loadWorksheetInfo.ColumnsToLoad = 12;

    // Set the LoadColumnOptions
    loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadFirstXColumns;
    
    // Other options
    // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
    // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadSpecifiedColumns;

    // load the workbook
    Workbook workbook = ExcelDataLoader.LoadWorkbook(path, loadWorksheetInfo);

    // if the workbook exists
    if ((NullHelper.Exists(workbook)) && (ListHelper.HasOneOrMoreItems(workbook.Worksheets)))
    {
        // get the index
        int index = workbook.GetWorksheetIndex(worksheetName);

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

                int cols = worksheet.Rows[1125].Columns.Count;]
    
                // Show a message if it works
                MessageBox.Show("There were " + String.Format("{0:n0}",  cols) + " columns found in the row index 1125.");
                
                // get values, code to verify rowNumbers and columnNumber are omitted for brevity. 
                // Always test for Rows.Count and Columns.Count in a real project.
                
                // Get the ColumnValue cast as a Decimal
                Decimal columnValue = worksheet.Rows[1125].Columns[4].DecimalValue;
                
                // Cast the column value as a string
                string temp = worksheet.Rows[x].Columns[y].StringValue;
                
                // Cast the ColumnValue as a bool value                
                bool active = worksheet.Rows[x].Columns[y].BoolValue;
                
                // Cast the column value as a nullable date
                DateTime? expirationDate = worksheet.Rows[x].Columns[y].DateValue;
            }
        }
    }
    
    There is another override to load multiple sheets at once. I will build a sample project when I get some time to build a sample spreadsheet I can give away.
    
    To load multiple sheets:
    
    List<LoadWorksheetInfo> loadWorkSheetsInfo = new List<LoadWorksheetInfo>();
    
    // Add each LoadWorksheetInfo
    workbook= ExcellDataLoader.LoadWorkbook(path, loadWorkSheetsInfo)
    
    I will build some helper methods to save writing as much code once I use this a little to know what is needed.
    
    
    
    
I am just starting testing now. My first test loaded a 12 column spreadsheet with 3,376 rows in just a few seconds.

More helper methods and features will be added. I will release a Nuget package once I finish my project I built this for.
Give me a day or two and the Nuget should be released as DataJuggler.Excelerate.

Feel free to mention any new features you think would be useful. I can't promise to do them all, but if it is a good fit for this project I will add it.

This code is all brand new, so use with caution until more testing has been done. First tests were promising.

