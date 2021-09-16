# DataJuggler.Excelerate
Excelerate uses EPPPlus version 4.5.3.3 (last free version), and it makes it easy to load Workbooks or Worksheets.

Nuget package version 1.1.0 was just published: DataJuggler.Excelerate

A class named CodeGenerator was just created, and now using the same CSharpClassWriter that code generates for DataTier.Net, I code generate
classes based on your header row.

I have a couple of clients that I build programs that automate combining columns from multiple Worksheets to form reports.

Rather than continue to write custom loaders, I really only need custom Exporters in most cases.

Here is a short video:
https://youtu.be/Sa-xroxPw_I

This short code snippet will load all the rows from a worksheet:

Snippet is from a Windows Form .Net 5 project, located in the Sample folder of this project. Very simple for now:

    using DataJuggler.UltimateHelper;
    using DataJuggler.Excelerate;
    using System;
    using System.Windows.Forms;

    // Set the text
    string path = WorksheetControl.Text;

    // Create a new instance of a 'LoadWorksheetInfo' object.
    LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

    // Set the SheetName
    oadWorksheetInfo.SheetName = SheetnameControl.SelectedObject.ToString();

    // Only load the first 12 columns for this test
    loadWorksheetInfo.ColumnsToLoad = 12;

    // Set the LoadColumnOptions
    loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadFirstXColumns;
    
    // other options
    // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
    // loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadSpecifiedColumns;

    // load the worksheet
    Worksheet worksheet = ExcelDataLoader.LoadWorksheet(path, loadWorksheetInfo);

    // if the worksheet exists
    if ((NullHelper.Exists(worksheet)) && (SheetnameControl.HasSelectedObject))
    {
        // if the rows collection was found
        if (worksheet.HasRows)
        {
            // Show a message as a test
            // MessageBox.Show("Worksheet Loaded", "Finished");

            // test only
            // int rows = worksheet.Rows.Count;

            // Show a message as a test
            // MessageBox.Show("There were " + String.Format("{0:n0}",  rows) + " rows found in the worksheet");

            // int cols = worksheet.Rows[1124].Columns.Count;

            // Show a message as a test
            // MessageBox.Show("There were " + String.Format("{0:n0}",  cols) + " columns found in the row index 1125.");

            // Get a nullable date
            // string columnValue = worksheet.Rows[1124].Columns[3].DateValue;

            // Show a message of the columnValue
            // MessageBox.Show("Column Value: " + columnValue);
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

I am now working on building a loader for the code generated classes to convert the data.

** I am available for hire if you need with any size C# project **

