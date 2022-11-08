<img height=192 width=192 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogoSmallWhite.png>

Live Demo

Blazor.Excelerate
https://excelerate.datajuggler.com 
Code Generate C# Classes From Excel Header Rows

# DataJuggler.Excelerate

# Update 11.8.2022:
I added some new properties to the row and Column object for use with the Grid for DataJuggler.Blazor.Components.

# Update 10.31.2022:
LoadWorksheetInfo.ExcludedColumnIndexes was added. This is a collection of integers
to not load. I may expand this to column names also as an option.

--

Excelerate uses EPPPlus version 4.5.3.3 (last free version), and it makes it easy to load Workbooks or Worksheets.

A class named CodeGenerator was just created, and now by inheriting from the same CSharpClassWriter that code generates for DataTier.Net, I code generate
classes based on your header row.

I have a couple of clients that I build programs that automate combining columns from multiple Worksheets to form reports.

Rather than continue to write custom loaders, I really only need custom Exporters in most cases.

Here is a short video:
https://youtu.be/Sa-xroxPw_I

This short code snippet will load all the rows from a worksheet:

Snippet is from a Windows Form .Net 6 project, located in the Sample folder of this project. Very simple for now:

# Load Worksheet Sample

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
    
    There is now a Code Generator class built into this project, to code generate a C# class from a header row. 
    The Code Generator has been updated to pass in a Row instance, to make loading the generate classes simple.
    
    This code is from a Windows Form .Net 5 project located in the sample:
    
    # Code Generation Sample
    
        // if the value for HasWorksheet is true
        if ((HasWorksheet) && (ListHelper.HasOneOrMoreItems(Worksheet.Rows)))
        {
            // The file I am using to test has 3 rows at the top above the header row. Take this out if I accidently check this in
            // worksheet.Rows.RemoveRange(0, 3);

            // Set the outputFolder
            string outputFolder = OutputFolderControl.Text;

            // Set the className (the name of the generated class)
            string className = "SalesTaxEntry";

            // Create a new instance of a CodeGenerator
            CodeGenerator codeGenerator = new CodeGenerator(worksheet, outputFolder, className);

            // Generate a class and set the Namespace
            bool success = codeGenerator.GenerateClassFromWorksheet("STATS.Objects");

            // Show the results
            MessageBox.Show("Success: " + success);
        }
    
    
    There is another override to load multiple sheets at once. I will build a sample project when I get some time to build a sample spreadsheet I can give away.
    
    To load multiple sheets:
    
    List<LoadWorksheetInfo> loadWorkSheetsInfo = new List<LoadWorksheetInfo>();
    
    // Add each LoadWorksheetInfo
    workbook = ExcellDataLoader.LoadWorkbook(path, loadWorkSheetsInfo)
    
    I will build some helper methods to save writing as much code once I use this a little to know what is needed.
    
    
    
My first test loaded a 12 column spreadsheet with 3,376 rows in just a few seconds.

I have a new project that uses this project as a good sample. Blazor.Excelerate will soon be an online 
way to create classes from a spreadsheet.

https://github.com/DataJuggler/Blazor.Excelerate

More helper methods and features will be added. The Nuget package has been released: DataJuggler.Excelerate.

Feel free to mention any new features you think would be useful. I can't promise to do them all, but if it is a good fit for this project I will add it.

This code is all brand new, so use with caution until more testing has been done. First tests have been promising.

I just finished adding a Load method, that is code generated when the classes are written.

** I am available for hire if you need help with any size C# / SQL Server project **

