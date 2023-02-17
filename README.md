<img height=192 width=192 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogoSmallWhite.png>

Here is a short video:
https://youtu.be/mj1h4_IRAQQ

I deleted my server, it was getting overrun by bots. I will release a new sample project later.

I will publish a new sample project soon.

Nuget package version 1.1.9 was just published: DataJuggler.Excelerate

A class named CodeGenerator was just created, and now by inheriting from the same CSharpClassWriter that code generates for DataTier.Net, I code generate
classes based on your header row.

I have a couple of clients that I build programs that automate combining columns from multiple Worksheets to form reports.

Rather than continue to write custom loaders, I really only need custom Exporters in most cases.

Here is a short video:
https://youtu.be/Sa-xroxPw_I

This short code snippet will load all the rows from a worksheet:

Snippet is from a Windows Form .Net 5 project, located in the Sample folder of this project. Very simple for now:

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

Live Demo

Code Generate C# Classes From Excel Header Rows
https://excelerate.datajuggler.com

Note: 
Blazor Excelerate comes with a sample workbook MemberData.xlsx

Random Members (20,000)<br>
Addresses (20,000)<br>
States (51)

Source Code For Above Site:

Blazor Excelerate
https://github.com/DataJuggler/Blazor.Excelerate



Working WinForms Demo

Use Excel As A Backend
https://github.com/DataJuggler/Excelerate.WinForms.Demo

Tutorial Practice

Step 1: 

Go To Blazor Excelerate
https://excelerate.datajuggler.com

Step 2: Download MemberData.xlsx

Step 3: Click the Upload Excel button and browse for MemberData.xlsx downloaded in Step 1.

Step 4: Set your Namespace name to 'Demo.Objects'

Step 5: Select Members sheet in the ComboBox

Step 6: Click the Generate Class Button.

Step 7: Download the zip file and extract the class Members.cs

Step 8: Repeat Steps 3 - 7 for the Address and States sheets.

Now you have the classes generated that were used to build the Excelerate WinForms Demo

Loading Data

This method loads all 3 worksheets: (to be continued, building Codopy.com as a code formatter. Site is not live yet).



More helper methods and features have been added. The Nuget package has been released: DataJuggler.Excelerate.

Feel free to mention any new features you think would be useful. I can't promise to do them all, but if it is a good fit for this project I will add it.

** I am available for hire if you need help with any size C# / SQL Server project **

