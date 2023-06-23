<img height=192 width=192 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogoSmallWhite.png>

Here is a short video:
https://youtu.be/mj1h4_IRAQQ

Live demo:
https://excelerate.datajuggler.com

Another very useful project built with Excelerate is:

Nuget package DataJuggler.SQLSnapshot.

Export an entire SQL Server database including all data rows to Excel with one line of code, passing in a connectionstring and a path to save.

Source Code
https://github.com/DataJuggler/SQLSnapshot

SQL Snapshot Desktop
https://github.com/DataJuggler/DemoSQLSnapshot
A Winforms app using SQL Snapshot.

Study the code in SQL Snapshot. I wrote SQL Snapshot in 1 day, with a few updates later. That is the power of DataJuggler.Net7 for SQL Server schema reading and data loading and using DataJuggler.Excelerate to write out the data rows to Excel.

Latest version is of Excelerate is 7.2.12 and has over 23,000 Nuget installs. 

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

    /// <summary>
    /// Load all 3 Lists Names, Addresses and States.
    /// </summary>
    public int LoadAllData()
    {
        // initial value
        int objectsLoaded = 0;
        
        // Ensure visible
        StatusLabel.Visible = true;
        Graph.Visible = true;
        
        // ExcelPath is a relative path: const string ExcelPath =  "Documents/MemberData.xlsx";
        string path = Path.GetFullPath(ExcelPath);
        
        // Set the Status
        StatusLabel.Text = "Loading data, please wait.";
        
        // Force a refresh here
        Refresh();
        Application.DoEvents();
        
        // if the path exists
        if (FileHelper.Exists(path))
        {
            // load the workbook
            Workbook = ExcelDataLoader.LoadAllData(path);
            
            // if the workbook exists and has 3 or more shorts
            if ((NullHelper.Exists(workbook)) && (ListHelper.HasXOrMoreItems(workbook.Worksheets, 3)))
            {
                // Get the indexes of each sheet
                membersIndex = workbook.GetWorksheetIndex("Members");
                addressIndex = workbook.GetWorksheetIndex("Address");
                statesIndex = workbook.GetWorksheetIndex("States");
                
                // verify all sheet indexes were found
                if ((membersIndex >= 0) && (addressIndex >= 0) && (statesIndex >= 0))
                {
                    // Get the counts
                    int membersCount = workbook.Worksheets[membersIndex].Rows.Count -1;
                    int addressesCount = workbook.Worksheets[addressIndex].Rows.Count -1;
                    int statesCount = workbook.Worksheets[statesIndex].Rows.Count -1;
                    
                    // Setup the Graph
                    Graph.Maximum = workbook.Worksheets[0].Rows.Count + addressesCount + statesCount;
                    Graph.Value = 0;
                    
                    // Force a refresh here
                    Refresh();
                    Application.DoEvents();
                    
                    // Load the Members
                    LoadMembers(workbook.Worksheets[membersIndex]);
                    
                    // Load the Addresses
                    LoadAddresses(workbook.Worksheets[addressIndex]);
                    
                    // Load the States
                    LoadStates(workbook.Worksheets[statesIndex]);
                    
                    // Set the StatusLabel
                    StatusLabel.Text = "All data has been loaded";
                }
            }
        }
        
        // return value
        return objectsLoaded;
    }
    


More helper methods and features have been added. The Nuget package has been released: DataJuggler.Excelerate.

Feel free to mention any new features you think would be useful. I can't promise to do them all, but if it is a good fit for this project I will add it.

** I am available for hire if you need help with any size C# / SQL Server project **

