
namespace DataJuggler.Excelerate.Sample
{

    #region class MainForm
    /// <summary>
    /// This is the designer for the MainForm.
    /// </summary>
    partial class MainForm
    {
        
        #region Private Variables
        private System.ComponentModel.IContainer components = null;
        private DataJuggler.Win.Controls.Button TestButton;
        private DataJuggler.Win.Controls.LabelTextBoxBrowserControl WorksheetControl;
        #endregion
        
        #region Methods
            
            #region Dispose(bool disposing)
            /// <summary>
            ///  Clean up any resources being used.
            /// </summary>
            /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
            protected override void Dispose(bool disposing)
            {
                if (disposing && (components != null))
                {
                    components.Dispose();
                }
                base.Dispose(disposing);
            }
            #endregion
            
            #region InitializeComponent()
            /// <summary>
            ///  Required method for Designer support - do not modify
            ///  the contents of this method with the code editor.
            /// </summary>
            private void InitializeComponent()
            {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.TestButton = new DataJuggler.Win.Controls.Button();
            this.WorksheetControl = new DataJuggler.Win.Controls.LabelTextBoxBrowserControl();
            this.SheetNameControl = new DataJuggler.Win.Controls.LabelTextBoxControl();
            this.SuspendLayout();
            // 
            // TestButton
            // 
            this.TestButton.BackColor = System.Drawing.Color.Transparent;
            this.TestButton.ButtonText = "Test";
            this.TestButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.TestButton.ForeColor = System.Drawing.Color.LemonChiffon;
            this.TestButton.Location = new System.Drawing.Point(592, 327);
            this.TestButton.Name = "TestButton";
            this.TestButton.Size = new System.Drawing.Size(144, 48);
            this.TestButton.TabIndex = 1;
            this.TestButton.Theme = DataJuggler.Win.Controls.Enumerations.ThemeEnum.Dark;
            this.TestButton.Click += new System.EventHandler(this.TestButton_Click);
            // 
            // WorksheetControl
            // 
            this.WorksheetControl.BackColor = System.Drawing.Color.Transparent;
            this.WorksheetControl.BrowseType = DataJuggler.Win.Controls.Enumerations.BrowseTypeEnum.File;
            this.WorksheetControl.ButtonColor = System.Drawing.Color.LemonChiffon;
            this.WorksheetControl.ButtonImage = ((System.Drawing.Image)(resources.GetObject("WorksheetControl.ButtonImage")));
            this.WorksheetControl.ButtonWidth = 48;
            this.WorksheetControl.DarkMode = false;
            this.WorksheetControl.DisabledLabelColor = System.Drawing.Color.Empty;
            this.WorksheetControl.Editable = true;
            this.WorksheetControl.EnabledLabelColor = System.Drawing.Color.Empty;
            this.WorksheetControl.Filter = null;
            this.WorksheetControl.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.WorksheetControl.HideBrowseButton = false;
            this.WorksheetControl.LabelBottomMargin = 0;
            this.WorksheetControl.LabelColor = System.Drawing.Color.LemonChiffon;
            this.WorksheetControl.LabelFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.WorksheetControl.LabelText = "Excel:";
            this.WorksheetControl.LabelTopMargin = 0;
            this.WorksheetControl.LabelWidth = 120;
            this.WorksheetControl.Location = new System.Drawing.Point(60, 46);
            this.WorksheetControl.Name = "WorksheetControl";
            this.WorksheetControl.OnTextChangedListener = null;
            this.WorksheetControl.OpenCallback = null;
            this.WorksheetControl.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.WorksheetControl.SelectedPath = null;
            this.WorksheetControl.Size = new System.Drawing.Size(676, 32);
            this.WorksheetControl.StartPath = null;
            this.WorksheetControl.TabIndex = 2;
            this.WorksheetControl.TextBoxBottomMargin = 0;
            this.WorksheetControl.TextBoxDisabledColor = System.Drawing.Color.Empty;
            this.WorksheetControl.TextBoxEditableColor = System.Drawing.Color.Empty;
            this.WorksheetControl.TextBoxFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.WorksheetControl.TextBoxTopMargin = 0;
            this.WorksheetControl.Theme = DataJuggler.Win.Controls.Enumerations.ThemeEnum.Dark;
            // 
            // SheetNameControl
            // 
            this.SheetNameControl.BackColor = System.Drawing.Color.Transparent;
            this.SheetNameControl.BottomMargin = 0;
            this.SheetNameControl.Editable = true;
            this.SheetNameControl.Encrypted = false;
            this.SheetNameControl.Font = new System.Drawing.Font("Verdana", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.SheetNameControl.LabelBottomMargin = 0;
            this.SheetNameControl.LabelColor = System.Drawing.Color.LemonChiffon;
            this.SheetNameControl.LabelFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.SheetNameControl.LabelText = "Sheetname:";
            this.SheetNameControl.LabelTopMargin = 0;
            this.SheetNameControl.LabelWidth = 120;
            this.SheetNameControl.Location = new System.Drawing.Point(60, 108);
            this.SheetNameControl.MultiLine = false;
            this.SheetNameControl.Name = "SheetNameControl";
            this.SheetNameControl.OnTextChangedListener = null;
            this.SheetNameControl.PasswordMode = false;
            this.SheetNameControl.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.SheetNameControl.Size = new System.Drawing.Size(360, 32);
            this.SheetNameControl.TabIndex = 3;
            this.SheetNameControl.TextBoxBottomMargin = 0;
            this.SheetNameControl.TextBoxDisabledColor = System.Drawing.Color.LightGray;
            this.SheetNameControl.TextBoxEditableColor = System.Drawing.Color.White;
            this.SheetNameControl.TextBoxFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.SheetNameControl.TextBoxTopMargin = 0;
            this.SheetNameControl.Theme = DataJuggler.Win.Controls.Enumerations.ThemeEnum.Dark;
            // 
            // MainForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.SheetNameControl);
            this.Controls.Add(this.WorksheetControl);
            this.Controls.Add(this.TestButton);
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excelerate";
            this.ResumeLayout(false);

            }
        #endregion

        #endregion

        private Win.Controls.LabelTextBoxControl SheetNameControl;
    }
    #endregion

}
