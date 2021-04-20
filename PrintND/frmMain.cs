using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Collections;
using  System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Data;
using TAXDOCPrt;
using MSXML2;
namespace PrintTiffDemo
{
	/// <summary>
	/// Summary description for frmMain.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
        private System.String ___thePrintTiffFolder="";
        private System.String ___thePrinterName="";
        private System.String ___thePagesList="";
        private System.Windows.Forms.TextBox txtPrintTiffFolder;
        private System.Windows.Forms.RadioButton radList;
        private System.Windows.Forms.Label lblPagesList;
        private System.Windows.Forms.RadioButton radOdd;
        private System.Windows.Forms.RadioButton radEven;
        private System.Windows.Forms.RadioButton radAll;
        private System.Windows.Forms.TextBox txtPagesList;
        private System.Windows.Forms.Label lblPrintTiffFolder;
        private System.Windows.Forms.Label lblPrinterName;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.Label lblAdditinal;
        private System.Windows.Forms.TextBox txtAdditional;
        private System.Windows.Forms.GroupBox grpPagesList;
        private System.Windows.Forms.Button btnRefreshPrintersList;
        private System.Windows.Forms.ComboBox cbxPrintersList;
        private System.Windows.Forms.Button btnPreview;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMain()
		{
			InitializeComponent();
            btnRefreshPrintersList_Click(btnRefreshPrintersList, null);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnRefreshPrintersList = new System.Windows.Forms.Button();
            this.txtPrintTiffFolder = new System.Windows.Forms.TextBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.grpPagesList = new System.Windows.Forms.GroupBox();
            this.radEven = new System.Windows.Forms.RadioButton();
            this.radOdd = new System.Windows.Forms.RadioButton();
            this.lblPagesList = new System.Windows.Forms.Label();
            this.txtPagesList = new System.Windows.Forms.TextBox();
            this.radAll = new System.Windows.Forms.RadioButton();
            this.radList = new System.Windows.Forms.RadioButton();
            this.lblPrintTiffFolder = new System.Windows.Forms.Label();
            this.lblPrinterName = new System.Windows.Forms.Label();
            this.lblFileName = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnPrint = new System.Windows.Forms.Button();
            this.lblAdditinal = new System.Windows.Forms.Label();
            this.txtAdditional = new System.Windows.Forms.TextBox();
            this.btnPreview = new System.Windows.Forms.Button();
            this.cbxPrintersList = new System.Windows.Forms.ComboBox();
            this.grpPagesList.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnRefreshPrintersList
            // 
            this.btnRefreshPrintersList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefreshPrintersList.Location = new System.Drawing.Point(416, 172);
            this.btnRefreshPrintersList.Name = "btnRefreshPrintersList";
            this.btnRefreshPrintersList.Size = new System.Drawing.Size(72, 24);
            this.btnRefreshPrintersList.TabIndex = 10;
            this.btnRefreshPrintersList.Text = "Обновить";
            this.btnRefreshPrintersList.Click += new System.EventHandler(this.btnRefreshPrintersList_Click);
            // 
            // txtPrintTiffFolder
            // 
            this.txtPrintTiffFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPrintTiffFolder.Location = new System.Drawing.Point(8, 28);
            this.txtPrintTiffFolder.Name = "txtPrintTiffFolder";
            this.txtPrintTiffFolder.Size = new System.Drawing.Size(396, 20);
            this.txtPrintTiffFolder.TabIndex = 1;
            this.txtPrintTiffFolder.TextChanged += new System.EventHandler(this.txtPrintTiffFolder_TextChanged);
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFolder.Location = new System.Drawing.Point(416, 24);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(72, 24);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "Обзор...";
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // grpPagesList
            // 
            this.grpPagesList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpPagesList.Controls.Add(this.radEven);
            this.grpPagesList.Controls.Add(this.radOdd);
            this.grpPagesList.Controls.Add(this.lblPagesList);
            this.grpPagesList.Controls.Add(this.txtPagesList);
            this.grpPagesList.Controls.Add(this.radAll);
            this.grpPagesList.Controls.Add(this.radList);
            this.grpPagesList.Location = new System.Drawing.Point(8, 200);
            this.grpPagesList.Name = "grpPagesList";
            this.grpPagesList.Size = new System.Drawing.Size(396, 128);
            this.grpPagesList.TabIndex = 11;
            this.grpPagesList.TabStop = false;
            this.grpPagesList.Text = "Печатать страницы";
            // 
            // radEven
            // 
            this.radEven.Location = new System.Drawing.Point(8, 84);
            this.radEven.Name = "radEven";
            this.radEven.Size = new System.Drawing.Size(104, 24);
            this.radEven.TabIndex = 2;
            this.radEven.Text = "Чётные";
            this.radEven.CheckedChanged += new System.EventHandler(this.radEven_CheckedChanged);
            // 
            // radOdd
            // 
            this.radOdd.Location = new System.Drawing.Point(8, 52);
            this.radOdd.Name = "radOdd";
            this.radOdd.Size = new System.Drawing.Size(104, 24);
            this.radOdd.TabIndex = 1;
            this.radOdd.Text = "Нечётные";
            this.radOdd.CheckedChanged += new System.EventHandler(this.radOdd_CheckedChanged);
            // 
            // lblPagesList
            // 
            this.lblPagesList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPagesList.Location = new System.Drawing.Point(136, 72);
            this.lblPagesList.Name = "lblPagesList";
            this.lblPagesList.Size = new System.Drawing.Size(252, 48);
            this.lblPagesList.TabIndex = 5;
            this.lblPagesList.Text = "номера или диапазоны страниц, разделённые запятыми (например: 1,3,5-12).";
            // 
            // txtPagesList
            // 
            this.txtPagesList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPagesList.Location = new System.Drawing.Point(136, 48);
            this.txtPagesList.Name = "txtPagesList";
            this.txtPagesList.Size = new System.Drawing.Size(252, 20);
            this.txtPagesList.TabIndex = 4;
            this.txtPagesList.TextChanged += new System.EventHandler(this.txtPagesList_TextChanged);
            // 
            // radAll
            // 
            this.radAll.Checked = true;
            this.radAll.Location = new System.Drawing.Point(8, 20);
            this.radAll.Name = "radAll";
            this.radAll.Size = new System.Drawing.Size(104, 24);
            this.radAll.TabIndex = 0;
            this.radAll.TabStop = true;
            this.radAll.Text = "Все";
            this.radAll.CheckedChanged += new System.EventHandler(this.radAll_CheckedChanged);
            // 
            // radList
            // 
            this.radList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.radList.Location = new System.Drawing.Point(136, 20);
            this.radList.Name = "radList";
            this.radList.Size = new System.Drawing.Size(252, 24);
            this.radList.TabIndex = 3;
            this.radList.Text = "Из списка:";
            this.radList.CheckedChanged += new System.EventHandler(this.radList_CheckedChanged);
            // 
            // lblPrintTiffFolder
            // 
            this.lblPrintTiffFolder.Location = new System.Drawing.Point(12, 8);
            this.lblPrintTiffFolder.Name = "lblPrintTiffFolder";
            this.lblPrintTiffFolder.Size = new System.Drawing.Size(144, 16);
            this.lblPrintTiffFolder.TabIndex = 0;
            this.lblPrintTiffFolder.Text = "Каталог шаблонов печати";
            // 
            // lblPrinterName
            // 
            this.lblPrinterName.Location = new System.Drawing.Point(12, 152);
            this.lblPrinterName.Name = "lblPrinterName";
            this.lblPrinterName.Size = new System.Drawing.Size(144, 16);
            this.lblPrinterName.TabIndex = 8;
            this.lblPrinterName.Text = "Принтер для печати";
            // 
            // lblFileName
            // 
            this.lblFileName.Location = new System.Drawing.Point(12, 56);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(144, 16);
            this.lblFileName.TabIndex = 3;
            this.lblFileName.Text = "Файл для печати";
            // 
            // txtFileName
            // 
            this.txtFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFileName.Location = new System.Drawing.Point(8, 76);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(396, 20);
            this.txtFileName.TabIndex = 4;
            this.txtFileName.TextChanged += new System.EventHandler(this.txtFileName_TextChanged);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFile.Location = new System.Drawing.Point(416, 72);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(72, 24);
            this.btnSelectFile.TabIndex = 5;
            this.btnSelectFile.Text = "Обзор...";
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.Location = new System.Drawing.Point(416, 204);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(72, 24);
            this.btnPrint.TabIndex = 12;
            this.btnPrint.Text = "Печать";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // lblAdditinal
            // 
            this.lblAdditinal.Location = new System.Drawing.Point(12, 104);
            this.lblAdditinal.Name = "lblAdditinal";
            this.lblAdditinal.Size = new System.Drawing.Size(280, 16);
            this.lblAdditinal.TabIndex = 6;
            this.lblAdditinal.Text = "Наименование налогового органа-получателя";
            // 
            // txtAdditional
            // 
            this.txtAdditional.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtAdditional.Location = new System.Drawing.Point(8, 124);
            this.txtAdditional.Name = "txtAdditional";
            this.txtAdditional.Size = new System.Drawing.Size(396, 20);
            this.txtAdditional.TabIndex = 7;
            // 
            // btnPreview
            // 
            this.btnPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPreview.Location = new System.Drawing.Point(416, 236);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(72, 24);
            this.btnPreview.TabIndex = 13;
            this.btnPreview.Text = "Просмотр";
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // cbxPrintersList
            // 
            this.cbxPrintersList.AllowDrop = true;
            this.cbxPrintersList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbxPrintersList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxPrintersList.Location = new System.Drawing.Point(8, 172);
            this.cbxPrintersList.Name = "cbxPrintersList";
            this.cbxPrintersList.Size = new System.Drawing.Size(396, 21);
            this.cbxPrintersList.TabIndex = 9;
            this.cbxPrintersList.SelectedIndexChanged += new System.EventHandler(this.cbxPrintersList_SelectedIndexChanged);
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(492, 321);
            this.Controls.Add(this.cbxPrintersList);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.lblAdditinal);
            this.Controls.Add(this.txtAdditional);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.txtPrintTiffFolder);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.lblFileName);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.lblPrinterName);
            this.Controls.Add(this.lblPrintTiffFolder);
            this.Controls.Add(this.grpPagesList);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.btnRefreshPrintersList);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(800, 360);
            this.MinimumSize = new System.Drawing.Size(500, 360);
            this.Name = "frmMain";
            this.Text = "Печать налоговых документов";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmMain_FormClosing);
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.grpPagesList.ResumeLayout(false);
            this.grpPagesList.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMain());
		}

        // Принтер
        private System.String varPrinterName{
            get{
                if(""==this.___thePrinterName){
                    throw new Exception("Выберите принтер");
                }
                return this.___thePrinterName;
            }
            set
            {
                cbxPrintersList.SelectedItem = this.___thePrinterName = value;
            }
        }
        private void cbxPrintersList_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            this.___thePrinterName=cbxPrintersList.SelectedItem.ToString();
        }
        private void btnRefreshPrintersList_Click(object sender, System.EventArgs e)
        {
            PrintDocument printDoc = new PrintDocument();
            PrinterSettings.StringCollection
              colInstalledPrinters=PrinterSettings.InstalledPrinters;
            System.Collections.IEnumerator
              curPrinter=colInstalledPrinters.GetEnumerator();
            for(cbxPrintersList.Items.Clear();curPrinter.MoveNext();cbxPrintersList.Items.Add(curPrinter.Current));

            


            if (printDoc.PrinterSettings.IsDefaultPrinter)
            {
                cbxPrintersList.SelectedItem = printDoc.PrinterSettings.PrinterName;
                //cbxPrintersList.SelectedItem=this.___thePrinterName;
            }


        }

        // Каталог с шаблонами печати (*.tiff)
        private System.String varPrintTiffFolder{
            get{
                if(""==this.___thePrintTiffFolder){
                    throw new Exception("Выберите каталог шаблонов печати");
                }
                return this.___thePrintTiffFolder;
            }set{
                this.txtPrintTiffFolder.Text=this.___thePrintTiffFolder=value;
            }
        }
        private void txtPrintTiffFolder_TextChanged(object sender, System.EventArgs e)
        {
            this.___thePrintTiffFolder=(sender as System.Windows.Forms.TextBox).Text;
        }
        private void btnSelectFolder_Click(object sender, System.EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog theFolderBrowserDialog=new System.Windows.Forms.FolderBrowserDialog();
            theFolderBrowserDialog.SelectedPath=this.___thePrintTiffFolder;
            if(System.Windows.Forms.DialogResult.OK==theFolderBrowserDialog.ShowDialog())
            {
                this.varPrintTiffFolder=theFolderBrowserDialog.SelectedPath;
            }
        }

        // Доп. параметры печати
        private System.String varPrintTiffParam(System.String name)
        {
            switch(name){
                case "Наименование налогового органа-получателя":
                return this.txtAdditional.Text;
            }
            throw new Exception("Параметр \""+name+"\" не поддерживается");
        }

        // Список страниц
        private System.String varPagesList{
            get{
                System.String thePagesList="";
                if(0==this.___thePagesList.IndexOf("."))
                {
                    thePagesList=this.___thePagesList.Substring(1);
                    System.Text.RegularExpressions.Regex thePagesListRexp=new System.Text.RegularExpressions.Regex("^\\d+(-\\d+)?(,\\d+(-\\d+)?)*$", System.Text.RegularExpressions.RegexOptions.Singleline);
                    if(!thePagesListRexp.IsMatch(thePagesList))
                    {
                        throw new System.Exception("Ошибка в списке листов");
                    }
                }
                else{
                    thePagesList=this.___thePagesList;
                }
                return thePagesList;
            }
        }
        private void radAll_CheckedChanged(object sender, System.EventArgs e)
        {
            this.___thePagesList="";
        }
        private void radOdd_CheckedChanged(object sender, System.EventArgs e)
        {
            this.___thePagesList="нечет";
        }
        private void radEven_CheckedChanged(object sender, System.EventArgs e)
        {
            this.___thePagesList="чет";
        }
        private void radList_CheckedChanged(object sender, System.EventArgs e)
        {
            txtPagesList_TextChanged(this.txtPagesList, null);
        }
        private void txtPagesList_TextChanged(object sender, System.EventArgs e)
        {
            if(this.radList.Checked)
            {
                this.___thePagesList="."+(sender as System.Windows.Forms.TextBox).Text;
            }
        }

        // Файл для печати
        private System.String varPrintFile
        {
            get{
                if(""==this.txtFileName.Text){
                    throw new Exception("Выберите файл для печати");
                }
                return this.txtFileName.Text;
            }
            set
            {
                this.txtFileName.Text=value;
            }
        }
        private void txtFileName_TextChanged(object sender, System.EventArgs e)
        {
            this.varPrintFile=(sender as System.Windows.Forms.TextBox).Text;
        }
        private void btnSelectFile_Click(object sender, System.EventArgs e)
        {
            
            System.Windows.Forms.OpenFileDialog theOpenFileDialog=new System.Windows.Forms.OpenFileDialog();
            theOpenFileDialog.InitialDirectory = varPrintFile;
            theOpenFileDialog.Filter = "текстовые файлы (*.txt)|*.txt|xml файлы (*.xml)|*.xml|все файлы (*.*)|*.*";
            theOpenFileDialog.FilterIndex = 2;
            theOpenFileDialog.RestoreDirectory = true;
            if (System.Windows.Forms.DialogResult.OK==theOpenFileDialog.ShowDialog())
            {
                this.varPrintFile=theOpenFileDialog.FileName;
            }
        }

        // Печать
        private void btnPrint_Click(object sender, System.EventArgs e)
        {
            //MSXML2.DOMDocument xmlReport=new MSXML2.DOMDocumentClass();
            DOMDocument xmlReport = new DOMDocument(); 
            try {
                //TAXDOCPrt.TAXDOCPrint2 TAXDOCPrint = new TAXDOCPrt.TAXDOCPrint2(); 
                TAXDOCPrintClass TAXDOCPrint=new TAXDOCPrintClass();
                //TAXDOCPrint TAXDOCPrint = (TAXDOCPrt.TAXDOCPrint2) Marshal.GetActiveObject("TAXDOCPrt.TAXDOCPrint2");
                
                String[] ParametersNames={"Наименование налогового органа-получателя"};
                for(int p=0; p<ParametersNames.Length; p++){
                    String ParametersName=ParametersNames[p];
                    TAXDOCPrint.SetPrintTiffParam(ParametersName, this.varPrintTiffParam(ParametersName));
                }
                TAXDOCPrint.PrintTiffFolder=this.varPrintTiffFolder;
                TAXDOCPrint.PrintFile(this.varPrintFile, this.varPagesList, this.varPrinterName, 1, xmlReport);
            }
            catch(System.Runtime.InteropServices.ExternalException ex){
                System.String theText=" :";
                if(0==(0x80040014^(ex.ErrorCode&0xFFFFFFFF)))
                {
                    MSXML2.IXMLDOMNodeList theErrors=xmlReport.selectNodes("//Error");
                    MSXML2.IXMLDOMNode theError;
                    while(null!=(theError=theErrors.nextNode()))
                    {
                        theText+="\r\n"+theError.text;
                    }
                }
                System.Windows.Forms.MessageBox.Show(this,ex.Message+theText,"Ошибка печати", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
            catch(System.Exception ex){
                System.Windows.Forms.MessageBox.Show(this, ex.Message, "Внимание", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        // Просмотр
        private void btnPreview_Click(object sender, System.EventArgs e)
        {
            DOMDocument xmlReport = new DOMDocument();
            //MSXML2.DOMDocument xmlReport=new MSXML2.DOMDocumentClass();
            try
            {
                TAXDOCPrintClass TAXDOCPrint = new TAXDOCPrintClass();

                //TAXDOCPrt.TAXDOCPrint2 TAXDOCPrint = new TAXDOCPrt.TAXDOCPrint2();
                //TAXDOCPrt.TAXDOCPrint TAXDOCPrint = (TAXDOCPrt.TAXDOCPrint)Marshal.GetActiveObject("TAXDOCPrt.TAXDOCPrint");
                String[] ParametersNames={"Наименование налогового органа-получателя"};
                for(int p=0; p<ParametersNames.Length; p++)
                {
                    String ParametersName=ParametersNames[p];
                    TAXDOCPrint.SetPrintTiffParam(ParametersName, this.varPrintTiffParam(ParametersName));
                }
                TAXDOCPrint.PrintTiffFolder=this.varPrintTiffFolder;
                TAXDOCPrint.PreviewFile(this.varPrintFile, this.Handle.ToInt32(), xmlReport);
            }
            catch(System.Runtime.InteropServices.ExternalException ex)
            {
                System.String theText=" :";
                if(0==(0x80040014^(ex.ErrorCode&0xFFFFFFFF)))
                {
                    MSXML2.IXMLDOMNodeList theErrors=xmlReport.selectNodes("//Error");
                    MSXML2.IXMLDOMNode theError;
                    while(null!=(theError=theErrors.nextNode()))
                    {
                        theText+="\r\n"+theError.text;
                    }
                }
                System.Windows.Forms.MessageBox.Show(this,ex.Message+theText,"Ошибка просмотра", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
            catch(System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(this, ex.Message, "Внимание", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            varPrintTiffFolder = Properties.Settings.Default.PathTemplate;
            varPrintFile = Properties.Settings.Default.PathFile;
            varPrinterName = Properties.Settings.Default.Printer;
            this.txtAdditional.Text = Properties.Settings.Default.KODNO;
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.PathTemplate = varPrintTiffFolder.ToString();
            Properties.Settings.Default.PathFile= varPrintFile;
            Properties.Settings.Default.Printer=varPrinterName;
            Properties.Settings.Default.KODNO= this.txtAdditional.Text;
            Properties.Settings.Default.Save();
        }
    }
}
