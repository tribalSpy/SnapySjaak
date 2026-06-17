namespace CMRPrint
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem extraToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem templateEditorToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exporterInfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem transportInfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem customerInfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadingPlacesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem batchPrintToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Panel panelCmrPreview;
        private System.Windows.Forms.Panel panelAdjust;
        private System.Windows.Forms.Button btnMoveUp;
        private System.Windows.Forms.Button btnMoveDown;
        private System.Windows.Forms.Button btnMoveLeft;
        private System.Windows.Forms.Button btnMoveRight;
        private System.Windows.Forms.Label lblAdjustHint;
        private System.Windows.Forms.Panel flowLayoutPanelFields;
        private System.Windows.Forms.Panel panelCustomer;
        private System.Windows.Forms.ComboBox customerComboBox;
        private System.Windows.Forms.Label lblCustomer;
        private System.Windows.Forms.TextBox txtConsignorName;
        private System.Windows.Forms.TextBox txtConsignorAddress;
        private System.Windows.Forms.TextBox txtConsignorCity;
        private System.Windows.Forms.TextBox txtConsignorCountry;
        private System.Windows.Forms.TextBox txtVatNumber;
        private System.Windows.Forms.Button btnImportExcel;
        private System.Windows.Forms.Button btnNewCustomer;
        private System.Windows.Forms.Button btnSaveCustomer;
        private System.Windows.Forms.Panel panelTemplates;
        private System.Windows.Forms.ComboBox cmbTemplates;
        private System.Windows.Forms.Button btnLoadTemplate;
        private System.Windows.Forms.Button btnSaveTemplate;
        private System.Windows.Forms.Button btnDeleteTemplate;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Button btnPrint;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog;
        private System.Windows.Forms.PrintDialog printDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            extraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            templateEditorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            exporterInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            transportInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            customerInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            loadingPlacesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            batchPrintToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            importExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            panelCmrPreview = new System.Windows.Forms.Panel();
            panelAdjust = new System.Windows.Forms.Panel();
            btnMoveUp = new System.Windows.Forms.Button();
            btnMoveDown = new System.Windows.Forms.Button();
            btnMoveLeft = new System.Windows.Forms.Button();
            btnMoveRight = new System.Windows.Forms.Button();
            lblAdjustHint = new System.Windows.Forms.Label();
            flowLayoutPanelFields = new System.Windows.Forms.Panel();
            panelCustomer = new System.Windows.Forms.Panel();
            customerComboBox = new System.Windows.Forms.ComboBox();
            lblCustomer = new System.Windows.Forms.Label();
            txtConsignorName = new System.Windows.Forms.TextBox();
            txtConsignorAddress = new System.Windows.Forms.TextBox();
            txtConsignorCity = new System.Windows.Forms.TextBox();
            txtConsignorCountry = new System.Windows.Forms.TextBox();
            txtVatNumber = new System.Windows.Forms.TextBox();
            btnImportExcel = new System.Windows.Forms.Button();
            btnNewCustomer = new System.Windows.Forms.Button();
            btnSaveCustomer = new System.Windows.Forms.Button();
            panelTemplates = new System.Windows.Forms.Panel();
            cmbTemplates = new System.Windows.Forms.ComboBox();
            btnLoadTemplate = new System.Windows.Forms.Button();
            btnSaveTemplate = new System.Windows.Forms.Button();
            btnDeleteTemplate = new System.Windows.Forms.Button();
            btnPreview = new System.Windows.Forms.Button();
            btnPrint = new System.Windows.Forms.Button();
            printDocument1 = new System.Drawing.Printing.PrintDocument();
            printPreviewDialog = new System.Windows.Forms.PrintPreviewDialog();
            printDialog = new System.Windows.Forms.PrintDialog();
            openFileDialog = new System.Windows.Forms.OpenFileDialog();

            menuStrip1.SuspendLayout();
            panelCustomer.SuspendLayout();
            panelTemplates.SuspendLayout();
            SuspendLayout();

            // menuStrip1
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[]
            {
                fileToolStripMenuItem,
                extraToolStripMenuItem
            });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(1400, 33);
            menuStrip1.TabIndex = 0;

            // fileToolStripMenuItem
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[]
            {
                importExcelToolStripMenuItem,
                exitToolStripMenuItem
            });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(54, 29);
            fileToolStripMenuItem.Text = "File";

            // importExcelToolStripMenuItem
            importExcelToolStripMenuItem.Name = "importExcelToolStripMenuItem";
            importExcelToolStripMenuItem.Size = new System.Drawing.Size(214, 30);
            importExcelToolStripMenuItem.Text = "Import Excel...";
            importExcelToolStripMenuItem.Click += new System.EventHandler(this.btnImportExcel_Click);

            // exitToolStripMenuItem
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new System.Drawing.Size(214, 30);
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);

            // extraToolStripMenuItem
            extraToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[]
            {
                templateEditorToolStripMenuItem,
                exporterInfoToolStripMenuItem,
                transportInfoToolStripMenuItem,
                customerInfoToolStripMenuItem,
                loadingPlacesToolStripMenuItem,
                batchPrintToolStripMenuItem
            });
            extraToolStripMenuItem.Name = "extraToolStripMenuItem";
            extraToolStripMenuItem.Size = new System.Drawing.Size(66, 29);
            extraToolStripMenuItem.Text = "Extra";

            // templateEditorToolStripMenuItem
            templateEditorToolStripMenuItem.Name = "templateEditorToolStripMenuItem";
            templateEditorToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            templateEditorToolStripMenuItem.Text = "Template Editor";
            templateEditorToolStripMenuItem.Click += new System.EventHandler(this.templateEditorToolStripMenuItem_Click);

            // exporterInfoToolStripMenuItem
            exporterInfoToolStripMenuItem.Name = "exporterInfoToolStripMenuItem";
            exporterInfoToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            exporterInfoToolStripMenuItem.Text = "Exporter Info";
            exporterInfoToolStripMenuItem.Click += new System.EventHandler(this.exporterInfoToolStripMenuItem_Click);

            // transportInfoToolStripMenuItem
            transportInfoToolStripMenuItem.Name = "transportInfoToolStripMenuItem";
            transportInfoToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            transportInfoToolStripMenuItem.Text = "Transport Info";
            transportInfoToolStripMenuItem.Click += new System.EventHandler(this.transportInfoToolStripMenuItem_Click);

            // customerInfoToolStripMenuItem
            customerInfoToolStripMenuItem.Name = "customerInfoToolStripMenuItem";
            customerInfoToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            customerInfoToolStripMenuItem.Text = "Customer Info";
            customerInfoToolStripMenuItem.Click += new System.EventHandler(this.customerInfoToolStripMenuItem_Click);

            // loadingPlacesToolStripMenuItem
            loadingPlacesToolStripMenuItem.Name = "loadingPlacesToolStripMenuItem";
            loadingPlacesToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            loadingPlacesToolStripMenuItem.Text = "Loading Places";
            loadingPlacesToolStripMenuItem.Click += new System.EventHandler(this.loadingPlacesToolStripMenuItem_Click);

            // batchPrintToolStripMenuItem
            batchPrintToolStripMenuItem.Name = "batchPrintToolStripMenuItem";
            batchPrintToolStripMenuItem.Size = new System.Drawing.Size(235, 30);
            batchPrintToolStripMenuItem.Text = "Batch Print CMRs";
            batchPrintToolStripMenuItem.Click += new System.EventHandler(this.batchPrintToolStripMenuItem_Click);

            // panelCmrPreview
            panelCmrPreview.BackColor = System.Drawing.Color.White;
            panelCmrPreview.AutoScroll = true;
            panelCmrPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panelCmrPreview.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
            panelCmrPreview.Location = new System.Drawing.Point(12, 40);
            panelCmrPreview.Name = "panelCmrPreview";
            panelCmrPreview.Size = new System.Drawing.Size(560, 690);
            panelCmrPreview.TabIndex = 1;
            panelCmrPreview.Paint += new System.Windows.Forms.PaintEventHandler(this.panelCmrPreview_Paint);

            // panelAdjust
            panelAdjust.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panelAdjust.Location = new System.Drawing.Point(12, 735);
            panelAdjust.Name = "panelAdjust";
            panelAdjust.Size = new System.Drawing.Size(560, 65);
            panelAdjust.TabIndex = 5;

            // lblAdjustHint
            lblAdjustHint.AutoSize = true;
            lblAdjustHint.Location = new System.Drawing.Point(10, 10);
            lblAdjustHint.Name = "lblAdjustHint";
            lblAdjustHint.Size = new System.Drawing.Size(273, 20);
            lblAdjustHint.TabIndex = 0;
            lblAdjustHint.Text = "Select a box, then move with arrows or keys";

            // btnMoveUp
            btnMoveUp.Location = new System.Drawing.Point(360, 5);
            btnMoveUp.Name = "btnMoveUp";
            btnMoveUp.Size = new System.Drawing.Size(48, 24);
            btnMoveUp.TabIndex = 1;
            btnMoveUp.Text = "Up";
            btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);

            // btnMoveLeft
            btnMoveLeft.Location = new System.Drawing.Point(305, 33);
            btnMoveLeft.Name = "btnMoveLeft";
            btnMoveLeft.Size = new System.Drawing.Size(55, 24);
            btnMoveLeft.TabIndex = 2;
            btnMoveLeft.Text = "Left";
            btnMoveLeft.Click += new System.EventHandler(this.btnMoveLeft_Click);

            // btnMoveDown
            btnMoveDown.Location = new System.Drawing.Point(360, 33);
            btnMoveDown.Name = "btnMoveDown";
            btnMoveDown.Size = new System.Drawing.Size(55, 24);
            btnMoveDown.TabIndex = 3;
            btnMoveDown.Text = "Down";
            btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);

            // btnMoveRight
            btnMoveRight.Location = new System.Drawing.Point(420, 33);
            btnMoveRight.Name = "btnMoveRight";
            btnMoveRight.Size = new System.Drawing.Size(60, 24);
            btnMoveRight.TabIndex = 4;
            btnMoveRight.Text = "Right";
            btnMoveRight.Click += new System.EventHandler(this.btnMoveRight_Click);

            panelAdjust.Controls.Add(lblAdjustHint);
            panelAdjust.Controls.Add(btnMoveUp);
            panelAdjust.Controls.Add(btnMoveLeft);
            panelAdjust.Controls.Add(btnMoveDown);
            panelAdjust.Controls.Add(btnMoveRight);

            // flowLayoutPanelFields
            flowLayoutPanelFields.AutoScroll = true;
            flowLayoutPanelFields.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            flowLayoutPanelFields.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            flowLayoutPanelFields.Location = new System.Drawing.Point(580, 220);
            flowLayoutPanelFields.Name = "flowLayoutPanelFields";
            flowLayoutPanelFields.Size = new System.Drawing.Size(808, 475);
            flowLayoutPanelFields.TabIndex = 2;

            // panelCustomer
            panelCustomer.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            panelCustomer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panelCustomer.Location = new System.Drawing.Point(580, 40);
            panelCustomer.Name = "panelCustomer";
            panelCustomer.Size = new System.Drawing.Size(808, 170);
            panelCustomer.TabIndex = 3;

            // lblCustomer
            lblCustomer.AutoSize = true;
            lblCustomer.Location = new System.Drawing.Point(10, 10);
            lblCustomer.Name = "lblCustomer";
            lblCustomer.Size = new System.Drawing.Size(79, 20);
            lblCustomer.TabIndex = 0;
            lblCustomer.Text = "Customer:";

            // customerComboBox
            customerComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            customerComboBox.Location = new System.Drawing.Point(100, 7);
            customerComboBox.Name = "customerComboBox";
            customerComboBox.Size = new System.Drawing.Size(500, 28);
            customerComboBox.TabIndex = 1;
            customerComboBox.SelectedIndexChanged += new System.EventHandler(this.customerComboBox_SelectedIndexChanged);

            // txtConsignorName
            txtConsignorName.Location = new System.Drawing.Point(100, 35);
            txtConsignorName.Name = "txtConsignorName";
            txtConsignorName.Size = new System.Drawing.Size(500, 27);
            txtConsignorName.TabIndex = 2;

            // txtConsignorAddress
            txtConsignorAddress.Location = new System.Drawing.Point(640, 35);
            txtConsignorAddress.Name = "txtConsignorAddress";
            txtConsignorAddress.Size = new System.Drawing.Size(360, 27);
            txtConsignorAddress.TabIndex = 3;

            // txtConsignorCity
            txtConsignorCity.Location = new System.Drawing.Point(1010, 35);
            txtConsignorCity.Name = "txtConsignorCity";
            txtConsignorCity.Size = new System.Drawing.Size(150, 27);
            txtConsignorCity.TabIndex = 4;

            // txtConsignorCountry
            txtConsignorCountry.Location = new System.Drawing.Point(1170, 35);
            txtConsignorCountry.Name = "txtConsignorCountry";
            txtConsignorCountry.Size = new System.Drawing.Size(90, 27);
            txtConsignorCountry.TabIndex = 5;

            // txtVatNumber
            txtVatNumber.Location = new System.Drawing.Point(1270, 35);
            txtVatNumber.Name = "txtVatNumber";
            txtVatNumber.Size = new System.Drawing.Size(100, 27);
            txtVatNumber.TabIndex = 6;

            // btnImportExcel
            btnImportExcel.Location = new System.Drawing.Point(10, 70);
            btnImportExcel.Name = "btnImportExcel";
            btnImportExcel.Size = new System.Drawing.Size(100, 30);
            btnImportExcel.TabIndex = 7;
            btnImportExcel.Text = "Import Excel";
            btnImportExcel.Click += new System.EventHandler(this.btnImportExcel_Click);

            // btnNewCustomer
            btnNewCustomer.Location = new System.Drawing.Point(120, 70);
            btnNewCustomer.Name = "btnNewCustomer";
            btnNewCustomer.Size = new System.Drawing.Size(90, 30);
            btnNewCustomer.TabIndex = 8;
            btnNewCustomer.Text = "New";
            btnNewCustomer.Click += new System.EventHandler(this.btnNewCustomer_Click);

            // btnSaveCustomer
            btnSaveCustomer.Location = new System.Drawing.Point(220, 70);
            btnSaveCustomer.Name = "btnSaveCustomer";
            btnSaveCustomer.Size = new System.Drawing.Size(90, 30);
            btnSaveCustomer.TabIndex = 9;
            btnSaveCustomer.Text = "Save";
            btnSaveCustomer.Click += new System.EventHandler(this.btnSaveCustomer_Click);

            panelCustomer.Controls.Add(lblCustomer);
            panelCustomer.Controls.Add(customerComboBox);
            panelCustomer.Controls.Add(txtConsignorName);
            panelCustomer.Controls.Add(txtConsignorAddress);
            panelCustomer.Controls.Add(txtConsignorCity);
            panelCustomer.Controls.Add(txtConsignorCountry);
            panelCustomer.Controls.Add(txtVatNumber);
            panelCustomer.Controls.Add(btnImportExcel);
            panelCustomer.Controls.Add(btnNewCustomer);
            panelCustomer.Controls.Add(btnSaveCustomer);

            // panelTemplates
            panelTemplates.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Bottom;
            panelTemplates.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            panelTemplates.Location = new System.Drawing.Point(580, 705);
            panelTemplates.Name = "panelTemplates";
            panelTemplates.Size = new System.Drawing.Size(808, 50);
            panelTemplates.TabIndex = 4;

            // cmbTemplates
            cmbTemplates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbTemplates.Location = new System.Drawing.Point(10, 10);
            cmbTemplates.Name = "cmbTemplates";
            cmbTemplates.Size = new System.Drawing.Size(400, 28);
            cmbTemplates.TabIndex = 0;

            // btnLoadTemplate
            btnLoadTemplate.Location = new System.Drawing.Point(420, 10);
            btnLoadTemplate.Name = "btnLoadTemplate";
            btnLoadTemplate.Size = new System.Drawing.Size(80, 28);
            btnLoadTemplate.TabIndex = 1;
            btnLoadTemplate.Text = "Load";
            btnLoadTemplate.Click += new System.EventHandler(this.btnLoadTemplate_Click);

            // btnSaveTemplate
            btnSaveTemplate.Location = new System.Drawing.Point(510, 10);
            btnSaveTemplate.Name = "btnSaveTemplate";
            btnSaveTemplate.Size = new System.Drawing.Size(80, 28);
            btnSaveTemplate.TabIndex = 2;
            btnSaveTemplate.Text = "Save";
            btnSaveTemplate.Click += new System.EventHandler(this.btnSaveTemplate_Click);

            // btnDeleteTemplate
            btnDeleteTemplate.Location = new System.Drawing.Point(600, 10);
            btnDeleteTemplate.Name = "btnDeleteTemplate";
            btnDeleteTemplate.Size = new System.Drawing.Size(80, 28);
            btnDeleteTemplate.TabIndex = 3;
            btnDeleteTemplate.Text = "Delete";
            btnDeleteTemplate.Click += new System.EventHandler(this.btnDeleteTemplate_Click);

            // btnPreview
            btnPreview.Location = new System.Drawing.Point(700, 10);
            btnPreview.Name = "btnPreview";
            btnPreview.Size = new System.Drawing.Size(100, 28);
            btnPreview.TabIndex = 4;
            btnPreview.Text = "Preview Print";
            btnPreview.Click += new System.EventHandler(this.btnPreview_Click);

            // btnPrint
            btnPrint.Location = new System.Drawing.Point(810, 10);
            btnPrint.Name = "btnPrint";
            btnPrint.Size = new System.Drawing.Size(100, 28);
            btnPrint.TabIndex = 5;
            btnPrint.Text = "Print CMR";
            btnPrint.Click += new System.EventHandler(this.btnPrint_Click);

            panelTemplates.Controls.Add(cmbTemplates);
            panelTemplates.Controls.Add(btnLoadTemplate);
            panelTemplates.Controls.Add(btnSaveTemplate);
            panelTemplates.Controls.Add(btnDeleteTemplate);
            panelTemplates.Controls.Add(btnPreview);
            panelTemplates.Controls.Add(btnPrint);

            // printDocument1
            printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);

            // printPreviewDialog
            printPreviewDialog.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            printPreviewDialog.ClientSize = new System.Drawing.Size(600, 500);
            printPreviewDialog.Document = printDocument1;
            printPreviewDialog.Name = "printPreviewDialog";
            printPreviewDialog.UseAntiAlias = true;

            // printDialog
            printDialog.Document = printDocument1;
            printDialog.UseEXDialog = true;

            // openFileDialog
            openFileDialog.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            openFileDialog.Title = "Select Excel file";

            // MainForm
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoScroll = true;
            ClientSize = new System.Drawing.Size(1450, 780);
            Controls.Add(menuStrip1);
            Controls.Add(panelCmrPreview);
            Controls.Add(panelAdjust);
            Controls.Add(panelCustomer);
            Controls.Add(flowLayoutPanelFields);
            Controls.Add(panelTemplates);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            KeyPreview = true;
            MainMenuStrip = menuStrip1;
            MaximizeBox = true;
            MinimumSize = new System.Drawing.Size(1350, 860);
            Name = "MainForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "CMR Print Manager - 24 Place CMR Document";
            WindowState = System.Windows.Forms.FormWindowState.Maximized;
            KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);

            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            panelCustomer.ResumeLayout(false);
            panelCustomer.PerformLayout();
            panelTemplates.ResumeLayout(false);
            panelTemplates.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
