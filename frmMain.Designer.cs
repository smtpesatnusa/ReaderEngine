
namespace ReaderEngine
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridViewProcessTime = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.deleteBtn = new System.Windows.Forms.Button();
            this.addBtn = new System.Windows.Forms.Button();
            this.dateTimePickerTimer = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridViewTransactionEmployee = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.timerRefresh = new System.Windows.Forms.Timer(this.components);
            this.ReaderEnginee = new System.Windows.Forms.NotifyIcon(this.components);
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.dateTimeNow = new System.Windows.Forms.ToolStripLabel();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.btnMail = new System.Windows.Forms.Button();
            this.btnProcess = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.dataGridViewESDLog = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.logDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refrehLabel = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProcessTime)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTransactionEmployee)).BeginInit();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewESDLog)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(15, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(215, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Process Time Scheduler";
            // 
            // dataGridViewProcessTime
            // 
            this.dataGridViewProcessTime.AllowUserToAddRows = false;
            this.dataGridViewProcessTime.AllowUserToDeleteRows = false;
            this.dataGridViewProcessTime.AllowUserToResizeColumns = false;
            this.dataGridViewProcessTime.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewProcessTime.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewProcessTime.Location = new System.Drawing.Point(3, 2);
            this.dataGridViewProcessTime.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridViewProcessTime.Name = "dataGridViewProcessTime";
            this.dataGridViewProcessTime.ReadOnly = true;
            this.dataGridViewProcessTime.RowHeadersWidth = 51;
            this.dataGridViewProcessTime.RowTemplate.Height = 24;
            this.dataGridViewProcessTime.Size = new System.Drawing.Size(231, 397);
            this.dataGridViewProcessTime.TabIndex = 1;
            this.dataGridViewProcessTime.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewProcessTime_CellContentClick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.deleteBtn);
            this.panel1.Controls.Add(this.addBtn);
            this.panel1.Controls.Add(this.dateTimePickerTimer);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.dataGridViewProcessTime);
            this.panel1.Location = new System.Drawing.Point(16, 101);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(237, 493);
            this.panel1.TabIndex = 2;
            // 
            // deleteBtn
            // 
            this.deleteBtn.Location = new System.Drawing.Point(136, 446);
            this.deleteBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.deleteBtn.Name = "deleteBtn";
            this.deleteBtn.Size = new System.Drawing.Size(85, 28);
            this.deleteBtn.TabIndex = 6;
            this.deleteBtn.Text = "Delete";
            this.deleteBtn.UseVisualStyleBackColor = true;
            this.deleteBtn.Click += new System.EventHandler(this.deleteBtn_Click);
            // 
            // addBtn
            // 
            this.addBtn.Location = new System.Drawing.Point(20, 446);
            this.addBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.addBtn.Name = "addBtn";
            this.addBtn.Size = new System.Drawing.Size(85, 28);
            this.addBtn.TabIndex = 4;
            this.addBtn.Text = "Add";
            this.addBtn.UseVisualStyleBackColor = true;
            this.addBtn.Click += new System.EventHandler(this.addBtn_Click);
            // 
            // dateTimePickerTimer
            // 
            this.dateTimePickerTimer.CustomFormat = "HH:mm";
            this.dateTimePickerTimer.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerTimer.Location = new System.Drawing.Point(91, 415);
            this.dateTimePickerTimer.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dateTimePickerTimer.Name = "dateTimePickerTimer";
            this.dateTimePickerTimer.ShowUpDown = true;
            this.dateTimePickerTimer.Size = new System.Drawing.Size(129, 22);
            this.dateTimePickerTimer.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 415);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "Time";
            // 
            // dataGridViewTransactionEmployee
            // 
            this.dataGridViewTransactionEmployee.AllowUserToAddRows = false;
            this.dataGridViewTransactionEmployee.AllowUserToDeleteRows = false;
            this.dataGridViewTransactionEmployee.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewTransactionEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewTransactionEmployee.Location = new System.Drawing.Point(259, 101);
            this.dataGridViewTransactionEmployee.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridViewTransactionEmployee.Name = "dataGridViewTransactionEmployee";
            this.dataGridViewTransactionEmployee.ReadOnly = true;
            this.dataGridViewTransactionEmployee.RowHeadersWidth = 51;
            this.dataGridViewTransactionEmployee.RowTemplate.Height = 24;
            this.dataGridViewTransactionEmployee.Size = new System.Drawing.Size(788, 499);
            this.dataGridViewTransactionEmployee.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(259, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "Transaction";
            // 
            // timerRefresh
            // 
            this.timerRefresh.Enabled = true;
            this.timerRefresh.Interval = 30000;
            this.timerRefresh.Tick += new System.EventHandler(this.timerRefresh_Tick);
            // 
            // ReaderEnginee
            // 
            this.ReaderEnginee.Icon = ((System.Drawing.Icon)(resources.GetObject("ReaderEnginee.Icon")));
            this.ReaderEnginee.Text = "notifyIcon1";
            this.ReaderEnginee.Visible = true;
            this.ReaderEnginee.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 639);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1059, 21);
            this.progressBar1.TabIndex = 20;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dateTimeNow});
            this.toolStrip1.Location = new System.Drawing.Point(0, 614);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1059, 25);
            this.toolStrip1.TabIndex = 21;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // dateTimeNow
            // 
            this.dateTimeNow.Name = "dateTimeNow";
            this.dateTimeNow.Size = new System.Drawing.Size(103, 22);
            this.dateTimeNow.Text = "dateTimeNow";
            // 
            // timer
            // 
            this.timer.Enabled = true;
            this.timer.Interval = 60000;
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // btnMail
            // 
            this.btnMail.Location = new System.Drawing.Point(651, 65);
            this.btnMail.Margin = new System.Windows.Forms.Padding(4);
            this.btnMail.Name = "btnMail";
            this.btnMail.Size = new System.Drawing.Size(185, 28);
            this.btnMail.TabIndex = 22;
            this.btnMail.Text = "Send eMail";
            this.btnMail.UseVisualStyleBackColor = true;
            this.btnMail.Click += new System.EventHandler(this.btnMail_Click);
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(844, 65);
            this.btnProcess.Margin = new System.Windows.Forms.Padding(4);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(199, 28);
            this.btnProcess.TabIndex = 23;
            this.btnProcess.Text = "Process Log";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(458, 65);
            this.btnExport.Margin = new System.Windows.Forms.Padding(4);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(185, 28);
            this.btnExport.TabIndex = 24;
            this.btnExport.Text = "Export Report";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // dataGridViewESDLog
            // 
            this.dataGridViewESDLog.AllowUserToAddRows = false;
            this.dataGridViewESDLog.AllowUserToDeleteRows = false;
            this.dataGridViewESDLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewESDLog.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewESDLog.Location = new System.Drawing.Point(259, 101);
            this.dataGridViewESDLog.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridViewESDLog.Name = "dataGridViewESDLog";
            this.dataGridViewESDLog.ReadOnly = true;
            this.dataGridViewESDLog.RowHeadersWidth = 51;
            this.dataGridViewESDLog.RowTemplate.Height = 24;
            this.dataGridViewESDLog.Size = new System.Drawing.Size(788, 499);
            this.dataGridViewESDLog.TabIndex = 25;
            this.dataGridViewESDLog.Visible = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.logDataToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1059, 28);
            this.menuStrip1.TabIndex = 26;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // logDataToolStripMenuItem
            // 
            this.logDataToolStripMenuItem.Name = "logDataToolStripMenuItem";
            this.logDataToolStripMenuItem.Size = new System.Drawing.Size(78, 24);
            this.logDataToolStripMenuItem.Text = "Logdata";
            this.logDataToolStripMenuItem.Click += new System.EventHandler(this.logDataToolStripMenuItem_Click);
            // 
            // refrehLabel
            // 
            this.refrehLabel.AutoSize = true;
            this.refrehLabel.Location = new System.Drawing.Point(985, 44);
            this.refrehLabel.Name = "refrehLabel";
            this.refrehLabel.Size = new System.Drawing.Size(58, 17);
            this.refrehLabel.TabIndex = 27;
            this.refrehLabel.TabStop = true;
            this.refrehLabel.Text = "Refresh";
            this.refrehLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.refrehLabel_LinkClicked);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1059, 660);
            this.Controls.Add(this.refrehLabel);
            this.Controls.Add(this.dataGridViewESDLog);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnMail);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dataGridViewTransactionEmployee);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reader Process";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProcessTime)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTransactionEmployee)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewESDLog)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridViewProcessTime;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DateTimePicker dateTimePickerTimer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button deleteBtn;
        private System.Windows.Forms.Button addBtn;
        private System.Windows.Forms.DataGridView dataGridViewTransactionEmployee;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Timer timerRefresh;
        private System.Windows.Forms.NotifyIcon ReaderEnginee;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel dateTimeNow;
        private System.Windows.Forms.Button btnMail;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.DataGridView dataGridViewESDLog;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem logDataToolStripMenuItem;
        public System.Windows.Forms.Timer timer;
        private System.Windows.Forms.LinkLabel refrehLabel;
    }
}

