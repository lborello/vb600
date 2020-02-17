namespace RFIDDemoCS
{
    partial class frmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnOpen = new System.Windows.Forms.Button();
            this.lblVersion = new System.Windows.Forms.Label();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tbpInventory = new System.Windows.Forms.TabPage();
            this.textBoxIP = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.lstView = new System.Windows.Forms.ListView();
            this.colNo = new System.Windows.Forms.ColumnHeader();
            this.colCount = new System.Windows.Forms.ColumnHeader();
            this.colEPC = new System.Windows.Forms.ColumnHeader();
            this.tabWrite = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.btnWrite = new System.Windows.Forms.Button();
            this.txtEPC = new System.Windows.Forms.TextBox();
            this.tbpConfig = new System.Windows.Forms.TabPage();
            this.traPowerLevel = new System.Windows.Forms.TrackBar();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnSet = new System.Windows.Forms.Button();
            this.lbldBm = new System.Windows.Forms.Label();
            this.lblAntPower = new System.Windows.Forms.Label();
            this.txtInvRounds = new System.Windows.Forms.TextBox();
            this.lblInventoryRounds = new System.Windows.Forms.Label();
            this.txtdwelltime = new System.Windows.Forms.TextBox();
            this.lbldwelltime = new System.Windows.Forms.Label();
            this.numTagStopCount = new System.Windows.Forms.NumericUpDown();
            this.lblTagStopCount = new System.Windows.Forms.Label();
            this.cmbOperationMode = new System.Windows.Forms.ComboBox();
            this.lblOperationMode = new System.Windows.Forms.Label();
            this.cmbResponseMode = new System.Windows.Forms.ComboBox();
            this.lblResponseMode = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tabMain.SuspendLayout();
            this.tbpInventory.SuspendLayout();
            this.tabWrite.SuspendLayout();
            this.tbpConfig.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(7, 230);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(60, 22);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "Open";
            this.btnOpen.Visible = false;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // lblVersion
            // 
            this.lblVersion.Location = new System.Drawing.Point(82, 255);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(112, 16);
            this.lblVersion.Text = " ";
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabWrite);
            this.tabMain.Controls.Add(this.tbpInventory);
            this.tabMain.Controls.Add(this.tbpConfig);
            this.tabMain.Location = new System.Drawing.Point(3, 3);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(225, 225);
            this.tabMain.TabIndex = 3;
            this.tabMain.SelectedIndexChanged += new System.EventHandler(this.tabMain_SelectedIndexChanged);
            // 
            // tbpInventory
            // 
            this.tbpInventory.Controls.Add(this.textBoxIP);
            this.tbpInventory.Controls.Add(this.btnSend);
            this.tbpInventory.Controls.Add(this.btnStart);
            this.tbpInventory.Controls.Add(this.lstView);
            this.tbpInventory.Location = new System.Drawing.Point(4, 25);
            this.tbpInventory.Name = "tbpInventory";
            this.tbpInventory.Size = new System.Drawing.Size(217, 196);
            this.tbpInventory.Text = "Inventario";
            // 
            // textBoxIP
            // 
            this.textBoxIP.Location = new System.Drawing.Point(137, 171);
            this.textBoxIP.Name = "textBoxIP";
            this.textBoxIP.Size = new System.Drawing.Size(93, 23);
            this.textBoxIP.TabIndex = 3;
            this.textBoxIP.Text = "192.168.92.98";
            this.textBoxIP.TextChanged += new System.EventHandler(this.textBoxIP_TextChanged);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(70, 171);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(65, 22);
            this.btnSend.TabIndex = 2;
            this.btnSend.Text = "Send";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(3, 171);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(65, 22);
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Inicio";
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // lstView
            // 
            this.lstView.Columns.Add(this.colNo);
            this.lstView.Columns.Add(this.colCount);
            this.lstView.Columns.Add(this.colEPC);
            this.lstView.Location = new System.Drawing.Point(0, 1);
            this.lstView.Name = "lstView";
            this.lstView.Size = new System.Drawing.Size(232, 165);
            this.lstView.TabIndex = 0;
            this.lstView.View = System.Windows.Forms.View.Details;
            this.lstView.SelectedIndexChanged += new System.EventHandler(this.lstView_SelectedIndexChanged);
            // 
            // colNo
            // 
            this.colNo.Text = "No";
            this.colNo.Width = 27;
            // 
            // colCount
            // 
            this.colCount.Text = "Cnt";
            this.colCount.Width = 35;
            // 
            // colEPC
            // 
            this.colEPC.Text = "EPC/ID";
            this.colEPC.Width = 200;
            // 
            // tabWrite
            // 
            this.tabWrite.Controls.Add(this.label1);
            this.tabWrite.Controls.Add(this.pictureBox1);
            this.tabWrite.Controls.Add(this.btnWrite);
            this.tabWrite.Controls.Add(this.txtEPC);
            this.tabWrite.Location = new System.Drawing.Point(4, 25);
            this.tabWrite.Name = "tabWrite";
            this.tabWrite.Size = new System.Drawing.Size(217, 196);
            this.tabWrite.Text = "Escribir";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(3, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 19);
            this.label1.Text = "EPC:";
            // 
            // btnWrite
            // 
            this.btnWrite.Location = new System.Drawing.Point(3, 64);
            this.btnWrite.Name = "btnWrite";
            this.btnWrite.Size = new System.Drawing.Size(184, 27);
            this.btnWrite.TabIndex = 2;
            this.btnWrite.Text = "Escribir";
            this.btnWrite.Click += new System.EventHandler(this.btnWrite_Click);
            // 
            // txtEPC
            // 
            this.txtEPC.Location = new System.Drawing.Point(3, 35);
            this.txtEPC.Name = "txtEPC";
            this.txtEPC.Size = new System.Drawing.Size(184, 23);
            this.txtEPC.TabIndex = 0;
            this.txtEPC.TextChanged += new System.EventHandler(this.txtEPC_TextChanged);
            this.txtEPC.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEPC_KeyDown);
            // 
            // tbpConfig
            // 
            this.tbpConfig.Controls.Add(this.traPowerLevel);
            this.tbpConfig.Controls.Add(this.btnRefresh);
            this.tbpConfig.Controls.Add(this.btnSet);
            this.tbpConfig.Controls.Add(this.lbldBm);
            this.tbpConfig.Controls.Add(this.lblAntPower);
            this.tbpConfig.Controls.Add(this.txtInvRounds);
            this.tbpConfig.Controls.Add(this.lblInventoryRounds);
            this.tbpConfig.Controls.Add(this.txtdwelltime);
            this.tbpConfig.Controls.Add(this.lbldwelltime);
            this.tbpConfig.Controls.Add(this.numTagStopCount);
            this.tbpConfig.Controls.Add(this.lblTagStopCount);
            this.tbpConfig.Controls.Add(this.cmbOperationMode);
            this.tbpConfig.Controls.Add(this.lblOperationMode);
            this.tbpConfig.Controls.Add(this.cmbResponseMode);
            this.tbpConfig.Controls.Add(this.lblResponseMode);
            this.tbpConfig.Location = new System.Drawing.Point(4, 25);
            this.tbpConfig.Name = "tbpConfig";
            this.tbpConfig.Size = new System.Drawing.Size(217, 196);
            this.tbpConfig.Text = "Configuracion";
            // 
            // traPowerLevel
            // 
            this.traPowerLevel.LargeChange = 2;
            this.traPowerLevel.Location = new System.Drawing.Point(10, 140);
            this.traPowerLevel.Maximum = 29;
            this.traPowerLevel.Minimum = 5;
            this.traPowerLevel.Name = "traPowerLevel";
            this.traPowerLevel.Size = new System.Drawing.Size(197, 23);
            this.traPowerLevel.TabIndex = 11;
            this.traPowerLevel.Value = 29;
            this.traPowerLevel.ValueChanged += new System.EventHandler(this.traPowerLevel_ValueChanged);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(124, 170);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(72, 24);
            this.btnRefresh.TabIndex = 14;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnSet
            // 
            this.btnSet.Location = new System.Drawing.Point(19, 170);
            this.btnSet.Name = "btnSet";
            this.btnSet.Size = new System.Drawing.Size(72, 24);
            this.btnSet.TabIndex = 13;
            this.btnSet.Text = "Set";
            this.btnSet.Click += new System.EventHandler(this.btnSet_Click);
            // 
            // lbldBm
            // 
            this.lbldBm.Location = new System.Drawing.Point(109, 122);
            this.lbldBm.Name = "lbldBm";
            this.lbldBm.Size = new System.Drawing.Size(76, 20);
            this.lbldBm.Text = "26dBm";
            // 
            // lblAntPower
            // 
            this.lblAntPower.Location = new System.Drawing.Point(3, 122);
            this.lblAntPower.Name = "lblAntPower";
            this.lblAntPower.Size = new System.Drawing.Size(100, 20);
            this.lblAntPower.Text = "Antenna Power";
            // 
            // txtInvRounds
            // 
            this.txtInvRounds.Location = new System.Drawing.Point(109, 97);
            this.txtInvRounds.Name = "txtInvRounds";
            this.txtInvRounds.Size = new System.Drawing.Size(71, 23);
            this.txtInvRounds.TabIndex = 9;
            this.txtInvRounds.Text = "8912";
            // 
            // lblInventoryRounds
            // 
            this.lblInventoryRounds.Location = new System.Drawing.Point(3, 100);
            this.lblInventoryRounds.Name = "lblInventoryRounds";
            this.lblInventoryRounds.Size = new System.Drawing.Size(113, 20);
            this.lblInventoryRounds.Text = "Inventory Rounds";
            // 
            // txtdwelltime
            // 
            this.txtdwelltime.Location = new System.Drawing.Point(109, 73);
            this.txtdwelltime.Name = "txtdwelltime";
            this.txtdwelltime.Size = new System.Drawing.Size(71, 23);
            this.txtdwelltime.TabIndex = 7;
            this.txtdwelltime.Text = "2000";
            // 
            // lbldwelltime
            // 
            this.lbldwelltime.Location = new System.Drawing.Point(3, 77);
            this.lbldwelltime.Name = "lbldwelltime";
            this.lbldwelltime.Size = new System.Drawing.Size(100, 20);
            this.lbldwelltime.Text = "dwelltime";
            // 
            // numTagStopCount
            // 
            this.numTagStopCount.Location = new System.Drawing.Point(109, 48);
            this.numTagStopCount.Maximum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numTagStopCount.Name = "numTagStopCount";
            this.numTagStopCount.Size = new System.Drawing.Size(71, 24);
            this.numTagStopCount.TabIndex = 5;
            this.numTagStopCount.ValueChanged += new System.EventHandler(this.numTagStopCount_ValueChanged);
            // 
            // lblTagStopCount
            // 
            this.lblTagStopCount.Location = new System.Drawing.Point(3, 52);
            this.lblTagStopCount.Name = "lblTagStopCount";
            this.lblTagStopCount.Size = new System.Drawing.Size(100, 20);
            this.lblTagStopCount.Text = "TagStop Count";
            // 
            // cmbOperationMode
            // 
            this.cmbOperationMode.Items.Add("CONTINUOUS");
            this.cmbOperationMode.Items.Add("NONCONTINUOUS");
            this.cmbOperationMode.Location = new System.Drawing.Point(109, 24);
            this.cmbOperationMode.Name = "cmbOperationMode";
            this.cmbOperationMode.Size = new System.Drawing.Size(98, 23);
            this.cmbOperationMode.TabIndex = 3;
            this.cmbOperationMode.SelectedIndexChanged += new System.EventHandler(this.cmbOperationMode_SelectedIndexChanged);
            // 
            // lblOperationMode
            // 
            this.lblOperationMode.Location = new System.Drawing.Point(3, 28);
            this.lblOperationMode.Name = "lblOperationMode";
            this.lblOperationMode.Size = new System.Drawing.Size(105, 20);
            this.lblOperationMode.Text = "Operation Mode";
            // 
            // cmbResponseMode
            // 
            this.cmbResponseMode.Items.Add("COMPACT");
            this.cmbResponseMode.Items.Add("NORMAL");
            this.cmbResponseMode.Location = new System.Drawing.Point(109, 1);
            this.cmbResponseMode.Name = "cmbResponseMode";
            this.cmbResponseMode.Size = new System.Drawing.Size(98, 23);
            this.cmbResponseMode.TabIndex = 1;
            this.cmbResponseMode.SelectedIndexChanged += new System.EventHandler(this.cmbResponseMode_SelectedIndexChanged);
            // 
            // lblResponseMode
            // 
            this.lblResponseMode.Location = new System.Drawing.Point(3, 4);
            this.lblResponseMode.Name = "lblResponseMode";
            this.lblResponseMode.Size = new System.Drawing.Size(100, 20);
            this.lblResponseMode.Text = "Response Mode";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(68, 230);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(60, 22);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(2, 227);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(48, 16);
            this.lblStatus.Text = "......";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(3, 97);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(184, 92);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(134, 230);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(60, 22);
            this.button1.TabIndex = 10;
            this.button1.Text = "Salir";
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(228, 278);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.tabMain);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.btnOpen);
            this.MinimizeBox = false;
            this.Name = "frmMain";
            this.Text = "Banco de archivos";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Closed += new System.EventHandler(this.frmMain_Closed);
            this.tabMain.ResumeLayout(false);
            this.tbpInventory.ResumeLayout(false);
            this.tabWrite.ResumeLayout(false);
            this.tbpConfig.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tbpInventory;
        private System.Windows.Forms.ListView lstView;
        private System.Windows.Forms.TabPage tbpConfig;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ColumnHeader colNo;
        private System.Windows.Forms.ColumnHeader colEPC;
        private System.Windows.Forms.ColumnHeader colCount;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cmbResponseMode;
        private System.Windows.Forms.Label lblResponseMode;
        private System.Windows.Forms.Label lblTagStopCount;
        private System.Windows.Forms.ComboBox cmbOperationMode;
        private System.Windows.Forms.Label lblOperationMode;
        private System.Windows.Forms.Label lblInventoryRounds;
        private System.Windows.Forms.TextBox txtdwelltime;
        private System.Windows.Forms.Label lbldwelltime;
        private System.Windows.Forms.NumericUpDown numTagStopCount;
        private System.Windows.Forms.TrackBar traPowerLevel;
        private System.Windows.Forms.Label lblAntPower;
        private System.Windows.Forms.TextBox txtInvRounds;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnSet;
        private System.Windows.Forms.Label lbldBm;
        private System.Windows.Forms.TextBox textBoxIP;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.TabPage tabWrite;
        private System.Windows.Forms.Button btnWrite;
        private System.Windows.Forms.TextBox txtEPC;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button1;       
    }
}

