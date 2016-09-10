namespace TimberlineLooper
{
    partial class TimberlineLooper_Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TimberlineLooper_Main));
            this.btnDoWork = new System.Windows.Forms.Button();
            this.chkBox1 = new System.Windows.Forms.CheckBox();
            this.chkBox2 = new System.Windows.Forms.CheckBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusText = new System.Windows.Forms.ToolStripStatusLabel();
            this.rTxtBoxStatus = new System.Windows.Forms.RichTextBox();
            this.tmrGather = new System.Windows.Forms.Timer(this.components);
            this.tmrProcess = new System.Windows.Forms.Timer(this.components);
            this.tmrBatch = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDoWork
            // 
            this.btnDoWork.Location = new System.Drawing.Point(272, 88);
            this.btnDoWork.Name = "btnDoWork";
            this.btnDoWork.Size = new System.Drawing.Size(75, 23);
            this.btnDoWork.TabIndex = 0;
            this.btnDoWork.Text = "Do Work";
            this.btnDoWork.UseVisualStyleBackColor = true;
            this.btnDoWork.Click += new System.EventHandler(this.button1_Click);
            // 
            // chkBox1
            // 
            this.chkBox1.AutoSize = true;
            this.chkBox1.Checked = true;
            this.chkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBox1.Location = new System.Drawing.Point(283, 42);
            this.chkBox1.Name = "chkBox1";
            this.chkBox1.Size = new System.Drawing.Size(58, 17);
            this.chkBox1.TabIndex = 1;
            this.chkBox1.Text = "Gather";
            this.chkBox1.UseVisualStyleBackColor = true;
            this.chkBox1.CheckedChanged += new System.EventHandler(this.chkBox1_CheckedChanged);
            // 
            // chkBox2
            // 
            this.chkBox2.AutoSize = true;
            this.chkBox2.Checked = true;
            this.chkBox2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBox2.Location = new System.Drawing.Point(283, 65);
            this.chkBox2.Name = "chkBox2";
            this.chkBox2.Size = new System.Drawing.Size(64, 17);
            this.chkBox2.TabIndex = 2;
            this.chkBox2.Text = "Process";
            this.chkBox2.UseVisualStyleBackColor = true;
            this.chkBox2.CheckedChanged += new System.EventHandler(this.chkBox2_CheckedChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(646, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "&Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusText});
            this.statusStrip1.Location = new System.Drawing.Point(0, 308);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(646, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // StatusText
            // 
            this.StatusText.Name = "StatusText";
            this.StatusText.Size = new System.Drawing.Size(45, 17);
            this.StatusText.Text = "Status: ";
            // 
            // rTxtBoxStatus
            // 
            this.rTxtBoxStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rTxtBoxStatus.Location = new System.Drawing.Point(12, 147);
            this.rTxtBoxStatus.Name = "rTxtBoxStatus";
            this.rTxtBoxStatus.ReadOnly = true;
            this.rTxtBoxStatus.Size = new System.Drawing.Size(622, 158);
            this.rTxtBoxStatus.TabIndex = 5;
            this.rTxtBoxStatus.Text = "";
            this.rTxtBoxStatus.TextChanged += new System.EventHandler(this.rTxtBoxStatus_TextChanged);
            // 
            // tmrGather
            // 
            this.tmrGather.Interval = 900000;
            this.tmrGather.Tick += new System.EventHandler(this.tmrGather_Tick);
            // 
            // tmrProcess
            // 
            this.tmrProcess.Interval = 900000;
            this.tmrProcess.Tick += new System.EventHandler(this.tmrProcess_Tick);
            // 
            // tmrBatch
            // 
            this.tmrBatch.Interval = 900000;
            this.tmrBatch.Tick += new System.EventHandler(this.tmrBatch_Tick);
            // 
            // TimberlineLooper_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(646, 330);
            this.Controls.Add(this.rTxtBoxStatus);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.chkBox2);
            this.Controls.Add(this.chkBox1);
            this.Controls.Add(this.btnDoWork);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(662, 369);
            this.MinimumSize = new System.Drawing.Size(662, 369);
            this.Name = "TimberlineLooper_Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Timberline Looper";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDoWork;
        private System.Windows.Forms.CheckBox chkBox1;
        private System.Windows.Forms.CheckBox chkBox2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel StatusText;
        private System.Windows.Forms.RichTextBox rTxtBoxStatus;
        private System.Windows.Forms.Timer tmrGather;
        private System.Windows.Forms.Timer tmrProcess;
        private System.Windows.Forms.Timer tmrBatch;
    }
}

