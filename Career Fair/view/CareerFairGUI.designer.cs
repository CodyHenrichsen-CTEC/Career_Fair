namespace Career_Fair.view
{
    partial class CareerFairGUI
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CareerFairGUI));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutSchedulerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusLabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.workshopFileButton = new System.Windows.Forms.Button();
            this.dateLabel = new System.Windows.Forms.Label();
            this.chooseStudentRequestButton = new System.Windows.Forms.Button();
            this.generateScheduleButton = new System.Windows.Forms.Button();
            this.helpProviderWFC = new System.Windows.Forms.HelpProvider();
            this.launchScheduleButton = new System.Windows.Forms.Button();
            this.loadRequestsButton = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(765, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutSchedulerToolStripMenuItem,
            this.helpFileToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // aboutSchedulerToolStripMenuItem
            // 
            this.aboutSchedulerToolStripMenuItem.Name = "aboutSchedulerToolStripMenuItem";
            this.aboutSchedulerToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
            this.aboutSchedulerToolStripMenuItem.Text = "About Scheduler";
            this.aboutSchedulerToolStripMenuItem.Click += new System.EventHandler(this.aboutSchedulerToolStripMenuItem_Click);
            // 
            // helpFileToolStripMenuItem
            // 
            this.helpFileToolStripMenuItem.Name = "helpFileToolStripMenuItem";
            this.helpFileToolStripMenuItem.Size = new System.Drawing.Size(162, 22);
            this.helpFileToolStripMenuItem.Text = "Help File";
            this.helpFileToolStripMenuItem.Click += new System.EventHandler(this.helpFileToolStripMenuItem_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(308, 221);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(40, 13);
            this.statusLabel.TabIndex = 2;
            this.statusLabel.Text = "Status:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // workshopFileButton
            // 
            this.workshopFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.workshopFileButton.Location = new System.Drawing.Point(61, 221);
            this.workshopFileButton.Name = "workshopFileButton";
            this.workshopFileButton.Size = new System.Drawing.Size(200, 50);
            this.workshopFileButton.TabIndex = 3;
            this.workshopFileButton.Text = "Select data file";
            this.workshopFileButton.UseVisualStyleBackColor = true;
            this.workshopFileButton.Click += new System.EventHandler(this.workshopFileButton_Click);
            // 
            // dateLabel
            // 
            this.dateLabel.AutoSize = true;
            this.dateLabel.Location = new System.Drawing.Point(28, 88);
            this.dateLabel.Name = "dateLabel";
            this.dateLabel.Size = new System.Drawing.Size(0, 13);
            this.dateLabel.TabIndex = 5;
            // 
            // chooseStudentRequestButton
            // 
            this.chooseStudentRequestButton.Enabled = false;
            this.chooseStudentRequestButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chooseStudentRequestButton.Location = new System.Drawing.Point(61, 280);
            this.chooseStudentRequestButton.Name = "chooseStudentRequestButton";
            this.chooseStudentRequestButton.Size = new System.Drawing.Size(200, 50);
            this.chooseStudentRequestButton.TabIndex = 6;
            this.chooseStudentRequestButton.Text = "Choose student request data";
            this.chooseStudentRequestButton.UseVisualStyleBackColor = true;
            this.chooseStudentRequestButton.Click += new System.EventHandler(this.chooseStudentRequestButton_Click);
            // 
            // generateScheduleButton
            // 
            this.generateScheduleButton.Enabled = false;
            this.generateScheduleButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateScheduleButton.Location = new System.Drawing.Point(61, 339);
            this.generateScheduleButton.Name = "generateScheduleButton";
            this.generateScheduleButton.Size = new System.Drawing.Size(200, 50);
            this.generateScheduleButton.TabIndex = 7;
            this.generateScheduleButton.Text = "Generate Career Fair Schedule";
            this.generateScheduleButton.UseVisualStyleBackColor = true;
            this.generateScheduleButton.Click += new System.EventHandler(this.generateScheduleButton_Click);
            // 
            // helpProviderWFC
            // 
            this.helpProviderWFC.HelpNamespace = "C:\\Users\\SQL-Dev\\Documents\\Visual Studio 2010\\Projects\\WFC Scheduler\\WFC Schedule" +
                "r\\Workshops Help File.pdf";
            // 
            // launchScheduleButton
            // 
            this.launchScheduleButton.Enabled = false;
            this.launchScheduleButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.launchScheduleButton.Location = new System.Drawing.Point(311, 313);
            this.launchScheduleButton.Name = "launchScheduleButton";
            this.launchScheduleButton.Size = new System.Drawing.Size(265, 64);
            this.launchScheduleButton.TabIndex = 8;
            this.launchScheduleButton.Text = "Launch Schedule";
            this.launchScheduleButton.UseVisualStyleBackColor = true;
            this.launchScheduleButton.Visible = false;
            this.launchScheduleButton.Click += new System.EventHandler(this.launchScheduleButton_Click);
            // 
            // loadRequestsButton
            // 
            this.loadRequestsButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadRequestsButton.Location = new System.Drawing.Point(49, 67);
            this.loadRequestsButton.Name = "loadRequestsButton";
            this.loadRequestsButton.Size = new System.Drawing.Size(681, 86);
            this.loadRequestsButton.TabIndex = 9;
            this.loadRequestsButton.Text = "Load requests and generate estimate";
            this.loadRequestsButton.UseVisualStyleBackColor = true;
            this.loadRequestsButton.Click += new System.EventHandler(this.loadRequestsButton_Click);
            // 
            // CareerFairGUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.MediumPurple;
            this.ClientSize = new System.Drawing.Size(765, 425);
            this.Controls.Add(this.loadRequestsButton);
            this.Controls.Add(this.launchScheduleButton);
            this.Controls.Add(this.generateScheduleButton);
            this.Controls.Add(this.chooseStudentRequestButton);
            this.Controls.Add(this.dateLabel);
            this.Controls.Add(this.workshopFileButton);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.menuStrip1);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CareerFairGUI";
            this.Opacity = 0.95D;
            this.Text = "Career Fair Scheduler";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button workshopFileButton;
        private System.Windows.Forms.Label dateLabel;
        private System.Windows.Forms.Button chooseStudentRequestButton;
        private System.Windows.Forms.Button generateScheduleButton;
        private System.Windows.Forms.ToolStripMenuItem aboutSchedulerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpFileToolStripMenuItem;
        private System.Windows.Forms.HelpProvider helpProviderWFC;
        private System.Windows.Forms.Button launchScheduleButton;
        private System.Windows.Forms.Button loadRequestsButton;
    }
}

