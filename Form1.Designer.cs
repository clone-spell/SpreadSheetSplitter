namespace SpreadSheetSplitter
{
    partial class Form1
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
            this.btnSplit = new System.Windows.Forms.Button();
            this.cbColumn = new System.Windows.Forms.ComboBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.cbSheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnChoseFile = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // btnSplit
            // 
            this.btnSplit.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnSplit.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSplit.Location = new System.Drawing.Point(78, 258);
            this.btnSplit.Name = "btnSplit";
            this.btnSplit.Size = new System.Drawing.Size(115, 31);
            this.btnSplit.TabIndex = 3;
            this.btnSplit.Text = "Split";
            this.btnSplit.UseVisualStyleBackColor = true;
            this.btnSplit.Click += new System.EventHandler(this.btnSplit_Click);
            // 
            // cbColumn
            // 
            this.cbColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbColumn.FormattingEnabled = true;
            this.cbColumn.Location = new System.Drawing.Point(12, 153);
            this.cbColumn.Name = "cbColumn";
            this.cbColumn.Size = new System.Drawing.Size(368, 25);
            this.cbColumn.TabIndex = 2;
            // 
            // btnClear
            // 
            this.btnClear.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnClear.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.Location = new System.Drawing.Point(199, 258);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(115, 31);
            this.btnClear.TabIndex = 0;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // cbSheet
            // 
            this.cbSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSheet.FormattingEnabled = true;
            this.cbSheet.Location = new System.Drawing.Point(12, 95);
            this.cbSheet.Name = "cbSheet";
            this.cbSheet.Size = new System.Drawing.Size(368, 25);
            this.cbSheet.TabIndex = 2;
            this.cbSheet.SelectedIndexChanged += new System.EventHandler(this.cbSheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Select Sheet";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(159, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select Column With Date";
            // 
            // btnChoseFile
            // 
            this.btnChoseFile.ForeColor = System.Drawing.SystemColors.GrayText;
            this.btnChoseFile.Location = new System.Drawing.Point(12, 34);
            this.btnChoseFile.Name = "btnChoseFile";
            this.btnChoseFile.Size = new System.Drawing.Size(368, 27);
            this.btnChoseFile.TabIndex = 5;
            this.btnChoseFile.Text = "Chose File";
            this.btnChoseFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnChoseFile.UseVisualStyleBackColor = true;
            this.btnChoseFile.Click += new System.EventHandler(this.btnChoseFile_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "Select Excel File";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 214);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(368, 19);
            this.progressBar.TabIndex = 6;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(12, 194);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(0, 17);
            this.lblProgress.TabIndex = 4;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AcceptButton = this.btnSplit;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 301);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnChoseFile);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbSheet);
            this.Controls.Add(this.cbColumn);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnSplit);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Spreadsheet Splitter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnSplit;
        private System.Windows.Forms.ComboBox cbColumn;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ComboBox cbSheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnChoseFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblProgress;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
    }
}

