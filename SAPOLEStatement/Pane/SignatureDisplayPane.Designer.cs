namespace SAPOLEStatement
{
    partial class SignatureDisplayPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlSignatures = new System.Windows.Forms.Panel();
            this.lblTotalSignatures = new System.Windows.Forms.Label();
            this.numSignatures = new System.Windows.Forms.NumericUpDown();
            this.btnPath = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.pnlButtons = new System.Windows.Forms.Panel();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnDeleteImage = new System.Windows.Forms.Button();
            this.btnLockImage = new System.Windows.Forms.Button();
            this.btnSaveImage = new System.Windows.Forms.Button();
            this.pnlSignatures.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numSignatures)).BeginInit();
            this.pnlButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlSignatures
            // 
            this.pnlSignatures.AutoScroll = true;
            this.pnlSignatures.AutoScrollMargin = new System.Drawing.Size(5, 5);
            this.pnlSignatures.Controls.Add(this.lblTotalSignatures);
            this.pnlSignatures.Controls.Add(this.numSignatures);
            this.pnlSignatures.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlSignatures.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnlSignatures.Location = new System.Drawing.Point(0, 0);
            this.pnlSignatures.Name = "pnlSignatures";
            this.pnlSignatures.Size = new System.Drawing.Size(365, 349);
            this.pnlSignatures.TabIndex = 60;
            // 
            // lblTotalSignatures
            // 
            this.lblTotalSignatures.AutoSize = true;
            this.lblTotalSignatures.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalSignatures.Location = new System.Drawing.Point(21, 26);
            this.lblTotalSignatures.Name = "lblTotalSignatures";
            this.lblTotalSignatures.Size = new System.Drawing.Size(139, 13);
            this.lblTotalSignatures.TabIndex = 46;
            this.lblTotalSignatures.Text = "Numbers of Signatures:";
            // 
            // numSignatures
            // 
            this.numSignatures.Enabled = false;
            this.numSignatures.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numSignatures.Location = new System.Drawing.Point(166, 24);
            this.numSignatures.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.numSignatures.Name = "numSignatures";
            this.numSignatures.Size = new System.Drawing.Size(40, 20);
            this.numSignatures.TabIndex = 45;
            this.numSignatures.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnPath
            // 
            this.btnPath.Location = new System.Drawing.Point(299, 71);
            this.btnPath.Name = "btnPath";
            this.btnPath.Size = new System.Drawing.Size(50, 23);
            this.btnPath.TabIndex = 59;
            this.btnPath.Text = "Browse";
            this.btnPath.UseVisualStyleBackColor = true;
            this.btnPath.Click += new System.EventHandler(this.btnPath_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(46, 71);
            this.txtPath.Name = "txtPath";
            this.txtPath.ReadOnly = true;
            this.txtPath.Size = new System.Drawing.Size(247, 20);
            this.txtPath.TabIndex = 58;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(13, 76);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(26, 13);
            this.lblPath.TabIndex = 57;
            this.lblPath.Text = "File:";
            // 
            // pnlButtons
            // 
            this.pnlButtons.Controls.Add(this.btnPath);
            this.pnlButtons.Controls.Add(this.txtPath);
            this.pnlButtons.Controls.Add(this.lblPath);
            this.pnlButtons.Controls.Add(this.lblInfo);
            this.pnlButtons.Controls.Add(this.btnDeleteImage);
            this.pnlButtons.Controls.Add(this.btnLockImage);
            this.pnlButtons.Controls.Add(this.btnSaveImage);
            this.pnlButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlButtons.Location = new System.Drawing.Point(0, 358);
            this.pnlButtons.Name = "pnlButtons";
            this.pnlButtons.Size = new System.Drawing.Size(365, 132);
            this.pnlButtons.TabIndex = 59;
            // 
            // lblInfo
            // 
            this.lblInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.Location = new System.Drawing.Point(8, 14);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(341, 54);
            this.lblInfo.TabIndex = 56;
            this.lblInfo.Text = "All signatures will be embedded into the word document, which will be converted t" +
                "o PDF format.  Signatures will not be saved in the word document.";
            // 
            // btnDeleteImage
            // 
            this.btnDeleteImage.Location = new System.Drawing.Point(193, 102);
            this.btnDeleteImage.Name = "btnDeleteImage";
            this.btnDeleteImage.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteImage.TabIndex = 43;
            this.btnDeleteImage.Text = "Clear";
            this.btnDeleteImage.UseVisualStyleBackColor = true;
            this.btnDeleteImage.Click += new System.EventHandler(this.btnDeleteImage_Click);
            // 
            // btnLockImage
            // 
            this.btnLockImage.Location = new System.Drawing.Point(11, 102);
            this.btnLockImage.Name = "btnLockImage";
            this.btnLockImage.Size = new System.Drawing.Size(75, 23);
            this.btnLockImage.TabIndex = 42;
            this.btnLockImage.Text = "Lock";
            this.btnLockImage.UseVisualStyleBackColor = true;
            this.btnLockImage.Visible = false;
            this.btnLockImage.Click += new System.EventHandler(this.btnLockImage_Click);
            // 
            // btnSaveImage
            // 
            this.btnSaveImage.Location = new System.Drawing.Point(274, 102);
            this.btnSaveImage.Name = "btnSaveImage";
            this.btnSaveImage.Size = new System.Drawing.Size(75, 23);
            this.btnSaveImage.TabIndex = 41;
            this.btnSaveImage.Text = "Save";
            this.btnSaveImage.UseVisualStyleBackColor = true;
            this.btnSaveImage.Click += new System.EventHandler(this.btnSaveImage_Click);
            // 
            // SignatureDisplayPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlSignatures);
            this.Controls.Add(this.pnlButtons);
            this.Name = "SignatureDisplayPane";
            this.Size = new System.Drawing.Size(365, 490);
            this.pnlSignatures.ResumeLayout(false);
            this.pnlSignatures.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numSignatures)).EndInit();
            this.pnlButtons.ResumeLayout(false);
            this.pnlButtons.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlSignatures;
        private System.Windows.Forms.Label lblTotalSignatures;
        private System.Windows.Forms.NumericUpDown numSignatures;
        private System.Windows.Forms.Button btnPath;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Panel pnlButtons;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnDeleteImage;
        private System.Windows.Forms.Button btnLockImage;
        private System.Windows.Forms.Button btnSaveImage;
    }
}
