namespace OfficeHelper
{
    partial class MainForm
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
            this.crystalTheme1 = new Telerik.WinControls.Themes.CrystalTheme();
            this.visualStudio2012DarkTheme1 = new Telerik.WinControls.Themes.VisualStudio2012DarkTheme();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.radButton1 = new Telerik.WinControls.UI.RadButton();
            this.radButton2 = new Telerik.WinControls.UI.RadButton();
            this.radProgressBar1 = new Telerik.WinControls.UI.RadProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.radLabel1 = new Telerik.WinControls.UI.RadLabel();
            this.radLabel2 = new Telerik.WinControls.UI.RadLabel();
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radProgressBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Word documents (*.docx)|";
            // 
            // radButton1
            // 
            this.radButton1.Location = new System.Drawing.Point(117, 146);
            this.radButton1.Name = "radButton1";
            this.radButton1.Size = new System.Drawing.Size(130, 69);
            this.radButton1.TabIndex = 0;
            this.radButton1.Text = "Select docx file !";
            this.radButton1.ThemeName = "VisualStudio2012Dark";
            this.radButton1.Click += new System.EventHandler(this.radButton1_Click);
            // 
            // radButton2
            // 
            this.radButton2.Location = new System.Drawing.Point(57, 252);
            this.radButton2.Name = "radButton2";
            this.radButton2.Size = new System.Drawing.Size(244, 50);
            this.radButton2.TabIndex = 1;
            this.radButton2.Text = "START";
            this.radButton2.ThemeName = "VisualStudio2012Dark";
            this.radButton2.Click += new System.EventHandler(this.radButton2_Click);
            // 
            // radProgressBar1
            // 
            this.radProgressBar1.Location = new System.Drawing.Point(-2, 319);
            this.radProgressBar1.Name = "radProgressBar1";
            this.radProgressBar1.Size = new System.Drawing.Size(367, 39);
            this.radProgressBar1.TabIndex = 2;
            this.radProgressBar1.ThemeName = "VisualStudio2012Dark";
            // 
            // radLabel1
            // 
            this.radLabel1.Location = new System.Drawing.Point(12, 12);
            this.radLabel1.Name = "radLabel1";
            this.radLabel1.Size = new System.Drawing.Size(332, 18);
            this.radLabel1.TabIndex = 3;
            this.radLabel1.Text = "Notation: Be careful that  your Microsoft Office must be activated";
            this.radLabel1.ThemeName = "VisualStudio2012Dark";
            // 
            // radLabel2
            // 
            this.radLabel2.Location = new System.Drawing.Point(15, 33);
            this.radLabel2.Name = "radLabel2";
            this.radLabel2.Size = new System.Drawing.Size(119, 18);
            this.radLabel2.TabIndex = 3;
            this.radLabel2.Text = "before you could start.";
            this.radLabel2.ThemeName = "VisualStudio2012Dark";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 355);
            this.Controls.Add(this.radLabel2);
            this.Controls.Add(this.radLabel1);
            this.Controls.Add(this.radProgressBar1);
            this.Controls.Add(this.radButton2);
            this.Controls.Add(this.radButton1);
            this.MaximumSize = new System.Drawing.Size(371, 388);
            this.Name = "MainForm";
            // 
            // 
            // 
            this.RootElement.ApplyShapeToControl = true;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "MainForm";
            this.ThemeName = "VisualStudio2012Dark";
            this.TransparencyKey = System.Drawing.Color.White;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.radButton1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radButton2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radProgressBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radLabel2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.Themes.CrystalTheme crystalTheme1;
        private Telerik.WinControls.Themes.VisualStudio2012DarkTheme visualStudio2012DarkTheme1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Telerik.WinControls.UI.RadButton radButton1;
        private Telerik.WinControls.UI.RadButton radButton2;
        private Telerik.WinControls.UI.RadProgressBar radProgressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Telerik.WinControls.UI.RadLabel radLabel1;
        private Telerik.WinControls.UI.RadLabel radLabel2;
    }
}
