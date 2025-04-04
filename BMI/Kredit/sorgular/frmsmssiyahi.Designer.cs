namespace BMI
{
    partial class frmsmssiyahi
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridmuracietler = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.txtaxtar = new DevExpress.XtraEditors.SimpleButton();
            this.btnexcel = new DevExpress.XtraEditors.SimpleButton();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridmuracietler)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.gridControl1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1405, 441);
            this.panel1.TabIndex = 0;
            // 
            // gridControl1
            // 
            this.gridControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridControl1.Location = new System.Drawing.Point(0, 0);
            this.gridControl1.MainView = this.gridmuracietler;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(1405, 441);
            this.gridControl1.TabIndex = 1;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridmuracietler});
            // 
            // gridmuracietler
            // 
            this.gridmuracietler.Appearance.OddRow.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridmuracietler.Appearance.OddRow.Options.UseFont = true;
            this.gridmuracietler.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridmuracietler.Appearance.Preview.Options.UseFont = true;
            this.gridmuracietler.GridControl = this.gridControl1;
            this.gridmuracietler.Name = "gridmuracietler";
            this.gridmuracietler.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
            this.gridmuracietler.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False;
            this.gridmuracietler.OptionsView.ColumnAutoWidth = false;
            this.gridmuracietler.OptionsView.ShowAutoFilterRow = true;
            this.gridmuracietler.OptionsView.ShowFooter = true;
            this.gridmuracietler.OptionsView.ShowGroupPanel = false;
            // 
            // txtaxtar
            // 
            this.txtaxtar.Appearance.Font = new System.Drawing.Font("Tahoma", 11F);
            this.txtaxtar.Appearance.Options.UseFont = true;
            this.txtaxtar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.txtaxtar.Location = new System.Drawing.Point(883, 479);
            this.txtaxtar.Name = "txtaxtar";
            this.txtaxtar.Size = new System.Drawing.Size(166, 27);
            this.txtaxtar.TabIndex = 5;
            this.txtaxtar.Text = "Sorğu borcalan";
            this.txtaxtar.Click += new System.EventHandler(this.txtaxtar_Click);
            // 
            // btnexcel
            // 
            this.btnexcel.Appearance.Font = new System.Drawing.Font("Tahoma", 11F);
            this.btnexcel.Appearance.Options.UseFont = true;
            this.btnexcel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnexcel.Enabled = false;
            this.btnexcel.Location = new System.Drawing.Point(1227, 479);
            this.btnexcel.Name = "btnexcel";
            this.btnexcel.Size = new System.Drawing.Size(166, 27);
            this.btnexcel.TabIndex = 6;
            this.btnexcel.Text = "Excel";
            this.btnexcel.Click += new System.EventHandler(this.btnexcel_Click);
            // 
            // simpleButton1
            // 
            this.simpleButton1.Appearance.Font = new System.Drawing.Font("Tahoma", 11F);
            this.simpleButton1.Appearance.Options.UseFont = true;
            this.simpleButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.simpleButton1.Location = new System.Drawing.Point(1055, 479);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.Size = new System.Drawing.Size(166, 27);
            this.simpleButton1.TabIndex = 7;
            this.simpleButton1.Text = "Sorğu Zamin";
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.label1.Font = new System.Drawing.Font("Tahoma", 10F);
            this.label1.Location = new System.Drawing.Point(11, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(265, 68);
            this.label1.TabIndex = 8;
            this.label1.Text = "Sorgu sütunun da nömrelerin düz ve ya ferqli qeyd edilmesi                    \r" +
    "\nonların sayları esasında müeyyen edilibMutleq nezerden kecirin";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.label1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 441);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(293, 77);
            this.panel2.TabIndex = 9;
            // 
            // frmsmssiyahi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1405, 518);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.simpleButton1);
            this.Controls.Add(this.btnexcel);
            this.Controls.Add(this.txtaxtar);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmsmssiyahi";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gecikməsi  olan borcalan veə zamin məlumatları";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridmuracietler)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridmuracietler;
        private DevExpress.XtraEditors.SimpleButton txtaxtar;
        private DevExpress.XtraEditors.SimpleButton btnexcel;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel2;
    }
}