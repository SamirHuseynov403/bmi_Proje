namespace BMI
{
    partial class frmehtiyyatlaruzresorgu
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
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridmuracietler)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.gridControl1);
            this.panel1.Location = new System.Drawing.Point(2, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1250, 445);
            this.panel1.TabIndex = 0;
            // 
            // gridControl1
            // 
            this.gridControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridControl1.Location = new System.Drawing.Point(0, 0);
            this.gridControl1.MainView = this.gridmuracietler;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(1250, 445);
            this.gridControl1.TabIndex = 22;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridmuracietler});
            // 
            // gridmuracietler
            // 
            this.gridmuracietler.Appearance.FocusedRow.BackColor = System.Drawing.Color.Transparent;
            this.gridmuracietler.Appearance.FocusedRow.BorderColor = System.Drawing.Color.Red;
            this.gridmuracietler.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.gridmuracietler.Appearance.FocusedRow.ForeColor = System.Drawing.Color.Red;
            this.gridmuracietler.Appearance.FocusedRow.Options.UseBackColor = true;
            this.gridmuracietler.Appearance.FocusedRow.Options.UseBorderColor = true;
            this.gridmuracietler.Appearance.FocusedRow.Options.UseFont = true;
            this.gridmuracietler.Appearance.FocusedRow.Options.UseForeColor = true;
            this.gridmuracietler.Appearance.OddRow.Font = new System.Drawing.Font("Tahoma", 11.25F);
            this.gridmuracietler.Appearance.OddRow.Options.UseFont = true;
            this.gridmuracietler.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.gridmuracietler.Appearance.Preview.Options.UseFont = true;
            this.gridmuracietler.Appearance.Row.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gridmuracietler.Appearance.Row.Options.UseFont = true;
            this.gridmuracietler.Appearance.SelectedRow.BackColor = System.Drawing.Color.Red;
            this.gridmuracietler.Appearance.SelectedRow.BorderColor = System.Drawing.Color.Red;
            this.gridmuracietler.Appearance.SelectedRow.ForeColor = System.Drawing.Color.Red;
            this.gridmuracietler.Appearance.SelectedRow.Options.UseBackColor = true;
            this.gridmuracietler.Appearance.SelectedRow.Options.UseBorderColor = true;
            this.gridmuracietler.Appearance.SelectedRow.Options.UseForeColor = true;
            this.gridmuracietler.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFullFocus;
            this.gridmuracietler.GridControl = this.gridControl1;
            this.gridmuracietler.Name = "gridmuracietler";
            this.gridmuracietler.OptionsView.ColumnAutoWidth = false;
            this.gridmuracietler.OptionsView.ShowAutoFilterRow = true;
            this.gridmuracietler.OptionsView.ShowFooter = true;
            this.gridmuracietler.OptionsView.ShowGroupPanel = false;
            // 
            // frmehtiyyatlaruzresorgu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1255, 604);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmehtiyyatlaruzresorgu";
            this.Text = "frmehtiyyatlaruzresorgu";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridmuracietler)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        public DevExpress.XtraGrid.GridControl gridControl1;
        public DevExpress.XtraGrid.Views.Grid.GridView gridmuracietler;
    }
}