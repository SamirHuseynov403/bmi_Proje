namespace BMI.Muhasibat
{
    partial class frm_reziden_ve_qeyri_rezident
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
            this.panelCalendar = new System.Windows.Forms.Panel();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.btn_sorgu = new System.Windows.Forms.Button();
            this.panelCalendar.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelCalendar
            // 
            this.panelCalendar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelCalendar.Controls.Add(this.monthCalendar1);
            this.panelCalendar.Controls.Add(this.btn_sorgu);
            this.panelCalendar.Location = new System.Drawing.Point(11, 11);
            this.panelCalendar.Name = "panelCalendar";
            this.panelCalendar.Size = new System.Drawing.Size(249, 238);
            this.panelCalendar.TabIndex = 28;
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.monthCalendar1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.monthCalendar1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.monthCalendar1.Location = new System.Drawing.Point(9, 9);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.ShowToday = false;
            this.monthCalendar1.TabIndex = 1;
            this.monthCalendar1.TitleBackColor = System.Drawing.SystemColors.Control;
            this.monthCalendar1.TitleForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.monthCalendar1.TodayDate = new System.DateTime(((long)(0)));
            this.monthCalendar1.TrailingForeColor = System.Drawing.Color.PaleTurquoise;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected_1);
            // 
            // btn_sorgu
            // 
            this.btn_sorgu.BackColor = System.Drawing.Color.Ivory;
            this.btn_sorgu.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_sorgu.Enabled = false;
            this.btn_sorgu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_sorgu.Location = new System.Drawing.Point(9, 195);
            this.btn_sorgu.Name = "btn_sorgu";
            this.btn_sorgu.Size = new System.Drawing.Size(228, 32);
            this.btn_sorgu.TabIndex = 3;
            this.btn_sorgu.Text = "Sorğu";
            this.btn_sorgu.UseVisualStyleBackColor = false;
            this.btn_sorgu.Click += new System.EventHandler(this.btn_sorgu_Click);
            // 
            // frm_reziden_ve_qeyri_rezident
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(271, 256);
            this.Controls.Add(this.panelCalendar);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_reziden_ve_qeyri_rezident";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Reziden ve qeyri rezident hesab qalıqları";
            this.panelCalendar.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelCalendar;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        public System.Windows.Forms.Button btn_sorgu;
    }
}