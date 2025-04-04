namespace BMI.UTimeMaster
{
    partial class frm_InputOutput
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.btn_Elave = new System.Windows.Forms.Button();
            this.dtg_InputOutput = new System.Windows.Forms.DataGridView();
            this.cl_AD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cl_Tarix = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cl_giris = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cl_cixis = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cl_gecikme = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtg_InputOutput)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1526, 30);
            this.panel1.TabIndex = 10;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox1.Image = global::BMI.Properties.Resources.systemshutdown_104277;
            this.pictureBox1.Location = new System.Drawing.Point(1490, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(28, 25);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.monthCalendar1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.monthCalendar1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.monthCalendar1.Location = new System.Drawing.Point(1179, 37);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.ShowToday = false;
            this.monthCalendar1.TabIndex = 23;
            this.monthCalendar1.TitleBackColor = System.Drawing.SystemColors.Control;
            this.monthCalendar1.TitleForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.monthCalendar1.TodayDate = new System.DateTime(((long)(0)));
            this.monthCalendar1.TrailingForeColor = System.Drawing.Color.PaleTurquoise;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // btn_Elave
            // 
            this.btn_Elave.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Elave.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btn_Elave.FlatAppearance.BorderSize = 2;
            this.btn_Elave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.PaleTurquoise;
            this.btn_Elave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Elave.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Elave.Location = new System.Drawing.Point(1179, 229);
            this.btn_Elave.Name = "btn_Elave";
            this.btn_Elave.Size = new System.Drawing.Size(169, 30);
            this.btn_Elave.TabIndex = 24;
            this.btn_Elave.Text = "Excel";
            this.btn_Elave.UseVisualStyleBackColor = true;
            this.btn_Elave.Click += new System.EventHandler(this.btn_Elave_Click);
            // 
            // dtg_InputOutput
            // 
            this.dtg_InputOutput.AllowUserToAddRows = false;
            this.dtg_InputOutput.AllowUserToDeleteRows = false;
            this.dtg_InputOutput.AllowUserToResizeColumns = false;
            this.dtg_InputOutput.AllowUserToResizeRows = false;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtg_InputOutput.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dtg_InputOutput.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtg_InputOutput.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cl_AD,
            this.cl_Tarix,
            this.cl_giris,
            this.cl_cixis,
            this.cl_gecikme});
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtg_InputOutput.DefaultCellStyle = dataGridViewCellStyle5;
            this.dtg_InputOutput.Location = new System.Drawing.Point(1, 37);
            this.dtg_InputOutput.Name = "dtg_InputOutput";
            this.dtg_InputOutput.ReadOnly = true;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtg_InputOutput.RowHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.dtg_InputOutput.Size = new System.Drawing.Size(1166, 649);
            this.dtg_InputOutput.TabIndex = 25;
            // 
            // cl_AD
            // 
            this.cl_AD.HeaderText = "Ad Soyad";
            this.cl_AD.Name = "cl_AD";
            this.cl_AD.ReadOnly = true;
            this.cl_AD.Width = 300;
            // 
            // cl_Tarix
            // 
            this.cl_Tarix.HeaderText = "Tarix";
            this.cl_Tarix.Name = "cl_Tarix";
            this.cl_Tarix.ReadOnly = true;
            this.cl_Tarix.Width = 200;
            // 
            // cl_giris
            // 
            this.cl_giris.HeaderText = "Giriş Saat";
            this.cl_giris.Name = "cl_giris";
            this.cl_giris.ReadOnly = true;
            this.cl_giris.Width = 200;
            // 
            // cl_cixis
            // 
            this.cl_cixis.HeaderText = "Çıxış Saat";
            this.cl_cixis.Name = "cl_cixis";
            this.cl_cixis.ReadOnly = true;
            this.cl_cixis.Width = 200;
            // 
            // cl_gecikme
            // 
            this.cl_gecikme.HeaderText = "Gecikmə";
            this.cl_gecikme.Name = "cl_gecikme";
            this.cl_gecikme.ReadOnly = true;
            this.cl_gecikme.Width = 200;
            // 
            // frm_InputOutput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1529, 698);
            this.Controls.Add(this.dtg_InputOutput);
            this.Controls.Add(this.btn_Elave);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frm_InputOutput";
            this.Text = "frm_InputOutput";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtg_InputOutput)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.Button btn_Elave;
        private System.Windows.Forms.DataGridView dtg_InputOutput;
        private System.Windows.Forms.DataGridViewTextBoxColumn cl_AD;
        private System.Windows.Forms.DataGridViewTextBoxColumn cl_Tarix;
        private System.Windows.Forms.DataGridViewTextBoxColumn cl_giris;
        private System.Windows.Forms.DataGridViewTextBoxColumn cl_cixis;
        private System.Windows.Forms.DataGridViewTextBoxColumn cl_gecikme;
    }
}