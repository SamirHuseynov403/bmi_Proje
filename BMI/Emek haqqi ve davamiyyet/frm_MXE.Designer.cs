namespace BMI.Emek_haqqi_ve_davamiyyet
{
    partial class frm_MXE
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.frm_MXE_dtg = new System.Windows.Forms.DataGridView();
            this.txt_SSN = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btn_Elave = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.cl_Info = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewButtonColumn();
            this.col_QNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Soyad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_SSN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Fin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_dax_tar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_Staj = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_M_Q_G = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.frm_MXE_dtg)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1526, 30);
            this.panel1.TabIndex = 6;
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
            // frm_MXE_dtg
            // 
            this.frm_MXE_dtg.AllowUserToAddRows = false;
            this.frm_MXE_dtg.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.frm_MXE_dtg.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.frm_MXE_dtg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.frm_MXE_dtg.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cl_Info,
            this.Column1,
            this.col_QNo,
            this.col_Soyad,
            this.col_SSN,
            this.col_Fin,
            this.col_dax_tar,
            this.col_Staj,
            this.col_M_Q_G});
            this.frm_MXE_dtg.Location = new System.Drawing.Point(12, 37);
            this.frm_MXE_dtg.MultiSelect = false;
            this.frm_MXE_dtg.Name = "frm_MXE_dtg";
            this.frm_MXE_dtg.ReadOnly = true;
            this.frm_MXE_dtg.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.frm_MXE_dtg.Size = new System.Drawing.Size(1286, 649);
            this.frm_MXE_dtg.TabIndex = 7;
            this.frm_MXE_dtg.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.frm_MXE_dtg_CellClick);
            this.frm_MXE_dtg.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.frm_MXE_dtg_CellContentClick);
            // 
            // txt_SSN
            // 
            this.txt_SSN.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_SSN.Location = new System.Drawing.Point(1307, 57);
            this.txt_SSN.Name = "txt_SSN";
            this.txt_SSN.Size = new System.Drawing.Size(112, 23);
            this.txt_SSN.TabIndex = 15;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(1304, 36);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(118, 18);
            this.label7.TabIndex = 14;
            this.label7.Text = "Q/N üzrə axtar:";
            // 
            // btn_Elave
            // 
            this.btn_Elave.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Elave.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btn_Elave.FlatAppearance.BorderSize = 2;
            this.btn_Elave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.PaleTurquoise;
            this.btn_Elave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Elave.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Elave.Location = new System.Drawing.Point(1304, 86);
            this.btn_Elave.Name = "btn_Elave";
            this.btn_Elave.Size = new System.Drawing.Size(196, 30);
            this.btn_Elave.TabIndex = 16;
            this.btn_Elave.Text = "Axtar";
            this.btn_Elave.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.button1.FlatAppearance.BorderSize = 2;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.PaleTurquoise;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Century Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1304, 122);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(196, 30);
            this.button1.TabIndex = 17;
            this.button1.Text = "Yenilə";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // cl_Info
            // 
            this.cl_Info.HeaderText = "İnfo";
            this.cl_Info.Name = "cl_Info";
            this.cl_Info.ReadOnly = true;
            this.cl_Info.Text = "...";
            this.cl_Info.UseColumnTextForButtonValue = true;
            this.cl_Info.Width = 40;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "MXÖ";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Text = "...";
            this.Column1.ToolTipText = "Məzuniyyət,Xəstəlik,Öz hesabına";
            this.Column1.UseColumnTextForButtonValue = true;
            this.Column1.Width = 40;
            // 
            // col_QNo
            // 
            this.col_QNo.HeaderText = "QeydNo";
            this.col_QNo.Name = "col_QNo";
            this.col_QNo.ReadOnly = true;
            this.col_QNo.Width = 140;
            // 
            // col_Soyad
            // 
            this.col_Soyad.HeaderText = "SAA";
            this.col_Soyad.Name = "col_Soyad";
            this.col_Soyad.ReadOnly = true;
            this.col_Soyad.Width = 320;
            // 
            // col_SSN
            // 
            this.col_SSN.HeaderText = "SSN";
            this.col_SSN.Name = "col_SSN";
            this.col_SSN.ReadOnly = true;
            this.col_SSN.Width = 150;
            // 
            // col_Fin
            // 
            this.col_Fin.HeaderText = "FIN";
            this.col_Fin.Name = "col_Fin";
            this.col_Fin.ReadOnly = true;
            this.col_Fin.Width = 140;
            // 
            // col_dax_tar
            // 
            this.col_dax_tar.HeaderText = "Daxil olma tarix";
            this.col_dax_tar.Name = "col_dax_tar";
            this.col_dax_tar.ReadOnly = true;
            this.col_dax_tar.Width = 150;
            // 
            // col_Staj
            // 
            this.col_Staj.HeaderText = "Staj";
            this.col_Staj.Name = "col_Staj";
            this.col_Staj.ReadOnly = true;
            // 
            // col_M_Q_G
            // 
            this.col_M_Q_G.HeaderText = "Məzuniyyət QG";
            this.col_M_Q_G.Name = "col_M_Q_G";
            this.col_M_Q_G.ReadOnly = true;
            this.col_M_Q_G.ToolTipText = "Məzuniyyət qalıq gün";
            this.col_M_Q_G.Width = 160;
            // 
            // frm_MXE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1529, 698);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btn_Elave);
            this.Controls.Add(this.txt_SSN);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.frm_MXE_dtg);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frm_MXE";
            this.Text = "frm_MXE";
            this.Load += new System.EventHandler(this.frm_MXE_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.frm_MXE_dtg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.DataGridView frm_MXE_dtg;
        private System.Windows.Forms.TextBox txt_SSN;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btn_Elave;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewButtonColumn cl_Info;
        private System.Windows.Forms.DataGridViewButtonColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_QNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Soyad;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_SSN;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Fin;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_dax_tar;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_Staj;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_M_Q_G;
    }
}