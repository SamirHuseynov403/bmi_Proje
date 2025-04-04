namespace Kredit_isler
{
    partial class Gelme_vaxti
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txb_gelme_sira = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.date_teyin = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.txb_raz_meb = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txb_raz_qeyd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txb_raz_saat = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txb_raz_ad = new System.Windows.Forms.TextBox();
            this.button9 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dgw_gelmetarix = new System.Windows.Forms.DataGridView();
            this.iDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.siraDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.adisoyadiDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.meblegDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gelmetarixiDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gelmesaatiDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.qeydDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.teyinolumuslarBindingSource = new System.Windows.Forms.BindingSource(this.components);
            //this.kreditDataSet8 = new Kredit_isler.KreditDataSet8();
            //this.teyin_olumuslarTableAdapter = new Kredit_isler.KreditDataSet8TableAdapters.Teyin_olumuslarTableAdapter();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgw_gelmetarix)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teyinolumuslarBindingSource)).BeginInit();
            //((System.ComponentModel.ISupportInitialize)(this.kreditDataSet8)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.txb_gelme_sira);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.date_teyin);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.txb_raz_meb);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txb_raz_qeyd);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txb_raz_saat);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.txb_raz_ad);
            this.panel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(381, 211);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // txb_gelme_sira
            // 
            this.txb_gelme_sira.Enabled = false;
            this.txb_gelme_sira.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txb_gelme_sira.Location = new System.Drawing.Point(97, 175);
            this.txb_gelme_sira.Name = "txb_gelme_sira";
            this.txb_gelme_sira.Size = new System.Drawing.Size(49, 20);
            this.txb_gelme_sira.TabIndex = 90;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(4, 181);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 20);
            this.label4.TabIndex = 89;
            this.label4.Text = "Sıra";
            // 
            // date_teyin
            // 
            this.date_teyin.CustomFormat = "";
            this.date_teyin.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.date_teyin.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.date_teyin.Location = new System.Drawing.Point(97, 77);
            this.date_teyin.Name = "date_teyin";
            this.date_teyin.Size = new System.Drawing.Size(99, 25);
            this.date_teyin.TabIndex = 88;
            this.date_teyin.Value = new System.DateTime(2021, 1, 10, 0, 0, 0, 0);
            this.date_teyin.ValueChanged += new System.EventHandler(this.date_teyin_ValueChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(162, 175);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 87;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txb_raz_meb
            // 
            this.txb_raz_meb.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txb_raz_meb.Location = new System.Drawing.Point(97, 45);
            this.txb_raz_meb.Name = "txb_raz_meb";
            this.txb_raz_meb.Size = new System.Drawing.Size(273, 25);
            this.txb_raz_meb.TabIndex = 86;
            this.txb_raz_meb.TextChanged += new System.EventHandler(this.txb_raz_meb_TextChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(4, 79);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(89, 20);
            this.label15.TabIndex = 84;
            this.label15.Text = "Gəlmə tarixi";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(4, 142);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 20);
            this.label3.TabIndex = 11;
            this.label3.Text = "Qeyd";
            // 
            // txb_raz_qeyd
            // 
            this.txb_raz_qeyd.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txb_raz_qeyd.Location = new System.Drawing.Point(97, 139);
            this.txb_raz_qeyd.Name = "txb_raz_qeyd";
            this.txb_raz_qeyd.Size = new System.Drawing.Size(273, 25);
            this.txb_raz_qeyd.TabIndex = 12;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(4, 110);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 20);
            this.label2.TabIndex = 9;
            this.label2.Text = "Saat";
            // 
            // txb_raz_saat
            // 
            this.txb_raz_saat.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txb_raz_saat.Location = new System.Drawing.Point(97, 107);
            this.txb_raz_saat.Name = "txb_raz_saat";
            this.txb_raz_saat.Size = new System.Drawing.Size(49, 25);
            this.txb_raz_saat.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "Məbləğ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(4, 16);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(73, 20);
            this.label9.TabIndex = 5;
            this.label9.Text = "Soyad Ad";
            // 
            // txb_raz_ad
            // 
            this.txb_raz_ad.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txb_raz_ad.Location = new System.Drawing.Point(97, 13);
            this.txb_raz_ad.Name = "txb_raz_ad";
            this.txb_raz_ad.Size = new System.Drawing.Size(273, 25);
            this.txb_raz_ad.TabIndex = 6;
            // 
            // button9
            // 
            this.button9.BackColor = System.Drawing.Color.PaleTurquoise;
            this.button9.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button9.Location = new System.Drawing.Point(12, 229);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(381, 47);
            this.button9.TabIndex = 90;
            this.button9.Text = "Əlavə et";
            this.button9.UseVisualStyleBackColor = false;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.dgw_gelmetarix);
            this.panel2.Location = new System.Drawing.Point(399, 12);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(755, 264);
            this.panel2.TabIndex = 91;
            // 
            // dgw_gelmetarix
            // 
            this.dgw_gelmetarix.AllowUserToAddRows = false;
            this.dgw_gelmetarix.AllowUserToDeleteRows = false;
            this.dgw_gelmetarix.AutoGenerateColumns = false;
            this.dgw_gelmetarix.BackgroundColor = System.Drawing.Color.White;
            this.dgw_gelmetarix.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.PaleTurquoise;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.Padding = new System.Windows.Forms.Padding(2, 3, 2, 3);
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgw_gelmetarix.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgw_gelmetarix.ColumnHeadersHeight = 28;
            this.dgw_gelmetarix.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iDDataGridViewTextBoxColumn,
            this.siraDataGridViewTextBoxColumn,
            this.adisoyadiDataGridViewTextBoxColumn,
            this.meblegDataGridViewTextBoxColumn,
            this.gelmetarixiDataGridViewTextBoxColumn,
            this.gelmesaatiDataGridViewTextBoxColumn,
            this.qeydDataGridViewTextBoxColumn});
            this.dgw_gelmetarix.DataSource = this.teyinolumuslarBindingSource;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.Padding = new System.Windows.Forms.Padding(2);
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.Gainsboro;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgw_gelmetarix.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgw_gelmetarix.EnableHeadersVisualStyles = false;
            this.dgw_gelmetarix.Location = new System.Drawing.Point(3, 3);
            this.dgw_gelmetarix.Name = "dgw_gelmetarix";
            this.dgw_gelmetarix.ReadOnly = true;
            this.dgw_gelmetarix.RowHeadersVisible = false;
            this.dgw_gelmetarix.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgw_gelmetarix.Size = new System.Drawing.Size(745, 254);
            this.dgw_gelmetarix.TabIndex = 0;
            // 
            // iDDataGridViewTextBoxColumn
            // 
            this.iDDataGridViewTextBoxColumn.DataPropertyName = "ID";
            this.iDDataGridViewTextBoxColumn.HeaderText = "ID";
            this.iDDataGridViewTextBoxColumn.Name = "iDDataGridViewTextBoxColumn";
            this.iDDataGridViewTextBoxColumn.ReadOnly = true;
            this.iDDataGridViewTextBoxColumn.Visible = false;
            // 
            // siraDataGridViewTextBoxColumn
            // 
            this.siraDataGridViewTextBoxColumn.DataPropertyName = "Sira";
            this.siraDataGridViewTextBoxColumn.HeaderText = "Sıra";
            this.siraDataGridViewTextBoxColumn.Name = "siraDataGridViewTextBoxColumn";
            this.siraDataGridViewTextBoxColumn.ReadOnly = true;
            this.siraDataGridViewTextBoxColumn.Width = 40;
            // 
            // adisoyadiDataGridViewTextBoxColumn
            // 
            this.adisoyadiDataGridViewTextBoxColumn.DataPropertyName = "Adi_soyadi";
            this.adisoyadiDataGridViewTextBoxColumn.HeaderText = "Soyadı Adı";
            this.adisoyadiDataGridViewTextBoxColumn.Name = "adisoyadiDataGridViewTextBoxColumn";
            this.adisoyadiDataGridViewTextBoxColumn.ReadOnly = true;
            this.adisoyadiDataGridViewTextBoxColumn.Width = 300;
            // 
            // meblegDataGridViewTextBoxColumn
            // 
            this.meblegDataGridViewTextBoxColumn.DataPropertyName = "Mebleg";
            this.meblegDataGridViewTextBoxColumn.HeaderText = "Məbləğ";
            this.meblegDataGridViewTextBoxColumn.Name = "meblegDataGridViewTextBoxColumn";
            this.meblegDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // gelmetarixiDataGridViewTextBoxColumn
            // 
            this.gelmetarixiDataGridViewTextBoxColumn.DataPropertyName = "Gelme_tarixi";
            this.gelmetarixiDataGridViewTextBoxColumn.HeaderText = "Tarix";
            this.gelmetarixiDataGridViewTextBoxColumn.Name = "gelmetarixiDataGridViewTextBoxColumn";
            this.gelmetarixiDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // gelmesaatiDataGridViewTextBoxColumn
            // 
            this.gelmesaatiDataGridViewTextBoxColumn.DataPropertyName = "Gelme_saati";
            this.gelmesaatiDataGridViewTextBoxColumn.HeaderText = "Saat";
            this.gelmesaatiDataGridViewTextBoxColumn.Name = "gelmesaatiDataGridViewTextBoxColumn";
            this.gelmesaatiDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // qeydDataGridViewTextBoxColumn
            // 
            this.qeydDataGridViewTextBoxColumn.DataPropertyName = "Qeyd";
            this.qeydDataGridViewTextBoxColumn.HeaderText = "Qeyd";
            this.qeydDataGridViewTextBoxColumn.Name = "qeydDataGridViewTextBoxColumn";
            this.qeydDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // teyinolumuslarBindingSource
            // 
            this.teyinolumuslarBindingSource.DataMember = "Teyin_olumuslar";
            //this.teyinolumuslarBindingSource.DataSource = this.kreditDataSet8;
            // 
            // kreditDataSet8
            // 
            //this.kreditDataSet8.DataSetName = "KreditDataSet8";
            //this.kreditDataSet8.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // teyin_olumuslarTableAdapter
            // 
            //this.teyin_olumuslarTableAdapter.ClearBeforeFill = true;
            // 
            // Gelme_vaxti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1159, 288);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Gelme_vaxti";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gelme_vaxti";
            this.Load += new System.EventHandler(this.Gelme_vaxti_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgw_gelmetarix)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teyinolumuslarBindingSource)).EndInit();
            //((System.ComponentModel.ISupportInitialize)(this.kreditDataSet8)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.TextBox txb_raz_qeyd;
        public System.Windows.Forms.TextBox txb_raz_ad;
        public System.Windows.Forms.DataGridView dgw_gelmetarix;
        private System.Windows.Forms.TextBox txb_raz_meb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DateTimePicker date_teyin;
        public System.Windows.Forms.TextBox txb_raz_saat;
        private System.Windows.Forms.TextBox txb_gelme_sira;
        private System.Windows.Forms.Label label4;
        //private KreditDataSet8 kreditDataSet8;
        private System.Windows.Forms.BindingSource teyinolumuslarBindingSource;
        //private KreditDataSet8TableAdapters.Teyin_olumuslarTableAdapter teyin_olumuslarTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn iDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn siraDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn adisoyadiDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn meblegDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn gelmetarixiDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn gelmesaatiDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn qeydDataGridViewTextBoxColumn;


    }
}