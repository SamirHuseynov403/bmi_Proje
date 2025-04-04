namespace BMI
{
    partial class zaminarayis
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(zaminarayis));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dateavtoMuqtarix = new System.Windows.Forms.DateTimePicker();
            this.dat1 = new System.Windows.Forms.TextBox();
            this.cmbGisamezmun = new System.Windows.Forms.ComboBox();
            this.dattar = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.txbmebleg = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txbzamin = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txbtarix = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txbzBorcalan = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_axtar = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label22 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_val = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.dataGridView1.ColumnHeadersHeight = 26;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dataGridView1.Location = new System.Drawing.Point(6, 4);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 35;
            this.dataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.ShowCellToolTips = false;
            this.dataGridView1.Size = new System.Drawing.Size(1196, 222);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // dateavtoMuqtarix
            // 
            this.dateavtoMuqtarix.CalendarMonthBackground = System.Drawing.Color.MistyRose;
            this.dateavtoMuqtarix.CustomFormat = "dd-MM-yyyy";
            this.dateavtoMuqtarix.Enabled = false;
            this.dateavtoMuqtarix.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateavtoMuqtarix.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateavtoMuqtarix.Location = new System.Drawing.Point(1171, 232);
            this.dateavtoMuqtarix.Name = "dateavtoMuqtarix";
            this.dateavtoMuqtarix.Size = new System.Drawing.Size(10, 22);
            this.dateavtoMuqtarix.TabIndex = 286;
            this.dateavtoMuqtarix.Value = new System.DateTime(2020, 6, 15, 0, 30, 37, 0);
            this.dateavtoMuqtarix.Visible = false;
            // 
            // dat1
            // 
            this.dat1.BackColor = System.Drawing.Color.White;
            this.dat1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dat1.Location = new System.Drawing.Point(1184, 234);
            this.dat1.Name = "dat1";
            this.dat1.Size = new System.Drawing.Size(18, 21);
            this.dat1.TabIndex = 285;
            this.dat1.Visible = false;
            // 
            // cmbGisamezmun
            // 
            this.cmbGisamezmun.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbGisamezmun.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbGisamezmun.BackColor = System.Drawing.Color.White;
            this.cmbGisamezmun.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cmbGisamezmun.DropDownHeight = 130;
            this.cmbGisamezmun.Font = new System.Drawing.Font("Ora Times", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmbGisamezmun.FormattingEnabled = true;
            this.cmbGisamezmun.IntegralHeight = false;
            this.cmbGisamezmun.Items.AddRange(new object[] {
            "Етибарнамя",
            "Шящ-дубл",
            "Билдириш",
            "Иддиа",
            "Мяктуб"});
            this.cmbGisamezmun.Location = new System.Drawing.Point(382, 242);
            this.cmbGisamezmun.Name = "cmbGisamezmun";
            this.cmbGisamezmun.Size = new System.Drawing.Size(165, 24);
            this.cmbGisamezmun.TabIndex = 284;
            this.cmbGisamezmun.Visible = false;
            // 
            // dattar
            // 
            this.dattar.CalendarMonthBackground = System.Drawing.Color.MistyRose;
            this.dattar.CustomFormat = "dd-MM-yyyy";
            this.dattar.Enabled = false;
            this.dattar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dattar.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dattar.Location = new System.Drawing.Point(1168, 232);
            this.dattar.Name = "dattar";
            this.dattar.Size = new System.Drawing.Size(10, 22);
            this.dattar.TabIndex = 283;
            this.dattar.Value = new System.DateTime(2020, 6, 15, 0, 30, 37, 0);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(490, 299);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(128, 23);
            this.button1.TabIndex = 282;
            this.button1.Text = "Yazdır";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // txbmebleg
            // 
            this.txbmebleg.BackColor = System.Drawing.Color.White;
            this.txbmebleg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txbmebleg.Location = new System.Drawing.Point(382, 301);
            this.txbmebleg.Name = "txbmebleg";
            this.txbmebleg.Size = new System.Drawing.Size(102, 21);
            this.txbmebleg.TabIndex = 280;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(322, 303);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 15);
            this.label4.TabIndex = 281;
            this.label4.Text = "Məbləğ";
            // 
            // txbzamin
            // 
            this.txbzamin.BackColor = System.Drawing.Color.White;
            this.txbzamin.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txbzamin.Location = new System.Drawing.Point(110, 301);
            this.txbzamin.Name = "txbzamin";
            this.txbzamin.Size = new System.Drawing.Size(208, 21);
            this.txbzamin.TabIndex = 278;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(13, 304);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 15);
            this.label3.TabIndex = 279;
            this.label3.Text = "Zamin";
            // 
            // txbtarix
            // 
            this.txbtarix.BackColor = System.Drawing.Color.White;
            this.txbtarix.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txbtarix.Location = new System.Drawing.Point(382, 272);
            this.txbtarix.Name = "txbtarix";
            this.txbtarix.Size = new System.Drawing.Size(102, 21);
            this.txbtarix.TabIndex = 276;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(342, 274);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 15);
            this.label1.TabIndex = 277;
            this.label1.Text = "Tarix";
            // 
            // txbzBorcalan
            // 
            this.txbzBorcalan.BackColor = System.Drawing.Color.White;
            this.txbzBorcalan.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txbzBorcalan.Location = new System.Drawing.Point(110, 272);
            this.txbzBorcalan.Name = "txbzBorcalan";
            this.txbzBorcalan.Size = new System.Drawing.Size(208, 21);
            this.txbzBorcalan.TabIndex = 274;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(13, 275);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 15);
            this.label2.TabIndex = 275;
            this.label2.Text = "Borcalan";
            // 
            // txt_axtar
            // 
            this.txt_axtar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_axtar.Location = new System.Drawing.Point(110, 245);
            this.txt_axtar.Name = "txt_axtar";
            this.txt_axtar.Size = new System.Drawing.Size(101, 21);
            this.txt_axtar.TabIndex = 273;
            // 
            // button2
            // 
            this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button2.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.button2.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.button2.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue;
            this.button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Aqua;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(217, 244);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(101, 22);
            this.button2.TabIndex = 272;
            this.button2.Text = "Sorğu";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.Location = new System.Drawing.Point(13, 246);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(85, 15);
            this.label22.TabIndex = 271;
            this.label22.Text = "Zamin pincod:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(495, 277);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 15);
            this.label5.TabIndex = 288;
            this.label5.Text = "Valyuta";
            // 
            // txt_val
            // 
            this.txt_val.BackColor = System.Drawing.Color.White;
            this.txt_val.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_val.Location = new System.Drawing.Point(547, 272);
            this.txt_val.Name = "txt_val";
            this.txt_val.Size = new System.Drawing.Size(71, 21);
            this.txt_val.TabIndex = 289;
            // 
            // zaminarayis
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1207, 329);
            this.Controls.Add(this.txt_val);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.dateavtoMuqtarix);
            this.Controls.Add(this.dat1);
            this.Controls.Add(this.cmbGisamezmun);
            this.Controls.Add(this.dattar);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txbmebleg);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txbzamin);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txbtarix);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txbzBorcalan);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txt_axtar);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "zaminarayis";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Zaminə məktub";
            this.Load += new System.EventHandler(this.zaminarayis_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DateTimePicker dateavtoMuqtarix;
        public System.Windows.Forms.TextBox dat1;
        public System.Windows.Forms.ComboBox cmbGisamezmun;
        private System.Windows.Forms.DateTimePicker dattar;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.TextBox txbmebleg;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox txbzamin;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox txbtarix;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox txbzBorcalan;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_axtar;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label5;
        public System.Windows.Forms.TextBox txt_val;
    }
}