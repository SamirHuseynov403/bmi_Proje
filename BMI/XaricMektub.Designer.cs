namespace BMI
{
    partial class XaricMektub
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
            this.label11 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtYol = new System.Windows.Forms.TextBox();
            this.btnTemizle = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.lblSgldml = new System.Windows.Forms.Label();
            this.cmbGonderyer = new System.Windows.Forms.ComboBox();
            this.cmbGisamezmun = new System.Windows.Forms.ComboBox();
            this.dtpTarix = new System.Windows.Forms.DateTimePicker();
            this.rctbMektubmetn = new System.Windows.Forms.RichTextBox();
            this.opnflYol = new System.Windows.Forms.OpenFileDialog();
            this.verilmis_krTableAdapter1 = new BMI.MuracietlerDataSetTableAdapters.Verilmis_krTableAdapter();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(23, 15);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(37, 16);
            this.label11.TabIndex = 75;
            this.label11.Text = "Tarix:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(23, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 16);
            this.label1.TabIndex = 75;
            this.label1.Text = "Göndərilən yer:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Narrow", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(23, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 16);
            this.label2.TabIndex = 75;
            this.label2.Text = "Qısa məzmun:";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.PaleTurquoise;
            this.panel1.Controls.Add(this.txtYol);
            this.panel1.Controls.Add(this.btnTemizle);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 533);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(822, 35);
            this.panel1.TabIndex = 77;
            // 
            // txtYol
            // 
            this.txtYol.BackColor = System.Drawing.Color.White;
            this.txtYol.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtYol.ForeColor = System.Drawing.Color.Blue;
            this.txtYol.Location = new System.Drawing.Point(251, 4);
            this.txtYol.Name = "txtYol";
            this.txtYol.Size = new System.Drawing.Size(289, 25);
            this.txtYol.TabIndex = 8;
            // 
            // btnTemizle
            // 
            this.btnTemizle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnTemizle.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnTemizle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTemizle.Location = new System.Drawing.Point(12, 4);
            this.btnTemizle.Name = "btnTemizle";
            this.btnTemizle.Size = new System.Drawing.Size(106, 26);
            this.btnTemizle.TabIndex = 7;
            this.btnTemizle.Text = "Təmizlə";
            this.btnTemizle.UseVisualStyleBackColor = true;
            this.btnTemizle.Click += new System.EventHandler(this.btnTemizle_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.White;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(544, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(106, 26);
            this.button2.TabIndex = 0;
            this.button2.Text = "Sənədi aç";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(655, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(155, 26);
            this.button1.TabIndex = 0;
            this.button1.Text = "Əlavə et";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblSgldml
            // 
            this.lblSgldml.AutoSize = true;
            this.lblSgldml.Location = new System.Drawing.Point(498, 15);
            this.lblSgldml.Name = "lblSgldml";
            this.lblSgldml.Size = new System.Drawing.Size(35, 13);
            this.lblSgldml.TabIndex = 78;
            this.lblSgldml.Text = "label3";
            // 
            // cmbGonderyer
            // 
            this.cmbGonderyer.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cmbGonderyer.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbGonderyer.BackColor = System.Drawing.Color.White;
            this.cmbGonderyer.Cursor = System.Windows.Forms.Cursors.Hand;
            this.cmbGonderyer.DropDownHeight = 211;
            this.cmbGonderyer.Font = new System.Drawing.Font("Ora Times", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmbGonderyer.FormattingEnabled = true;
            this.cmbGonderyer.IntegralHeight = false;
            this.cmbGonderyer.Items.AddRange(new object[] {
            "Азсыьорта АСЪ",
            "Азярселл Телеъом ММЪ",
            "Азярфон",
            "Бакъелл",
            "Бакы Апеллйасийа Мящкямяси",
            "Баш Дювлят Йол Полис Идарясиня",
            "БМИ Дубай",
            "БМИ Тещран",
            "БТИ",
            "Верэи",
            "ДЙП",
            "ДСМФ",
            "Дювлят Эюмрцк Комитяси",
            "ДЯДРХ",
            "Йасамал район Мящкямяси",
            "Меэа сыьорта",
            "Меэасервиъе ММЪ",
            "Милликарт ММЪ",
            "Мяркязи Банк",
            "Нязарят Палатасы",
            "Сабунчу район Мящкямяси",
            "Сураханы район Мящкямяси",
            "Сябаил район Мящкямяси",
            "Халг Сыьорта",
            "Хятайи район Мящкямяси",
            "Щярби Прокурорлуг"});
            this.cmbGonderyer.Location = new System.Drawing.Point(124, 42);
            this.cmbGonderyer.Name = "cmbGonderyer";
            this.cmbGonderyer.Size = new System.Drawing.Size(635, 24);
            this.cmbGonderyer.Sorted = true;
            this.cmbGonderyer.TabIndex = 79;
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
            this.cmbGisamezmun.Location = new System.Drawing.Point(124, 72);
            this.cmbGisamezmun.Name = "cmbGisamezmun";
            this.cmbGisamezmun.Size = new System.Drawing.Size(635, 24);
            this.cmbGisamezmun.TabIndex = 80;
            // 
            // dtpTarix
            // 
            this.dtpTarix.CalendarMonthBackground = System.Drawing.Color.White;
            this.dtpTarix.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dtpTarix.Font = new System.Drawing.Font("Ora Times", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dtpTarix.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpTarix.Location = new System.Drawing.Point(124, 15);
            this.dtpTarix.Name = "dtpTarix";
            this.dtpTarix.Size = new System.Drawing.Size(200, 24);
            this.dtpTarix.TabIndex = 81;
            // 
            // rctbMektubmetn
            // 
            this.rctbMektubmetn.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rctbMektubmetn.AutoWordSelection = true;
            this.rctbMektubmetn.BackColor = System.Drawing.Color.White;
            this.rctbMektubmetn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rctbMektubmetn.Font = new System.Drawing.Font("Azeri Bookman Lat", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rctbMektubmetn.Location = new System.Drawing.Point(26, 102);
            this.rctbMektubmetn.Name = "rctbMektubmetn";
            this.rctbMektubmetn.Size = new System.Drawing.Size(770, 415);
            this.rctbMektubmetn.TabIndex = 82;
            this.rctbMektubmetn.Text = "";
            this.rctbMektubmetn.WordWrap = false;
            // 
            // verilmis_krTableAdapter1
            // 
            this.verilmis_krTableAdapter1.ClearBeforeFill = true;
            // 
            // XaricMektub
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(822, 568);
            this.Controls.Add(this.rctbMektubmetn);
            this.Controls.Add(this.dtpTarix);
            this.Controls.Add(this.cmbGisamezmun);
            this.Controls.Add(this.cmbGonderyer);
            this.Controls.Add(this.lblSgldml);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label11);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "XaricMektub";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Xaric olan məktublar";
            this.Load += new System.EventHandler(this.XaricMektub_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button2;
        public System.Windows.Forms.Label lblSgldml;
        public System.Windows.Forms.ComboBox cmbGonderyer;
        public System.Windows.Forms.ComboBox cmbGisamezmun;
        public System.Windows.Forms.DateTimePicker dtpTarix;
        public System.Windows.Forms.RichTextBox rctbMektubmetn;
        private System.Windows.Forms.Button btnTemizle;
        private System.Windows.Forms.OpenFileDialog opnflYol;
        private System.Windows.Forms.TextBox txtYol;
        public System.Windows.Forms.Button button1;
        private MuracietlerDataSetTableAdapters.Verilmis_krTableAdapter verilmis_krTableAdapter1;
    }
}