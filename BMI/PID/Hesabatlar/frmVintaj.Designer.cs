namespace BMI.PID.Hesabatlar
{
    partial class frmVintaj
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
            this.dtHesabatTarixi = new System.Windows.Forms.DateTimePicker();
            this.dtDovrİlk = new System.Windows.Forms.DateTimePicker();
            this.dtDovrSon = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_sorgu = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.dtHesabatTarixi);
            this.panel1.Controls.Add(this.dtDovrİlk);
            this.panel1.Controls.Add(this.dtDovrSon);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btn_sorgu);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(302, 199);
            this.panel1.TabIndex = 28;
            // 
            // dtHesabatTarixi
            // 
            this.dtHesabatTarixi.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtHesabatTarixi.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtHesabatTarixi.Location = new System.Drawing.Point(97, 100);
            this.dtHesabatTarixi.Name = "dtHesabatTarixi";
            this.dtHesabatTarixi.Size = new System.Drawing.Size(104, 23);
            this.dtHesabatTarixi.TabIndex = 34;
            this.dtHesabatTarixi.Value = new System.DateTime(2025, 3, 31, 0, 0, 0, 0);
            // 
            // dtDovrİlk
            // 
            this.dtDovrİlk.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtDovrİlk.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDovrİlk.Location = new System.Drawing.Point(34, 40);
            this.dtDovrİlk.Name = "dtDovrİlk";
            this.dtDovrİlk.Size = new System.Drawing.Size(104, 23);
            this.dtDovrİlk.TabIndex = 33;
            this.dtDovrİlk.Value = new System.DateTime(2022, 3, 1, 0, 0, 0, 0);
            // 
            // dtDovrSon
            // 
            this.dtDovrSon.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtDovrSon.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtDovrSon.Location = new System.Drawing.Point(163, 40);
            this.dtDovrSon.Name = "dtDovrSon";
            this.dtDovrSon.Size = new System.Drawing.Size(104, 23);
            this.dtDovrSon.TabIndex = 32;
            this.dtDovrSon.Value = new System.DateTime(2025, 3, 31, 0, 0, 0, 0);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(45, 73);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(211, 17);
            this.label4.TabIndex = 30;
            this.label4.Text = "Hesabat tarixi (ayın son iş günü)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(144, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(13, 17);
            this.label2.TabIndex = 29;
            this.label2.Text = "-";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(90, 169);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 15);
            this.label1.TabIndex = 27;
            this.label1.Text = "Hazırlanır gözləyin...";
            this.label1.Visible = false;
            // 
            // btn_sorgu
            // 
            this.btn_sorgu.BackColor = System.Drawing.Color.Ivory;
            this.btn_sorgu.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_sorgu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_sorgu.Location = new System.Drawing.Point(59, 129);
            this.btn_sorgu.Name = "btn_sorgu";
            this.btn_sorgu.Size = new System.Drawing.Size(178, 32);
            this.btn_sorgu.TabIndex = 26;
            this.btn_sorgu.Text = "Ümumi sorğu";
            this.btn_sorgu.UseVisualStyleBackColor = false;
            this.btn_sorgu.Click += new System.EventHandler(this.btn_sorgu_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(110, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 17);
            this.label3.TabIndex = 20;
            this.label3.Text = "Dövr aralığı";
            // 
            // frmVintaj
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(324, 224);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmVintaj";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vintaj hesabatı";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_sorgu;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtDovrSon;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dtHesabatTarixi;
        private System.Windows.Forms.DateTimePicker dtDovrİlk;
    }
}