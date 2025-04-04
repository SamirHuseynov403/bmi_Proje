namespace BMI.Esas
{
    partial class frm_setup
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_setup));
            this.button1 = new System.Windows.Forms.Button();
            this.pnl_istifadeciler = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dtg_cedvel = new System.Windows.Forms.DataGridView();
            this.pnl_aletler = new System.Windows.Forms.Panel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.pnl_formlar = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbl_alet = new System.Windows.Forms.Label();
            this.lbl_form = new System.Windows.Forms.Label();
            this.lbl_ad = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtg_cedvel)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.PaleTurquoise;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1197, 42);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(196, 29);
            this.button1.TabIndex = 0;
            this.button1.Text = "Formları bazaya at";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pnl_istifadeciler
            // 
            this.pnl_istifadeciler.AutoScroll = true;
            this.pnl_istifadeciler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnl_istifadeciler.Location = new System.Drawing.Point(6, 19);
            this.pnl_istifadeciler.Name = "pnl_istifadeciler";
            this.pnl_istifadeciler.Size = new System.Drawing.Size(235, 414);
            this.pnl_istifadeciler.TabIndex = 1;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.SeaShell;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(1197, 77);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(196, 29);
            this.button2.TabIndex = 2;
            this.button2.Text = "Işçi dərəcəsi";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pnl_istifadeciler);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 74);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(247, 441);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "İstifadəçilər";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtg_cedvel);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(265, 74);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(926, 441);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Alətlər";
            // 
            // dtg_cedvel
            // 
            this.dtg_cedvel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtg_cedvel.Location = new System.Drawing.Point(6, 23);
            this.dtg_cedvel.Name = "dtg_cedvel";
            this.dtg_cedvel.Size = new System.Drawing.Size(914, 410);
            this.dtg_cedvel.TabIndex = 0;
            this.dtg_cedvel.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtg_cedvel_CellClick);
            // 
            // pnl_aletler
            // 
            this.pnl_aletler.AutoScroll = true;
            this.pnl_aletler.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnl_aletler.Location = new System.Drawing.Point(1197, 458);
            this.pnl_aletler.Name = "pnl_aletler";
            this.pnl_aletler.Size = new System.Drawing.Size(30, 16);
            this.pnl_aletler.TabIndex = 1;
            this.pnl_aletler.Visible = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.pnl_formlar);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(1197, 493);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(10, 17);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Formlar";
            this.groupBox3.Visible = false;
            // 
            // pnl_formlar
            // 
            this.pnl_formlar.AutoScroll = true;
            this.pnl_formlar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnl_formlar.Location = new System.Drawing.Point(6, 19);
            this.pnl_formlar.Name = "pnl_formlar";
            this.pnl_formlar.Size = new System.Drawing.Size(914, 100);
            this.pnl_formlar.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.lbl_alet);
            this.panel1.Controls.Add(this.lbl_form);
            this.panel1.Controls.Add(this.lbl_ad);
            this.panel1.Location = new System.Drawing.Point(12, 42);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1009, 29);
            this.panel1.TabIndex = 6;
            // 
            // lbl_alet
            // 
            this.lbl_alet.AutoSize = true;
            this.lbl_alet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_alet.Location = new System.Drawing.Point(550, 0);
            this.lbl_alet.Name = "lbl_alet";
            this.lbl_alet.Size = new System.Drawing.Size(18, 20);
            this.lbl_alet.TabIndex = 2;
            this.lbl_alet.Text = "1";
            this.lbl_alet.Visible = false;
            // 
            // lbl_form
            // 
            this.lbl_form.AutoSize = true;
            this.lbl_form.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_form.Location = new System.Drawing.Point(248, 2);
            this.lbl_form.Name = "lbl_form";
            this.lbl_form.Size = new System.Drawing.Size(18, 20);
            this.lbl_form.TabIndex = 1;
            this.lbl_form.Text = "1";
            this.lbl_form.Visible = false;
            // 
            // lbl_ad
            // 
            this.lbl_ad.AutoSize = true;
            this.lbl_ad.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ad.Location = new System.Drawing.Point(9, 4);
            this.lbl_ad.Name = "lbl_ad";
            this.lbl_ad.Size = new System.Drawing.Size(18, 20);
            this.lbl_ad.TabIndex = 0;
            this.lbl_ad.Text = "1";
            // 
            // frm_setup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1443, 692);
            this.Controls.Add(this.pnl_aletler);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_setup";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Setup";
            this.Load += new System.EventHandler(this.frm_setup_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dtg_cedvel)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel pnl_istifadeciler;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Panel pnl_aletler;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Panel pnl_formlar;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lbl_ad;
        private System.Windows.Forms.Label lbl_form;
        private System.Windows.Forms.Label lbl_alet;
        private System.Windows.Forms.DataGridView dtg_cedvel;
    }
}