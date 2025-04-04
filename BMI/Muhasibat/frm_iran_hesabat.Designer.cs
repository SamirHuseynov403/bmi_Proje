namespace BMI.Muhasibat
{
    partial class frm_iran_hesabat
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_iran_hesabat));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_evvelki_ay = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_tarix_son = new System.Windows.Forms.TextBox();
            this.txt_tarix_evvel = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txt_evvelki_ay);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txt_tarix_son);
            this.panel1.Controls.Add(this.txt_tarix_evvel);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(322, 180);
            this.panel1.TabIndex = 28;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(10, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(161, 17);
            this.label2.TabIndex = 29;
            this.label2.Text = "Əvvəlki ayın son iş günü";
            // 
            // txt_evvelki_ay
            // 
            this.txt_evvelki_ay.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_evvelki_ay.Location = new System.Drawing.Point(193, 70);
            this.txt_evvelki_ay.Name = "txt_evvelki_ay";
            this.txt_evvelki_ay.Size = new System.Drawing.Size(116, 23);
            this.txt_evvelki_ay.TabIndex = 28;
            this.txt_evvelki_ay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_evvelki_ay.TextChanged += new System.EventHandler(this.txt_evvelki_ay_TextChanged);
            this.txt_evvelki_ay.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_evvelki_ay_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(10, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 17);
            this.label1.TabIndex = 27;
            this.label1.Text = "Hesabat ayının son iş günü";
            // 
            // txt_tarix_son
            // 
            this.txt_tarix_son.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_tarix_son.Location = new System.Drawing.Point(193, 38);
            this.txt_tarix_son.Name = "txt_tarix_son";
            this.txt_tarix_son.Size = new System.Drawing.Size(116, 23);
            this.txt_tarix_son.TabIndex = 24;
            this.txt_tarix_son.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_tarix_son.TextChanged += new System.EventHandler(this.txt_tarix_son_TextChanged);
            this.txt_tarix_son.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_tarix_son_KeyDown);
            // 
            // txt_tarix_evvel
            // 
            this.txt_tarix_evvel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_tarix_evvel.Location = new System.Drawing.Point(193, 6);
            this.txt_tarix_evvel.Name = "txt_tarix_evvel";
            this.txt_tarix_evvel.Size = new System.Drawing.Size(116, 23);
            this.txt_tarix_evvel.TabIndex = 23;
            this.txt_tarix_evvel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_tarix_evvel.TextChanged += new System.EventHandler(this.txt_tarix_evvel_TextChanged);
            this.txt_tarix_evvel.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_tarix_evvel_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(10, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(169, 17);
            this.label3.TabIndex = 26;
            this.label3.Text = "Hesabat ayının ilk iş günü";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Ivory;
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(13, 110);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(296, 29);
            this.button1.TabIndex = 22;
            this.button1.Text = "Sorğu";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(384, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(92, 178);
            this.dataGridView1.TabIndex = 29;
            this.dataGridView1.Visible = false;
            // 
            // frm_iran_hesabat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(346, 202);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_iran_hesabat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Iran hesabatı kredit";
            this.Load += new System.EventHandler(this.frm_iran_hesabat_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_tarix_son;
        private System.Windows.Forms.TextBox txt_tarix_evvel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_evvelki_ay;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}