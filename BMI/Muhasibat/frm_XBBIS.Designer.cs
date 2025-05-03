namespace BMI.Muhasibat
{
    partial class frm_XBBIS
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_XBBIS));
            this.panel1 = new System.Windows.Forms.Panel();
            this.btn_sorgu = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_dov_son = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_dov_evvel = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.btn_sorgu);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txt_dov_son);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txt_dov_evvel);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(202, 150);
            this.panel1.TabIndex = 26;
            // 
            // btn_sorgu
            // 
            this.btn_sorgu.BackColor = System.Drawing.Color.Ivory;
            this.btn_sorgu.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_sorgu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_sorgu.Location = new System.Drawing.Point(10, 103);
            this.btn_sorgu.Name = "btn_sorgu";
            this.btn_sorgu.Size = new System.Drawing.Size(178, 32);
            this.btn_sorgu.TabIndex = 26;
            this.btn_sorgu.Text = "Ümumi sorğu";
            this.btn_sorgu.UseVisualStyleBackColor = false;
            this.btn_sorgu.Click += new System.EventHandler(this.btn_sorgu_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(54, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 17);
            this.label2.TabIndex = 25;
            this.label2.Text = "Dövrün sonu:";
            // 
            // txt_dov_son
            // 
            this.txt_dov_son.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_dov_son.Location = new System.Drawing.Point(40, 74);
            this.txt_dov_son.Name = "txt_dov_son";
            this.txt_dov_son.Size = new System.Drawing.Size(121, 23);
            this.txt_dov_son.TabIndex = 2;
            this.txt_dov_son.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_dov_son.TextChanged += new System.EventHandler(this.txt_dov_son_TextChanged);
            this.txt_dov_son.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_dov_son_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(54, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 17);
            this.label3.TabIndex = 20;
            this.label3.Text = "Dövrün əvvəli:";
            // 
            // txt_dov_evvel
            // 
            this.txt_dov_evvel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_dov_evvel.Location = new System.Drawing.Point(40, 28);
            this.txt_dov_evvel.Name = "txt_dov_evvel";
            this.txt_dov_evvel.Size = new System.Drawing.Size(121, 23);
            this.txt_dov_evvel.TabIndex = 1;
            this.txt_dov_evvel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txt_dov_evvel.TextChanged += new System.EventHandler(this.txt_dov_evvel_TextChanged);
            this.txt_dov_evvel.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_dov_evvel_KeyDown);
            // 
            // frm_XBBIS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(225, 173);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_XBBIS";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XBBIS";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox txt_dov_son;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox txt_dov_evvel;
        private System.Windows.Forms.Button btn_sorgu;
    }
}