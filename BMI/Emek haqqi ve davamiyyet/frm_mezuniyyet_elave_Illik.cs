using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.Emek_haqqi_ve_davamiyyet
{
    public partial class frm_mezuniyyet_elave_Illik : Form
    {
        public frm_mezuniyyet_elave_Illik()
        {
            InitializeComponent();
        }
        cl_Database db = new cl_Database();

        private void InsertMezuniyyetIllikYeni()
        {
            try
            {
                string regnum = txb_QeydNo.Text;
                string saa = txb_SAA.Text;
                string il = txb_Il.Text;
                string mez_new = txb_MSayIluzre.Text;


                // Metoda düzgün tiplərlə ötürürük
                db.InsertIscininIllikMez(regnum, saa, il, mez_new, db.today);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Elave_Click(object sender, EventArgs e)
        {
            InsertMezuniyyetIllikYeni();
            this.Close();
        }
    }
}
