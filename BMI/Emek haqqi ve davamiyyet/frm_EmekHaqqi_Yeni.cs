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
    public partial class frm_EmekHaqqi_Yeni : Form
    {
        public frm_EmekHaqqi_Yeni()
        {
            InitializeComponent();
        }
        cl_Database db = new cl_Database();
        private void InsertEmekHaqqiYeni()
        {
            try
            {
                string regnum = txb_QeydNo.Text;
                string saa = txb_SAA.Text;
                string il = txb_Il.Text;
                string ay_il=cmb_Ay_Il.Text;
                decimal emehHaqqiNew = Convert.ToDecimal(txb_EmekHaqqi_New.Text);
                

                // Metoda düzgün tiplərlə ötürürük
                db.InsertIscininYeniEH(regnum, saa, il,ay_il, emehHaqqiNew, db.today);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void FillComboBox()
        {
            int currentYear = DateTime.Now.Year; // Cari il
            int currentMonth = DateTime.Now.Month; // Cari ay

            for (int month = currentMonth; month <= 12; month++) // Yalnız cari aydan ilin sonuna qədər
            {
                string monthYear = new DateTime(currentYear, month, 1).ToString("MM-yyyy");
                cmb_Ay_Il.Items.Add(monthYear);
            }

            cmb_Ay_Il.SelectedIndex = 0; // İlk elementi seçili etmək
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Elave_Click(object sender, EventArgs e)
        {
            InsertEmekHaqqiYeni();
            Emek_haqqi_ve_davamiyyet.frm_Isciler frm =new Emek_haqqi_ve_davamiyyet.frm_Isciler();
            frm.LoadDataToGrid();
            this.Close();
        }

        private void frm_EmekHaqqi_Yeni_Load(object sender, EventArgs e)
        {
            FillComboBox();
        }
    }
}
