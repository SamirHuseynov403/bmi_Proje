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
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;

namespace BMI.Emek_haqqi_ve_davamiyyet
{
    public partial class frm_Elaveler : Form
    {
        public frm_Elaveler()
        {
            InitializeComponent();
        }
        cl_Database db = new cl_Database();
        private void InsertMaasElaveFromGrid()
        {
            try
            {
                // DataGridView-dən məlumatları oxuyuruq
                foreach (DataGridViewRow row in dtg_Elaveler.Rows)
                {
                    // Boş və ya əlavə olunan boş sətirləri nəzərə almaq üçün
                    if (row.IsNewRow) continue;

                    // SAA və Cari Hesab DataGridView-in 0 və 1-ci sütunlarından oxunur
                    string saa = row.Cells[0].Value?.ToString() ?? string.Empty;
                    string cari_hesab = row.Cells[1].Value?.ToString() ?? string.Empty;

                    // Digər sütunları da oxumaq üçün
                    decimal avans = Convert.ToDecimal(row.Cells[2].Value ?? 0);
                    decimal mukafat = Convert.ToDecimal(row.Cells[3].Value ?? 0);
                    decimal elave_eh = Convert.ToDecimal(row.Cells[4].Value ?? 0);
                    decimal mezuniyyet_haqqi = Convert.ToDecimal(row.Cells[5].Value ?? 0);
                    decimal xestelik_vereqesi = Convert.ToDecimal(row.Cells[6].Value ?? 0);
                    decimal emr07 = Convert.ToDecimal(row.Cells[7].Value ?? 0);
                    decimal komp_odenisi = Convert.ToDecimal(row.Cells[8].Value ?? 0);
                    decimal hediye = Convert.ToDecimal(row.Cells[9].Value ?? 0);
                    decimal msss = Convert.ToDecimal(row.Cells[10].Value ?? 0);
                    decimal mad98_2_1 = Convert.ToDecimal(row.Cells[11].Value ?? 0);
                    decimal mad98_2_3 = Convert.ToDecimal(row.Cells[12].Value ?? 0);
                    decimal hyh = Convert.ToDecimal(row.Cells[13].Value ?? 0);

                    // Cari tarixdən ay və ili əldə edirik (format MM-yyyy)
                    string cari_il_ay = cmb_Ay_Il.Text;

                    // Məlumatı bazaya əlavə edirik
                    db.InsertMaasElave(cari_il_ay, saa, cari_hesab, avans, mukafat, elave_eh, mezuniyyet_haqqi, xestelik_vereqesi,
                                       emr07, komp_odenisi, hediye, msss, mad98_2_1, mad98_2_3, hyh);
                }

                MessageBox.Show("Məlumatlar uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadDataToGrid()
        {
            List<Isciler> iscilerList = db.GetIsciler(); // Oracle-dan məlumatları alırıq

            dtg_Elaveler.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var isciler in iscilerList)
            {
                dtg_Elaveler.Rows.Add(isciler.SAA,isciler.CariHesab);
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        private void frm_Avans_Load(object sender, EventArgs e)
        {
            // Yeni sətir və sütun əlavə etməyə icazə vermə
            dtg_Elaveler.AllowUserToAddRows = false;
            dtg_Elaveler.AllowUserToDeleteRows = false;

            dtg_Elaveler.Columns[0].Frozen = true;
            dtg_Elaveler.Columns[1].Frozen = true;

            // 0 və 1-ci sütunları redaktəyə bağla
            dtg_Elaveler.Columns[0].ReadOnly = true;
            dtg_Elaveler.Columns[1].ReadOnly = true;

            // 2, 3 və 4-cü sütunları redaktəyə aç
            dtg_Elaveler.Columns[2].ReadOnly = false;
            dtg_Elaveler.Columns[3].ReadOnly = false;
            dtg_Elaveler.Columns[4].ReadOnly = false;

            FillComboBox();
            dtg_Elaveler.AutoGenerateColumns = false;
            LoadDataToGrid();
        }
        private void FillComboBox()
        {
            int currentYear = DateTime.Now.Year; // Cari il
            int currentMonth = DateTime.Now.Month; // Cari ay

            for (int i = 0; i < 12; i++) // 12 ay irəli getmək üçün
            {
                int month = (currentMonth + i) % 12;
                int year = currentYear + (currentMonth + i - 1) / 12;

                if (month == 0)
                {
                    month = 12;
                    year--; // Dekabr üçün il düzəlişi
                }

                string monthYear = new DateTime(year, month, 1).ToString("MM-yyyy");
                cmb_Ay_Il.Items.Add(monthYear);
            }

            cmb_Ay_Il.SelectedIndex = 0; // İlk elementi seçili etmək
        }

        private void btn_Elave_Click(object sender, EventArgs e)
        {
            InsertMaasElaveFromGrid();
        }
    }
}
