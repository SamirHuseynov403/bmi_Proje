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
    public partial class frm_Isciler : Form
    {
        cl_Database db=new cl_Database();
        public frm_Isciler()
        {
            InitializeComponent();
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        private void InsertIsciler()
        {
            try
            {
                string saa = txt_Ad.Text;
                string pincode = txt_Fin.Text;
                string regnum = txt_Qno.Text;
                string ssn = txt_SSN.Text;
                string cariHesab = txt_carihesab.Text;
                string department = txt_Dep.Text;
                string vezife = txt_Vezife.Text;
                DateTime daxilOlmaTarixi = DateTime.Parse(txt_Dax_tar.Text);

                db.InsertIsciler(saa, pincode, regnum, ssn, cariHesab, department, vezife, daxilOlmaTarixi);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void LoadDataToGrid()
        {
            List<Isciler> iscilerList = db.GetIsciler(); // Oracle-dan məlumatları alırıq

            dtg_Isci_dyg.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var isciler in iscilerList)
            {
                dtg_Isci_dyg.Rows.Add(null, null, isciler.Regnum, isciler.SAA, isciler.Pincode, isciler.SSN, isciler.Departament, isciler.Vezife, isciler.CariHesab,
                    isciler.DaxilOlmaTarixi.ToString("dd-MM-yyyy"),isciler.MezuniyyetGS,isciler.EmekHaqqi);
            }
        }
        private void SelectIsciler()
        {
            List<Isciler> iscilerList = db.GetIsciler();
            dtg_Isci_dyg.DataSource = iscilerList;
        }
        private void dtg_genisliyi()
        {
            dtg_Isci_dyg.Columns[0].Width = 100;
            dtg_Isci_dyg.Columns[1].Width = 250;
            dtg_Isci_dyg.Columns[2].Width = 100;
            dtg_Isci_dyg.Columns[3].Width = 150;
            dtg_Isci_dyg.Columns[4].Width = 200;
            dtg_Isci_dyg.Columns[5].Width = 150;
            dtg_Isci_dyg.Columns[6].Width = 150;
            dtg_Isci_dyg.Columns[7].Width = 150;
            dtg_Isci_dyg.Columns[8].Width = 100;
            dtg_Isci_dyg.Columns[9].Width = 130;

            dtg_Isci_dyg.Columns[0].HeaderText = "Qeyd No";
            dtg_Isci_dyg.Columns[1].HeaderText = "Adı";
            dtg_Isci_dyg.Columns[2].HeaderText = "FİN";
            dtg_Isci_dyg.Columns[3].HeaderText = "SSN";
            dtg_Isci_dyg.Columns[4].HeaderText = "Hesab";
            dtg_Isci_dyg.Columns[5].HeaderText = "DP";
            dtg_Isci_dyg.Columns[6].HeaderText = "Vəzifə";
            dtg_Isci_dyg.Columns[7].HeaderText = "Q/O Tarixi";
            dtg_Isci_dyg.Columns[8].HeaderText = "Məzuniyyət";
            dtg_Isci_dyg.Columns[9].HeaderText = "Əmək haqqı";


        }
        private void btn_Elave_Click(object sender, EventArgs e)
        {
            InsertIsciler();
            SelectIsciler();
            dtg_genisliyi();
            
        }
        private void frm_Isciler_Load(object sender, EventArgs e)
        {
            dtg_Isci_dyg.AutoGenerateColumns = false;
            LoadDataToGrid();
            
        }
        private void dtg_Isci_dyg_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;
            if (e.ColumnIndex == 0)
            {
                frm_mezuniyyet_elave_Illik frm = new frm_mezuniyyet_elave_Illik();

                if (dtg_Isci_dyg.Rows[e.RowIndex].Cells[2].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm.txb_QeydNo.Text = dtg_Isci_dyg.Rows[e.RowIndex].Cells[2].Value.ToString();
                }

                if (dtg_Isci_dyg.Rows[e.RowIndex].Cells[3].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm.txb_SAA.Text = dtg_Isci_dyg.Rows[e.RowIndex].Cells[3].Value.ToString();
                }

                frm.ShowDialog();
            }
            if (e.ColumnIndex == 1)
            {
                frm_EmekHaqqi_Yeni frm1 = new frm_EmekHaqqi_Yeni();

                if (dtg_Isci_dyg.Rows[e.RowIndex].Cells[2].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm1.txb_QeydNo.Text = dtg_Isci_dyg.Rows[e.RowIndex].Cells[2].Value.ToString();
                }

                if (dtg_Isci_dyg.Rows[e.RowIndex].Cells[3].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm1.txb_SAA.Text = dtg_Isci_dyg.Rows[e.RowIndex].Cells[3].Value.ToString();
                }

                frm1.ShowDialog();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadDataToGrid();
        }
    }
}
