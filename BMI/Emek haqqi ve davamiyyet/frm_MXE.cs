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
    public partial class frm_MXE : Form
    {
        public frm_MXE()
        {
            InitializeComponent();
        }
        cl_Database db = new cl_Database();
        cl_isciler isciler1 = new cl_isciler();
        

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void frm_MXE_Load(object sender, EventArgs e)
        {
            frm_MXE_dtg.AutoGenerateColumns = false;
            LoadDataToGrid();
        }
        private void LoadDataToGrid()
        {
            List<Isciler> iscilerList = db.GetIsciler(); // Oracle-dan məlumatları alırıq

            frm_MXE_dtg.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var isciler in iscilerList)
            {
                frm_MXE_dtg.Rows.Add(null, null, isciler.Regnum, isciler.SAA, isciler.SSN, isciler.Pincode,isciler.DaxilOlmaTarixi.ToString("dd-MM-yyyy"),isciler.Staj,isciler.MezuniyyetQaliq);
            }
        }

        private void frm_MXE_dtg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
            if (e.ColumnIndex == 1)
            {
                frm_MXO frm = new frm_MXO();
                if (frm_MXE_dtg.Rows[e.RowIndex].Cells[2].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm.txb_QeydNo.Text = frm_MXE_dtg.Rows[e.RowIndex].Cells[2].Value.ToString();
                }

                if (frm_MXE_dtg.Rows[e.RowIndex].Cells[3].Value != null) // 2-ci sütun (SAA və ya başqa sütun)
                {
                    frm.txb_SAA.Text = frm_MXE_dtg.Rows[e.RowIndex].Cells[3].Value.ToString();
                }

                frm.ShowDialog();
            }
            else if (e.ColumnIndex == 0)
            {

            }
        }

        private void frm_MXE_dtg_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
