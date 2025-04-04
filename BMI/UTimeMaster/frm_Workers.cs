using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.UTimeMaster
{
    public partial class frm_Workers : Form
    {
        public frm_Workers()
        {
            InitializeComponent();
        }
        void bazadancek()
        {
            string connString = "Host=172.16.0.5;Port=7496;Username=postgres;Password=123456;Database=biotime";



            using (NpgsqlConnection conn = new NpgsqlConnection(connString))
            {
                try
                {
                    conn.Open();
                    string sql = @"select em.first_name,d.dept_name,p.position_name from personnel_employee em,personnel_position p,personnel_department d
                        where em.position_id=CAST(p.position_code AS integer) and em.department_id=CAST(d.dept_code AS integer) and em.enroll_sn =''";
                    NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    // ✅ Əvvəlcə təmizləyirik
                    dtg_isciler.Rows.Clear();

                   
                    // ✅ Məlumatları əlavə edirik
                    foreach (DataRow row in dt.Rows)
                    {
                        dtg_isciler.Rows.Add(
                            row["first_name"].ToString(),
                            row["dept_name"].ToString(),
                            row["position_name"].ToString()
                        );
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message);
                }
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void frm_Workers_Load(object sender, EventArgs e)
        {
            bazadancek();
        }
    }
}
