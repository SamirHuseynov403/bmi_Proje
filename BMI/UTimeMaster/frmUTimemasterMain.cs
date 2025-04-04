using BMI.Emek_haqqi_ve_davamiyyet;
using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using BMI.Muhasibat;
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
    public partial class frmUTimemasterMain : Form
    {
        cl_yanasmalar cl = new cl_yanasmalar();
        cl_Database db = new cl_Database();
        Aletler aletler = new Aletler();
        public frmUTimemasterMain()
        {
            InitializeComponent();
        }
        private Dictionary<Type, Form> formInstances = new Dictionary<Type, Form>();
        private Form currentForm = null;
        void bazadancek()
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(cl.connStringUTimemaster))
            {
                try
                {
                    conn.Open();
                    string sql = @"SELECT 
                        pe.first_name,
                        it.punch_time::date AS work_date,   
                        TO_CHAR(it.punch_time AT TIME ZONE 'UTC', 'HH12:MI AM') AS giris, 
	                    TO_CHAR(it.upload_time, 'HH24:MI AM') AS upload_time, 
	                    CASE 
                            WHEN (it.punch_time AT TIME ZONE 'UTC')::time > TIME '09:00:00' 
                            THEN ROUND(
                                EXTRACT(EPOCH FROM ((it.punch_time AT TIME ZONE 'UTC')::time - TIME '09:00:00')) / 60
                            ) || ' dəqiqə'
                            ELSE ''
                        END AS gecikme
                    FROM 
                        iclock_transaction it
                    JOIN 
                        personnel_employee pe 
                    ON 
                        it.emp_code = pe.emp_code
                    WHERE 
                        it.punch_time::date = '2025-02-24'/*24-28,03*/
                        AND TO_CHAR(it.punch_time AT TIME ZONE 'UTC', 'AM') = 'AM'
                    ORDER BY it.punch_time ";
                    NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    //dataGridView1.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message);
                }
            }
        }
        public void OpenFormInPanel(Type formType)
        {
            // Əgər əvvəl açılmış form varsa və tipi eynidirsə, sadəcə göstəririk
            if (currentForm != null && currentForm.GetType() == formType)
            {
                currentForm.Show();
                return;
            }

            // Əgər əvvəl bir form açıqdırsa, onu gizlədirik
            if (currentForm != null)
            {
                currentForm.Hide();
            }

            // Form artıq yaradılıbsa, onu istifadə edirik, əks halda yenisini yaradırıq
            if (!formInstances.ContainsKey(formType))
            {
                Form newForm = (Form)Activator.CreateInstance(formType);
                newForm.TopLevel = false;
                newForm.FormBorderStyle = FormBorderStyle.None;
                newForm.Dock = DockStyle.Fill;

                formInstances[formType] = newForm;
            }

            currentForm = formInstances[formType];

            pnl_orta.Controls.Clear();
            pnl_orta.Controls.Add(currentForm);
            currentForm.Show();
        }
        private void frmUTimemasterMain_Load(object sender, EventArgs e)
        {
            //bazadancek();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Isciler_Click(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof(frm_Workers));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof(frm_InputOutput));
        }
    }
}
