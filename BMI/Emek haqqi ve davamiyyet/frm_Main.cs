using BMI.Emek_haqqi_ve_davamiyyet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.Emeliyyat
{
    public partial class frm_Main : Form
    {
        private Dictionary<Type, Form> formInstances = new Dictionary<Type, Form>();
        private Form currentForm = null;
        public frm_Main()
        {
            InitializeComponent();
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
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof( frm_Isciler)); // Burada çağırılan form dəyişdirilə bilər
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_MXE_Click(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof(frm_MXE)); // Burada çağırılan form dəyişdirilə bilər
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof(frm_emek_haqq_stat));
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFormInPanel(typeof(frm_Elaveler));
        }
    }
}
