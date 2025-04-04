using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMI
{
    public partial class daxil_olam_mektub : Form
    {
        public daxil_olam_mektub()
        {
            InitializeComponent();
        }
        public daxil_olan_mektub frmMktbdaxadd;
        public string icraci_kod = string.Empty;

        private void button1_Click(object sender, EventArgs e)
        {
            if (lblSgldml.Text == "insert")
            {
                if (cmbGonteshgilat.Text.Trim().Length > 0)
                {
                    DialogResult dr = MessageBox.Show("Məktubun qeydiyyata alınsın ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        frmMktbdaxadd.dxlmektb.Daxiltarix = dtpDaxiloltarix.Value.ToShortDateString();
                        frmMktbdaxadd.dxlmektb.Teshkilatadi = cmbGonteshgilat.Text.Trim();
                        frmMktbdaxadd.dxlmektb.GonTarix = dtpGontarix.Value.ToShortDateString();
                        frmMktbdaxadd.dxlmektb.Mektubnom = cmbMektubnom.Text.Trim();
                        frmMktbdaxadd.dxlmektb.icracikodu = icraci_kod;
                        frmMktbdaxadd.dxlmektb.Daxilmktbinsert();

                        //btnTemizle.PerformClick();
                        frmMktbdaxadd.btnyenile.PerformClick();
                    }
                }
                else
                {
                    MessageBox.Show("Göndərən təşgilatın adı yazılmayıb", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (lblSgldml.Text == "update")
            {
                if (cmbGonteshgilat.Text.Trim().Length > 0)
                {
                    DialogResult dr = MessageBox.Show("Məktuba düzəliş olunsun ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        frmMktbdaxadd.dxlmektb.Daxiltarix = dtpDaxiloltarix.Value.ToShortDateString();
                        frmMktbdaxadd.dxlmektb.Teshkilatadi = cmbGonteshgilat.Text.Trim();
                        frmMktbdaxadd.dxlmektb.GonTarix = dtpGontarix.Value.ToShortDateString();
                        frmMktbdaxadd.dxlmektb.Mektubnom = cmbMektubnom.Text.Trim();
                        frmMktbdaxadd.dxlmektb.Daxilmktbupdate();

                        frmMktbdaxadd.btnyenile.PerformClick();
                    }
                }
                else
                {
                    MessageBox.Show("Göndərən təşgilatın adı yazılmayıb", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }
}
