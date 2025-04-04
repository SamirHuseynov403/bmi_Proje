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
    public partial class daxil_olan_mektub : Form
    {
        public  daxil_olan_mektub()
        {
            InitializeComponent();
        }


        public  mektubdaxil dxlmektb;
        public Anakredit frmana4;
        public daxil_olam_mektub frmMktbdadd;
        string icracikod = string.Empty;
        public string icraci_kod { get; set; }
        public mektubdaxil mktb;

      
        private void daxil_olan_mektub_Load(object sender, EventArgs e)
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            dxlmektb = new mektubdaxil();
            dxlmektb.Xmloxu();
            dxlmektb.il = tarixIl;
            dxlmektb.Daxilmektubgos(dgwMektubdaxil);

            

        }

        public DataTable DataSource { get; set; }

        private void button4_Click(object sender, EventArgs e)
        {
            frmMktbdadd = new daxil_olam_mektub();
            frmMktbdadd.frmMktbdaxadd = this;
            frmMktbdadd.lblSgldml.Text = "insert";
            frmMktbdadd.ShowDialog();
            frmMktbdadd.icraci_kod = icraci_kod;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            lblSglkod.Text = string.Empty;

            if (txtGeynom.Text.Trim().Length > 0)
            {
                lblSglkod.Text = " d.nom1=" + txtGeynom.Text + "";
            }

            if (msktxtDtarixev.Text.Length == 10 && msktxtDtarixson.Text.Length == 10)
            {
                try
                {
                    DateTime Dtarixev = DateTime.Parse(msktxtDtarixev.Text);
                    DateTime Dtarixson = DateTime.Parse(msktxtDtarixson.Text);

                    if (Dtarixson >= Dtarixev)
                    {
                        if (lblSglkod.Text.Length > 0)
                        {
                            lblSglkod.Text += " and d.dax_tarix >=TO_DATE('" + msktxtDtarixev.Text + "','dd-MM-yyyy') and d.dax_tarix <= TO_DATE('" + msktxtDtarixson.Text + "', 'dd-MM-yyyy')";
                        }
                        else
                        {
                            lblSglkod.Text = " d.dax_tarix >=TO_DATE('" + msktxtDtarixev.Text + "','dd-MM-yyyy') and d.dax_tarix <= TO_DATE('" + msktxtDtarixson.Text + "', 'dd-MM-yyyy')";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Daxil olma tarixi üzrə \nDövrün son tarixi - Dövrün əvvəlindən böyük olmalıdır !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                catch (FormatException)
                {
                    MessageBox.Show("Daxil olma tarixi üzrə \nDövrün son tarixi - Dövrün əvvəlindən böyük olmalıdır !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if (txtGonteshgilat.Text.Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and upper(d.idare_adi) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
                else
                {
                    lblSglkod.Text = "upper(d.idare_adi) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
            }

            if (msktxtGtarixev.Text.Length == 10 && msktxtGtarixson.Text.Length == 10)
            {
                try
                {
                    DateTime Gtarixev = DateTime.Parse(msktxtGtarixev.Text);
                    DateTime Gtarixson = DateTime.Parse(msktxtGtarixson.Text);

                    if (Gtarixson >= Gtarixev)
                    {
                        if (lblSglkod.Text.Length > 0)
                        {
                            lblSglkod.Text += " and d.gon_tarix >=TO_DATE('" + msktxtGtarixev.Text + "','dd-MM-yyyy') and d.gon_tarix <= TO_DATE('" + msktxtGtarixson.Text + "', 'dd-MM-yyyy')";
                        }
                        else
                        {
                            lblSglkod.Text = "d.gon_tarix >=TO_DATE('" + msktxtGtarixev.Text + "','dd-MM-yyyy') and d.gon_tarix <= TO_DATE('" + msktxtGtarixson.Text + "', 'dd-MM-yyyy')";
                        }
                    }
                    else
                    {
                        MessageBox.Show("Göndərilən tarix üzrə \nDövrün son tarixi - Dövrün əvvəlindən böyük olmalıdır !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                catch (FormatException)
                {
                    MessageBox.Show("Göndərilən tarix üzrə \nDövrün son tarixi - Dövrün əvvəlindən böyük olmalıdır !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if (cmbMektubnom.Text.Trim().Length > 0)
            {
                if (cmbMektubnom.Text == "null")
                {
                    if (lblSglkod.Text.Length > 0)
                    {
                        lblSglkod.Text += " and d.dax_nom is null";
                    }
                    else
                    {
                        lblSglkod.Text = "d.dax_nom is null";
                    }
                }
                else
                {
                    if (lblSglkod.Text.Length > 0)
                    {
                        lblSglkod.Text += " and d.dax_nom like '%" + cmbMektubnom.Text + "%'";
                    }
                    else
                    {
                        lblSglkod.Text = " d.dax_nom like '%" + cmbMektubnom.Text + "%'";
                    }
                }
            }
            if (txtIcracikod.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and d.mek_unvan=" + txtIcracikod.Text + "";
                }
                else
                {
                    lblSglkod.Text = "d.mek_unvan=" + txtIcracikod.Text + "";
                }
            }

            if (txtIl.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and d.il=" + txtIl.Text + "";
                }
                else
                {
                    lblSglkod.Text = "d.il=" + txtIl.Text + "";
                }
            }

            if (lblSglkod.Text.Length > 0)
            {
                dxlmektb.DaxilmektubAxtarish(dgwMektubdaxil, lblSglkod.Text);
                //tlsplblSetirsay.Text = string.Empty;
                //tlsplblSetirsay.Text = dgwMektubdaxil.RowCount.ToString();
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dxlmektb.siranom = dgwMektubdaxil.CurrentRow.Cells[0].Value.ToString();
            dxlmektb.icracikodu = dgwMektubdaxil.CurrentRow.Cells[5].Value.ToString();
            dxlmektb.il = dgwMektubdaxil.CurrentRow.Cells[7].Value.ToString();
            icracikod = dgwMektubdaxil.CurrentRow.Cells[4].Value.ToString();

            if (icraci_kod == icracikod)
            {
                DialogResult dr = MessageBox.Show("" + dxlmektb.siranom + " -nömrəli məktubun silinsin?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    dxlmektb.Daxilmktbdelete();
                    btnyenile.PerformClick();
                }
            }
            else
            {
                MessageBox.Show("Bu mətubu yalnız həmin İcraçı silə bilər...!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnyenile_Click(object sender, EventArgs e)
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            mktb = new mektubdaxil();
            mktb.Xmloxu();
            mktb.il = tarixIl;
            mktb.Daxilmektubgos(dgwMektubdaxil);

        }
    }
}
