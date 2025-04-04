using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Data.OleDb;

namespace BMI
{
    public partial class Xarix_olan_mektub : Form
    {
        public Xarix_olan_mektub()
        {
            InitializeComponent();
        }
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;
        public XaricMektub frmMktadd;
        public string il = string.Empty;
        public Mektub mktb;
        public Form1 frmana4;
        string icracikod = string.Empty;
        public string icraci_kod { get; set; }

        private void Xarix_olan_mektub_Load(object sender, EventArgs e)
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            mktb = new Mektub();
            mktb.Xmloxu();
            mktb.il = tarixIl;
            mktb.Xaricmektub(dgwMektubxaric);
            //goster();

        }

        private void goster()
        {
            //try
            //{
                //dgw.Refresh();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
            //Orcom = new OracleCommand("select m.qey_nom, m.tarix, m.gon_yer, m.qisa_mez, m.icraci, m.il from xaric_mektub m Where m.il=" + il + " order by m.kod desc", Orcon);

            //select m.qey_nom, m.tarix,m.gon_yer, m.qisa_mez, m.icraci, n.f_i_o, m.il from odb.xaric_mektub m left outer join odb.nameoi n on n.code = m.icraci Where m.il= " + il + "order by m.kod desc 
            Orcom = new OracleCommand("select m.qey_nom, m.tarix, m.gon_yer, m.qisa_mez, m.icraci, n.f_i_o, m.il from odb.xaric_mektub m left outer join odb.nameoi n on n.code = m.icraci Where m.il= "+il+" order by m.kod desc", Orcon);
            Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dgwMektubxaric.DataSource = Ordt;
                Orcon.Close();
                //dgwMektubxaric.DefaultCellStyle.Encoding = Encoding.UTF8; // DataGridView'deki hücreler için karakter kodlaması ayarı

                dgwMektubxaric.Columns[0].HeaderText = "Məktubun №";
                dgwMektubxaric.Columns[0].Width = 125;

                dgwMektubxaric.Columns[1].HeaderText = "Tarix";
                dgwMektubxaric.Columns[1].Width = 100;

                dgwMektubxaric.Columns[2].HeaderText = "Göndərilən yer";
                dgwMektubxaric.Columns[2].Width = 350;

                dgwMektubxaric.Columns[3].HeaderText = "Qısa məzmun";
                dgwMektubxaric.Columns[3].Width = 420;

                dgwMektubxaric.Columns[4].HeaderText = "İcraçı kodu";
                dgwMektubxaric.Columns[4].Width = 110;

                dgwMektubxaric.Columns[5].HeaderText = "İcraçı";
                dgwMektubxaric.Columns[5].Width = 190;

                dgwMektubxaric.Columns[6].HeaderText = "İl";
                dgwMektubxaric.Columns[6].Width = 80;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button4_Click(object sender, EventArgs e)
        {
            XaricMektub xrcmek = new XaricMektub();
            xrcmek.frmMktbxrc = this;
            xrcmek.lblSgldml.Text = "insert";
            xrcmek.icraci_kod = icraci_kod;
            xrcmek.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string test = dgwMektubxaric.CurrentRow.Cells[0].Value.ToString();
            icracikod = dgwMektubxaric.CurrentRow.Cells[4].Value.ToString();
            if (icraci_kod == icracikod)
            {
                DialogResult dr = MessageBox.Show("" + test + " -nömrəli məktubun silinsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                    Orcon.Open();

                    OracleCommand sorgu = new OracleCommand("delete odb.xaric_mektub x where x.qey_nom='" + dgwMektubxaric.CurrentRow.Cells[0].Value.ToString() + "'", Orcon);
                    sorgu.ExecuteNonQuery();
                    Orcon.Close();
                    MessageBox.Show("Məktub uğurla silindi...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnYenile.PerformClick();
                }

            }

            else
            {
                MessageBox.Show("Bu mətubu yalnız həmin İcraçı silə bilər...!", "Diqqət", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }



            //mktb.Xmgeydiyyatnom = dgwMektubxaric.CurrentRow.Cells[0].Value.ToString();
            //icracikod = dgwMektubxaric.CurrentRow.Cells[4].Value.ToString();

            //if (frmana4.lblicracikodu.Text == icracikod && frmana4.tlsplblTarix.Text == dgwMektubxaric.CurrentRow.Cells[1].Value.ToString().Substring(0, 10))
            //{
            //    DialogResult dr = MessageBox.Show("" + mktb.Xmgeydiyyatnom + " -nömrəli məktubun silinməsini təsdiq edirsiniz ?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //    if (dr == DialogResult.Yes)
            //    {
            //        mktb.Xaricmektdelete();
            //        btnYenile.PerformClick();
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Bu mətubu silmək hüququnuz yoxdur...!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }

        public OracleConnection Orcon { get; set; }

        private void btnYenile_Click(object sender, EventArgs e)
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            mktb = new Mektub();
            mktb.Xmloxu();
            mktb.il = tarixIl;
            mktb.Xaricmektub(dgwMektubxaric);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            mktb.Xmgeydiyyatnom = dgwMektubxaric.CurrentRow.Cells[0].Value.ToString();
            mktb.Xaricmektuboxu();

            frmMktadd = new XaricMektub();
            frmMktadd.frmMktbxrc = this;
            frmMktadd.dtpTarix.Value = DateTime.Parse(mktb.Xaricmektbtarix);
            frmMktadd.cmbGonderyer.Text = mktb.Xmgonderilenyer;
            frmMktadd.cmbGisamezmun.Text = mktb.Xmgisamezmun;
            frmMktadd.rctbMektubmetn.Text = mktb.Xmmektubmetn;

            frmMktadd.lblSgldml.Text = "update";
            frmMktadd.button1.Text = "Düzəliş et";

            icracikod = dgwMektubxaric.CurrentRow.Cells[4].Value.ToString();

            if (icraci_kod!=icracikod)
            {
                frmMktadd.button1.Enabled = false;
            }
            frmMktadd.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            lblSglkod.Text = string.Empty;

            if (textBox1.Text.Trim().Length > 5)
            {
                lblSglkod.Text = "m.qey_nom='" + textBox1.Text + "'";
            }

            if (txtGonteshgilat.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and upper(m.gon_yer) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
                else
                {
                    lblSglkod.Text = "upper(m.gon_yer) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
            }

            if (textBox1.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and upper(m.qisa_mez) like '%" + textBox1.Text.ToUpper() + "%'";
                }
                else
                {
                    lblSglkod.Text = "upper(m.qisa_mez) like '%" + textBox1.Text.ToUpper() + "%'";
                }
            }

            if (txtIcracikodu.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and m.icraci =" + txtIcracikodu.Text + "";
                }
                else
                {
                    lblSglkod.Text = "m.icraci =" + txtIcracikodu.Text + "";
                }
            }

            if (txtIl.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and m.il=" + txtIl.Text + "";
                }
                else
                {
                    lblSglkod.Text = "m.il=" + txtIl.Text + "";
                }
            }

            if (lblSglkod.Text.Length > 0)
            {
                mktb.XaricmektubAxtarish(dgwMektubxaric, lblSglkod.Text);
                //tlsplblSetirsay.Text = string.Empty;
                //tlsplblSetirsay.Text = dgwMektubxaric.RowCount.ToString();
            }
        }
    }
}
