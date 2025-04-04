using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class geden_hevale : Form
    {
        public geden_hevale()
        {
            InitializeComponent();
        }
        public BindingSource mybing;
        public string icraci_kod { get; set; }
        string icracikod = string.Empty;
        public hevalegeden frmMktbxrc;
        public Mektub mktb;
        private void listtes()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select * From odb.geden_hevale where to_char( tarix,'yyyy')='2022'", Ocon);

            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odt = new DataTable();
            Oda.Fill(Odt);
            dgwgedenhevale.DataSource = Odt;
            Ocon.Close();
        }
        public void listelegedenhevale()
        {
            //where to_char( tarix,'yyyy')='"+tarixIl+"'
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank,ICRA,n.f_i_o From odb.geden_hevale left outer join odb.nameoi n on n.code = icra where to_char( tarix,'yyyy')='" + tarixIl + "'order by hev_nom desc  ", Ocon);
            //order by to_number(substr(hev_nom,6,4)) desc
            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odt = new DataTable();
            Oda.Fill(Odt);
            dgwgedenhevale.DataSource = Odt;
            Ocon.Close();
            dgwgedenhevale.Columns[0].HeaderText = "Həvalə №";
            dgwgedenhevale.Columns[0].Width = 125;
            dgwgedenhevale.Columns[1].HeaderText = "Hesab №";
            dgwgedenhevale.Columns[1].Width = 150;
            dgwgedenhevale.Columns[2].HeaderText = "Adı";
            dgwgedenhevale.Columns[2].Width = 300;
            dgwgedenhevale.Columns[3].HeaderText = "Rezident tipi ";
            dgwgedenhevale.Columns[3].Width = 125;
            dgwgedenhevale.Columns[4].HeaderText = "Məbləğ №";
            dgwgedenhevale.Columns[4].Width = 125;
            dgwgedenhevale.Columns[5].HeaderText = "Valyuta növü ";
            dgwgedenhevale.Columns[5].Width = 125;
            dgwgedenhevale.Columns[6].HeaderText = "Tarix ";
            dgwgedenhevale.Columns[6].Width = 100;
            dgwgedenhevale.Columns[7].HeaderText = "Ölkə mənşəyi ";
            dgwgedenhevale.Columns[7].Width = 150;
            dgwgedenhevale.Columns[8].HeaderText = "Ölkə ";
            dgwgedenhevale.Columns[8].Width = 150;
            dgwgedenhevale.Columns[9].HeaderText = "Həvalə tipi ";
            dgwgedenhevale.Columns[9].Width = 125;
            dgwgedenhevale.Columns[10].HeaderText = "Göndərən tipi";
            dgwgedenhevale.Columns[10].Width = 125;
            dgwgedenhevale.Columns[11].HeaderText = "Alan bank";
            dgwgedenhevale.Columns[11].Width = 200;
            dgwgedenhevale.Columns[12].HeaderText = "İcraçı";
            dgwgedenhevale.Columns[12].Width = 200;
            dgwgedenhevale.Columns[13].HeaderText = "İcraçı adı";
            dgwgedenhevale.Columns[13].Width = 200;
          
        }
        private void geden_hevale_Load(object sender, EventArgs e)
        {
            listelegedenhevale();
            //listtes();
            //mktb = new Mektub();
            //mktb.listelegedenhevale(dgwgedenhevale);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            hevalegeden hvlged = new hevalegeden();
            hvlged.icraci_kod = icraci_kod;
            hvlged.lblSgldml.Text = "insert";
            hvlged.ShowDialog();
            listelegedenhevale();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string test = dgwgedenhevale.CurrentRow.Cells[0].Value.ToString();
            icracikod = dgwgedenhevale.CurrentRow.Cells[12].Value.ToString();
            if (icraci_kod == icracikod)
            {
                DialogResult dr = MessageBox.Show("" + test + " -nömrəli həvalə silinsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                    Ocon.Open();

                    OracleCommand sorgu = new OracleCommand("delete odb.geden_hevale x where x.hev_nom='" + dgwgedenhevale.CurrentRow.Cells[0].Value.ToString() + "'", Ocon);
                    sorgu.ExecuteNonQuery();
                    Ocon.Close();
                    MessageBox.Show("Həvalə uğurla silindi...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    listelegedenhevale();
               }

            }

            else
            {
                MessageBox.Show("Bu həvaləni yalnız həmin İcraçı silə bilər...!", "Diqqət", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void btnYenile_Click(object sender, EventArgs e)
        {
            listelegedenhevale();
            //listtes();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //string test = dgwgedenhevale.CurrentRow.Cells[0].Value.ToString();
            icracikod = dgwgedenhevale.CurrentRow.Cells[12].Value.ToString();
            if (icraci_kod == icracikod)
            {
            hevalegeden hvlged = new hevalegeden();
            hvlged.gedhev = this;
            string test = dgwgedenhevale.CurrentRow.Cells[0].Value.ToString();
            if (test.Length > 0)
            {
                hvlged.txbgedHNo.Text = dgwgedenhevale.CurrentRow.Cells[0].Value.ToString();
                hvlged.txbgedHesNo.Text = dgwgedenhevale.CurrentRow.Cells[1].Value.ToString();
                hvlged.txbgedSAA.Text = dgwgedenhevale.CurrentRow.Cells[2].Value.ToString();
                hvlged.cmbgedveten.Text = dgwgedenhevale.CurrentRow.Cells[3].Value.ToString();
                hvlged.txbgedmebleg.Text = dgwgedenhevale.CurrentRow.Cells[4].Value.ToString();
                hvlged.cmbgedvalyuta.Text = dgwgedenhevale.CurrentRow.Cells[5].Value.ToString();
                hvlged.dtpTarix.Text = dgwgedenhevale.CurrentRow.Cells[6].Value.ToString();
                hvlged.txbgedalanb.Text = dgwgedenhevale.CurrentRow.Cells[7].Value.ToString();
                hvlged.txbgedtesnifat.Text = dgwgedenhevale.CurrentRow.Cells[9].Value.ToString();
                //hvlged.cmbgedvalyuta.Text = dgwgedenhevale.CurrentRow.Cells[8].Value.ToString();
                //hvlged.cmbgedMhes.Text = dgwgedenhevale.CurrentRow.Cells[10].Value.ToString();
                hvlged.textBox1.Text = dgwgedenhevale.CurrentRow.Cells[11].Value.ToString();
                hvlged.cmbgedGonderen.Text = dgwgedenhevale.CurrentRow.Cells[10].Value.ToString();

                hvlged.button2.Text = "Düzəliş et";
                //HEV_NOM,HES_NOM,SAA,Mebleg,Val_tip,tarix,icrachi

                hvlged.lblSgldml.Text = "update";
                hvlged.ShowDialog();
            }
            }
            else
            {
                MessageBox.Show("Yalnız həmin icraçı düzəliş edə bilər !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
                lblSglkod.Text = string.Empty;

            if (textBox1.Text.Trim().Length > 5)
            {
                lblSglkod.Text = "saa='" + textBox1.Text + "'";
            }

            if (txtGonteshgilat.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and upper(hev_nom) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
                else
                {
                    lblSglkod.Text = "upper(hev_nom) like '%" + txtGonteshgilat.Text.ToUpper() + "%'";
                }
            }

            if (textBox1.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and (saa) like '%" + textBox1.Text.ToUpper() + "%'";
                }
                else
                {
                    lblSglkod.Text = "(saa) like '%" + textBox1.Text.ToUpper() + "%'";
                }
            }

            if (txtIcracikodu.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and icra =" + txtIcracikodu.Text + "";
                }
                else
                {
                    lblSglkod.Text = "icra =" + txtIcracikodu.Text + "";
                }
            }

            if (txtIl.Text.Trim().Length > 0)
            {
                if (lblSglkod.Text.Length > 0)
                {
                    lblSglkod.Text += " and to_char( tarix,'yyyy')=" + txtIl.Text + "";
                }
                else
                {
                    lblSglkod.Text = "to_char( tarix,'yyyy')='" + txtIl.Text + "'";
                }
            }
            //
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank,ICRA,n.f_i_o From odb.geden_hevale left outer join odb.nameoi n on n.code = icra where " + lblSglkod.Text + "   ", Ocon);
            //order by to_number(substr(hev_nom,6,4)) desc
            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odt = new DataTable();
            Oda.Fill(Odt);
            dgwgedenhevale.DataSource = Odt;
            Ocon.Close();
            dgwgedenhevale.Columns[0].HeaderText = "Həvalə №";
            dgwgedenhevale.Columns[0].Width = 125;
            dgwgedenhevale.Columns[1].HeaderText = "Hesab №";
            dgwgedenhevale.Columns[1].Width = 150;
            dgwgedenhevale.Columns[2].HeaderText = "Adı";
            dgwgedenhevale.Columns[2].Width = 300;
            dgwgedenhevale.Columns[3].HeaderText = "Rezident tipi ";
            dgwgedenhevale.Columns[3].Width = 125;
            dgwgedenhevale.Columns[4].HeaderText = "Məbləğ №";
            dgwgedenhevale.Columns[4].Width = 125;
            dgwgedenhevale.Columns[5].HeaderText = "Valyuta növü ";
            dgwgedenhevale.Columns[5].Width = 125;
            dgwgedenhevale.Columns[6].HeaderText = "Tarix ";
            dgwgedenhevale.Columns[6].Width = 100;
            dgwgedenhevale.Columns[7].HeaderText = "Ölkə mənşəyi ";
            dgwgedenhevale.Columns[7].Width = 150;
            dgwgedenhevale.Columns[8].HeaderText = "Ölkə ";
            dgwgedenhevale.Columns[8].Width = 150;
            dgwgedenhevale.Columns[9].HeaderText = "Həvalə tipi ";
            dgwgedenhevale.Columns[9].Width = 125;
            dgwgedenhevale.Columns[10].HeaderText = "Göndərən tipi";
            dgwgedenhevale.Columns[10].Width = 125;
            dgwgedenhevale.Columns[11].HeaderText = "Alan bank";
            dgwgedenhevale.Columns[11].Width = 200;
            dgwgedenhevale.Columns[12].HeaderText = "İcraçı";
            dgwgedenhevale.Columns[12].Width = 200;
            dgwgedenhevale.Columns[13].HeaderText = "İcraçı adı";
            dgwgedenhevale.Columns[13].Width = 200;

        }
    }
}
