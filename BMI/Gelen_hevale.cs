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
    public partial class Gelen_hevale : Form
    {
        public Gelen_hevale()
        {
            InitializeComponent();
        }
        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod { get; set; }
        string icracikod = string.Empty;
        public Mektub mktb;

        private void listelegelen()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();

        }

        private void listtes()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select * From odb.gelen_hevale where to_char( tarix,'yyyy')='2018'", Ocon);

            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odt = new DataTable();
            Oda.Fill(Odt);
            dgwgeLenhevale.DataSource = Odt;
            Ocon.Close();
        }

        private void Gelen_hevale_Load(object sender, EventArgs e)
        {
            //listtes();
            listeleLedenhevale();
        }

        public void listeleLedenhevale()
        {
            //where to_char( tarix,'yyyy')='"+tarixIl+"'
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,hev_tip,gel_olke,gon_tip,al_bank,icra,n.f_i_o From odb.gelen_hevale left outer join odb.nameoi n on n.code = icra where to_char( tarix,'yyyy')='" + tarixIl + "'  ", Ocon);
            //order by to_number(substr(hev_nom,6,4)) desc
            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odta = new DataTable();
            Oda.Fill(Odta);
            dgwgeLenhevale.DataSource = Odta;
            Ocon.Close();
            dgwgeLenhevale.Columns[0].HeaderText = "Həvalə №";
            dgwgeLenhevale.Columns[0].Width = 125;
            dgwgeLenhevale.Columns[1].HeaderText = "Hesab №";
            dgwgeLenhevale.Columns[1].Width = 150;
            dgwgeLenhevale.Columns[2].HeaderText = "Adı";
            dgwgeLenhevale.Columns[2].Width = 200;
            dgwgeLenhevale.Columns[3].HeaderText = "Rezident tipi ";
            dgwgeLenhevale.Columns[3].Width = 125;
            dgwgeLenhevale.Columns[4].HeaderText = "Məbləğ №";
            dgwgeLenhevale.Columns[4].Width = 125;
            dgwgeLenhevale.Columns[5].HeaderText = "Valyuta növü ";
            dgwgeLenhevale.Columns[5].Width = 125;
            dgwgeLenhevale.Columns[6].HeaderText = "Tarix ";
            dgwgeLenhevale.Columns[6].Width = 100;
            dgwgeLenhevale.Columns[7].HeaderText = "Ölkə mənşəyi ";
            dgwgeLenhevale.Columns[7].Width = 150;
            
            dgwgeLenhevale.Columns[9].HeaderText = "Həvalə tipi ";
            dgwgeLenhevale.Columns[9].Width = 125;
            dgwgeLenhevale.Columns[8].HeaderText = "Ölkə ";
            dgwgeLenhevale.Columns[8].Width = 150;

            dgwgeLenhevale.Columns[10].HeaderText = "Göndərən tipi";
            dgwgeLenhevale.Columns[10].Width = 125;
            dgwgeLenhevale.Columns[11].HeaderText = "Alan bank";
            dgwgeLenhevale.Columns[11].Width = 200;
            dgwgeLenhevale.Columns[12].HeaderText = "İcraçı";
            dgwgeLenhevale.Columns[12].Width = 200;
            dgwgeLenhevale.Columns[13].HeaderText = "İcraçı adı";
            dgwgeLenhevale.Columns[13].Width = 200;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Gel_hev hvlged = new Gel_hev();
            hvlged.icraci_kod = icraci_kod;
            hvlged.lblSgldml.Text = "insert";
            hvlged.ShowDialog();
            listeleLedenhevale();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Gel_hev hvlged = new Gel_hev();
            hvlged.gedhev = this;
            string test = dgwgeLenhevale.CurrentRow.Cells[0].Value.ToString();
            if (test.Length > 0)
            {
                hvlged.txbgedHNo.Text = dgwgeLenhevale.CurrentRow.Cells[0].Value.ToString();
                hvlged.txbgedHesNo.Text = dgwgeLenhevale.CurrentRow.Cells[1].Value.ToString();
                hvlged.txbgedSAA.Text = dgwgeLenhevale.CurrentRow.Cells[2].Value.ToString();
                hvlged.txbgedmebleg.Text = dgwgeLenhevale.CurrentRow.Cells[4].Value.ToString();
                hvlged.cmbgedveten.Text = dgwgeLenhevale.CurrentRow.Cells[3].Value.ToString();
                hvlged.cmbgedvalyuta.Text = dgwgeLenhevale.CurrentRow.Cells[5].Value.ToString();
                hvlged.dtpTarix.Text = dgwgeLenhevale.CurrentRow.Cells[6].Value.ToString();
                hvlged.txbgedmense.Text = dgwgeLenhevale.CurrentRow.Cells[7].Value.ToString();
                hvlged.txbgedtesnifat.Text = dgwgeLenhevale.CurrentRow.Cells[9].Value.ToString();
                hvlged.cmbgedGonderen.Text = dgwgeLenhevale.CurrentRow.Cells[10].Value.ToString();
                hvlged.txbgedalanb.Text = dgwgeLenhevale.CurrentRow.Cells[11].Value.ToString();
                hvlged.cmbgedMhes.Text = dgwgeLenhevale.CurrentRow.Cells[8].Value.ToString();

                hvlged.button1.Text = "Düzəliş et";
                hvlged.lblSgldml.Text = "update";
                //HEV_NOM,HES_NOM,SAA,Mebleg,Val_tip,tarix,icrachi

                hvlged.ShowDialog();
            }
            else
            {
                MessageBox.Show("Düzəliş olunacaq məlumat seçilməyib !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string test = dgwgeLenhevale.CurrentRow.Cells[0].Value.ToString();
            icracikod = dgwgeLenhevale.CurrentRow.Cells[12].Value.ToString();
            if (icraci_kod == icracikod)
            {
                DialogResult dr = MessageBox.Show("" + test + " -nömrəli həvalə silinsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                    Ocon.Open();

                    OracleCommand sorgu = new OracleCommand("delete odb.gelen_hevale x where x.hev_nom='" + dgwgeLenhevale.CurrentRow.Cells[0].Value.ToString() + "'", Ocon);
                    sorgu.ExecuteNonQuery();
                    Ocon.Close();
                    MessageBox.Show("Həvalə uğurla silindi...", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                }

            }

            else
            {
                MessageBox.Show("Bu həvaləni yalnız həmin İcraçı silə bilər...!", "Diqqət", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
                listeleLedenhevale();
        }

        private void btnYenile_Click(object sender, EventArgs e)
        {
            listeleLedenhevale();
        }

        private void dgwgeLenhevale_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
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

            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
             OracleCommand Ocom = new OracleCommand("Select hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,hev_tip,gel_olke,gon_tip,al_bank,icra,n.f_i_o From odb.gelen_hevale left outer join odb.nameoi n on n.code = icra where " + lblSglkod.Text + "  ", Ocon);
            //order by to_number(substr(hev_nom,6,4)) desc
            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odta = new DataTable();
            Oda.Fill(Odta);
            dgwgeLenhevale.DataSource = Odta;
            Ocon.Close();
            dgwgeLenhevale.Columns[0].HeaderText = "Həvalə №";
            dgwgeLenhevale.Columns[0].Width = 125;
            dgwgeLenhevale.Columns[1].HeaderText = "Hesab №";
            dgwgeLenhevale.Columns[1].Width = 150;
            dgwgeLenhevale.Columns[2].HeaderText = "Adı";
            dgwgeLenhevale.Columns[2].Width = 200;
            dgwgeLenhevale.Columns[3].HeaderText = "Rezident tipi ";
            dgwgeLenhevale.Columns[3].Width = 125;
            dgwgeLenhevale.Columns[4].HeaderText = "Məbləğ №";
            dgwgeLenhevale.Columns[4].Width = 125;
            dgwgeLenhevale.Columns[5].HeaderText = "Valyuta növü ";
            dgwgeLenhevale.Columns[5].Width = 125;
            dgwgeLenhevale.Columns[6].HeaderText = "Tarix ";
            dgwgeLenhevale.Columns[6].Width = 100;
            dgwgeLenhevale.Columns[7].HeaderText = "Ölkə mənşəyi ";
            dgwgeLenhevale.Columns[7].Width = 150;

            dgwgeLenhevale.Columns[9].HeaderText = "Həvalə tipi ";
            dgwgeLenhevale.Columns[9].Width = 125;
            dgwgeLenhevale.Columns[8].HeaderText = "Ölkə ";
            dgwgeLenhevale.Columns[8].Width = 150;

            dgwgeLenhevale.Columns[10].HeaderText = "Göndərən tipi";
            dgwgeLenhevale.Columns[10].Width = 125;
            dgwgeLenhevale.Columns[11].HeaderText = "Alan bank";
            dgwgeLenhevale.Columns[11].Width = 200;
            dgwgeLenhevale.Columns[12].HeaderText = "İcraçı";
            dgwgeLenhevale.Columns[12].Width = 200;
            dgwgeLenhevale.Columns[13].HeaderText = "İcraçı adı";
            dgwgeLenhevale.Columns[13].Width = 200;

            try
            {
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //if (lblSglkod.Text.Length > 0)
            //{
            //    mktb.gelenhevAxtarish(dgwgeLenhevale, lblSglkod.Text);
            //    //tlsplblSetirsay.Text = string.Empty;
            //    //tlsplblSetirsay.Text = dgwMektubxaric.RowCount.ToString();
            //}
            }
            catch (Exception)
            {
                
                
            }
        }

       // public object mybing { get; set; }
    }

}
