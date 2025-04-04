using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel.FinancialFunctions;

namespace BMI
{
    public partial class frmmuracietler : Form
    {
        public frmmuracietler()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");

        OleDbCommand islek;
        OracleCommand komyazdir;
        OracleCommand komutmurraz;
        public string icraci_kod = string.Empty;
        DataTable muracietler = new DataTable();
        string tarixIl = DateTime.Now.Date.Year.ToString();
        string tarixay = DateTime.Now.Date.Month.ToString();
        string tarixgun = DateTime.Now.Date.Day.ToString();
       public void listelemuraciet()
        {

            try
            {


                muracietler.Clear();
                string tarixIl = txbil.Text;
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State==ConnectionState.Closed)
                {
                    con.Open();
                }
                
                
                OracleDataAdapter isleme = new OracleDataAdapter("select Sira,Tarix,Adi,Muraciet_mebleg,Raziliq_mebleg,Etiraz_sebebi,Qeyd,Komite_qerar,Qerar_qeyd,Danisiq,Telefon,Melumat,Gelme_tarixi,Saat,İcraci,il from odb.Muracietler where il>='"+ iltarix + "' order by sira Desc ", con);
                isleme.Fill(muracietler);
                gridControl1.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }
            
                gridmuracietler.Columns[0].Caption = "Sira";
                gridmuracietler.Columns[0].Width = 80;
                gridmuracietler.Columns[1].Caption = "Tarix";
                gridmuracietler.Columns[1].Width = 110;
                gridmuracietler.Columns[2].Caption = "Adı";
                gridmuracietler.Columns[2].Width = 220;
                gridmuracietler.Columns[3].Caption = "Müraciət məbləği";
                gridmuracietler.Columns[3].Width = 100;
                gridmuracietler.Columns[4].Caption = "Razılıq məbləği";
                gridmuracietler.Columns[4].Width = 100;
                gridmuracietler.Columns[5].Caption = "Etiraz səbəbi";
                gridmuracietler.Columns[5].Width = 250;
                gridmuracietler.Columns[6].Caption = "Qeyd";
                gridmuracietler.Columns[6].Width = 150;
                gridmuracietler.Columns[7].Caption = "Komitə qərarı";
                gridmuracietler.Columns[7].Width = 200;
                gridmuracietler.Columns[8].Caption = "Qərar qeyd";
                gridmuracietler.Columns[8].Width = 100;
                gridmuracietler.Columns[9].Caption = "Danışıq";
                gridmuracietler.Columns[9].Width = 100;
                gridmuracietler.Columns[10].Caption = "Telefon";
                gridmuracietler.Columns[10].Width = 200;
                gridmuracietler.Columns[11].Caption = "Məlumat";
                gridmuracietler.Columns[11].Width = 100;
                gridmuracietler.Columns[12].Caption = "Gəlmə tarixi";
                gridmuracietler.Columns[12].Width = 100;
                gridmuracietler.Columns[13].Caption = "Saat";
                gridmuracietler.Columns[13].Width = 70;
                gridmuracietler.Columns[14].Caption = "İcraçı";
                gridmuracietler.Columns[14].Width = 100;
                gridmuracietler.Columns[15].Caption = "Il";
                gridmuracietler.Columns[15].Width = 50;

            }
            catch (Exception)
            {


            }
            finally { };


        }

        private void frmmuracietler_Load(object sender, EventArgs e)
        {
            txbil.Text = DateTime.Now.Date.Year.ToString();
            dateEdit1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            listelemuraciet();
            muracietsay();
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            
        }

        private void gridmuracietler_DoubleClick(object sender, EventArgs e)
        {
            
        }
        void listelemuraciettest()
        {

            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            con.Open();
            muracietler.Clear();
            OracleDataAdapter isleme = new OracleDataAdapter("select * from odb.Muracietler  where il=>'" + txbil + "'", con);
            isleme.Fill(muracietler);
            gridControl1.DataSource = muracietler;
            con.Close();
            gridmuracietler.Columns[0].Caption = "Sira";
            gridmuracietler.Columns[0].Width = 60;
            gridmuracietler.Columns[1].Caption = "Tarix";
            gridmuracietler.Columns[1].Width = 110;
            gridmuracietler.Columns[2].Caption = "Adı";
            gridmuracietler.Columns[2].Width = 320;
            gridmuracietler.Columns[3].Caption = "Valyuta";
            gridmuracietler.Columns[3].Width = 100;
            gridmuracietler.Columns[4].Caption = "Müraciət məbləği";
            gridmuracietler.Columns[4].Width = 150;
            gridmuracietler.Columns[5].Caption = "Razılıq məbləği";
            gridmuracietler.Columns[5].Width = 150;
            gridmuracietler.Columns[6].Caption = "Təminat";
            gridmuracietler.Columns[6].Width = 250;
            gridmuracietler.Columns[7].Caption = "Etiraz səbəbi";
            gridmuracietler.Columns[7].Width = 250;
            gridmuracietler.Columns[8].Caption = "Qeyd";
            gridmuracietler.Columns[8].Width = 250;
            gridmuracietler.Columns[9].Caption = "Komitə qərarı";
            gridmuracietler.Columns[9].Width = 250;
            gridmuracietler.Columns[10].Caption = "Qərar qeyd";
            gridmuracietler.Columns[10].Width = 250;
            gridmuracietler.Columns[11].Caption = "Telefon";
            gridmuracietler.Columns[11].Width = 250;
            gridmuracietler.Columns[12].Caption = "Danışıq";
            gridmuracietler.Columns[12].Width = 250;
            gridmuracietler.Columns[13].Caption = "Məlumat";
            gridmuracietler.Columns[13].Width = 250;
            gridmuracietler.Columns[14].Caption = "Gəlmə tarixi";
            gridmuracietler.Columns[14].Width = 120;
            gridmuracietler.Columns[15].Caption = "Saat";
            gridmuracietler.Columns[15].Width = 70;
            gridmuracietler.Columns[16].Caption = "İl";
            gridmuracietler.Columns[16].Width = 250;

        
}

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            frmmuracietyazdir frmm = new frmmuracietyazdir();
            frmm.Show();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            frmmuracietyazdir frmm = new frmmuracietyazdir();
            frmm.txbmurad.Text = gridmuracietler.GetFocusedRowCellValue("Adı").ToString();
            frmm.Show();
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            listelemuraciet();
        }

        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            //txbmurad.Text = gridmuracietler.GetFocusedRowCellValue.[0].Value.ToString();
            
        }

        private void gridmuracietler_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            //txbmursira.Text = gridmuracietler.GetFocusedRowCellValue("Adı").ToString();
        }

        private void simpleButton1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
        private void muracietsay()

        {
            try
            {
                int test = 2022;
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                string proid;
                string query = "select sira from odb.Muracietler where il='" + iltarix + "' order by  sira Desc";
                
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                OracleCommand cmd = new OracleCommand(query, con);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString()) + 1;
                    proid = id.ToString();
                }
                else if (Convert.IsDBNull(dr))
                {
                    proid = ("1");
                }
                else
                {
                    proid = ("1");
                }
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                
                txbmursira.Text = proid.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void temizlemuraciet()
        {
            try
            {
                txbmurad.Text = "";
                txbmurmeb.Text = "";
                cmbmurdan.Text = "";
                cmbkrverilmesi.Text = "";
                cmbteminati.Text = "";
                cmbetirazseb.Text = "";
                txbkomiteqeyd.Text = "";
                txbumumiqeyd.Text = "";
                txbmurtel.Text = "";
                txtraziliq.Text = "";

            }
            catch (Exception)
            {


            }


        }
        private void BtnEkle_Click_1(object sender, EventArgs e)
        {
            try
            {
                muracietsay();
                int tmeb = Convert.ToInt32(txbmurmeb.Text);
                int traz = Convert.ToInt32(txtraziliq.Text);
                int tsira = Convert.ToInt32(txbmursira.Text);
                // dateTimePicker4.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                string tarixIl = DateTime.Now.Date.Year.ToString();
                
                con.Open();
                komyazdir = new OracleCommand("insert into odb.Muracietler(Sira,Tarix,adi,Muraciet_mebleg,Raziliq_mebleg,Etiraz_sebebi,Qeyd,Komite_qerar,Qerar_qeyd,danisiq,Telefon,Melumat,İcraci,il) values ( '" + tsira + "', '" + dateEdit1.Text + "','" + txbmurad.Text + "','" + tmeb + "','" + traz + "','" + cmbetirazseb.Text + "', '" + txbumumiqeyd.Text + "', '" + cmbteminati.Text + "', '" + txbkomiteqeyd.Text + "','" + cmbmurdan.Text + "','" + txbmurtel.Text + "','" + cmbkrverilmesi.Text + "','" + icraci_kod + "','" + tarixIl + "')", con);
                komyazdir.ExecuteNonQuery();
                con.Close();
                listelemuraciet();
                
                temizlemuraciet();
            }
            catch (Exception)
            {


            }
            finally { }

        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {
            try
            {
                double murmeb = Convert.ToDouble(txbmurmeb.Text);
                double murraz = Convert.ToDouble(txtraziliq.Text);
                string sorgu5 = "Update odb.Muracietler set tarix='" + dateEdit1.Text + "', Adi='" + txbmurad.Text + "',Muraciet_mebleg='" + murmeb + "',Raziliq_mebleg='" + murraz + "',Etiraz_sebebi='" + cmbetirazseb.Text + "',Qeyd='" + txbumumiqeyd.Text + "', Komite_qerar='" + cmbteminati.Text + "',Qerar_qeyd='" + txbkomiteqeyd.Text + "',Danisiq='" + cmbmurdan.Text + "',Telefon='" + txbmurtel.Text + "',Melumat='" + cmbkrverilmesi.Text + "' where Sira=" + txbsiranoelave.Text + " and il='" + txbil.Text + "'";
                komyazdir = new OracleCommand(sorgu5, con);
                con.Open();
                komyazdir.ExecuteNonQuery();
                con.Close();
                XtraMessageBox.Show("Məlumatlar yeniləndi", "Qeyd");
                muracietler.Clear();
                listelemuraciet();
                temizlemuraciet();
                muracietsay();
            }
            catch (Exception)
            {
                throw;
            }
            finally { }
        }

        private void gridmuracietler_DoubleClick_1(object sender, EventArgs e)
        {
            
            
        }

        private void BtnListele_Click(object sender, EventArgs e)
        {
            temizlemuraciet();
        }

        private void gridmuracietler_KeyUp(object sender, KeyEventArgs e)
        {
            
        }

        private void gridmuracietler_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataRow row = gridmuracietler.GetDataRow(gridmuracietler.FocusedRowHandle);
                txbmurad.Text = row[2].ToString();
                txbmurmeb.Text = row[3].ToString();
                txtraziliq.Text = row[4].ToString();
                cmbetirazseb.Text = row[5].ToString();
                cmbteminati.Text = row[7].ToString();
                txbkomiteqeyd.Text = row[8].ToString();
                txbumumiqeyd.Text = row[6].ToString();
                cmbmurdan.Text = row[9].ToString();
                txbmurtel.Text = row[10].ToString();
                cmbkrverilmesi.Text = row[11].ToString();
                txbsiranoelave.Text = row[0].ToString();
                txbil.Text = row[15].ToString();
            }
            catch (Exception)
            {


            }
            finally { }
            
        }

        private void BtnSil_Click(object sender, EventArgs e)
        {// and il=" + txbil.Text + "
            con.Open();//
            OracleCommand sorgu = new OracleCommand("delete  from odb.Muracietler where sira=" + txbsiranoelave.Text + " ", con);
            sorgu.ExecuteNonQuery();
            con.Close();
            XtraMessageBox.Show("Məlumat silindi","Qeyd");
            muracietler.Clear();
            listelemuraciet();
            temizlemuraciet();
            muracietsay();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            listelemuraciet();
        }

        private void simpleButton3_Click_1(object sender, EventArgs e)
        {
            frmmusteridanisiq frmdan = new frmmusteridanisiq();
            frmdan.Show();
        }
    }
}
