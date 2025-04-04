using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class frmmuracietyazdir : Form
    {
        public frmmuracietyazdir()
        {
            InitializeComponent();
        }
        public frmmuracietler frmmur;
        string tarixIl = DateTime.Now.Date.Year.ToString();
        string tarixay = DateTime.Now.Date.Month.ToString();
        string tarixgun = DateTime.Now.Date.Day.ToString();
        OleDbCommand islek;
        OleDbCommand isleketiraz;
        OleDbCommand islekraz;
        DataTable muracietler = new DataTable();
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");

        private void silmek()
        {
            try
            {
                frmmuracietler frm1 = new frmmuracietler();
                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete  from Muracietler where sira=" + frm1.gridmuracietler.Columns[0].ToString() + "", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat silindi");
                muracietler.Clear();
                listelemuraciet();
                temizlemuraciet();
                //muracietsay();

                // btnmurelave.Enabled = true;
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
                cmbmurmel.Text = "";
                cmbmurqerar.Text = "";
                cmbmurseb.Text = "";
                cmbmurtem.Text = "";
                txbmurqerqeyd.Text = "";
                txbmurqeyd.Text = "";
                txbmurtel.Text = "";
                txtraziliq.Text = "";

            }
            catch (Exception)
            {


            }


        }
        void listelemuraciet()
        {

            try
            {

                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                frmmuracietler frm1 = new frmmuracietler();
                baglanti.Open();
                muracietler.Clear();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Muracietler  where il='" + iltarix + "'order by sira desc", baglanti);
                isleme.Fill(muracietler);
                frm1.gridControl1.DataSource = muracietler;

                baglanti.Close();
                frm1.gridmuracietler.Columns[0].Caption = "Sira";
                frm1.gridmuracietler.Columns[0].Width = 60;
                frm1.gridmuracietler.Columns[1].Caption = "Tarix";
                frm1.gridmuracietler.Columns[1].Width = 110;
                frm1.gridmuracietler.Columns[2].Caption = "Adı";
                frm1.gridmuracietler.Columns[2].Width = 320;
                frm1.gridmuracietler.Columns[3].Caption = "Valyuta";
                frm1.gridmuracietler.Columns[3].Width = 100;
                frm1.gridmuracietler.Columns[4].Caption = "Müraciət məbləği";
                frm1.gridmuracietler.Columns[4].Width = 150;
                frm1.gridmuracietler.Columns[5].Caption = "Razılıq məbləği";
                frm1.gridmuracietler.Columns[5].Width = 150;
                frm1.gridmuracietler.Columns[6].Caption = "Təminat";
                frm1.gridmuracietler.Columns[6].Width = 250;
                frm1.gridmuracietler.Columns[7].Caption = "Etiraz səbəbi";
                frm1.gridmuracietler.Columns[7].Width = 250;
                frm1.gridmuracietler.Columns[8].Caption = "Qeyd";
                frm1.gridmuracietler.Columns[8].Width = 250;
                frm1.gridmuracietler.Columns[9].Caption = "Komitə qərarı";
                frm1.gridmuracietler.Columns[9].Width = 250;
                frm1.gridmuracietler.Columns[10].Caption = "Qərar qeyd";
                frm1.gridmuracietler.Columns[10].Width = 250;
                frm1.gridmuracietler.Columns[11].Caption = "Danışıq";
                frm1.gridmuracietler.Columns[11].Width = 250;
                frm1.gridmuracietler.Columns[12].Caption = "Telefon";
                frm1.gridmuracietler.Columns[12].Width = 250;
                frm1.gridmuracietler.Columns[13].Caption = "Məlumat";
                frm1.gridmuracietler.Columns[13].Width = 250;
                frm1.gridmuracietler.Columns[14].Caption = "Gəlmə tarixi";
                frm1.gridmuracietler.Columns[14].Width = 120;
                frm1.gridmuracietler.Columns[15].Caption = "Saat";
                frm1.gridmuracietler.Columns[15].Width = 70;
                frm1.gridmuracietler.Columns[16].Caption = "İcraçı";
                frm1.gridmuracietler.Columns[16].Width = 250;

            }
            catch (Exception)
            {

                throw;
            }


        }

        private void frmmuracietyazdir_Load(object sender, EventArgs e)
        {
            if (Convert.ToInt16(tarixay) < 10)
            {
                txttarix.Text = tarixgun + "-0" + tarixay + "-" + tarixIl;
            }
            else if (Convert.ToInt16(tarixay) >= 10)
            {
                txttarix.Text = tarixgun + "-" + tarixay + "-" + tarixIl;
            }
            //muracietsay();
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            try
            {
                //DateTime.Today.ToShortDateString();
                //,valyuta,Muraciet_mebleg,Raziliq_mebleg,teminat,Etiraz_sebebi,qeyd,Komite_qerar,danisiq,telefon,melumat,İcraci
               // muracietsay();
                //@valyuta,@Muraciet_mebleg,@Raziliq_mebleg,@teminat,@Etiraz_sebebi,@qeyd,@Komite_qerar,@danisiq,@telefon,@melumat,@İcraci
                //dateTimePicker4.CustomFormat = "dd-MM-yyyy";
                string sorgu = "Insert into Muracietler(sira,tarix,adi,valyuta,Muraciet_mebleg,Raziliq_mebleg,teminat,Etiraz_sebebi,qeyd,Komite_qerar,danisiq,telefon,melumat) values (@sira,@tarix,@adi,@valyuta,@Muraciet_mebleg,@Raziliq_mebleg,@teminat,@Etiraz_sebebi,@qeyd,@Komite_qerar,@danisiq,@telefon,@melumat)";
                islek = new OleDbCommand(sorgu, baglanti);
                islek.Parameters.AddWithValue("@sira", txbmursira.Text);
                islek.Parameters.AddWithValue("@tarix", txttarix.Text);
                islek.Parameters.AddWithValue("@adi", txbmurad.Text);
                islek.Parameters.AddWithValue("@valyuta", cmbmurvaly.Text);
                islek.Parameters.AddWithValue("@Muraciet_mebleg", txbmurmeb.Text);
                islek.Parameters.AddWithValue("@Raziliq_mebleg", txtraziliq.Text);
                islek.Parameters.AddWithValue("@teminat", cmbmurtem.Text);
                islek.Parameters.AddWithValue("@Etiraz_sebebi", cmbmurseb.Text);
                islek.Parameters.AddWithValue("@qeyd", txbmurqeyd.Text);
                islek.Parameters.AddWithValue("@Komite_qerar", cmbmurqerar.Text);
                islek.Parameters.AddWithValue("@Qerar_qeyd", txbmurqerqeyd.Text);
                islek.Parameters.AddWithValue("@danisiq", cmbmurdan.Text);
                islek.Parameters.AddWithValue("@telefon", txbmurtel.Text);
                islek.Parameters.AddWithValue("@melumat", cmbmurmel.Text);
                //islek.Parameters.AddWithValue("@icraci", lblmuricraci.Text);


                baglanti.Open();
                islek.ExecuteNonQuery();
                baglanti.Close();
                //frmmur.btnyenile.PerformClick();
                //textBox14.Text = "0";
                //MessageBox.Show("Məlumatlar daxil edildi");
                //string cmbqrtext = cmbmurqerar.SelectedItem.ToString();
                if (cmbmurqerar.SelectedItem == "Etiraz")
                {
                    //dtpTarix.CustomFormat = "dd-MM-yyyy";
                    string sorguetiraz = "Insert into Etiraz(sira,tarix,adi,muraciet_mebleg,teminat,etiraz_sebebi,qeyd) values (@sira,@tarix,@adi,@muraciet_mebleg,@teminat,@etiraz_sebebi,@qeyd)";
                    isleketiraz = new OleDbCommand(sorguetiraz, baglanti);
                    isleketiraz.Parameters.AddWithValue("@sira", txbmursira.Text);
                    isleketiraz.Parameters.AddWithValue("@tarix", txtraziliq.Text);
                    isleketiraz.Parameters.AddWithValue("@adi", txbmurad.Text);
                    isleketiraz.Parameters.AddWithValue("@muraciet_mebleg", txbmurmeb.Text);
                    isleketiraz.Parameters.AddWithValue("@teminat", cmbmurtem.Text);
                    isleketiraz.Parameters.AddWithValue("@etiraz_sebebi", cmbmurseb.Text);
                    isleketiraz.Parameters.AddWithValue("@qeyd", txbmurqeyd.Text);
                    baglanti.Open();
                    isleketiraz.ExecuteNonQuery();
                    baglanti.Close();

                }
                else
                {
                    //dateTimePicker4.CustomFormat = "dd-MM-yyyy";
                    string sorguraz = "Insert into verilmis_kr(sira,tarix,adi_soyadi,muraciet_mebleg,komite_qerar,qerar_qeyd,danisiq,telefon,melumat) values (@sira,@tarix,@adi_soyadi,@muraciet_mebleg,@komite_qerar,@qerar_qeyd,@danisiq,@telefon,@melumat)";
                    islekraz = new OleDbCommand(sorguraz, baglanti);
                    islekraz.Parameters.AddWithValue("@sira", txbmursira.Text);
                    islekraz.Parameters.AddWithValue("@tarix", txttarix.Text);
                    islekraz.Parameters.AddWithValue("@adi_soyadi", txbmurad.Text);
                    islekraz.Parameters.AddWithValue("@muraciet_mebleg", txtraziliq.Text);
                    islekraz.Parameters.AddWithValue("@komite_qerar", cmbmurqerar.Text);
                    islekraz.Parameters.AddWithValue("@qerar_qeyd", txbmurqeyd.Text);
                    islekraz.Parameters.AddWithValue("@danisiq", cmbmurdan.Text);
                    islekraz.Parameters.AddWithValue("@telefon", txbmurtel.Text);
                    islekraz.Parameters.AddWithValue("@melumat", "Gözləmədə");

                    baglanti.Open();
                    islekraz.ExecuteNonQuery();
                    baglanti.Close();

                }
                muracietler.Clear();
                temizlemuraciet();
                
                //muracietsay();

                
            }
            catch (Exception)
            {

                throw;
            }
            txtraziliq.Text = "0";
            
            
            
        }

        private void BtnSil_Click(object sender, EventArgs e)
        {
            silmek();
            listelemuraciet();
        }
    }
}
