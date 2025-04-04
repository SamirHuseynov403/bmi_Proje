using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Data.OleDb;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using BMI;
using Excel = Microsoft.Office.Interop.Excel;


namespace Kredit_isler
{
    public partial class Formesas : Form
    {
        
        //OleDbConnection baglanti;
        //OleDbDataAdapter da;
        OleDbCommand komut;
        OracleCommand orkomut;
        OracleCommand komyazdir;
        OleDbCommand komut2;
        OleDbCommand komut3;
        OleDbCommand komutmur;
        OleDbCommand komutmurraz;
        OleDbCommand komutmurikisi;
        OleDbCommand islek;       
        OleDbCommand isleketiraz;
        OleDbCommand islekraz;
        OleDbCommand SMSkomut;
        OleDbCommand saykomut;
        //DataSet ds;
        public string icraci_kod = string.Empty;
        public static string gelmevaxt_adi = "";
        public static string gelmevaxt_tarix = "";
        public static int raz_mebleg;
        public static int teyin_sirasi;
        //DateTime.Today.ToShortDateString();
        TimeSpan fark;

        double farkGun;

        string tarixIl = DateTime.Now.Date.Year.ToString();
        string tarixay = DateTime.Now.Date.Month.ToString();
        string tarixgun = DateTime.Now.Date.Day.ToString();
        
        
        public Formesas()
        {
            InitializeComponent();
        }
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source="\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb"
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Kredit isler\\Kredit isler\\bin\\Debug\\Muracietler.accdb
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.1\kred_sob\Kredit.accdb
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source="C:\Kredit isler\Kredit isler\bin\Debug\Muracietler.accdb"
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\bin\Debug\Muracietler.accdb
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        DataTable cedvel = new DataTable();
        DataTable cedvelraz = new DataTable();
        DataTable teyinler = new DataTable();
        DataTable muracietler = new DataTable();
        DataTable cedvelSMS = new DataTable();
        DataTable cedvelSMSbos = new DataTable();

        DataTable yoxla1 = new DataTable();
        
         void listele()
        {
            try
            {
                //con.Open();
                ////cedvel.Clear();
                //OracleDataAdapter isleme = new OracleDataAdapter("select * from etiraz", con);
                //isleme.Fill(cedvel);
                //dataGridView1.DataSource = cedvel;
                //con.Close();
                baglanti.Open();
                //cedvel.Clear();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from etiraz", baglanti);
                isleme.Fill(cedvel);
                dataGridView1.DataSource = cedvel;

                baglanti.Close();
                dataGridView1.Columns[2].HeaderText = "Sira";
                dataGridView1.Columns[2].Width = 60;
                dataGridView1.Columns[3].HeaderText = "Tarix";
                dataGridView1.Columns[3].Width = 110;
                dataGridView1.Columns[4].HeaderText = "Adı";
                dataGridView1.Columns[4].Width = 320;


                dataGridView1.Columns[5].HeaderText = "Razılıq məbləği";
                dataGridView1.Columns[5].Width = 150;
                dataGridView1.Columns[6].HeaderText = "Təminat";
                dataGridView1.Columns[6].Width = 250;
                dataGridView1.Columns[7].HeaderText = "Etiraz səbəbi";
                dataGridView1.Columns[7].Width = 350;
                dataGridView1.Columns[8].HeaderText = "Qeyd";
                dataGridView1.Columns[8].Width = 350;
                //dataGridView1.Columns[9].HeaderText = "Komitə qərarı";
                //dataGridView1.Columns[9].Width = 250;


                //dataGridView1.Columns[16].HeaderText = "İcraçı";
                //dataGridView1.Columns[16].Width = 250;
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

         public void listeSMS()
         {
            //con.Open();
            ////cedvel.Clear();
            //OracleDataAdapter isleme = new OracleDataAdapter("select sira, Adi_soyadi,Telefon,tarix,sms_tarixi,sms_sayi from Verilmis_kr where Melumat='Gözləmədə' ", baglanti);
            //isleme.Fill(cedvelSMS);
            //dataGridView4.DataSource = cedvelSMS;
            //con.Close();

            //SMSExcel smexc = new SMSExcel();
            baglanti.Open();
             //cedvel.Clear();
             OleDbDataAdapter isleme = new OleDbDataAdapter("select sira, Adi_soyadi,Telefon,tarix,sms_tarixi,sms_sayi from Verilmis_kr where Melumat='Gözləmədə' ", baglanti);
             isleme.Fill(cedvelSMS);
             dataGridView4.DataSource = cedvelSMS;
             baglanti.Close();
             dataGridView4.Columns[0].HeaderText = "Sıra";
             dataGridView4.Columns[0].Width = 50;
             dataGridView4.Columns[1].HeaderText = "Adı";
             dataGridView4.Columns[1].Width = 200;
             dataGridView4.Columns[2].HeaderText = "Telefon";
             dataGridView4.Columns[2].Width = 150;
             dataGridView4.Columns[3].HeaderText = "Verilmə tarixi";
             dataGridView4.Columns[3].Width = 100;
             dataGridView4.Columns[4].HeaderText = "Sms tarixi";
             dataGridView4.Columns[4].Width = 100;
             dataGridView4.Columns[5].HeaderText = "Sms sayı";
             dataGridView4.Columns[5].Width = 100;
             
         }

         void testbugun()
         {
             baglanti.Open();
             OleDbDataAdapter yoxla = new OleDbDataAdapter("select * from muracietler",baglanti);
             yoxla.Fill(yoxla1);
             dataGridView2.DataSource = yoxla1;
             baglanti.Close();
         }


         void listelemuraciet()
         {

             try
             {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                //con.Open();
                //muracietler.Clear();
                //OracleDataAdapter isleme = new OracleDataAdapter("select * from Muracietler where il='" + iltarix + "'", con);
                //isleme.Fill(muracietler);
                //dataGridView2.DataSource = muracietler;

                //con.Close();

                
                 baglanti.Open();
                 muracietler.Clear();
                 OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Muracietler where il='" + iltarix + "'", baglanti);
                 isleme.Fill(muracietler);
                 dataGridView2.DataSource = muracietler;

                 baglanti.Close();
                 dataGridView2.Columns[0].HeaderText = "Sira";
                 dataGridView2.Columns[0].Width = 60;
                 dataGridView2.Columns[1].HeaderText = "Tarix";
                 dataGridView2.Columns[1].Width = 110;
                 dataGridView2.Columns[2].HeaderText = "Adı";
                 dataGridView2.Columns[2].Width = 320;
                 dataGridView2.Columns[3].HeaderText = "Valyuta";
                 dataGridView2.Columns[3].Width = 100;
                 dataGridView2.Columns[4].HeaderText = "Müraciət məbləği";
                 dataGridView2.Columns[4].Width = 150;
                 dataGridView2.Columns[5].HeaderText = "Razılıq məbləği";
                 dataGridView2.Columns[5].Width = 150;
                 dataGridView2.Columns[6].HeaderText = "Təminat";
                 dataGridView2.Columns[6].Width = 250;
                 dataGridView2.Columns[7].HeaderText = "Etiraz səbəbi";
                 dataGridView2.Columns[7].Width = 250;
                 dataGridView2.Columns[8].HeaderText = "Qeyd";
                 dataGridView2.Columns[8].Width = 250;
                 dataGridView2.Columns[9].HeaderText = "Komitə qərarı";
                 dataGridView2.Columns[9].Width = 250;
                 dataGridView2.Columns[10].HeaderText = "Qərar qeyd";
                 dataGridView2.Columns[10].Width = 250;
                 dataGridView2.Columns[11].HeaderText = "Danışıq";
                 dataGridView2.Columns[11].Width = 250;
                 dataGridView2.Columns[12].HeaderText = "Telefon";
                 dataGridView2.Columns[12].Width = 250;
                 dataGridView2.Columns[13].HeaderText = "Məlumat";
                 dataGridView2.Columns[13].Width = 250;
                 dataGridView2.Columns[14].HeaderText = "Gəlmə tarixi";
                 dataGridView2.Columns[14].Width = 120;
                 dataGridView2.Columns[15].HeaderText = "Saat";
                 dataGridView2.Columns[15].Width = 70;
                 dataGridView2.Columns[16].HeaderText = "İcraçı";
                 dataGridView2.Columns[16].Width = 250;
                

            }
             catch (Exception)
             {
                 
                 
             }
            finally { }
                 
            
         }

        void listelemuraciettest()
        {

            try
            {
                //
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select * from odb.Muracietler where il='2022'", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;
                Orcon.Close();

                
                //con.Open();
                //muracietler.Clear();
                //OracleDataAdapter isleme = new OracleDataAdapter("select * from Muracietler where il='" + iltarix + "'", con);
                //isleme.Fill(muracietler);
                //dataGridView2.DataSource = muracietler;

                //con.Close();
                dataGridView2.Columns[0].HeaderText = "Sira";
                dataGridView2.Columns[0].Width = 60;
                dataGridView2.Columns[1].HeaderText = "Tarix";
                dataGridView2.Columns[1].Width = 110;
                dataGridView2.Columns[2].HeaderText = "Adı";
                dataGridView2.Columns[2].Width = 320;
                dataGridView2.Columns[3].HeaderText = "Valyuta";
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[4].HeaderText = "Müraciət məbləği";
                dataGridView2.Columns[4].Width = 150;
                dataGridView2.Columns[5].HeaderText = "Razılıq məbləği";
                dataGridView2.Columns[5].Width = 150;
                dataGridView2.Columns[6].HeaderText = "Təminat";
                dataGridView2.Columns[6].Width = 250;
                dataGridView2.Columns[7].HeaderText = "Etiraz səbəbi";
                dataGridView2.Columns[7].Width = 250;
                dataGridView2.Columns[8].HeaderText = "Qeyd";
                dataGridView2.Columns[8].Width = 250;
                dataGridView2.Columns[9].HeaderText = "Komitə qərarı";
                dataGridView2.Columns[9].Width = 250;
                dataGridView2.Columns[10].HeaderText = "Qərar qeyd";
                dataGridView2.Columns[10].Width = 250;
                dataGridView2.Columns[11].HeaderText = "Danışıq";
                dataGridView2.Columns[11].Width = 250;
                dataGridView2.Columns[12].HeaderText = "Telefon";
                dataGridView2.Columns[12].Width = 250;
                dataGridView2.Columns[13].HeaderText = "Məlumat";
                dataGridView2.Columns[13].Width = 250;
                dataGridView2.Columns[14].HeaderText = "Gəlmə tarixi";
                dataGridView2.Columns[14].Width = 120;
                dataGridView2.Columns[15].HeaderText = "Saat";
                dataGridView2.Columns[15].Width = 70;
                dataGridView2.Columns[16].HeaderText = "İcraçı";
                dataGridView2.Columns[16].Width = 250;

            }
            catch (Exception)
            {


            }
            finally { }


        }

        public void listeleraziliq()
         {
             int tes = 2022;
             try
             {
                 cedvelraz.Clear();

                //con.Open();
                //OracleDataAdapter isleme = new OracleDataAdapter("select * from verilmis_kr where Melumat='Gözləmədə'", con);
                ////OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where Melumat='' ", con);
                //isleme.Fill(cedvelraz);
                //dtg_raz.DataSource = cedvelraz;

                //con.Close();

                baglanti.Open();
                 OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where Melumat='Gözləmədə'", baglanti);
                 //OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where Melumat='' ", baglanti);
                 isleme.Fill(cedvelraz);
                 dtg_raz.DataSource = cedvelraz;

                 baglanti.Close();
                 dtg_raz.Columns[1].HeaderText = "Sira";
                 dtg_raz.Columns[1].Width = 60;
                 dtg_raz.Columns[2].HeaderText = "Tarix";
                 dtg_raz.Columns[2].Width = 110;
                 dtg_raz.Columns[3].HeaderText = "Adı";
                 dtg_raz.Columns[3].Width = 320;
                 //dtg_raz.Columns[3].HeaderText = "Valyuta";
                 //dtg_raz.Columns[3].Width = 100;
                 dtg_raz.Columns[4].HeaderText = "Müraciət məbləği";
                 dtg_raz.Columns[4].Width = 150;
                 dtg_raz.Columns[5].HeaderText = "Komitə qərarı";
                 dtg_raz.Columns[5].Width = 150;
                 dtg_raz.Columns[6].HeaderText = "Qərar qeyd";
                 dtg_raz.Columns[6].Width = 250;
                 dtg_raz.Columns[7].HeaderText = "Danışıq";
                 dtg_raz.Columns[7].Width = 250;
                 dtg_raz.Columns[8].HeaderText = "Telefon";
                 dtg_raz.Columns[8].Width = 250;
                 dtg_raz.Columns[9].HeaderText = "Məlumat";
                 dtg_raz.Columns[9].Width = 250;
                 dtg_raz.Columns[10].HeaderText = "Gəlmə tarixi";
                 dtg_raz.Columns[10].Width = 110;
                 dtg_raz.Columns[11].HeaderText = "Saat";
                 dtg_raz.Columns[11].Width = 110;
                 dtg_raz.Columns[12].HeaderText = "Qeyd";
                 dtg_raz.Columns[12].Width = 250;
                 //dtg_raz.Columns[13].HeaderText = "İcraçı";
                 //dtg_raz.Columns[13].Width = 250;
             
             }
             catch (Exception)
             {
                 
                 throw;
             }
             

         }
         private void listeleteyinler()
        {
            try
            {

                //con.Open();
                //OracleDataAdapter isleme = new OracleDataAdapter("select * from Teyin_olumuslar", con);
                //isleme.Fill(teyinler);
                //dataGridView3.DataSource = teyinler;

                //con.Close();

                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Teyin_olumuslar", baglanti);
                isleme.Fill(teyinler);
                dataGridView3.DataSource = teyinler;

                baglanti.Close();
            }
            catch (Exception)
            {
                
                
            }
             
         }

         private void sqlmeblegtopla()
         {
             //baglanti.Open();
             //OleDbCommand isleme = new OleDbCommand("select where Melumat=verildi sum(Muraciet_mebleg) from Verilmis_kr", baglanti);
             //textBox13.Text = isleme.ExecuteScalar().ToString();
             //baglanti.Close();
         }


        //string i_trh = Convert.ToDateTime(TextBox1.Text).ToString("MM/dd/yyyy");
        public void ChangeKeyboardAZE()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("az-AZE");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage); InputLanguage.CurrentInputLanguage = l;
        }
        private void etirazsay()
        {
            try
            {
                string proid;
                //string query = "select sira from etiraz order by sira Desc";
                //con.Open();
                //OracleCommand cmd = new OracleCommand(query, con);
                //OracleDataReader dr = cmd.ExecuteReader();

                string query = "select sira from etiraz order by sira Desc";
                baglanti.Open();
                OleDbCommand cmd = new OleDbCommand(query, baglanti);
                OleDbDataReader dr = cmd.ExecuteReader();
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
                baglanti.Close();
                txb_sira.Text = proid.ToString();
            }
            catch (Exception)
            {
                
                throw;
            }
            
            
        }
        
        private void raziliqsay()
        {
            try
            {
                int test = 2022;
                string tarixIl = DateTime.Now.Date.Year.ToString();
                //label23.Text = tarixIl;
                int iltarix = Convert.ToInt32(tarixIl);
                string proid;
                //string query = "select sira from Verilmis_kr where il='" + iltarix + "' order by sira Desc";
                //con.Open();
                //OracleCommand cmd = new OracleCommand(query, con);
                //OracleDataReader dr = cmd.ExecuteReader();

                string query = "select sira from Verilmis_kr where il='" + iltarix + "' order by sira Desc";
                baglanti.Open();
                OleDbCommand cmd = new OleDbCommand(query, baglanti);
                OleDbDataReader dr = cmd.ExecuteReader();
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
                baglanti.Close();
                //con.Close();
                texb_raz_sira.Text = proid.ToString();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        

        private void muracietsay()

        {
            try
            {
                int test = 2022;
                string tarixIl = DateTime.Now.Date.Year.ToString();
                //label23.Text = tarixIl;
                int iltarix = Convert.ToInt32(tarixIl);
                string proid;

                //string query = "select sira from Muracietler where il='" + iltarix + "' order by  sira Desc";
                //con.Open();
                //OracleCommand cmd = new OracleCommand(query, con);
                //OracleDataReader dr = cmd.ExecuteReader();

                string query = "select sira from Muracietler where il='" + iltarix + "' order by  sira Desc";
                baglanti.Open();
                OleDbCommand cmd = new OleDbCommand(query, baglanti);
                OleDbDataReader dr = cmd.ExecuteReader();
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
                baglanti.Close();
               // con.Close();
                txbmursira.Text = proid.ToString();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void muracietsayartimsiz()
        {
            try
            {
                string proid;
                string query = "select sira from Muracietler order by sira Desc";

                //con.Open();
                //OracleCommand cmd = new OracleCommand(query, con);
                //OracleDataReader dr = cmd.ExecuteReader();

                baglanti.Open();
                OleDbCommand cmd = new OleDbCommand(query, baglanti);
                OleDbDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString());
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
                baglanti.Close();
                //con.Close();
                txbmursira.Text = proid.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void teyinolunmussay()
        {
            try
            {
                string proid;

                //string query = "select sira from Teyin_olumuslar order by sira Desc";
                //con.Open();
                //OracleCommand cmd = new OracleCommand(query, con);
                //OracleDataReader dr = cmd.ExecuteReader();

                string query = "select sira from Teyin_olumuslar order by sira Desc";
                baglanti.Open();
                OleDbCommand cmd = new OleDbCommand(query, baglanti);
                OleDbDataReader dr = cmd.ExecuteReader();
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
                baglanti.Close();
                con.Close();
                txb_tey_sira.Text = proid.ToString();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void sayı()
        {

            //int satirsayi ;

            //OleDbCommand orcmd = new OleDbCommand("select max(ID) from etiraz", baglanti);
            //baglanti.Open();
            //satirsayi = Convert.ToInt32(orcmd.ExecuteScalar());
            //baglanti.Close();
            //txb_sira.Text=satirsayi.ToString();
            //////string cariil = DateTime.Now.Date.Year.ToString();
            ////int birartir = satirsayi + 1;
            ////string nom = birartir.ToString();
            ////txb_sira.Text = nom;
        }
            private void sayıraziliq()
        {
            try
            {
                int satirsayi = -1;

                //OracleCommand orcmd = new OracleCommand("select count(*) from etiraz", con);
                //con.Open();

                OleDbCommand orcmd = new OleDbCommand("select count(*) from etiraz", baglanti);
                baglanti.Open();
                satirsayi = Convert.ToInt32(orcmd.ExecuteScalar());
                baglanti.Close();
                con.Close();
                //string cariil = DateTime.Now.Date.Year.ToString();
                int birartir = satirsayi;
                string nom = birartir.ToString();
                text_et_saylar.Text = nom;
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

            private void combodoldur()
            { 
            
}

        private void temizle()
            {
                try
                {
                    txb_ad.Text = "";
                    txb_mebleg.Text = "";
                    cmb_sebeb.Text = "";
                    cmb_teminat.Text = "";
                    txb_qeyd.Text = "";
                    textBox5.Text = "";
                    txb_qeyd.Text = "";
                }
                catch (Exception)
                {
                    
                    
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
                textBox14.Text = "";

            }
            catch (Exception)
            {


            }


        }

        private void temizleraziliq()
        {
            try
            {
                txb_raz_ad.Text = "";
                txb_raz_meb.Text = "";
                txb_raz_danisiq.Text = "";
                cmbrazqerar.Text = "";
                txb_raz_danisiq.Text = "";
                txb_raz_qeyd.Text = "";
                txb_raz_axtaris.Text = "";
                txb_raz_tel.Text = "";
                cmbrazqerar.Text = "";
            }
            catch (Exception)
            {
                
                
            }
            

        }


        private void tabPage1_Click(object sender, EventArgs e)
        {
            
        }

        private void testload()
        {
            //listele();
            //listeSMS();
            //sayı();
            //listeleraziliq();
            //listeleteyinler();
            //sayıraziliq();
            //etirazsay();
            //raziliqsay();
            //teyinolunmussay();
            //tarix_goster();
            //listeleteyinler();
            //etiraz_topla();
            //etiraza_at();
            //raziliq_topla();
            //saytopla();
            
            //dtmtarix2.Value = DateTime.Now;
            //dateTimePicker1.Value = DateTime.Now;
            //dateTimePicker2.Value = DateTime.Now;
            //dtpTarix.Value = DateTime.Now;
            //dateTimePicker4.Value = DateTime.Now;
            //Form1 frm1 = new Form1();
            //lblmuricraci.Text = frm1.lblTamad.Text;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (Convert.ToInt16( tarixay)<10)
            {
                textBox15.Text = tarixgun + "-0" + tarixay + "-" + tarixIl;
            }
            else if (Convert.ToInt16( tarixay)>=10)
            {
                textBox15.Text = tarixgun + "-" + tarixay + "-" + tarixIl;
            }
            
            ChangeKeyboardAZE();
            listelemuraciet();
            //listelemuraciettest();
            muracietsay();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                dtpTarix.CustomFormat = "dd-MM-yyyy";
                string sorgu = "Insert into Etiraz(sira,tarix,adi,muraciet_mebleg,teminat,etiraz_sebebi,qeyd) values (@sira,@tarix,@adi,@muraciet_mebleg,@teminat,@etiraz_sebebi,@qeyd)";
                komut = new OleDbCommand(sorgu, baglanti);
                //komut = new OracleCommand(sorgu, con);
                komut.Parameters.AddWithValue("@sira", txb_sira.Text);
                komut.Parameters.AddWithValue("@tarix", dtpTarix);
                komut.Parameters.AddWithValue("@adi", txb_ad.Text);
                komut.Parameters.AddWithValue("@muraciet_mebleg", txb_mebleg.Text);
                komut.Parameters.AddWithValue("@teminat", cmb_teminat.Text);
                komut.Parameters.AddWithValue("@etiraz_sebebi", cmb_sebeb.Text);
                komut.Parameters.AddWithValue("@qeyd", txb_qeyd.Text);
                baglanti.Open();
                con.Close();
                komut.ExecuteNonQuery();
                baglanti.Close();
                //con.Close();
                cedvel.Clear();
                listele();
                sayı();
                etirazsay();
                temizle();
            }
            catch (Exception)
            {
                
               
            }
            


        }

        private void temizle_teyin()
        {
            textBox9.Text = "";
            textBox8.Text = "";
            textBox3.Text = "";
            textBox10.Text = "";
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            
            
        }

        private void S(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txb_sira.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                dtpTarix.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                txb_ad.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                txb_mebleg.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                cmb_teminat.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                cmb_sebeb.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                txb_qeyd.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();

               
            }
            catch (Exception)
            {
                
                
            }
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //con.Open();
                //OracleCommand sorgu = new OracleCommand("delete  from etiraz where teminat='" + dataGridView1.CurrentRow.Cells[4].Value.ToString() + "' and Adi ='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'", con);
                //sorgu.ExecuteNonQuery();
                //con.Close();

                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete  from etiraz where teminat='" + dataGridView1.CurrentRow.Cells[4].Value.ToString() + "' and Adi ='" + dataGridView1.CurrentRow.Cells[2].Value.ToString() + "'", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat silindi");
                cedvel.Clear();
                listele();
                temizle();
                sayı();
            }
            catch (Exception)
            {
                
                throw;
            }
            
            

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //con.Open();
                //OracleDataAdapter axtar = new OracleDataAdapter("select * from Etiraz where  Adi like '%" + textBox5.Text + "%'", con);
                //DataTable tablo2 = new DataTable();
                //axtar.Fill(tablo2);
                //dataGridView1.DataSource = tablo2;
                //con.Close();

                baglanti.Open();
                OleDbDataAdapter axtar = new OleDbDataAdapter("select * from Etiraz where  Adi like '%" + textBox5.Text + "%'", baglanti);
                DataTable tablo2 = new DataTable();
                axtar.Fill(tablo2);
                dataGridView1.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {
                
                throw;
            }
                
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txb_raz_axtaris_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //con.Open();
                //OracleDataAdapter axtar = new OracleDataAdapter("select * from verilmis_kr where Adi_soyadi like '%" + txb_raz_axtaris.Text + "%'", con);
                //DataTable tablo2 = new DataTable();
                //axtar.Fill(tablo2);
                //dtg_raz.DataSource = tablo2;
                //con.Close();

                baglanti.Open();
                OleDbDataAdapter axtar = new OleDbDataAdapter("select * from verilmis_kr where Adi_soyadi like '%" + txb_raz_axtaris.Text + "%'", baglanti);
                DataTable tablo2 = new DataTable();
                axtar.Fill(tablo2);
                dtg_raz.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("tapilmadi");
            } 
        }

        private void button15_Click(object sender, EventArgs e)
        {
            cedvelraz.Clear();
            listeleraziliq();
            temizleraziliq();
            button9.Visible = false;
            btn_razelave.Enabled = true;
            raziliq_topla();
            saytopla();
            listeSMS();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                //dtmtarix2.CustomFormat = "dd-MM-yyyy";
                //string sorgu = "Insert into verilmis_kr(sira,tarix,adi_soyadi,muraciet_mebleg,komite_qerar,qerar_qeyd,danisiq,telefon) values (@sira,@tarix,@adi_soyadi,@muraciet_mebleg,@komite_qerar,@qerar_qeyd,@danisiq,@telefon)";
                //orkomut = new OracleCommand(sorgu, con);

                string sorgu = "Insert into verilmis_kr(sira,tarix,adi_soyadi,muraciet_mebleg,komite_qerar,qerar_qeyd,danisiq,telefon) values (@sira,@tarix,@adi_soyadi,@muraciet_mebleg,@komite_qerar,@qerar_qeyd,@danisiq,@telefon)";
                komut = new OleDbCommand(sorgu, baglanti);
                komut.Parameters.AddWithValue("@sira", texb_raz_sira.Text);
                komut.Parameters.AddWithValue("@tarix", dtpTarix.Value.ToString("dd-MM-yyyy"));
                komut.Parameters.AddWithValue("@adi_soyadi", txb_raz_ad.Text);
                komut.Parameters.AddWithValue("@muraciet_mebleg", txb_raz_meb.Text);
                komut.Parameters.AddWithValue("@komite_qerar", cmbrazqerar.Text);
                komut.Parameters.AddWithValue("@qerar_qeyd", txb_raz_qeyd.Text);
                komut.Parameters.AddWithValue("@danisiq", txb_raz_danisiq.Text);
                komut.Parameters.AddWithValue("@telefon", txb_raz_tel.Text);
                baglanti.Open();
                komut.ExecuteNonQuery();
                //orkomut.ExecuteNonQuery();
                //con.Close();
                cedvelraz.Clear();
                listeleraziliq();
                raziliqsay();
                temizleraziliq();
            }
            catch (Exception)
            {
                
                
            }
            
        }

        private void btn_raz_sil_Click(object sender, EventArgs e)
        {
            try
            {
                //con.Open();
                //OracleCommand sorgu = new OracleCommand("delete  from verilmis_kr where sira=" + dtg_raz.CurrentRow.Cells[1].Value.ToString() + "", con);
                //sorgu.ExecuteNonQuery();
                //con.Close();

                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete  from verilmis_kr where sira=" + dtg_raz.CurrentRow.Cells[1].Value.ToString() + "", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                //con.Close();
                MessageBox.Show("Məlumat silindi");
                cedvelraz.Clear();
                listeleraziliq();
                temizleraziliq();
                sayıraziliq();
                btn_razelave.Enabled = true;
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                string cavab;

                gelmevaxt_adi = dtg_raz.CurrentRow.Cells[3].Value.ToString();
                teyin_sirasi = Convert.ToInt32(dtg_raz.CurrentRow.Cells[1].Value.ToString());
                cavab = dtg_raz.CurrentRow.Cells[4].Value.ToString();
                raz_mebleg = Convert.ToInt32(cavab);
                Gelme_vaxti glmvaxt = new Gelme_vaxti();
                glmvaxt.ShowDialog();
            }
            catch (Exception)
            {
                
                throw;
            }
            
             
            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                dateTimePicker2.CustomFormat = "dd-MM-yyyy";
            //BURDA QALMISIQ TEYIN EDILMISLERI COMBO TRUE OLANDA YENIDEN GUN TEYIN ELESIN
            if (checkBox1.Checked == true)
            {
                string sorgu1 = "update Teyin_olumuslar set Melumat=@melumat,Gelme_tarixi=@tarix,Gelme_saati=@saat where Sira=" + textBox7.Text + "";
                komut2 = new OleDbCommand(sorgu1, baglanti);
                komut2.Parameters.AddWithValue("@melumat", comboBox4.Text);
                komut2.Parameters.AddWithValue("@tarix", dateTimePicker2);
                komut2.Parameters.AddWithValue("@saat", textBox3.Text);
                baglanti.Open();
                komut2.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat yeniləndi");
                listeleraziliq();
                listeleteyinler();
                checkBox1.Checked = false;
                dateTimePicker2.Enabled = false;
                textBox3.Enabled = false;
                textBox10.Enabled = false;
            }

            else
            {
            string sorgu = "update Verilmis_kr set Melumat=@melumat where Sira=" + textBox7.Text + "";
            komut = new OleDbCommand(sorgu,baglanti);
            //komut.Parameters.AddWithValue("@Sira",sayis);
            komut.Parameters.AddWithValue("@melumat", comboBox4.Text);
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Məlumat yeniləndi");
            listeleraziliq();
            //listeleteyinler();
            checkBox1.Checked = false;
            dateTimePicker2.Enabled = false;
            textBox3.Enabled = false;
            textBox10.Enabled = false;
            }
            }
            catch (Exception)
            {
                
                throw;
            }
        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.etirazTableAdapter.FillBy(this.muracietlerDataSet1.Etiraz);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }
        public void tarix_goster()
        {
            try
            {
                //dateTimePicker1.CustomFormat = ("dd-MM-yyyy");
                teyinler.Clear();
                baglanti.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("select * from Teyin_olumuslar where Gelme_tarixi like '" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "'", baglanti);
                adtr.Fill(teyinler);
                dataGridView3.DataSource = teyinler;
                baglanti.Close();
                //listeleteyinler();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            tarix_goster();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            
        }
        private void goster()
        {
            if (checkBox1.Checked==true)
	{
        dateTimePicker2.Enabled = true;
        textBox3.Enabled = true;
        textBox10.Enabled = true;
	}
            else if (checkBox1.Checked == false)
            {

                dateTimePicker2.Enabled = false;
                textBox3.Enabled = false;
                textBox10.Enabled = false;
            }

        }
        private void dateTimePicker1_ValueChanged_1(object sender, EventArgs e)
        {
            tarix_goster();
        }

        private void dtmtarix2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                teyinolunmussay();
                // reqem oldugu ucun tek dirnaglari silmisem textbox7 olan yerde----------------------
                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete from Teyin_olumuslar where sira=" + textBox7.Text + "", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat silindi");
                teyinler.Clear();
                //listeleteyinler();
                tarix_goster();
            }
            catch (Exception)
            {
                
                
            }
            
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void fillByToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {
                //this.teyin_olumuslarTableAdapter1.FillBy(this.kreditDataSet7.Teyin_olumuslar);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            goster();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void text_et_saylar_TextChanged(object sender, EventArgs e)
        {

        }
        private void etiraz_topla()
        {
            try
            {
                int toplam = 0;//www.yazilimkodlama.com
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    toplam += Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
                }
                maskedTextBox1.Text = toplam.ToString();
            }
            catch (Exception)
            {
                
                
            }
            

        }

        private void raziliq_topla()
        {
            //int toplam = 0;//www.yazilimkodlama.com
            //for (int i = 0; i < dtg_raz.Rows.Count; ++i)
            //{
            //    toplam += Convert.ToInt32(dtg_raz.Rows[i].Cells[4].Value);
            //}
            //maskedTextBox3.Text = toplam.ToString();

        }
        
        private void saytopla()
        {

            int kayitsayisi;
            kayitsayisi = dtg_raz.RowCount;
            textBox11.Text = kayitsayisi.ToString();
        }

        void etiraza_at()
        {

            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{

            //    fark = Convert.ToDateTime(dataGridView1.Rows[i].Cells[1].Value.ToString()) - Convert.ToDateTime(DateTime.Now.ToShortDateString());

            //    farkGun = fark.TotalDays;  //www.gorselprogramlama.com

            //    if (farkGun <= 3) dataGridView1.Rows[i].Cells[7].Value=textBox6.Text;

                //else if (farkGun > 3 && farkGun < 7) dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;

                //www.gorselprogramlama.com

           // }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Muracietler.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Muracietler.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "Muracietler";
                    //  worksheet.Cells.Font.Size = 13;

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                        worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                        worksheet.Cells[1, i].Font.Bold = true;
                    }
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                            worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                        }
                    }
                    app.Visible = true;
                }
            }
            }
            catch (Exception)
            {
                
                throw;
            }
        }
            

        private void button8_Click(object sender, EventArgs e)
        {
            
        }

        private void dtg_raz_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txb_raz_ad.Text = dtg_raz.CurrentRow.Cells[3].Value.ToString();
                txb_raz_meb.Text = dtg_raz.CurrentRow.Cells[4].Value.ToString();
                cmbrazqerar.Text = dtg_raz.CurrentRow.Cells[5].Value.ToString();
                txb_raz_qeyd.Text = dtg_raz.CurrentRow.Cells[6].Value.ToString();
                txb_raz_danisiq.Text = dtg_raz.CurrentRow.Cells[7].Value.ToString();
                txb_raz_tel.Text = dtg_raz.CurrentRow.Cells[8].Value.ToString();
                textBox4.Text = dtg_raz.CurrentRow.Cells[1].Value.ToString();

                btn_razelave.Enabled = false;
            }
            catch (Exception)
            {
                
                throw;
            }
            
            
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            button9.Visible = false;
        }

        private void exceleatsms()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = desktopFolder + "\\SMSexc.xls";

            if (dataGridView1.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\SMSexc.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\SMSexc.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "SMSexc.xls";
                    //  worksheet.Cells.Font.Size = 13;

                    //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    //{
                    //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    //    worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                    //    worksheet.Cells[1, i].Font.Bold = true;
                    //}
                    for (int i = 0; i < dataGridView4.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView4.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView4.Rows[i].Cells[j].Value;
                            worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                        }
                    }
                    app.Visible = true;
                    workbook.SaveAs(desktopFolder, Type.Missing);
                    app.Quit();
                }
            }
        }

        private void btn_raz_temiz_Click(object sender, EventArgs e)
        {

            try
            {
                if (dtg_raz.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Muracietler.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Muracietler.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "Muracietler";
                    //  worksheet.Cells.Font.Size = 13;

                    for (int i = 1; i < dtg_raz.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dtg_raz.Columns[i - 1].HeaderText;
                        worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                        worksheet.Cells[1, i].Font.Bold = true;
                    }
                    for (int i = 0; i < dtg_raz.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtg_raz.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dtg_raz.Rows[i].Cells[j].Value;
                            worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                        }
                    }
                    app.Visible = true;
                }
            }
            }
            catch (Exception)
            {
                
                throw;
            }
        }

        private void smssayyaz()
        {
        
        }

        private void smstarixiyaz()
        {
            try
            {
                
                for (int i =0; i < dataGridView4.Rows.Count; i++)
                {
                    int sayplus = 0;
                    textBox4.Text = dataGridView4.Rows[i].Cells[0].Value.ToString();
                    string sayitop =dataGridView4.Rows[i].Cells[5].Value.ToString();
                    if (sayitop == "")
                    {
                        sayitop = "0";
                    }
                    else
                        sayitop.ToString();
                    sayplus = Convert.ToInt32(sayitop.ToString()) + 1;
                    string sorgu = "update Verilmis_kr set Sms_tarixi=@Sms_tarixi,sms_sayi=@smssay where Sira=" + textBox4.Text + "";
                    SMSkomut = new OleDbCommand(sorgu, baglanti);
                    
                    SMSkomut.Parameters.AddWithValue("@Sms_tarixi", dtmtarix2.Value.ToString("dd-MM-yyyy"));
                    SMSkomut.Parameters.AddWithValue("@smssay", sayplus.ToString());
                    baglanti.Open();
                    SMSkomut.ExecuteNonQuery();
                    baglanti.Close();
                    //exceleatsms();
               }
                
                
            }

            catch (Exception)
            {
                
                
            }
            
        }
            //try
            //{
            //    Microsoft.Office.Interop.Excel.Application etiraz = new Microsoft.Office.Interop.Excel.Application();
            //    etiraz.Visible = true;
            //    Microsoft.Office.Interop.Excel.Workbook kitab = etiraz.Workbooks.Add(System.Reflection.Missing.Value);
            //    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitab.Sheets[1];
            //    for (int i = 0; i < dtg_raz.Columns.Count; i++)
            //    {
            //        Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
            //        myrange.Value2 = dtg_raz.Columns[i].HeaderText;
            //    }

            //    for (int i = 0; i < dtg_raz.Columns.Count; i++)
            //    {
            //        for (int j = 0; j < dtg_raz.Rows.Count; j++)
            //        {
            //            Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
            //            myrange.Value2 = dtg_raz[i, j].Value;
            //        }
            //    }
            //}
            //catch (Exception)
            //{
                
                
            //}
            
       // }

        

        private void btn_raz_duzelis_Click(object sender, EventArgs e)
        {
            try
            {
                string sorgu1 = "update Verilmis_kr set danisiq=@danis,melumat=@melumat,Qerar_qeyd=@qeyd,adi_soyadi=@ad,Muraciet_mebleg=@mebleg,Komite_qerar=@qerar,Telefon=@telefon where Sira=" + textBox4.Text + "";
                komut2 = new OleDbCommand(sorgu1, baglanti);
                komut2.Parameters.AddWithValue("@danis", txb_raz_danisiq.Text);
                komut2.Parameters.AddWithValue("@melumat", comboBox1.Text);
                komut2.Parameters.AddWithValue("@Qerar_qeyd", txb_raz_qeyd.Text);
                komut2.Parameters.AddWithValue("@ad", txb_raz_ad.Text);
                komut2.Parameters.AddWithValue("@mebleg", txb_raz_meb.Text);
                komut2.Parameters.AddWithValue("@qerar", cmbrazqerar.Text);
                komut2.Parameters.AddWithValue("@Telefon", txb_raz_tel.Text);
                baglanti.Open();
                komut2.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat yeniləndi");
                
                temizleraziliq();
                //raziliq_topla();
                //saytopla();
                razili_qalan();
                listeleraziliq();
                btn_razelave.Enabled = true;
            }
            catch (Exception)
            {
                
                throw;
            }
            

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            listele();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string sorgu1 = "update Teyin_olumuslar set Melumat=@melumat,Gelme_tarixi=@tarix,Gelme_saati=@saat where Sira=" + textBox7.Text + "";
                komut2 = new OleDbCommand(sorgu1, baglanti);
                komut2.Parameters.AddWithValue("@melumat", comboBox4.Text);
                komut2.Parameters.AddWithValue("@tarix", dateTimePicker2);
                komut2.Parameters.AddWithValue("@saat", textBox3.Text);
                baglanti.Open();
                komut2.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat yeniləndi");
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }
        private void razili_qalan()
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr", baglanti);
                isleme.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;

                baglanti.Close();
                raziliq_topla();
                saytopla();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where melumat is null or melumat=' ' ", baglanti);
                isleme.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;

                baglanti.Close();
                raziliq_topla();
                saytopla();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where melumat ='Verildi' ", baglanti);
                isleme.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;

                baglanti.Close();
                raziliq_topla();
                saytopla();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where melumat ='Müddət bitdi' ", baglanti);
                isleme.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;

                baglanti.Close();
                raziliq_topla();
                saytopla();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr ", baglanti);
                isleme.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;

                baglanti.Close();
                raziliq_topla();
                int kayitsayisi;
                kayitsayisi = dtg_raz.RowCount;
                textBox11.Text = kayitsayisi.ToString();
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                cedvelraz.Clear();
                baglanti.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("select * from Verilmis_kr where tarix <= tr1", baglanti);
                adtr.SelectCommand.Parameters.AddWithValue("tr1", dateTimePicker3.Value.ToString("dd-MM-yyyy"));
                adtr.Fill(cedvelraz);
                dtg_raz.DataSource = cedvelraz;
                baglanti.Close();
                raziliq_topla();
                saytopla();
            }
            catch (Exception)
            {
                
                
            }
            

            //cedvelraz.Clear();
            //baglanti.Open();
            //OleDbDataAdapter isleme = new OleDbDataAdapter("select * from verilmis_kr where tarix < getdate("+dateTimePicker3+") ", baglanti);
            //isleme.Fill(cedvelraz);
            //dtg_raz.DataSource = cedvelraz;

            //baglanti.Close();
            //raziliq_topla();
            //saytopla();
            ////listeleraziliq();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //Mektub mktb = new Mektub();
            //mktb.Show();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            teyinolunmussay();
            //listeleteyinler();
            tarix_goster();
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            //Rating rtg = new Rating();
            //rtg.Show();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                teyinolunmussay();
                string sorgu = "update Teyin_olumuslar set Gelme_tarixi=@Tarix,gelme_Saati=@saat,Melumat=@Melumat where Sira=" + textBox7.Text + "";
                string sorguRAZ = "update Verilmis_kr set Gelme_tarixi=@Tarix,Saat=@saat,Melumat=@Melumat where Sira=" + textBox7.Text + "";
                komut = new OleDbCommand(sorgu, baglanti);
                komut3 = new OleDbCommand(sorguRAZ, baglanti);

                komut3.Parameters.AddWithValue("@Tarix", dateTimePicker2.Value.ToString("dd-MM-yyyy"));
                komut3.Parameters.AddWithValue("@saat", textBox3.Text);
                komut3.Parameters.AddWithValue("@Melumat", comboBox4.Text);


                // komut.Parameters.AddWithValue("@sira", textBox7.Text);
                //komut.Parameters.AddWithValue("@adi", txb_raz_ad.Text);
                // komut.Parameters.AddWithValue("@mebleg", txb_raz_meb.Text);
                komut.Parameters.AddWithValue("@gelme_tarixi", dateTimePicker2.Value.ToString("dd-MM-yyyy"));
                komut.Parameters.AddWithValue("@gelme_saati", textBox3.Text);
                komut.Parameters.AddWithValue("@Melumat", comboBox4.Text);


                baglanti.Open();
                komut.ExecuteNonQuery();
                komut3.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat yeniləndi");
                //listeleteyinler();
                tarix_goster();
            
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void dataGridView3_DoubleClick(object sender, EventArgs e)
        {
            textBox7.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            textBox9.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
            textBox8.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            txb_tey_sira.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            dateTimePicker2.Enabled = true;
            textBox3.Enabled = true;
            comboBox4.Enabled = true;
            button21.Enabled = true;
            button11.Enabled = true;
            button14.Enabled = false;
        }

        private void txb_raz_axtaris_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            txb_raz_axtaris.Text = "";
        }

        private void button22_Click(object sender, EventArgs e)
        {
            
        }

        private void panel21_Paint(object sender, PaintEventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void radioButtonetiraz_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void panel16_Paint(object sender, PaintEventArgs e)
        {

        }

        

        private void cmbmurqerar_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbmurqerar.Text == "Etiraz")
                {
                    txbmurqerqeyd.Enabled = false;
                    cmbmurdan.Enabled = false;
                    cmbmurmel.Enabled = false;
                    cmbmurseb.Enabled = true;
                }
                else
                {
                    txbmurqerqeyd.Enabled = true;
                    cmbmurdan.Enabled = true;
                    cmbmurmel.Enabled = true;
                    cmbmurseb.Enabled = false;
                }
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete  from Muracietler where sira=" + dataGridView2.CurrentRow.Cells[0].Value.ToString() + "", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat silindi");
                muracietler.Clear();
                listelemuraciet();
                temizlemuraciet();
                muracietsay();

                btnmurelave.Enabled = true;
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                string sorgu5 = "Update Muracietler set Adi=@Adi,Valyuta=@Valyuta,Muraciet_mebleg=@Muraciet_mebleg,Raziliq_mebleg=@Raziliq_mebleg,Teminat=@Teminat,Etiraz_sebebi=@Etiraz_sebebi,Qeyd=@Qeyd,Komite_qerar=@Komite_qerar,Qerar_qeyd=@Qerar_qeyd,Danisiq=@Danisiq,Telefon=@Telefon,Melumat=@Melumat where Sira=" + txbmursira.Text + " and Adi='"+txbmurad.Text+"'";
                komutmurraz = new OleDbCommand(sorgu5, baglanti);
                //komutmur.Parameters.AddWithValue("@sira", txbmursira.Text);
                //komutmur.Parameters.AddWithValue("@tarix", dtpmur);
                komutmurraz.Parameters.AddWithValue("@Adi", txbmurad.Text);
                komutmurraz.Parameters.AddWithValue("@Valyuta", cmbmurvaly.Text);
                komutmurraz.Parameters.AddWithValue("@Muraciet_mebleg", txbmurmeb.Text);
                komutmurraz.Parameters.AddWithValue("@Raziliq_mebleg", textBox14.Text);
                komutmurraz.Parameters.AddWithValue("@Teminat", cmbmurtem.Text);
                komutmurraz.Parameters.AddWithValue("@Etiraz_sebebi", cmbmurseb.Text);
                komutmurraz.Parameters.AddWithValue("@Qeyd", txbmurqeyd.Text);
                komutmurraz.Parameters.AddWithValue("@Komite_qerar", cmbmurqerar.Text);
                komutmurraz.Parameters.AddWithValue("@Qerar_qeyd", txbmurqerqeyd.Text);
                komutmurraz.Parameters.AddWithValue("@Danisiq", cmbmurdan.Text);
                komutmurraz.Parameters.AddWithValue("@Telefon", txbmurtel.Text);
                komutmurraz.Parameters.AddWithValue("@Melumat", cmbmurmel.Text);
                baglanti.Open();
                komutmurraz.ExecuteNonQuery();
                baglanti.Close();

                string sorguraz = "Update verilmis_kr set adi_soyadi=@adi_soyadi,muraciet_mebleg=@muraciet_mebleg,komite_qerar=@komite_qerar,qerar_qeyd=@qerar_qeyd,danisiq=@danisiq,telefon=@telefon,melumat=@melumat where Sira=" + txbmursira.Text + "and adi_soyadi='"+txbmurad.Text+"'";
                islekraz = new OleDbCommand(sorguraz, baglanti);
                
                
                islekraz.Parameters.AddWithValue("@adi_soyadi", txbmurad.Text);
                islekraz.Parameters.AddWithValue("@muraciet_mebleg", textBox14.Text);
                islekraz.Parameters.AddWithValue("@komite_qerar", cmbmurqerar.Text);
                islekraz.Parameters.AddWithValue("@qerar_qeyd", txbmurqerqeyd.Text);
                islekraz.Parameters.AddWithValue("@danisiq", cmbmurdan.Text);
                islekraz.Parameters.AddWithValue("@telefon", txbmurtel.Text);
                islekraz.Parameters.AddWithValue("@melumat", "Sonradan etiraz");

                baglanti.Open();
                islekraz.ExecuteNonQuery();
                baglanti.Close();

                //komutmur.Parameters.AddWithValue("@Gelme_tarixi", txb_raz_tel.Text);
                //komutmur.Parameters.AddWithValue("@Saat", txb_raz_tel.Text);
                //komutmur.Parameters.AddWithValue("@İcraci", lblmuricraci.Text);
                
                MessageBox.Show("Məlumat yeniləndi");
                listelemuraciet();
                temizlemuraciet();
                //raziliq_topla();
                //saytopla();
                //razili_qalan();
                btnmurelave.Enabled = true;
                muracietsay();
           
            }
            catch (Exception)
            {
                
                throw;
            }

            
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                txbmursira.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
                //komutmur.Parameters.AddWithValue("@tarix", dtpmur);
                txbmurad.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                cmbmurvaly.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
                txbmurmeb.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
                //komutmur.Parameters.AddWithValue("@Raziliq_mebleg", txb_raz_qeyd.Text);
                cmbmurtem.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
                cmbmurseb.Text = dataGridView2.CurrentRow.Cells[7].Value.ToString();
                txbmurqeyd.Text = dataGridView2.CurrentRow.Cells[8].Value.ToString();
                cmbmurqerar.Text = dataGridView2.CurrentRow.Cells[09].Value.ToString();
                txbmurqerqeyd.Text = dataGridView2.CurrentRow.Cells[10].Value.ToString();
                cmbmurdan.Text = dataGridView2.CurrentRow.Cells[11].Value.ToString();
                txbmurtel.Text = dataGridView2.CurrentRow.Cells[12].Value.ToString();
                cmbmurmel.Text = dataGridView2.CurrentRow.Cells[13].Value.ToString();
                textBox14.Text = dataGridView2.CurrentRow.Cells[5].Value.ToString();
                //komutmur.Parameters.AddWithValue("@Gelme_tarixi", txb_raz_tel.Text);
                //komutmur.Parameters.AddWithValue("@Saat", txb_raz_tel.Text);
                btnmurelave.Enabled = false;
                txbmurqerqeyd.Enabled = true;
                cmbmurdan.Enabled = true;
                cmbmurmel.Enabled = true;
                cmbmurseb.Enabled = true;
                cmbmurseb.Enabled = true;
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void button29_Click_1(object sender, EventArgs e)
        {
           // testbugun();
            muracietsayartimsiz();
            listelemuraciet();
            //muracietler.Clear();
            //listelemuraciet();
            //temizlemuraciet();
            //btnmurelave.Enabled = true;
            //muracietsay();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.RowCount > 0)
                {
                    if (!System.IO.File.Exists(Application.StartupPath + "\\muraciet.xls"))
                    {
                        MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\muraciet.xls");
                        Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                        worksheet.Name = "muraciet";
                        //  worksheet.Cells.Font.Size = 13;

                        //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                        //{
                        //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                        //    worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                        //    worksheet.Cells[1, i].Font.Bold = true;
                        //}
                        for (int i = 0; i < dataGridView2.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                            }
                        }
                        app.Visible = true;
                    }
                }
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }
        public string Xmgeydiyyatnom = string.Empty;
        private void yazdirtest()
        {
            int tmeb =Convert.ToInt16( txbmurmeb.Text);
            int traz= Convert.ToInt16(textBox14.Text);
            int tsira = Convert.ToInt16(txbmursira.Text);
            dateTimePicker4.Text=DateTime.Now.Date.ToString("dd-MM-yyyy");
            string tarixIl = DateTime.Now.Date.Year.ToString();

            con.Open();

            komyazdir = new OracleCommand("insert into odb.Muracietler(Sira,Tarix,adi,Muraciet_mebleg,Raziliq_mebleg,Teminat,Etiraz_sebebi,Qeyd,Komite_qerar,Qerar_qeyd,danisiq,Telefon,Melumat,il) values ( '" + tsira + "', '" + dateTimePicker4.Text + "','" + txbmurad.Text + "','" + tmeb + "','" + traz + "','" + cmbmurtem.Text + "','" + cmbmurseb.Text + "', '" + txbmurqeyd.Text + "', '" + cmbmurqerar.Text + "', '" + txbmurqerqeyd.Text + "','" + cmbmurdan.Text + "','" + txbmurtel.Text + "','" + cmbmurmel.Text + "','"+tarixIl+"')", con);
            komyazdir.ExecuteNonQuery();
            con.Close();
        }
        

        private void btnmurelave_Click(object sender, EventArgs e)
        {
            try
            {
                //DateTime.Today.ToShortDateString();
                //,valyuta,Muraciet_mebleg,Raziliq_mebleg,teminat,Etiraz_sebebi,qeyd,Komite_qerar,danisiq,telefon,melumat,İcraci
                muracietsay();
                //@valyuta,@Muraciet_mebleg,@Raziliq_mebleg,@teminat,@Etiraz_sebebi,@qeyd,@Komite_qerar,@danisiq,@telefon,@melumat,@İcraci
                dateTimePicker4.CustomFormat = "dd-MM-yyyy";
                string sorgu = "Insert into Muracietler(sira,tarix,adi,valyuta,Muraciet_mebleg,Raziliq_mebleg,teminat,Etiraz_sebebi,qeyd,Komite_qerar,danisiq,telefon,melumat) values (@sira,@tarix,@adi,@valyuta,@Muraciet_mebleg,@Raziliq_mebleg,@teminat,@Etiraz_sebebi,@qeyd,@Komite_qerar,@danisiq,@telefon,@melumat)";
                islek = new OleDbCommand(sorgu, baglanti);
                islek.Parameters.AddWithValue("@sira", txbmursira.Text);
                islek.Parameters.AddWithValue("@tarix", textBox15.Text);
                islek.Parameters.AddWithValue("@adi", txbmurad.Text);
                islek.Parameters.AddWithValue("@valyuta", cmbmurvaly.Text);
                islek.Parameters.AddWithValue("@Muraciet_mebleg", txbmurmeb.Text);
                islek.Parameters.AddWithValue("@Raziliq_mebleg", textBox14.Text);
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
                //textBox14.Text = "0";
                //MessageBox.Show("Məlumatlar daxil edildi");
                //string cmbqrtext = cmbmurqerar.SelectedItem.ToString();
                if (cmbmurqerar.SelectedItem == "Etiraz")
                {
                    dtpTarix.CustomFormat = "dd-MM-yyyy";
                    string sorguetiraz = "Insert into Etiraz(sira,tarix,adi,muraciet_mebleg,teminat,etiraz_sebebi,qeyd) values (@sira,@tarix,@adi,@muraciet_mebleg,@teminat,@etiraz_sebebi,@qeyd)";
                    isleketiraz = new OleDbCommand(sorguetiraz, baglanti);
                    isleketiraz.Parameters.AddWithValue("@sira", txbmursira.Text);
                    isleketiraz.Parameters.AddWithValue("@tarix", textBox15.Text);
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
                    dateTimePicker4.CustomFormat = "dd-MM-yyyy";
                    string sorguraz = "Insert into verilmis_kr(sira,tarix,adi_soyadi,muraciet_mebleg,komite_qerar,qerar_qeyd,danisiq,telefon,melumat) values (@sira,@tarix,@adi_soyadi,@muraciet_mebleg,@komite_qerar,@qerar_qeyd,@danisiq,@telefon,@melumat)";
                    islekraz = new OleDbCommand(sorguraz, baglanti);
                    islekraz.Parameters.AddWithValue("@sira", txbmursira.Text);
                    islekraz.Parameters.AddWithValue("@tarix", textBox15.Text);
                    islekraz.Parameters.AddWithValue("@adi_soyadi", txbmurad.Text);
                    islekraz.Parameters.AddWithValue("@muraciet_mebleg", textBox14.Text);
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
                listelemuraciet();
                //muracietsay();
            
            }
            catch (Exception)
            {
                
                throw;
            }
            textBox14.Text = "0";
        }

        private void button24_Click(object sender, EventArgs e)
        {
            //SMSExcel  smsxl = new SMSExcel();
            //smsxl.Show();
            //yazdirtest();
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                string tarix = row.Cells[1].Value.ToString().Substring(0,10);
                komyazdir = new OracleCommand("insert into odb.Muracietler(Sira,Tarix,adi,Muraciet_mebleg,Raziliq_mebleg,Etiraz_sebebi,Qeyd,Komite_qerar,Qerar_qeyd,danisiq,Telefon,Melumat,İcraci,il) values ( '" + row.Cells[0].Value + "', '" + tarix + "','" + row.Cells[2].Value + "','" + row.Cells[4].Value + "','" + row.Cells[5].Value + "', '" + row.Cells[7].Value + "', '" + row.Cells[8].Value + "', '" + row.Cells[9].Value + "','" + row.Cells[10].Value + "','" + row.Cells[11].Value + "','" + row.Cells[12].Value + "','" + row.Cells[13].Value + "','" + row.Cells[16].Value + "','" + row.Cells[17].Value + "')", con);
                komyazdir.ExecuteNonQuery();
            }
            con.Close();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        private void txbmurad_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    baglanti.Open();
            //    OleDbDataAdapter axtar = new OleDbDataAdapter("select * from Muracietler where  Adi like '%" + txbmurad.Text + "%'", baglanti);
            //    DataTable tablo2 = new DataTable();
            //    axtar.Fill(tablo2);
            //    dataGridView2.DataSource = tablo2;
            //    baglanti.Close();
            //}
            //catch (Exception)
            //{

            //    MessageBox.Show("tapılmadı");
            //}
            
        }

        private void txb_ad_TextChanged(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                OleDbDataAdapter axtar = new OleDbDataAdapter("select * from Etiraz where  Adi like '%" + txb_ad.Text + "%'", baglanti);
                DataTable tablo2 = new DataTable();
                axtar.Fill(tablo2);
                dataGridView1.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("tapılmadı");
            }
            
        }

        private void txb_raz_ad_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button31_Click(object sender, EventArgs e)
        {
            cedvelraz.Clear();
            baglanti.Open();
            OleDbDataAdapter isleme = new OleDbDataAdapter("select * from muracietler where Komite_qerar <>'Etiraz' ", baglanti);
            isleme.Fill(cedvelraz);
            dtg_raz.DataSource = cedvelraz;

            baglanti.Close();
            raziliq_topla();
            saytopla();
        }

        private void panelust_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.RowCount > 0)
                {
                    if (!System.IO.File.Exists(Application.StartupPath + "\\Muracietler.xls"))
                    {
                        MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Muracietler.xls");
                        Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                        worksheet.Name = "Muracietler";
                        //  worksheet.Cells.Font.Size = 13;

                        for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                        {
                            worksheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                            worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                            worksheet.Cells[1, i].Font.Bold = true;
                        }
                        for (int i = 0; i < dataGridView2.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView2.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                                worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                            }
                        }
                        app.Visible = true;
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        private void panel14_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button33_Click(object sender, EventArgs e)
        {
            
            smstarixiyaz();
            exceleatsms();
            
        }

        //string sorgu1 = "update Verilmis_kr set danisiq=@danis,melumat=@melumat,Qerar_qeyd=@qeyd,adi_soyadi=@ad,Muraciet_mebleg=@mebleg,Komite_qerar=@qerar,Telefon=@telefon where Sira=" + textBox4.Text + "";
        //        komut2 = new OleDbCommand(sorgu1, baglanti);
        //        komut2.Parameters.AddWithValue("@danis", txb_raz_danisiq.Text);
        //        komut2.Parameters.AddWithValue("@melumat", comboBox1.Text);
        //        komut2.Parameters.AddWithValue("@Qerar_qeyd", txb_raz_qeyd.Text);
        //        komut2.Parameters.AddWithValue("@ad", txb_raz_ad.Text);
        //        komut2.Parameters.AddWithValue("@mebleg", txb_raz_meb.Text);
        //        komut2.Parameters.AddWithValue("@qerar", cmbrazqerar.Text);
        //        komut2.Parameters.AddWithValue("@Telefon", txb_raz_tel.Text);
        //        baglanti.Open();
        //        komut2.ExecuteNonQuery();
        //        baglanti.Close();
        //        MessageBox.Show("Məlumat yeniləndi");

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label48_Click(object sender, EventArgs e)
        {

        }

        private void button35_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            string sql = "SELECT sira, Adi_soyadi,Telefon,tarix,sms_tarixi,sms_sayi FROM Verilmis_kr Where Melumat='Gözləmədə' and tariX BETWEEN @tar1 and @tar2";
            DataTable dt = new DataTable();
            OleDbDataAdapter adp = new OleDbDataAdapter(sql, baglanti);
            adp.SelectCommand.Parameters.AddWithValue("@tar1", dateTimePicker3.Value);
            adp.SelectCommand.Parameters.AddWithValue("@tar2", dateTimePicker5.Value);
            adp.Fill(dt);
            baglanti.Close();
            dataGridView4.DataSource = dt;
        }

        private void button34_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            string sql = "SELECT sira, Adi_soyadi,Telefon,tarix,sms_tarixi,sms_sayi FROM Verilmis_kr Where Melumat='Gözləmədə' and Sms_tarixi BETWEEN @tar1 and @tar2";
            DataTable dt = new DataTable();
            OleDbDataAdapter adp = new OleDbDataAdapter(sql, baglanti);
            adp.SelectCommand.Parameters.AddWithValue("@tar1", dateTimePicker3.Value);
            adp.SelectCommand.Parameters.AddWithValue("@tar2", dateTimePicker5.Value);
            adp.Fill(dt);
            baglanti.Close();
            dataGridView4.DataSource = dt;
        }

        private void button36_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            //cedvel.Clear();
            OleDbDataAdapter isleme = new OleDbDataAdapter("select sira, Adi_soyadi,Telefon,tarix,sms_tarixi,sms_sayi from Verilmis_kr where Melumat='Gözləmədə' and sms_sayi is null or Sms_sayi=' ' ", baglanti);
            isleme.Fill(cedvelSMSbos);
            dataGridView4.DataSource = cedvelSMSbos;
            baglanti.Close();
            dataGridView4.Columns[0].HeaderText = "Sıra";
            dataGridView4.Columns[0].Width = 50;
            dataGridView4.Columns[1].HeaderText = "Adı";
            dataGridView4.Columns[1].Width = 200;
            dataGridView4.Columns[2].HeaderText = "Telefon";
            dataGridView4.Columns[2].Width = 150;
            dataGridView4.Columns[3].HeaderText = "Verilmə tarixi";
            dataGridView4.Columns[3].Width = 100;
            dataGridView4.Columns[4].HeaderText = "Sms tarixi";
            dataGridView4.Columns[4].Width = 100;
            dataGridView4.Columns[5].HeaderText = "Sms sayı";
            dataGridView4.Columns[5].Width = 100;
        }

        private void dtg_raz_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button37_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            string sql = "select * from Muracietler where tariX BETWEEN @tar1 and @tar2";
            DataTable dt = new DataTable();
            OleDbDataAdapter adp = new OleDbDataAdapter(sql, baglanti);
            adp.SelectCommand.Parameters.AddWithValue("@tar1", dateTimePicker7.Value);
            adp.SelectCommand.Parameters.AddWithValue("@tar2", dateTimePicker6.Value);
            adp.Fill(dt);
            baglanti.Close();
            dataGridView2.DataSource = dt;
        }

        private void button38_Click(object sender, EventArgs e)
        {
            try
            {

                baglanti.Open();
                OleDbDataAdapter axtar = new OleDbDataAdapter("select * from Muracietler where adi like '%" + textBox13.Text + "%' or sira like '" + textBox13.Text + "'", baglanti);
                DataTable tablo2 = new DataTable();
                axtar.Fill(tablo2);
                dataGridView2.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {
                //MessageBox.Show("tapılmadı");

            }
            finally { }
        }
    }
    
}
