using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class frmmenzil : Form
    {
        public frmmenzil()
        {
            InitializeComponent();
        }
        public Anakredit anakkr;
        public Form1 form1_kod;
        public OracleConnection Orcon;
        public Mektub mktb;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;

        public string muqtipi { get; set; }
        public string adi { get; set; }
        public string ks { get; set; }
        public string hesab { get; set; }
        public string mebleg { get; set; }
        public string faiz { get; set; }
        public string vkfaiz { get; set; }
        public string ehtfaiz { get; set; }
        public string subhes { get; set; }
        public string muddet { get; set; }
        public string seriyano { get; set; }
        public string ver_orqan { get; set; }
        public string ver_tar { get; set; }
        public string mobil { get; set; }
        public string unvan { get; set; }
        public string ayliq { get; set; }
        public string fifd { get; set; }
        public string teyinat { get; set; }
        public string olke { get; set; }
        public string sudahes { get; set; }
        public string faizhes { get; set; }
        public string vkhes { get; set; }
        public string vkfaizhes { get; set; }
        public string odgunu { get; set; }
        public string mebyazi { get; set; }
        public string valyuta { get; set; }
        public string carihes { get; set; }
        public string tamhesab { get; set; }
        public string subkod_qeyd { get; set; }
        public string girov_deyeri { get; set; }
        public string test { get; set; }
        public string mektub_No { get; set; }

        public string icraciadizaminlikde { get; set; }
        public string icraci_adi { get; set; }

        public string zam1adi { get; set; }
        public string zam1pass { get; set; }
        public string zam1tel { get; set; }
        public string zam1unvan { get; set; }
        public string zam1Pastar { get; set; }
        public string zam1Pasorqan { get; set; }
        public string zam1olkesi { get; set; }

        public string zam2adi { get; set; }
        public string zam2pass { get; set; }
        public string zam2tel { get; set; }
        public string zam2unvan { get; set; }
        public string zam2Pastar { get; set; }
        public string zam2Pasorqan { get; set; }
        public string zam2olkesi { get; set; }

        public string zam3adi { get; set; }
        public string zam3pass { get; set; }
        public string zam3tel { get; set; }
        public string zam3unvan { get; set; }
        public string zam3Pastar { get; set; }
        public string zam3Pasorqan { get; set; }
        public string zam3olkesi { get; set; }

        public string zaminmuqn1 { get; set; }
        public string zaminmuqn2 { get; set; }
        public string zaminmuqn3 { get; set; }

        public string tamtarix { get; set; }

        public string comboadi { get; set; }
        public string teyinatadi { get; set; }

        
        public string mek_no { get; set; }
        public string umumitar { get; set; }

        //cixaris melumatlari
        public string C_seriyaNo { get; set; }
        public string C_obadi { get; set; }
        public string C_girovunvan { get; set; }
        public string C_umsahe { get; set; }
        public string C_yasayis { get; set; }
        public string C_yardimci { get; set; }
        public string C_reyesno { get; set; }
        public string C_qeydno { get; set; }
        public string C_tarix { get; set; }
        public string C_otaqsay { get; set; }
        public string C_kitabno { get; set; }
        public string C_vereqno { get; set; }
        public string C_sahibi { get; set; }
        public string C_zamNo { get; set; }
        public string C_pasport { get; set; }
        public string C_pastarix { get; set; }
        public string C_pasorqan { get; set; }
        public string C_tel { get; set; }
        public string C_olke { get; set; }
        public string C_qeydunvan { get; set; }
        public string C_zamhe_yox { get; set; }
        public string C_gorovferqlieyni { get; set; }
        public string C_ipotekaNo { get; set; }
        public string C_ipotekanovu { get; set; }

        public string C_seriyaNo2 { get; set; }
        public string C_obadi2 { get; set; }
        public string C_girovunvan2 { get; set; }
        public string C_umsahe2 { get; set; }
        public string C_yasayis2 { get; set; }
        public string C_yardimci2 { get; set; }
        public string C_reyesno2 { get; set; }
        public string C_qeydno2 { get; set; }
        public string C_tarix2 { get; set; }
        public string C_otaqsay2 { get; set; }
        public string C_kitabno2 { get; set; }
        public string C_vereqno2 { get; set; }
        public string C_sahibi2 { get; set; }
        public string C_zamNo2 { get; set; }
        public string C_pasport2 { get; set; }
        public string C_pastarix2 { get; set; }
        public string C_pasorqan2 { get; set; }
        public string C_tel2 { get; set; }
        public string C_olke2 { get; set; }
        public string C_qeydunvan2 { get; set; }
        public string C_zamhe_yox2 { get; set; }
        public string C_ipotekaNo2 { get; set; }
        public string C_ipotekanovu2 { get; set; }
        public string C_gorovferqlieyni2 { get; set; }

        frmbircixaris frm1 = new frmbircixaris();
        frmikicixaris frm2 = new frmikicixaris();
        string VAL_AD;
        private void valyuta_adi()
        {
            if (valyuta == "00")
            {

                cmbval.Text = "AZN";
                VAL_AD = "AZN";
            }
            else if (valyuta == "01")
            {

                cmbval.Text = "USD";
                VAL_AD = "USD";
            }
            else if (valyuta == "02")
            {

                cmbval.Text = "AVRO";
                VAL_AD = "AVRO";
            }

        }
        int kr_say;
        int ser_say;
        public int zam_say;
        private void olkeadi()
        {
            if (olke == "AZ")
            {
                cmbolke.Text = "Azərbaycan Respublikası";
            }
            else if (olke == "IRN")
            {
                cmbolke.Text = " İran İslam Respublikası";
            }


        }
        public void muqavile_nom()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            string satirsayi = "2021";


            ////string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand komut = new OracleCommand();
            komut.Connection = con;
            komut.CommandText = "select kr_avtomobil,kr_serencam,kr_zaminler,il from odb.muqavile_nomreleri where il='" + tarixIl + "'";
            OracleDataReader dr = komut.ExecuteReader();


            while (dr.Read())
            {

                zam_say = Convert.ToInt32(dr["kr_zaminler"].ToString()) + 1;



                kr_say = Convert.ToInt32(dr["kr_avtomobil"].ToString()) + 1;
                txtmuqno.Text = kr_say.ToString();
                ser_say = Convert.ToInt32(dr["kr_serencam"].ToString()) + 1;
                //txbkavtoSerencam.Text = ser_say.ToString();


            }


            con.Close();


        }
        private void teyadi()
        {
            if (teyinat == "2001")
            {
                cmbtey.Text = "mənzil təmiri";
            }
            else if (teyinat == "2002")
            {
                cmbtey.Text = " Avtomobil alınması";
            }
            else if (teyinat == "2003")
            {
                cmbtey.Text = " məişət əşyalarının alınması";
            }

        }
        private string yaziyaCevir(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için            
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";

            string[] birler = { "", " bir ", " iki ", " üç ", " dörd ", " beş ", " altı ", " yeddi ", " səkkiz ", " doqquz " };
            string[] onlar = { "", " on ", " iyirmi ", " otuz ", " qırx ", " əlli ", " altmış ", " yetmiş ", " səksən ", " doxsan " };
            string[] binler = { "katrilyon", "trilyon", " milyard ", " milyon ", " min ", "" }; //KATRİLYON'un önüne ekleme yapılarak artırabilir.

            int grupSayisi = 6; //sayıdaki 3'lü grup sayısı. katrilyon içi 6. (1.234,00 daki grup sayısı 2'dir.)
            //KATRİLYON'un başına ekleyeceğiniz her değer için grup sayısını artırınız.

            lira = lira.PadLeft(grupSayisi * 3, '0'); //sayının soluna '0' eklenerek sayı 'grup sayısı x 3' basakmaklı yapılıyor.            

            string grupDegeri;

            for (int i = 0; i < grupSayisi * 3; i += 3) //sayı 3'erli gruplar halinde ele alınıyor.
            {
                grupDegeri = "";

                if (lira.Substring(i, 1) != "0")
                    grupDegeri += birler[Convert.ToInt32(lira.Substring(i, 1))] + " yüz "; //yüzler                

                if (grupDegeri == " biryüz ") //biryüz düzeltiliyor.
                    grupDegeri = " yüz ";

                grupDegeri += onlar[Convert.ToInt32(lira.Substring(i + 1, 1))]; //onlar

                grupDegeri += birler[Convert.ToInt32(lira.Substring(i + 2, 1))]; //birler                

                if (grupDegeri != "") //binler
                    grupDegeri += binler[i / 3];

                if (grupDegeri == " birmin ") //birbin düzeltiliyor.
                    grupDegeri = " min ";

                yazi += grupDegeri;
            }

            if (yazi != "")
                yazi += " manat ";

            int yaziUzunlugu = yazi.Length;

            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];

            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];

            if (yazi.Length > yaziUzunlugu)
                yazi += " qəpik.";
            else
                yazi += "sıfır qəpik.";

            return yazi;

        }

        private string yaziyaCevirqepiksiz(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için            
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";

            string[] birler = { "", " bir ", " iki ", " üç ", " dörd ", " beş ", " altı ", " yeddi ", " səkkiz ", " doqquz " };
            string[] onlar = { "", " on ", " iyirmi ", " otuz ", " qırx ", " əlli ", " altmış ", " yetmiş ", " səksən ", " doxsan " };
            string[] binler = { "katrilyon", "trilyon", " milyard ", " milyon ", " min ", "" }; //KATRİLYON'un önüne ekleme yapılarak artırabilir.

            int grupSayisi = 6; //sayıdaki 3'lü grup sayısı. katrilyon içi 6. (1.234,00 daki grup sayısı 2'dir.)
            //KATRİLYON'un başına ekleyeceğiniz her değer için grup sayısını artırınız.

            lira = lira.PadLeft(grupSayisi * 3, '0'); //sayının soluna '0' eklenerek sayı 'grup sayısı x 3' basakmaklı yapılıyor.            

            string grupDegeri;

            for (int i = 0; i < grupSayisi * 3; i += 3) //sayı 3'erli gruplar halinde ele alınıyor.
            {
                grupDegeri = "";

                if (lira.Substring(i, 1) != "0")
                    grupDegeri += birler[Convert.ToInt32(lira.Substring(i, 1))] + " yüz "; //yüzler                

                if (grupDegeri == " biryüz ") //biryüz düzeltiliyor.
                    grupDegeri = " yüz ";

                grupDegeri += onlar[Convert.ToInt32(lira.Substring(i + 1, 1))]; //onlar

                grupDegeri += birler[Convert.ToInt32(lira.Substring(i + 2, 1))]; //birler                

                if (grupDegeri != "") //binler
                    grupDegeri += binler[i / 3];

                if (grupDegeri == " birmin ") //birbin düzeltiliyor.
                    grupDegeri = " min ";

                yazi += grupDegeri;
            }

            if (yazi != "")
                yazi += "";

            int yaziUzunlugu = yazi.Length;

            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];

            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];

            if (yazi.Length > yaziUzunlugu)
                yazi += "";
            else
                yazi += "";

            return yazi;

        }
        string c_lar;
        private void word_at()
        {
            string meknoal;
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                string proid;

                int mek_no_arti = 0;
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                //string query = "select qey_nom ,il from odb.xaric_mektub where il='2022' order by qey_nom desc";
                string query = "select max(-to_number(substr(t.qey_nom,5,5)))mn from odb.xaric_mektub t where t.il='" + tarixIl + "'";
                OracleCommand cmd = new OracleCommand(query, Orcon);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString());
                    proid = id.ToString();
                    mek_no_arti = id + 1;

                    //hevsecal = proid;
                }
                else if (Convert.IsDBNull(dr))
                {
                    proid = ("1");
                }
                else
                {
                    proid = ("1");
                }
                Orcon.Close();
                //txbkavtoModel.Text =tarixIl+"-"+ mek_no_arti.ToString();
                mek_no = tarixIl + "-" + mek_no_arti.ToString();
                meknoal = tarixIl + "-" + mek_no_arti.ToString();
            }
            catch (Exception)
            {

                throw;
            }



            if (!System.IO.File.Exists(Application.StartupPath + "\\KreditAvtoMuq.docx"))
            {
                MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            Zaminler zmnlar = new Zaminler();
            string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txtmeb.Text));
            string ayyazile = yaziyaCevir(Convert.ToDecimal(txtayliq.Text));
            string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtmud.Text));
            string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtfaiz.Text));
            string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtvkfaiz.Text));
            string girovyaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txtgirovdey.Text));
            
            string ilsonreqem = DateTime.Now.Date.Year.ToString();
            if (ilsonreqem.Substring(4,1)=="1"|| ilsonreqem.Substring(4, 1) == "2" || ilsonreqem.Substring(4, 1) == "5" || ilsonreqem.Substring(4, 1) == "8" || ilsonreqem.Substring(4, 1) == "7")
            {
                c_lar = "ci";
            }
            else if (ilsonreqem.Substring(4, 1) == "3" || ilsonreqem.Substring(4, 1) == "4")
            {
                c_lar = "cü";
            }
            else if (ilsonreqem.Substring(4, 1) == "6" )
            {
                c_lar = "cı";
            }
            else if (ilsonreqem.Substring(4, 1) == "9" || ilsonreqem.Substring(4, 1) == "10")
            {
                c_lar = "cu";
            }

            var ilsonrasi = c_lar;
            var adi = txtborcalan.Text;
            var passport = txtaze.Text;
            var pasvertarix = datepasvertarix.Text;
            var Orqan = cmbverorqan.Text;
            var Unvan = txtunvan.Text;
            var CariHesab = txtcarihes.Text;
            var Valyuta = cmbval.Text;
            //var Serencam = txbkavtoSerencam.Text;
            var MuqNo = txtmuqno.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cmbtey.Text;
            var Olke = cmbolke.Text;
            var mebleg = txtmeb.Text + " " + cmbval.Text;
            var muddet = txtmud.Text;
            var faiz = txtfaiz.Text;
            var vkfaiz = txtvkfaiz.Text;
            var telf = txttel.Text;
            var ayliq = txtayliq.Text;
            var fifd = txtfifd.Text + "%";
            var vsudahes = sudahes;
            var vfaizhes = faizhes;
            var vvkhes = vkhes;
            var vvkhesfaiz = vkfaizhes;
            var vodgun = odgunu;
            var eht_faiz = ehtfaiz;
            var girov_mebleg = txtgirovdey.Text;
            //girov1 melumatlari
            var MseriyaNo = C_seriyaNo;
            var huqobadi = C_obadi;
            var umumisah = C_umsahe;
            var yasayis = C_yasayis;
            var yardimci = C_yardimci;
            var reyestrNo = C_reyesno;
            var qeydNo = C_qeydno;
            var mtarix = C_tarix;
            var otaqsay = C_otaqsay;
            var kitabno = C_kitabno;
            var vereqno = C_vereqno;
            var msahibadi = C_sahibi;
            var girovunvani = C_girovunvan;
            var zamno = C_zamNo;
            var sahibunvan = C_qeydno;
            var mpasport = C_pasport;
            var mpastarixi = C_pastarix;
            var mverorqan = C_vereqno;
            var mtelefon = C_tel;
            var molke = C_olke;
            var ipNo = C_ipotekaNo;
            var ipNovu = C_ipotekanovu;

            //girov1 melumatlari
            var MseriyaNo2 = C_seriyaNo2;
            var huqobadi2 = C_obadi2;
            var umumisah2 = C_umsahe2;
            var yasayis2 = C_yasayis2;
            var yardimci2 = C_yardimci2;
            var reyestrNo2 = C_reyesno2;
            var qeydNo2 = C_qeydno2;
            var mtarix2 = C_tarix2;
            var otaqsay2 = C_otaqsay2;
            var kitabno2 = C_kitabno2;
            var vereqno2 = C_vereqno2;
            var msahibadi2 = C_sahibi2;
            var girovunvani2 = C_girovunvan2;
            var zamno2 = C_zamNo2;
            var sahibunvan2 = C_qeydno2;
            var mpasport2 = C_pasport2;
            var mpastarixi2 = C_pastarix2;
            var mverorqan2 = C_vereqno2;
            var mtelefon2 = C_tel2;
            var molke2 = C_olke2;
            var ipNo2 = C_ipotekaNo2;
            var ipNovu2 = C_ipotekanovu2;

            var zam1ad = zam1adi;
            var zam1pas = zam1pass;
            var zam1telef = zam1tel;
            var zam1unvani = zam1unvan;
            var zam1pastarixi = zam1Pastar;
            var zam1pasorqani = zam1Pasorqan;
            var zam1olke = zam1olkesi;

            var zam2 = zam2adi;
            var zam2pas = zam2pass;
            var zam2telef = zam2tel;
            var zam2unvani = zam2unvan;
            var zam2pastarixi = zam2Pastar;
            var zam2pasorqani = zam2Pasorqan;
            var zam2olke = zam2olkesi;

            var zam3 = zam3adi;
            var zam3pas = zam3pass;
            var zam3telef = zam3tel;
            var zam3unvani = zam3unvan;
            var zam3pastarixi = zam3Pastar;
            var zam3pasorqani = zam3Pasorqan;
            var zam3olke = zam3olkesi;

            var zammuqno1 = zaminmuqn1;
            var zammuqno2 = zaminmuqn2;
            var zammuqno3 = zaminmuqn3;
            var mebyaziile = yaziile;
            var mevayyazi = ayyazile;
            var tamcari = tamhesab;
            //var icraci = icraci_adi;
            var icraci = icraciadizaminlikde;
            var mektubNo = meknoal;


            // TODO: Word Export
            var wordapp = new Microsoft.Office.Interop.Word.Application();
            wordapp.Visible = false;

            try
            {
                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\KreditAvtoMuq.docx");

                ReplaceWordStub("{ilci}", ilsonrasi, wordDocument);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                ReplaceWordStub("{CariHesab}", tamcari, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                //ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg + "(" + mebyaziile + ")", wordDocument);
                ReplaceWordStub("{muddet}", muddet + " ay " + "(" + muddetyazi + ")", wordDocument);
                ReplaceWordStub("{faiz}", faiz + "% " + "(" + faizyazi + ")", wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz + "%" + " (" + vkfaizyazi + ")", wordDocument);
                ReplaceWordStub("{tel}", telf, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq + VAL_AD + "(" + mevayyazi + ")", wordDocument);
                ReplaceWordStub("{fifd}", fifd, wordDocument);
                ReplaceWordStub("{sudahes}", vsudahes, wordDocument);
                ReplaceWordStub("{faizhes}", vfaizhes, wordDocument);
                ReplaceWordStub("{vkhes}", vvkhes, wordDocument);
                ReplaceWordStub("{vkfaizhes}", vvkhesfaiz, wordDocument);
                ReplaceWordStub("{odgunu}", odgunu, wordDocument);
                ReplaceWordStub("{ehtfaiz}", eht_faiz, wordDocument);
                ReplaceWordStub("{girovmeb}", txtgirovdey.Text + " AZN " + "(" + girovyaziile + ")", wordDocument);
                // avto melumatlari
                ReplaceWordStub("{rey_kitab_no1}", C_kitabno, wordDocument);
                            
                ReplaceWordStub("{rey_vereq_no1}", C_vereqno, wordDocument);
                ReplaceWordStub("{rey_no1}", C_reyesno, wordDocument);
                ReplaceWordStub("{rey_qeyd_no1}", C_qeydno, wordDocument);
                ReplaceWordStub("{rey_seriya1}", C_seriyaNo, wordDocument);
                ReplaceWordStub("{rey_tarix1}", C_tel, wordDocument);
                ReplaceWordStub("{otq_say}", C_otaqsay, wordDocument);
                ReplaceWordStub("{umumi_sah}", C_umsahe, wordDocument);
                ReplaceWordStub("{yasayis_sah}", C_yasayis, wordDocument);
                ReplaceWordStub("{yardimci_sah}", C_yardimci, wordDocument);
                ReplaceWordStub("{menzil_unvan}", C_girovunvan, wordDocument);
                ReplaceWordStub("{saa2}", C_sahibi, wordDocument);
                ReplaceWordStub("{adres2}", C_qeydunvan, wordDocument);
                ReplaceWordStub("{tel2}", C_tel, wordDocument);
                ReplaceWordStub("{mektub}", mektubNo, wordDocument);
                ReplaceWordStub("{iptk_no}", ipNo, wordDocument);

                ReplaceWordStub("{icraci}", icraci, wordDocument);

                ReplaceWordStub("{zam1ad}", zam1ad, wordDocument);
                ReplaceWordStub("{zam1pas}", zam1pas, wordDocument);
                ReplaceWordStub("{zam1telef}", zam1tel, wordDocument);
                ReplaceWordStub("{zam1unvani}", zam1unvan, wordDocument);
                ReplaceWordStub("{zam1ptarix}", zam1pastarixi, wordDocument);
                ReplaceWordStub("{zam1porqan}", zam1pasorqani, wordDocument);
                ReplaceWordStub("{zam1olke}", zam1olke, wordDocument);

                ReplaceWordStub("{zam2ad}", zam2, wordDocument);
                ReplaceWordStub("{zam2pas}", zam2pas, wordDocument);
                ReplaceWordStub("{zam2telef}", zam2tel, wordDocument);
                ReplaceWordStub("{zam2unvani}", zam2unvan, wordDocument);
                ReplaceWordStub("{zam2ptarix}", zam2pastarixi, wordDocument);
                ReplaceWordStub("{zam2porqan}", zam2pasorqani, wordDocument);
                ReplaceWordStub("{zam2olke}", zam2olke, wordDocument);

                ReplaceWordStub("{zam3ad}", zam3, wordDocument);
                ReplaceWordStub("{zam3pas}", zam3pas, wordDocument);
                ReplaceWordStub("{zam3telef}", zam3tel, wordDocument);
                ReplaceWordStub("{zam3unvani}", zam3unvan, wordDocument);
                ReplaceWordStub("{zam3ptarix}", zam3pastarixi, wordDocument);
                ReplaceWordStub("{zam3porqan}", zam3pasorqani, wordDocument);
                ReplaceWordStub("{zam3olke}", zam3olke, wordDocument);

                ReplaceWordStub("{zammuq1}", zammuqno1, wordDocument);
                ReplaceWordStub("{zammuq2}", zammuqno2, wordDocument);
                ReplaceWordStub("{zammuq3}", zammuqno3, wordDocument);


                wordapp.Visible = true;
                string muqadi = MuqNo + " " + adi + " " + tamtarix;
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                wordDocument.SaveAs(Application.StartupPath + muqadi);




            }
            catch
            {
                //MessageBox.Show("Sehv oldu");
            }

        }
        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document WordDocument)
        {
            var range = WordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }
        private void doldur()
        {
            if (comboadi == "Progress")
            {
                string adhazir;
                string unvanhazir;
                //DateTime bugun = DateTime.Now.ToShortDateString();
                //label19.Text = DateTime.Now.ToShortDateString();
                //teyadi();
                //olkeadi();
                valyuta_adi();
                //

                ///
                muqavile_nom();
                //muqavile_nom1();

                adhazir = adi.ToLower();

                txtborcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);

                //txbzBorcalan.Text = txbzBorcalan.Text.ToUpper();

                txtmeb.Text = mebleg;
                txtaze.Text = seriyano;
                txttel.Text = mobil;
                unvanhazir = unvan.ToLower();
                txtunvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);
                txtvkfaiz.Text = vkfaiz;
                txtfaiz.Text = faiz;
                cmbverorqan.Text = ver_orqan;
                txtcarihes.Text = carihes;
                txtayliq.Text = ayliq;
                txtgirovdey.Text = girov_deyeri;
                datepasvertarix.Text = ver_tar.Substring(0,10);
                cmbverorqan.Text = ver_orqan;
                txtfifd.Text = fifd;
                datetarix.Text = ver_tar.Substring(0, 10);
                cmbtey.Text = teyinatadi;
                olkeadi();
                int gun = Convert.ToInt32(muddet);
                int gunsay = gun / 30;
                txtmud.Text = gunsay.ToString();
                //cboxkzTeyinat.Text = teyinat; 
                if (valyuta == "00")
                {
                    //textBox1.Text = "AZN";
                    cmbval.Text = "AZN";
                }
                else if (valyuta == "01")
                {
                    //textBox1.Text = "USD";
                    cmbval.Text = "USD";
                }
                else if (valyuta == "02")
                {
                    //textBox1.Text = "AVRO";
                    cmbval.Text = "AVRO";
                }
                //textBox3.Text = textBox1.Text;
            }
        }
        private void cmbgirovsayi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbgirovsayi.Text == "Bir əmlak girovu ilə")
            {
                frm1.Text = "Bir əmlak girovu ilə";
                frm1.Size = new Size(777, 380);
                frm1.ShowDialog();
            }
            if (cmbgirovsayi.Text == "İki əmlak girovu ilə")
            {
                frm1.Text = "İki əmlak girovu ilə";
                frm1.Size = new Size(778, 710);
                frm1.grpikigirov.Enabled = true;
                frm1.ShowDialog();
                

            }
        }

        private void frmmenzil_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            doldur();
            this.datetarix.Properties.DisplayFormat.FormatString = "dd/MM/yyyy";
        }

        private void cmbteminat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtayliq.Text == "")
            {
                MessageBox.Show("Aylıq ödəniş qeyd edilməyib!!!", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {


                int zaminlarsay = zam_say;


                Zaminler zmnlar = new Zaminler();
                zmnlar.zmnlartarix = umumitar;
                //zmnlar.muqtipi = label24.Text;
                //zmnlar.muqtipi = girovnovucombo;
                zmnlar.muqtipi = muqtipi;
                zmnlar.comboadi = cmbteminat.Text;
                zmnlar.muqtipi = "menzil";
                zmnlar.qeydnozamin = subkod_qeyd;
                zmnlar.groupBox1.Visible = false;
                zmnlar.groupBox2.Visible = false;
                zmnlar.groupBox3.Visible = false;

                //zmnlar.kataloqtarixi = secilmistarix;
                if (cmbteminat.Text == "Bir nəfərin zəmanəti")
                {
                    zmnlar.groupBox1.Visible = true;

                    zmnlar.txb_zam1No.Text = zaminlarsay.ToString();

                    zmnlar.ShowDialog();

                    zmnlar.txbzam1adi.Text = zam1adi;
                    zmnlar.txbzam1Pasport.Text = zam1pass;
                    zmnlar.txbzam1Telefon.Text = zam1tel;
                    zmnlar.txbzam1Unvan.Text = zam1unvan;
                    zmnlar.muqtipi = test;
                    zmnlar.zmnlartarix = umumitar;

                }
                else if (cmbteminat.Text == "İki nəfərin zəmanəti")
                {
                    zmnlar.groupBox1.Visible = true;
                    zmnlar.groupBox2.Visible = true;
                    zmnlar.txb_zam1No.Text = zaminlarsay.ToString();
                    zaminlarsay = zaminlarsay + 1;
                    zmnlar.txb_zam2No.Text = zaminlarsay.ToString();

                    zmnlar.ShowDialog();

                    zmnlar.txbzam1adi.Text = zam1adi;
                    zmnlar.txbzam1Pasport.Text = zam1pass;
                    zmnlar.txbzam1Telefon.Text = zam1tel;
                    zmnlar.txbzam1Unvan.Text = zam1unvan;
                    zmnlar.muqtipi = test;
                    zmnlar.zmnlartarix = umumitar;

                    zmnlar.txbzam2Zamin.Text = zam2adi;
                    zmnlar.txbzam2Pasport.Text = zam2pass;
                    zmnlar.txbzam2Telefon.Text = zam2tel;
                    zmnlar.txbzam2Unvan.Text = zam2unvan;


                }
                else if (cmbteminat.Text == "Üç nəfərin zəmanəti")
                {
                    zmnlar.groupBox1.Visible = true;
                    zmnlar.groupBox2.Visible = true;
                    zmnlar.groupBox3.Visible = true;
                    zmnlar.txb_zam1No.Text = zaminlarsay.ToString();
                    zaminlarsay = zaminlarsay + 1;
                    zmnlar.txb_zam2No.Text = zaminlarsay.ToString();
                    zaminlarsay = zaminlarsay + 1;
                    zmnlar.txb_zam3No.Text = zaminlarsay.ToString();

                    zmnlar.ShowDialog();

                    zmnlar.txbzam1adi.Text = zam1adi;
                    zmnlar.txbzam1Pasport.Text = zam1pass;
                    zmnlar.txbzam1Telefon.Text = zam1tel;
                    zmnlar.txbzam1Unvan.Text = zam1unvan;
                    zmnlar.muqtipi = test;
                    zmnlar.zmnlartarix = umumitar;

                    zmnlar.txbzam2Zamin.Text = zam2adi;
                    zmnlar.txbzam2Pasport.Text = zam2pass;
                    zmnlar.txbzam2Telefon.Text = zam2tel;
                    zmnlar.txbzam2Unvan.Text = zam2unvan;

                    zmnlar.txbzam3Zamin.Text = zam3adi;
                    zmnlar.txbzam3Pasport.Text = zam3pass;
                    zmnlar.txbzam3Telefon.Text = zam3tel;
                    zmnlar.txbzam3Unvan.Text = zam3unvan;
                }

            }
        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {

        }
    }
}
