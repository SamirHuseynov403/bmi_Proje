using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Globalization;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.Runtime.InteropServices;

namespace BMI
{
    public partial class Avtomobil : Form
    {
        private readonly string templatefilname;
        public Avtomobil()
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

        public string test { get; set; }
        public string mek_no { get; set; }
        public string umumitar { get; set; }

        public static string gedenbilgi = "", kr1 = "", kr2 = "", kr3 = "", kr4 = "", kr5 = "", kr6 = "", kr7 = "";
        public static string db1 = "", db2 = "", db3 = "", db4 = "", db5 = "", db6 = "", db7 = "", c1 = "", sub1 = "";
        public static string meb1 = "", meb2 = "", meb3 = "", meb4 = "", meb5 = "", meb6 = "", meb7 = "";
        public static string teyinat1 = "", teyinat2 = "", teyinat3 = "", teyinat4 = "", teyinat5 = "", teyinat6 = "", teyinat7 = "";

        private readonly string avtomuq = @"C:\BMI_\BMI\bin\Debug\qiziltablozaminsiz.docx";
        private readonly string avtobirzamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_Birzamin.docx";
        private readonly string avtoikizamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_ikizamin.docx";
        private readonly string avtouczamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_Uczamin.docx";

        string VAL_AD;

        private void valyuta_adi()
        {
            if (valyuta == "00")
            {
                
                textBox3.Text = "AZN";
                VAL_AD = "AZN";
            }
            else if (valyuta == "01")
            {
                
                textBox3.Text = "USD";
                VAL_AD = "USD";
            }
            else if (valyuta == "02")
            {
                
                textBox3.Text = "AVRO";
                VAL_AD = "AVRO";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            avtocevirme();
            excele_at_avto();
        }
        private void olkeadi()
        {
            if (olke == "AZ")
            {
                cboxavtoOlke.Text = "Azərbaycan Respublikası";
            }
            else if (olke == "IRN")
            {
                cboxavtoOlke.Text = " İran İslam Respublikası";
            }
        }

        string sigortahesabi;
        private void avtocevirme()
        {
            string kassa_hesazn, kassa_hesusd, kassa_hesavro, x_h_kassaazn, x_h_kassausd, x_h_kassaavro, x_h_kreditazn, x_h_kreditusd, x_h_kreditavro,sigorta;
            string x_h_kredit_azn_QR, x_h_kredit_usd_QR, x_h_kredit_avro_QR;

            kassa_hesazn = "10010000000000100000";
            kassa_hesusd = "10020010000000100000";
            kassa_hesavro = "10020020000000100000";

            x_h_kassaazn = "67010000000000600000";
            x_h_kassausd = "67020010000000600000";
            x_h_kassaavro = "67020020000000600000";

            x_h_kreditazn = "67034000010000600000";
            sigorta = sigortahesabi;
            
            x_h_kreditusd = "67044010010000600000";
            x_h_kreditavro = "67044020010000600000";

            x_h_kredit_azn_QR = "67020010000000600000";
            x_h_kredit_usd_QR = "67044010050000600000";
            x_h_kredit_avro_QR = "67044020020000600000";

            teyinat1 = DateTime.Now.ToShortDateString() + " il tarixli kredit müq.əsəasən " + txbkavtoMuddet.Text + " ay müddətinə kredit verilir";
            
            db1 = sudahes;
            db2 = kassa_hesazn;
            if (checkBox1.Checked==true)
	{
		    db2 = kassa_hesazn;
            db3 = kassa_hesazn;
            db4 = carihes;
            db5 = carihes;
            db6 = carihes;

            kr1 = carihes;
            kr2 = carihes;
            kr3 = sigorta;
            kr4 = kassa_hesazn;
            sub1 = subhes;
            kr5 = x_h_kassaazn;
            kr6 = x_h_kreditazn;

            meb1 = mebleg;
            double krmeb = Convert.ToInt32(mebleg) / 100;
            double krxh = krmeb / 2;
            meb2 = krmeb.ToString();
            meb3=sig_meb.Text;
            meb4 = mebleg;
            meb5 = krxh.ToString();
            meb6 = krxh.ToString();

            teyinat2 = DateTime.Now.ToShortDateString() + " tar.kr.müq.əs 1% x/h tut ";
            teyinat3 = "Avto sığ.üçün " + txbkavtoMuddet.Text + " ay " + adi;
            teyinat4 = "Kredit verilməsi ilə əlaqədar";
            teyinat5 = "Kassa 0,5 % x/h tutulur";
            teyinat6 = "kred.ver.əlaqədar 0,5 % x/h";
	}
            else
	{
            db3 = carihes;
            db4 = carihes;
            db5 = carihes;

            kr1 = carihes;
            kr2 = carihes;
            kr3 = kassa_hesazn;
            sub1 = subhes;
            kr4 = x_h_kassaazn;
            kr5 = x_h_kreditazn;

            meb1 = mebleg;
            int krmeb = Convert.ToInt32(mebleg) / 100;
            int krxh = krmeb / 2;
            meb2 = krmeb.ToString();
            meb3 = mebleg;
            meb4 = krxh.ToString();
            meb5 = krxh.ToString();

            teyinat2 = DateTime.Now.ToShortDateString() + " tar.kr.müq.əs 1% x/h tut ";
            teyinat3 = "Kredit verilməsi ilə əlaqədar";
            teyinat4 = "Kassa 0,5 % x/h tutulur";
            teyinat5 = "kred.ver.əlaqədar 0,5 % x/h";
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

        private void excele_at_avto()
        {

            string sablonExcelDosyaYolu = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Kredit pr.xls";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(sablonExcelDosyaYolu);

            // Veriyi belirli hücrelere yazdır
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1]; // İlk çalışma sayfası
            worksheet.Cells[2, 2] = "0";
            worksheet.Cells[2, 3] = sub1;
            worksheet.Cells[2, 4] = db1;
            worksheet.Cells[3, 4] = db2;
            worksheet.Cells[4, 4] = db3;
            worksheet.Cells[5, 4] = db4;
            worksheet.Cells[5, 8] = "0";
            worksheet.Cells[6, 4] = db5;
            worksheet.Cells[7, 4] = db6;
            worksheet.Cells[2, 6] = kr1;
            worksheet.Cells[3, 6] = kr2;
            worksheet.Cells[4, 6] = kr3;
            worksheet.Cells[5, 6] = kr4;
            worksheet.Cells[6, 6] = kr5;
            worksheet.Cells[7, 6] = kr6;
            worksheet.Cells[2, 7] = meb1;
            worksheet.Cells[3, 7] = meb2;
            worksheet.Cells[4, 7] = meb3;
            worksheet.Cells[5, 7] = meb4;
            worksheet.Cells[6, 7] = meb5;
            worksheet.Cells[7, 7] = meb6;
            worksheet.Cells[2, 10] = teyinat1;
            worksheet.Cells[3, 10] = teyinat2;
            worksheet.Cells[4, 10] = teyinat3;
            worksheet.Cells[5, 10] = teyinat4;
            worksheet.Cells[6, 10] = teyinat5;
            worksheet.Cells[7, 10] = teyinat6;

            workbook.Save(); // Değişiklikleri kaydet
            workbook.Close(); // Excel dosyasını kapat
            excelApp.Quit(); // Excel uygulamasını kapat

            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
        }

        private void doldur()
        {
            if (comboadi == "Progress")
            {
                string adhazir;
                string unvanhazir;
                valyuta_adi();
                muqavile_nom();
                adhazir = adi.ToLower();
                txbkavtoBorcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);
                txbkavtoMebleg.Text = mebleg;
                txbkavtoPasport.Text = seriyano;
                txbkavtoTelefon.Text = mobil;
                unvanhazir = unvan.ToLower();
                txbkavtoUnvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);
                txbkavtoVKFaiz.Text = vkfaiz;
                txbkavtoFaiz.Text = faiz;
                cboxavtoOrqan.Text = ver_orqan;
                txbkavtoCarihesab.Text = carihes;
                txbkavtoAyliq.Text = ayliq;
                txb_avtodeyer.Text = girov_deyeri;
                txbkavtoFİFD.Text = fifd;
                dateTimePicker1.Text = ver_tar;
                int gun = Convert.ToInt32(muddet);
                int gunsay = gun / 30;
                txbkavtoMuddet.Text = gunsay.ToString();
                if (valyuta == "00")
                {
                    textBox1.Text = "AZN";
                    cboxavtoValyuta.Text = "AZN";
                }
                else if (valyuta == "01")
                {
                    textBox1.Text = "USD";
                    cboxavtoValyuta.Text = "USD";
                }
                else if (valyuta == "02")
                {
                    textBox1.Text = "AVRO";
                    cboxavtoValyuta.Text = "AVRO";
                }
                textBox3.Text = textBox1.Text;
            }
        }

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
                string query = "select max(-to_number(substr(t.qey_nom,5,5)))mn from odb.xaric_mektub t where t.il='" + tarixIl + "'";
                OracleCommand cmd = new OracleCommand(query, Orcon);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString());
                    proid = id.ToString();
                    mek_no_arti = id + 1;
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
            string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbkavtoMebleg.Text));
            string ayyazile = yaziyaCevir(Convert.ToDecimal(txbkavtoAyliq.Text));
            string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbkavtoMuddet.Text));
            string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbkavtoFaiz.Text));
            string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbkavtoVKFaiz.Text));
            string girovyaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txb_avtodeyer.Text));

            var adi = txbkavtoBorcalan.Text;
            var passport = txbkavtoPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = cboxavtoOrqan.Text;
            var Unvan = txbkavtoUnvan.Text;
            var CariHesab = txbkavtoCarihesab.Text;
            var Valyuta = cboxavtoValyuta.Text;
            var Serencam = txbkavtoSerencam.Text;
            var MuqNo = txbkavtoMuqNo.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cboxavtoTeyinat.Text;
            var Olke = cboxavtoOlke.Text;
            var mebleg = txbkavtoMebleg.Text + " " + cboxavtoValyuta.Text;
            var muddet = txbkavtoMuddet.Text;
            var faiz = txbkavtoFaiz.Text;
            var vkfaiz = txbkavtoVKFaiz.Text;
            var telf = txbkavtoTelefon.Text;
            var ayliq = txbkavtoAyliq.Text;
            var fifd = txbkavtoFİFD.Text + " " + textBox5.Text;
            var vsudahes = sudahes;
            var vfaizhes = faizhes;
            var vvkhes = vkhes;
            var vvkhesfaiz = vkfaizhes;
            var vodgun = odgunu;
            var eht_faiz = ehtfaiz;
            var girov_mebleg = txb_avtodeyer.Text;
            var modeli = txbkavtoModel.Text;
            var muherrik = txbkavtoMuherrik.Text;
            var ban = txbkavtoBan.Text;
            var rengi = txbkavtoRengi.Text;
            var ili = txbkavtoİli.Text;

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
            var icraci = icraciadizaminlikde;
            var mektubNo = meknoal;
            
            var wordapp = new Microsoft.Office.Interop.Word.Application();
            wordapp.Visible = false;

            try
            {
                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\KreditAvtoMuq.docx");
                
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                ReplaceWordStub("{CariHesab}", tamcari, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                ReplaceWordStub("{Serencam}", Serencam, wordDocument);
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
                ReplaceWordStub("{girovmeb}", txb_avtodeyer.Text + " AZN " + "(" + girovyaziile + ")", wordDocument);
                ReplaceWordStub("{modeli}", modeli, wordDocument);
                ReplaceWordStub("{muherrik}", muherrik, wordDocument);
                ReplaceWordStub("{ban}", ban, wordDocument);
                ReplaceWordStub("{rengi}", rengi, wordDocument);
                ReplaceWordStub("{ili}", ili, wordDocument);
                ReplaceWordStub("{mektub}", mektubNo, wordDocument);

                ReplaceWordStub("{icraci}", icraci, wordDocument);

                ReplaceWordStub("{zam1ad}", zam1, wordDocument);
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
                wordDocument.SaveAs(Application.StartupPath + muqadi);
                
            }
            catch
            {
            }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document WordDocument)
        {
            var range = WordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Zaminler zmnlar = new Zaminler();
            zmnlar.ShowDialog();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void Avtomobil_Load(object sender, EventArgs e)
        {
            doldur();
        }

        private void Avtomobil_Load_1(object sender, EventArgs e)
        {
            dateavtoMuqtarix.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            label27.Text = icraci_kod;
            doldur();
            olkeadi();
            muracietsayartimsiz();
            muqavile_nom();
        }
        
        int kr_say;
        int ser_say;
        public int zam_say;
        public void muqavile_nom()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            string satirsayi = "2021";

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
                txbkavtoMuqNo.Text = kr_say.ToString();
                ser_say = Convert.ToInt32(dr["kr_serencam"].ToString()) + 1;
                txbkavtoSerencam.Text = ser_say.ToString();
            }
            con.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                sig_meb.Enabled = true;
                txbsighes.Text = "35090000001344900000";
                sigortahesabi = txbsighes.Text;
            }
            else if (checkBox1.Checked == false)
            {
                sig_meb.Enabled = false;
                txbsighes.Text = "";
                sigortahesabi = txbsighes.Text;
            }   
        }
        int kecencavab = 0;
        private void meknoal()
    {
            int test=45;
        string tarixIl = DateTime.Now.Date.Year.ToString();
        Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        Orcon.Open();
        Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('DYP', TO_DATE('" + dateavtoMuqtarix.Text + "','dd-MM-yyyy'), 'avto gir sal', '"+icraci_kod+"',  " + tarixIl + ")", Orcon);
        Orcom.ExecuteNonQuery();
        Orcon.Close();
    }

        private void btnAvtoYazdir_Click(object sender, EventArgs e)
        {
            if (txbsighes.Text=="")
            {
                MessageBox.Show("Siğorta hesabı qeyd edilməyib");
            }
            else 
           {
                if (txbsighes.Text != "")
                {
                    meknoal();
                    muracietsayartimsiz();
                    word_at();
                    avtocevirme();
                    excele_at_avto();
                }
            }
        }
        private void muracietsayartimsiz()
        {
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                string proid;
                string hevsecal;
                int mek_no_arti=0;
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                //string query = "select qey_nom ,il from odb.xaric_mektub where il='2022' order by qey_nom desc";
                string query = "select max(-to_number(substr(t.qey_nom,5,5)))mn from odb.xaric_mektub t where t.il='"+tarixIl+"'";
                OracleCommand cmd = new OracleCommand(query, Orcon);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString());
                    proid = id.ToString();
                    mek_no_arti = id + 1;
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
                mek_no = tarixIl+ "-" +mek_no_arti.ToString();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string zam1 { get; set; }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int zaminlarsay = zam_say;

            Zaminler zmnlar = new Zaminler();

            zmnlar.zmnlartarix = umumitar;
            zmnlar.muqtipi = label31.Text;
            zmnlar.muqtipi = muqtipi;
            zmnlar.comboadi = comboBox1.Text;
            zmnlar.muqtipi = label31.Text;
            zmnlar.qeydnozamin = subkod_qeyd;
            zmnlar.groupBox1.Visible = false;
            zmnlar.groupBox2.Visible = false;
            zmnlar.groupBox3.Visible = false;
            zmnlar.label23.Text = label31.Text;
            if (comboBox1.Text == "Bir nəfərin zəmanəti")
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
            else if (comboBox1.Text == "İki nəfərin zəmanəti")
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
            else if (comboBox1.Text == "Üç nəfərin zəmanəti")
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
}

