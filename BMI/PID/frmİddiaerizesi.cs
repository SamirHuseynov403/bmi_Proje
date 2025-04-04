using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class frmİddiaerizesi : Form
    {
        public frmİddiaerizesi()
        {
            InitializeComponent();
        }
        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;
        public string mek_no { get; set; }
        public string mek_nomaliyye { get; set; }
        public string tamtarix { get; set; }
        string olke = "";
        string ilkin_muqtarix = "";
        string VAL_AD;
        string muqavilNo = "";
        string arayTar = "";
        string teminatsay = "";
        string arayisad = "";
        string arayisNo = "";
        string muqavileili = "";

        private void groupControl1_Paint(object sender, PaintEventArgs e)
        {

        }
        
        private void doldur()
        {
            decimal summa = 0;
            decimal krmeb = 0;
            decimal qaliqsumma = 0;
            decimal topesasqaliq = 0;
            decimal topesasqaliqcixsonra = 0;
            decimal summa_19 = 0;
            decimal cemqaliq = 0;
            string valyuta = "";
            valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
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
            txtayliq.BackColor = Color.White;
            txtadi.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            dateTimePicker1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();//.Substring(0, 10);
            ilkin_muqtarix = dateTimePicker1.Text;
            summa = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[7].Value.ToString());
            krmeb = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[6].Value.ToString());
            qaliqsumma = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[8].Value.ToString());
            summa_19 = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[29].Value.ToString());
            topesasqaliq = krmeb - (summa+summa_19);
            txtsubkod.Text= dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtsuda.Text= dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtodenilenesas.Text = Decimal.Round(topesasqaliq, 2).ToString();
            txtvkborc.Text= Decimal.Round(summa + summa_19, 2).ToString();
            decimal vkfaiz =Convert.ToDecimal( dataGridView1.CurrentRow.Cells[34].Value.ToString());
            decimal faiz = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[33].Value.ToString());
            decimal topfaiz = vkfaiz + faiz;
            txtfaizborc.Text = topfaiz.ToString();
            cemqaliq = summa + summa_19+topfaiz;
            txttopborc.Text = cemqaliq.ToString();
            txtpassp.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            cmbverorqan.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            string testtarix = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            if (testtarix=="")
            {
                txtpastar.Text = "";
            }

            else
            {
                txtpastar.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0,10);
            }

            txtkrmeb.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            textEdit17.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            //txtrusum.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            txtayliq.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            if (dataGridView2.Rows.Count==1)
            {
                txtzam1.Text = dataGridView2.Rows[0].Cells[0].Value.ToString();
                txtzamunvan1.Text = dataGridView2.Rows[0].Cells[2].Value.ToString();
            }
            if (dataGridView2.Rows.Count == 2)
            {
                txtzam1.Text = dataGridView2.Rows[0].Cells[0].Value.ToString();
                txtzamunvan1.Text = dataGridView2.Rows[0].Cells[2].Value.ToString();
                txtzam2.Text = dataGridView2.Rows[1].Cells[0].Value.ToString();
                txtzamunvan2.Text = dataGridView2.Rows[1].Cells[2].Value.ToString();
            }
            if (dataGridView2.Rows.Count == 3)
            {
                txtzam1.Text = dataGridView2.Rows[0].Cells[0].Value.ToString();
                txtzamunvan1.Text = dataGridView2.Rows[0].Cells[2].Value.ToString();
                txtzam2.Text = dataGridView2.Rows[1].Cells[0].Value.ToString();
                txtzamunvan2.Text = dataGridView2.Rows[1].Cells[2].Value.ToString();
                txtzam3.Text = dataGridView2.Rows[2].Cells[0].Value.ToString();
                txtzamunvan3.Text = dataGridView2.Rows[2].Cells[2].Value.ToString();
            }


            if (Convert.ToDouble(txtayliq.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString()) == Convert.ToDouble(txtkrmeb.Text))
            {
                txtayliq.Text = "";
                txtayliq.BackColor = Color.Red;
            }



            txtfaiz.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            txtvkfaiz.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();

            int muddeti = Convert.ToInt32(dataGridView1.CurrentRow.Cells[13].Value.ToString()) / 30;
            txtmuddet.Text = muddeti.ToString();
        //    txbzTelefon.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            txtunvan.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();

            tarixitapsoz();
            tarixitapsozilkintarix();
            tarixitapsozcaritarix();
            string adhazir;
            string unvanhazir;
            adhazir = txtadi.Text.ToLower();
            unvanhazir = txtunvan.Text.ToLower();
            txtadi.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);
            txtunvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);

        }

        private void zamindoldurzamin()
        {

  
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
        private void word_at()
        {
            try
            {

                if (!System.IO.File.Exists(Application.StartupPath + "\\İddia.doc"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txtkrmeb.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txtayliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtmuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtfaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtvkfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var mekno = mek_no;
                var mekno1 = mek_nomaliyye;
                var mektar = tarix_sozcaritarix;
                var caritar = tarix_sozcaritarix;
                var adi = txtadi.Text;
                var passport = txtpassp.Text;
                var pasvertarix = txtpastar.Text;
                var Orqan = cmbverorqan.Text;
                var Unvan = txtunvan.Text;
                //var CariHesab = txbzCarihesab.Text;
                var Valyuta = cmbval.Text;
                ///var Serencam = txbzSerencam.Text;
                var MuqNo = txtmuqno.Text;
                var odenmeli = txtodenilmeli.Text;
                var odenilmis = txtodenilmis.Text;
                var tamborc = txttopborc.Text;
                var vkborc = txtvkborc.Text;
                var vkfaizborc = txtfaizborc.Text;
                var rusummeb = txtrusum.Text;
                var mehkeme = cmbrayonıar.Text;
                var Muqtarix = tarix_sozilkintarix;

                //var Teyinat = cboxkzTeyinat.Text;
                //var Olke = cboxkzOlke.Text;
                var mebleg = txtkrmeb.Text; //+ " " + cboxkzValyuta.Text;
                var muddet = txtmuddet.Text;
                var faiz = txtfaiz.Text;
                var vkfaiz = txtvkfaiz.Text;
                var cedfaiz = txtfaizmebleg.Text;
                var ayliq = txtayliq.Text;
                //var fifd = txbzFIFD.Text + " " + textBox5.Text;
                var topcedmeb = txttxttoplamcedvel.Text;
                var wayesas = txtayesas.Text;
                var wayfaiz = txtayfaiz.Text;
                var waytop = txtaytoplam.Text;
                var sontarix = txtsontarix.Text;
                var sonmebleg = txtsonmebleg.Text;

                var caves = txtodenilenesas.Text;
                var cavfaiz = txtodfaiz.Text;
                var cavcem = txtodcem.Text;

                //var vodgun = odgunu;
                //var eht_faiz = ehtfaiz;
                var zam1 = txtzam1.Text;
                var zam1No = txtzamno1.Text;
                var zam1unvan = txtzamunvan1.Text;

                var zam2 = txtzam2.Text;
                var zam2No = txtzam1.Text;
                var zam2unvan = txtzamunvan2.Text;

                var zam3 = txtzam3.Text;
                var zam3No = txtzam1.Text;
                var zam3unvan = txtzamunvan3.Text;

                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;



                //var icraci = icraci_adi;
                //var icraciadi = icraciadizaminlikde;



            //TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\İddia.doc");
                //var wordDocument = wordapp.Documents.Open(zamin1);
                ReplaceWordStub("{mekNo}", mekno, wordDocument);
                ReplaceWordStub("{mekNo1}", mekno1, wordDocument);
                ReplaceWordStub("{krtarix}", caritar, wordDocument);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                //ReplaceWordStub("{CariHesab}", tamcari, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                //ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{mehkeme}", mehkeme, wordDocument);
                ReplaceWordStub("{yetirmeli}", odenmeli, wordDocument);
                ReplaceWordStub("{yetirmis}", odenilmis, wordDocument);
                ReplaceWordStub("{topborc}", tamborc, wordDocument);
                ReplaceWordStub("{vkborc}", vkborc, wordDocument);
                ReplaceWordStub("{cedfaiz}", cedfaiz, wordDocument);
                ReplaceWordStub("{topcedmebleg}", topcedmeb, wordDocument);
                ReplaceWordStub("{vkfaizmeb}", vkfaizborc, wordDocument);
                ReplaceWordStub("{rusum}", rusummeb, wordDocument);
                ReplaceWordStub("{ayesas}", wayesas, wordDocument);
                ReplaceWordStub("{ayfaiz}", wayfaiz, wordDocument);
                ReplaceWordStub("{aycemi}", waytop, wordDocument);
                ReplaceWordStub("{snod}", sontarix, wordDocument);
                ReplaceWordStub("{snodmeb}", sonmebleg, wordDocument);

                ReplaceWordStub("{cavesas}", caves, wordDocument);
                ReplaceWordStub("{cavfaiz}", cavfaiz, wordDocument);
                ReplaceWordStub("{cavcem}", cavcem, wordDocument);

                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                //ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                //ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg + "(" + mebyaziile + ")", wordDocument);
                ReplaceWordStub("{muddet}", muddet + " ay " , wordDocument);
                ReplaceWordStub("{faiz}", faiz + "% " , wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz + "%" , wordDocument);
                // ReplaceWordStub("{tel}", telf, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq + VAL_AD + "(" + mevayyazi + ")", wordDocument);

                //ReplaceWordStub("{icraci}", icraciadi, wordDocument);

                ReplaceWordStub("{zam1ad}", zam1, wordDocument);
                ReplaceWordStub("{zam1unvan}", zam1unvan, wordDocument);
                ReplaceWordStub("{zammuq1}", zam1No, wordDocument);

                ReplaceWordStub("{zam2ad}", zam2, wordDocument);


                ReplaceWordStub("{zam3ad}", zam3, wordDocument);


                wordapp.Visible = true;
                string muqadi = MuqNo + " " + adi + " " + tamtarix;
                //string muqadi = "yoxlanis";
                //string muqadi = adi;
                //wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                string testt = @"\\192.168.0.5\12345\Personal\Anar_Is\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\Məhkəmə sənədlərii\" + txtadi.Text + "\\";

                wordDocument.SaveAs(testt + muqadi + ".doc");
                //wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".doc");


            }
            catch (Exception)
            {


            }
            finally { }

        }
        private void word_at_ikizamin()
        {
            try
            {

                if (!System.IO.File.Exists(Application.StartupPath + "\\İddia_ikizamin.doc"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txtkrmeb.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txtayliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtmuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtfaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtvkfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var mekno = mek_no;
                var mekno1 = mek_nomaliyye;
                var mektar = tarix_sozcaritarix;
                var caritar = tarix_sozcaritarix;
                var adi = txtadi.Text;
                var passport = txtpassp.Text;
                var pasvertarix = txtpastar.Text;
                var Orqan = cmbverorqan.Text;
                var Unvan = txtunvan.Text;
                //var CariHesab = txbzCarihesab.Text;
                var Valyuta = cmbval.Text;
                ///var Serencam = txbzSerencam.Text;
                var MuqNo = txtmuqno.Text;
                var odenmeli = txtodenilmeli.Text;
                var odenilmis = txtodenilmis.Text;
                var tamborc = txttopborc.Text;
                var vkborc = txtvkborc.Text;
                var vkfaizborc = txtfaizborc.Text;
                var rusummeb = txtrusum.Text;
                var mehkeme = cmbrayonıar.Text;
                var Muqtarix = tarix_sozilkintarix;

                //var Teyinat = cboxkzTeyinat.Text;
                //var Olke = cboxkzOlke.Text;
                var mebleg = txtkrmeb.Text; //+ " " + cboxkzValyuta.Text;
                var muddet = txtmuddet.Text;
                var faiz = txtfaiz.Text;
                var vkfaiz = txtvkfaiz.Text;
                var cedfaiz = txtfaizmebleg.Text;
                var ayliq = txtayliq.Text;
                //var fifd = txbzFIFD.Text + " " + textBox5.Text;
                var topcedmeb = txttxttoplamcedvel.Text;
                var wayesas = txtayesas.Text;
                var wayfaiz = txtayfaiz.Text;
                var waytop = txtaytoplam.Text;
                var sontarix = txtsontarix.Text;
                var sonmebleg = txtsonmebleg.Text;

                var caves = txtodenilenesas.Text;
                var cavfaiz = txtodfaiz.Text;
                var cavcem = txtodcem.Text;

                //var vodgun = odgunu;
                //var eht_faiz = ehtfaiz;
                var zam1 = txtzam1.Text;
                var zam1No = txtzamno1.Text;
                var zam1unvan = txtzamunvan1.Text;

                var zam2 = txtzam2.Text;
                var zam2No = txtzam1.Text;
                var zam2unvan = txtzamunvan2.Text;

                var zam3 = txtzam3.Text;
                var zam3No = txtzam1.Text;
                var zam3unvan = txtzamunvan3.Text;

                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;

                //var icraci = icraci_adi;
                //var icraciadi = icraciadizaminlikde;



                // TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\İddia_ikizamin.doc");
                //var wordDocument = wordapp.Documents.Open(zamin1);
                ReplaceWordStub("{mekNo}", mekno, wordDocument);
                ReplaceWordStub("{mekNo1}", mekno1, wordDocument);
                ReplaceWordStub("{krtarix}", caritar, wordDocument);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                //ReplaceWordStub("{CariHesab}", tamcari, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                //ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{mehkeme}", mehkeme, wordDocument);
                ReplaceWordStub("{yetirmeli}", odenmeli, wordDocument);
                ReplaceWordStub("{yetirmis}", odenilmis, wordDocument);
                ReplaceWordStub("{topborc}", tamborc, wordDocument);
                ReplaceWordStub("{vkborc}", vkborc, wordDocument);
                ReplaceWordStub("{cedfaiz}", cedfaiz, wordDocument);
                ReplaceWordStub("{topcedmebleg}", topcedmeb, wordDocument);
                ReplaceWordStub("{vkfaizmeb}", vkfaizborc, wordDocument);
                ReplaceWordStub("{rusum}", rusummeb, wordDocument);
                ReplaceWordStub("{ayesas}", wayesas, wordDocument);
                ReplaceWordStub("{ayfaiz}", wayfaiz, wordDocument);
                ReplaceWordStub("{aycemi}", waytop, wordDocument);
                ReplaceWordStub("{snod}", sontarix, wordDocument);
                ReplaceWordStub("{snodmeb}", sonmebleg, wordDocument);

                ReplaceWordStub("{cavesas}", caves, wordDocument);
                ReplaceWordStub("{cavfaiz}", cavfaiz, wordDocument);
                ReplaceWordStub("{cavcem}", cavcem, wordDocument);

                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                //ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                //ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg + "(" + mebyaziile + ")", wordDocument);
                ReplaceWordStub("{muddet}", muddet + " ay " , wordDocument);
                ReplaceWordStub("{faiz}", faiz + "% " , wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz + "%" , wordDocument);
                // ReplaceWordStub("{tel}", telf, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq + VAL_AD + "(" + mevayyazi + ")", wordDocument);

                //ReplaceWordStub("{icraci}", icraciadi, wordDocument);

                ReplaceWordStub("{zam1ad}", zam1, wordDocument);
                ReplaceWordStub("{zam1unvan}", zam1unvan, wordDocument);
                ReplaceWordStub("{zammuq1}", zam1No, wordDocument);

                ReplaceWordStub("{zam2ad}", zam2, wordDocument);


                ReplaceWordStub("{zam3ad}", zam3, wordDocument);


                wordapp.Visible = true;
                string muqadi = MuqNo + " " + adi + " " + tamtarix;
                //string muqadi = "yoxlanis";
                //string muqadi = adi;
                // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                //wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".doc");
                string testt = @"\\192.168.0.5\12345\Personal\Anar_Is\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\Məhkəmə sənədlərii\" + txtadi.Text + "\\";

                wordDocument.SaveAs(testt + muqadi + ".doc");

            }
            catch (Exception)
            {


            }
            finally { }
        }
        
        private  void KlasorOlusturmaIslemi()
        {
            
            DirectoryInfo directoryInfo = Directory.
            CreateDirectory(@"\\192.168.0.5\12345\Personal\Anar_Is\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\Məhkəmə sənədlərii\"+ txtadi.Text);
            //Directory.CreateDirectory(@"fs:\12345\Personal\Samir\test");//C sürücüsü altında klasör oluşturma
            //Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"Salih\DirectorySinifi");//Projenin Exe sinin bulunduğu kısımda klasör oluşturma.
        }
            private void word_at_uczamin()
        {
            try
            {
                
                if (!System.IO.File.Exists(Application.StartupPath + "\\İddia_uczamin.doc"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txtkrmeb.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txtayliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtmuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtfaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txtvkfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var mekno = mek_no;
                var mekno1 = mek_nomaliyye;
                var mektar = tarix_sozcaritarix;
                var caritar = tarix_sozcaritarix;
                var adi = txtadi.Text;
                var passport = txtpassp.Text;
                var pasvertarix = txtpastar.Text;
                var Orqan = cmbverorqan.Text;
                var Unvan = txtunvan.Text;
                //var CariHesab = txbzCarihesab.Text;
                var Valyuta = cmbval.Text;
                ///var Serencam = txbzSerencam.Text;
                var MuqNo = txtmuqno.Text;
                var odenmeli = txtodenilmeli.Text;
                var odenilmis = txtodenilmis.Text;
                var tamborc = txttopborc.Text;
                var vkborc = txtvkborc.Text;
                var vkfaizborc = txtfaizborc.Text;
                var rusummeb = txtrusum.Text;
                var mehkeme = cmbrayonıar.Text;
                var Muqtarix = tarix_sozilkintarix;

                //var Teyinat = cboxkzTeyinat.Text;
                //var Olke = cboxkzOlke.Text;
                var mebleg = txtkrmeb.Text; //+ " " + cboxkzValyuta.Text;
                var muddet = txtmuddet.Text;
                var faiz = txtfaiz.Text;
                var vkfaiz = txtvkfaiz.Text;
                var cedfaiz = txtfaizmebleg.Text;
                var ayliq = txtayliq.Text;
                //var fifd = txbzFIFD.Text + " " + textBox5.Text;
                var topcedmeb = txttxttoplamcedvel.Text;
                var wayesas = txtayesas.Text;
                var wayfaiz = txtayfaiz.Text;
                var waytop = txtaytoplam.Text;
                var sontarix = txtsontarix.Text;
                var sonmebleg = txtsonmebleg.Text;

                var caves = txtodenilenesas.Text;
                var cavfaiz = txtodfaiz.Text;
                var cavcem = txtodcem.Text;

                //var vodgun = odgunu;
                //var eht_faiz = ehtfaiz;
                var zam1 = txtzam1.Text;
                var zam1No = txtzamno1.Text;
                var zam1unvan = txtzamunvan1.Text;

                var zam2 = txtzam2.Text;
                var zam2No = txtzam1.Text;
                var zam2unvan = txtzamunvan2.Text;

                var zam3 = txtzam3.Text;
                var zam3No = txtzam1.Text;
                var zam3unvan = txtzamunvan3.Text;

                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;

                //var icraci = icraci_adi;
                //var icraciadi = icraciadizaminlikde;



                // TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\İddia_uczamin.doc");
                //var wordDocument = wordapp.Documents.Open(zamin1);
                ReplaceWordStub("{mekNo}", mekno, wordDocument);
                ReplaceWordStub("{mekNo1}", mekno1, wordDocument);
                ReplaceWordStub("{krtarix}", caritar, wordDocument);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                //ReplaceWordStub("{CariHesab}", tamcari, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                //ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{mehkeme}", mehkeme, wordDocument);
                ReplaceWordStub("{yetirmeli}", odenmeli, wordDocument);
                ReplaceWordStub("{yetirmis}", odenilmis, wordDocument);
                ReplaceWordStub("{topborc}", tamborc, wordDocument);
                ReplaceWordStub("{vkborc}", vkborc, wordDocument);
                ReplaceWordStub("{cedfaiz}", cedfaiz, wordDocument);
                ReplaceWordStub("{topcedmebleg}", topcedmeb, wordDocument);
                ReplaceWordStub("{vkfaizmeb}", vkfaizborc, wordDocument);
                ReplaceWordStub("{rusum}", rusummeb, wordDocument);
                ReplaceWordStub("{ayesas}", wayesas, wordDocument);
                ReplaceWordStub("{ayfaiz}", wayfaiz, wordDocument);
                ReplaceWordStub("{aycemi}", waytop, wordDocument);
                ReplaceWordStub("{snod}", sontarix, wordDocument);
                ReplaceWordStub("{snodmeb}", sonmebleg, wordDocument);

                ReplaceWordStub("{cavesas}", caves, wordDocument);
                ReplaceWordStub("{cavfaiz}", cavfaiz, wordDocument);
                ReplaceWordStub("{cavcem}", cavcem, wordDocument);

                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                //ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                //ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg + "(" + mebyaziile + ")", wordDocument);
                ReplaceWordStub("{muddet}", muddet + " ay " , wordDocument);
                ReplaceWordStub("{faiz}", faiz + "% " , wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz + "%" , wordDocument);
                // ReplaceWordStub("{tel}", telf, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq + VAL_AD + "(" + mevayyazi + ")", wordDocument);

                //ReplaceWordStub("{icraci}", icraciadi, wordDocument);

                ReplaceWordStub("{zam1ad}", zam1, wordDocument);
                ReplaceWordStub("{zam1unvan}", zam1unvan, wordDocument);
                ReplaceWordStub("{zammuq1}", zam1No, wordDocument);

                ReplaceWordStub("{zam2ad}", zam2, wordDocument);


                ReplaceWordStub("{zam3ad}", zam3, wordDocument);


                wordapp.Visible = true;
                string muqadi = MuqNo + " " + adi + " " + tamtarix;
                //string muqadi = "yoxlanis";
                //string muqadi = adi;
                // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                //Directory.CreateDirectory("\\fs\\12345\\Personal\\Anar_Is\\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\\Məhkəmə sənədlərii");
                //   string folderInfo = folderBrowserDialog1.SelectedPath + "\\fs\\12345\\Personal\\Anar_Is\\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\\Məhkəmə sənədlərii" + txtFolderName.Text;

                // Directory.CreateDirectory(folderInfo);
                string testt = @"\\192.168.0.5\12345\Personal\Anar_Is\2012 Bas ofis hesabatlari ve Yeni prudensial (kredit uzre)\Məhkəmə sənədlərii\" + txtadi.Text + "\\";

                wordDocument.SaveAs(testt + muqadi + ".doc");
                //wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".doc");


            }
            catch (Exception)
            {


            }
            finally { }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document WordDocument)
        {
            var range = WordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }
        private void rusum()
        {
            double rusum=0.00;
            double topborc = Convert.ToDouble(txttopborc.Text);
            if (topborc<=1000)
            {
                rusum = 30;
            }
            else if (topborc > 1000 && topborc <= 10000)
            {
                rusum = ((topborc - 1000) * 1/100)+30;
            }
            else if (topborc > 10000 && topborc <= 100000)
            {
                rusum = ((topborc - 10000) * 0.3 / 100) + 120;
            }
            txtrusum.Text = Math.Round( rusum,2).ToString();
        }
        private void meknoal()
        {
            string gun, ay, il;
            dateiddiatar.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            textBox2.Text = gun + "-" + ay + "-" + il;
            int test = 45;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();
            string gonyer = cmbrayonıar.Text + "məhkəməsi ";
            string mezm = " İddia ərizəsi";
            string tammezmun = txtadi.Text + mezm;
            //'avto gir azad Cars'  
            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('"+gonyer+"', TO_DATE('" + dateiddiatar.Text + "','dd-MM-yyyy'), '" + tammezmun + "'   , '" + icraci_kod + "',  " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();
            
        }
        private void meknoalmaliye()
        {
            string gun, ay, il;
            dateiddiatar.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            textBox2.Text = gun + "-" + ay + "-" + il;
            int test = 45;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();
            string gonyer = cmbrayonıar.Text + "məhkəməsi ";
            string mezm = " Maliyyə arayışı";
            string tammezmun = txtadi.Text + mezm;
            //'avto gir azad Cars'  
            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('" + gonyer + "', TO_DATE('" + dateiddiatar.Text + "','dd-MM-yyyy'), '" + tammezmun + "'   , '" + icraci_kod + "',  " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();

        }
        private void muracietsayartimsiz()
        {
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                string proid;
                string hevsecal;
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
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void muracietsayartimsizmaliyyə()
        {
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                string proid;
                string hevsecal;
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
                mek_nomaliyye = tarixIl + "-" + mek_no_arti.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void ayitap()
        {
            DateTime krtarixi = Convert.ToDateTime(dateTimePicker1.Text);
            DateTime iddiatarixi = Convert.ToDateTime(dateiddiatar.Text);
            int kryil = krtarixi.Year;
            int idyil = iddiatarixi.Year;
            int kray = krtarixi.Month;
            int iday = iddiatarixi.Month;
            int iller = idyil-kryil;
            int haziray;
            if (idyil>kryil)
            {
                haziray = 0;
                haziray = iday - kray+12;
                //int duzil = iller * 12;
                //int duzilkr, duzilid;
                //duzilkr = duzil + kray;
                //duzilid = duzil + iday;
                // haziray = duzilid - duzilkr;

            }
            else
            {
                haziray = 0;
                haziray = iday - kray;
            }
            
            label1.Text = haziray.ToString();
            //double topedeme = Convert.ToDouble(haziray) * Convert.ToDouble(txtayliq.Text);
            
            double faizler = (Convert.ToDouble(txtmuddet.Text) * Convert.ToDouble(txtayliq.Text));
            //txtfaizmebleg.Text = (faizler - Convert.ToDouble(txtkrmeb.Text)).ToString();
            //txttxttoplamcedvel.Text = (Convert.ToDouble(txtfaizmebleg.Text) + Convert.ToDouble(txtkrmeb.Text)).ToString();

            //TimeSpan fark = Convert.ToDateTime(dateTimePicker1.Text) - Convert.ToDateTime(dateTimePicker2.Text);

            //int fark1 = Convert.ToInt32(fark.TotalDays);

            //if ((fark1 * (-1)) == fark1) fark1 -= 1;

            //else fark1 += 1;
            //label1.Text = (fark1 / 30).ToString();
            //label1.Text = "Ay : " + fark1 / 30 + "\nGün : " + fark1.ToString() + "\nSaat : " + (fark1 * 60).ToString() + "\nDakika : " + (fark1 * 60) * 60;
        }
        private void kataloqgetir_zamin_ayliq()
        {//t.subschkre t.licschkre
            try
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon, r.name_regnom," +
                    " t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre," +
                    "t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq," +
                    " r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari from odb.licschkre t," +
                    "odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k " +
                    "where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom " +
                    "and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  " +
                    "and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;
                Orcon.Close();
                dataGridView2.Columns[0].HeaderText = "Zamin adı";
                dataGridView2.Columns[0].Width = 300;

                dataGridView2.Columns[1].HeaderText = "Zamin seriya No";
                dataGridView2.Columns[1].Width = 200;

                dataGridView2.Columns[2].HeaderText = "Zamin Ünvan";
                dataGridView2.Columns[2].Width = 200;

                dataGridView2.Columns[3].HeaderText = "Zamin telefon";
                dataGridView2.Columns[3].Width = 200;
            }
            catch (Exception)
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);

                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text+ "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;
                Orcon.Close();
                dataGridView2.Columns[0].HeaderText = "Zamin adı";
                dataGridView2.Columns[0].Width = 300;

                dataGridView2.Columns[1].HeaderText = "Zamin seriya No";
                dataGridView2.Columns[1].Width = 200;

                dataGridView2.Columns[2].HeaderText = "Zamin Ünvan";
                dataGridView2.Columns[2].Width = 200;

                dataGridView2.Columns[3].HeaderText = "Zamin telefon";
                dataGridView2.Columns[3].Width = 200;
            }

            finally
            { }


        }
        
        private void kataloqgetir_()
        {//t.subschkre t.licschkre
            try
            {
				//select  z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz , sum(z3.oden) oden_um from(select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20  and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre=&r1 and t.subschkre=&t2 and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk  order by substr(z.kredit,10,6)   ) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6) ) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar --and z1.sk=z3.sk group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar
				//dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
                //OracleCommand Orcom2 = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "'  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,(select x.pogprospro_ish, x.nacpro_ish-x.pogpro_ish  from view_nacpogprokre_all x where t.licschpkre=x.licschppkre and t.subschkre=x.subschkre) test_faiz,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                OracleCommand Orcom = new OracleCommand("select distinct z1.*,z2.*from(select r.name_regnom, t.subschkre sk, substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open, t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre + g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz,t.procstav_19, t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq,r.senedin_verilme_tarixi ver_tar,r.telefon, r.registrac,r.grajdanstvo,t.licschkre kre,t.licschpkre pkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az || m.nr || m.bank || m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r, odb.licsch m, odb.srokpogprockre k where substr(t.licschkre, 10, 6)= '"+ txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre and t.subschkre = k.subschkre and m.licsch = k.licsch_3(+) and t.date_close is null) z1,(select x.date_oper,x.licschpkre, x.subschkre, x.nacprospro_ish - x.pogprospro_ish , x.nacpro_ish - x.pogpro_ish ferq from view_nacpogprokre_all x where x.date_oper = to_date(sysdate)) z2  where z1.pkre = z2.licschpkre and z1.sk = z2.subschkre ", Orcon);

                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Adı";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "KS";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Sub-qeyd";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Hesab No";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "ver.tarixi";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Təyinat";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Məbləğ";
                dataGridView1.Columns[6].Width = 120;

                dataGridView1.Columns[7].HeaderText = "Məbləğ AZN";
                dataGridView1.Columns[7].Width = 120;

                dataGridView1.Columns[8].HeaderText = "Aylıq";
                dataGridView1.Columns[8].Width = 100;

                dataGridView1.Columns[9].HeaderText = "Fifd";
                dataGridView1.Columns[9].Width = 100;

                dataGridView1.Columns[10].HeaderText = "Faiz";
                dataGridView1.Columns[10].Width = 100;

                dataGridView1.Columns[11].HeaderText = "V/K %";
                dataGridView1.Columns[11].Width = 110;

                dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
                dataGridView1.Columns[12].Width = 120;

                dataGridView1.Columns[13].HeaderText = "Müddət";
                dataGridView1.Columns[13].Width = 120;

                dataGridView1.Columns[14].HeaderText = "Seriya No";
                dataGridView1.Columns[14].Width = 120;

                dataGridView1.Columns[15].HeaderText = "Verən orqan";
                dataGridView1.Columns[15].Width = 200;

                dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
                dataGridView1.Columns[16].Width = 300;

                dataGridView1.Columns[17].HeaderText = "Mobil";
                dataGridView1.Columns[17].Width = 300;

                dataGridView1.Columns[18].HeaderText = "Ünvan";
                dataGridView1.Columns[18].Width = 100;

                dataGridView1.Columns[19].HeaderText = "Cari hesab";
                dataGridView1.Columns[19].Width = 100;

                dataGridView1.Columns[33].HeaderText = "Faiz";
                dataGridView1.Columns[33].Width = 300;

                dataGridView1.Columns[34].HeaderText = "VK faiz";
                dataGridView1.Columns[34].Width = 200;

                //dataGridView1.Columns[22].HeaderText = "Zamin Ünvan";
                //dataGridView1.Columns[22].Width = 200;

                //dataGridView1.Columns[23].HeaderText = "Zamin telefon";
                //dataGridView1.Columns[23].Width = 200;

                //dataGridView1.Columns[30].HeaderText = "tam hesab";
                //dataGridView1.Columns[30].Width = 200;
            }
            catch (Exception)
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,(select x.pogprospro_ish, x.nacpro_ish-x.pogpro_ish  from view_nacpogprokre_all x where t.licschpkre=x.licschppkre and t.subschkre=x.subschkre) test_faiz,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);

                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga ,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);

                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text+ "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Adı";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "KS";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Sub-qeyd";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Hesab No";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "ver.tarixi";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Təyinat";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Məbləğ";
                dataGridView1.Columns[6].Width = 120;

                dataGridView1.Columns[7].HeaderText = "Məbləğ AZN";
                dataGridView1.Columns[7].Width = 120;

                dataGridView1.Columns[8].HeaderText = "Aylıq";
                dataGridView1.Columns[8].Width = 100;

                dataGridView1.Columns[9].HeaderText = "Fifd";
                dataGridView1.Columns[9].Width = 100;

                dataGridView1.Columns[10].HeaderText = "Faiz";
                dataGridView1.Columns[10].Width = 100;

                dataGridView1.Columns[11].HeaderText = "V/K %";
                dataGridView1.Columns[11].Width = 110;

                dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
                dataGridView1.Columns[12].Width = 120;

                dataGridView1.Columns[13].HeaderText = "Müddət";
                dataGridView1.Columns[13].Width = 120;

                dataGridView1.Columns[14].HeaderText = "Seriya No";
                dataGridView1.Columns[14].Width = 120;

                dataGridView1.Columns[15].HeaderText = "Verən orqan";
                dataGridView1.Columns[15].Width = 200;

                dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
                dataGridView1.Columns[16].Width = 300;

                dataGridView1.Columns[17].HeaderText = "Mobil";
                dataGridView1.Columns[17].Width = 300;

                dataGridView1.Columns[18].HeaderText = "Ünvan";
                dataGridView1.Columns[18].Width = 100;

                dataGridView1.Columns[19].HeaderText = "Cari hesab";
                dataGridView1.Columns[19].Width = 100;

                //dataGridView1.Columns[20].HeaderText = "Zamin adı";
                //dataGridView1.Columns[20].Width = 300;

                //dataGridView1.Columns[21].HeaderText = "Zamin seriya No";
                //dataGridView1.Columns[21].Width = 200;

                //dataGridView1.Columns[22].HeaderText = "Zamin Ünvan";
                //dataGridView1.Columns[22].Width = 200;

                //dataGridView1.Columns[23].HeaderText = "Zamin telefon";
                //dataGridView1.Columns[23].Width = 200;

            }

            finally
            { }


        }
        private void kataloqgetir_cedveluzre()
        {//t.subschkre t.licschkre
            //try
            //{
                //select  z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz , sum(z3.oden) oden_um from(select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20  and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre=&r1 and t.subschkre=&t2 and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk  order by substr(z.kredit,10,6)   ) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6) ) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar --and z1.sk=z3.sk group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
                //OracleCommand Orcom2 = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "'  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,(select x.pogprospro_ish, x.nacpro_ish-x.pogpro_ish  from view_nacpogprokre_all x where t.licschpkre=x.licschppkre and t.subschkre=x.subschkre) test_faiz,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                //OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre ='"+txtsuda.Text+"' and t.subschkre ='"+ txtsubkod.Text + "' and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);
				OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre ='" + txtsuda.Text + "' and t.subschkre ='" + txtsubkod.Text + "' and k.licschkre=t.licschkre(+) and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn, max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6)) z4 where z1.qn=z2.qn(+)  and z1.sk=z2.sk(+) and z1.qn=z3.qn(+)  and z1.qn=z4.qn(+) and z4.son_tar=z3.son_tar(+)group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView3.DataSource = Ordt;
                Orcon.Close();
                //dataGridView3.Columns[0].HeaderText = "Adı";
                //dataGridView3.Columns[0].Width = 350;

                //dataGridView3.Columns[1].HeaderText = "KS";
                //dataGridView3.Columns[1].Width = 55;

                //dataGridView3.Columns[2].HeaderText = "Sub-qeyd";
                //dataGridView3.Columns[2].Width = 55;

                //dataGridView3.Columns[3].HeaderText = "Hesab No";
                //dataGridView3.Columns[3].Width = 200;

                //dataGridView3.Columns[4].HeaderText = "ver.tarixi";
                //dataGridView3.Columns[4].Width = 120;

                //dataGridView3.Columns[5].HeaderText = "Təyinat";
                //dataGridView3.Columns[5].Width = 120;

                //dataGridView3.Columns[6].HeaderText = "Məbləğ";
                //dataGridView3.Columns[6].Width = 120;

                //dataGridView3.Columns[7].HeaderText = "Məbləğ AZN";
                //dataGridView3.Columns[7].Width = 120;

                //dataGridView3.Columns[8].HeaderText = "Aylıq";
                //dataGridView3.Columns[8].Width = 100;

                //dataGridView3.Columns[9].HeaderText = "Fifd";
                //dataGridView3.Columns[9].Width = 100;

                //dataGridView3.Columns[10].HeaderText = "Faiz";
                //dataGridView3.Columns[10].Width = 100;

                //dataGridView3.Columns[11].HeaderText = "V/K %";
                //dataGridView3.Columns[11].Width = 110;

                //dataGridView3.Columns[12].HeaderText = "Ehtiyyat %";
                //dataGridView3.Columns[12].Width = 120;

                //dataGridView3.Columns[13].HeaderText = "Müddət";
                //dataGridView3.Columns[13].Width = 120;

                //dataGridView3.Columns[14].HeaderText = "Seriya No";
                //dataGridView3.Columns[14].Width = 120;

                //dataGridView3.Columns[15].HeaderText = "Verən orqan";
                //dataGridView3.Columns[15].Width = 200;

                //dataGridView3.Columns[16].HeaderText = "Verilmə tarixi";
                //dataGridView3.Columns[16].Width = 300;

                //dataGridView3.Columns[17].HeaderText = "Mobil";
                //dataGridView3.Columns[17].Width = 300;

                //dataGridView3.Columns[18].HeaderText = "Ünvan";
                //dataGridView3.Columns[18].Width = 100;

                //dataGridView3.Columns[19].HeaderText = "Cari hesab";
                //dataGridView3.Columns[19].Width = 100;

                //dataGridView3.Columns[33].HeaderText = "Faiz";
                //dataGridView3.Columns[33].Width = 300;

                //dataGridView3.Columns[34].HeaderText = "VK faiz";
                //dataGridView3.Columns[34].Width = 200;

                //dataGridView1.Columns[22].HeaderText = "Zamin Ünvan";
                //dataGridView1.Columns[22].Width = 200;

                //dataGridView1.Columns[23].HeaderText = "Zamin telefon";
                //dataGridView1.Columns[23].Width = 200;

                //dataGridView1.Columns[30].HeaderText = "tam hesab";
                //dataGridView1.Columns[30].Width = 200;
            
            
        }
        private void testvklar()
        {//t.subschkre t.licschkre
         //try
         //{
         //select  z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz , sum(z3.oden) oden_um from(select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20  and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre=&r1 and t.subschkre=&t2 and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk  order by substr(z.kredit,10,6)   ) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6) ) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar --and z1.sk=z3.sk group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar
         //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            OracleCommand Orcom = new OracleCommand("select t.licsch_19,t.licschppkre from odb.licschkre t where  t.licsch_19>0 and t.date_close is null  ", Orcon);
            //OracleCommand Orcom2 = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "'  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,(select x.pogprospro_ish, x.nacpro_ish-x.pogpro_ish  from view_nacpogprokre_all x where t.licschpkre=x.licschppkre and t.subschkre=x.subschkre) test_faiz,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
            //OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre ='"+txtsuda.Text+"' and t.subschkre ='"+ txtsubkod.Text + "' and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);
            //OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre ='" + txtsuda.Text + "' and t.subschkre ='" + txtsubkod.Text + "' and k.licschkre=t.licschkre(+) and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn, max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6)) z4 where z1.qn=z2.qn(+)  and z1.sk=z2.sk(+) and z1.qn=z3.qn(+)  and z1.qn=z4.qn(+) and z4.son_tar=z3.son_tar(+)group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();
            //dataGridView3.Columns[0].HeaderText = "Adı";
            //dataGridView3.Columns[0].Width = 350;

            //dataGridView3.Columns[1].HeaderText = "KS";
            //dataGridView3.Columns[1].Width = 55;

            //dataGridView3.Columns[2].HeaderText = "Sub-qeyd";
            //dataGridView3.Columns[2].Width = 55;

            //dataGridView3.Columns[3].HeaderText = "Hesab No";
            //dataGridView3.Columns[3].Width = 200;

            //dataGridView3.Columns[4].HeaderText = "ver.tarixi";
            //dataGridView3.Columns[4].Width = 120;

            //dataGridView3.Columns[5].HeaderText = "Təyinat";
            //dataGridView3.Columns[5].Width = 120;

            //dataGridView3.Columns[6].HeaderText = "Məbləğ";
            //dataGridView3.Columns[6].Width = 120;

            //dataGridView3.Columns[7].HeaderText = "Məbləğ AZN";
            //dataGridView3.Columns[7].Width = 120;

            //dataGridView3.Columns[8].HeaderText = "Aylıq";
            //dataGridView3.Columns[8].Width = 100;

            //dataGridView3.Columns[9].HeaderText = "Fifd";
            //dataGridView3.Columns[9].Width = 100;

            //dataGridView3.Columns[10].HeaderText = "Faiz";
            //dataGridView3.Columns[10].Width = 100;

            //dataGridView3.Columns[11].HeaderText = "V/K %";
            //dataGridView3.Columns[11].Width = 110;

            //dataGridView3.Columns[12].HeaderText = "Ehtiyyat %";
            //dataGridView3.Columns[12].Width = 120;

            //dataGridView3.Columns[13].HeaderText = "Müddət";
            //dataGridView3.Columns[13].Width = 120;

            //dataGridView3.Columns[14].HeaderText = "Seriya No";
            //dataGridView3.Columns[14].Width = 120;

            //dataGridView3.Columns[15].HeaderText = "Verən orqan";
            //dataGridView3.Columns[15].Width = 200;

            //dataGridView3.Columns[16].HeaderText = "Verilmə tarixi";
            //dataGridView3.Columns[16].Width = 300;

            //dataGridView3.Columns[17].HeaderText = "Mobil";
            //dataGridView3.Columns[17].Width = 300;

            //dataGridView3.Columns[18].HeaderText = "Ünvan";
            //dataGridView3.Columns[18].Width = 100;

            //dataGridView3.Columns[19].HeaderText = "Cari hesab";
            //dataGridView3.Columns[19].Width = 100;

            //dataGridView3.Columns[33].HeaderText = "Faiz";
            //dataGridView3.Columns[33].Width = 300;

            //dataGridView3.Columns[34].HeaderText = "VK faiz";
            //dataGridView3.Columns[34].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "Zamin Ünvan";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[23].HeaderText = "Zamin telefon";
            //dataGridView1.Columns[23].Width = 200;

            //dataGridView1.Columns[30].HeaderText = "tam hesab";
            //dataGridView1.Columns[30].Width = 200;


        }
        private void kataloqgetir_cedveluzre_topfaiz()
        {//t.subschkre t.licschkre
         //try
         //{
         //select  z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz , sum(z3.oden) oden_um from(select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20  and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre=&r1 and t.subschkre=&t2 and t.licschkre=k.licschkre and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk  order by substr(z.kredit,10,6)   ) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6) ) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn  and z1.sk=z2.sk and z1.qn=z3.qn  and z1.qn=z4.qn and z3.son_tar=z4.son_tar --and z1.sk=z3.sk group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar
         //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
            //OracleCommand Orcom2 = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "'  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, m.az||m.nr||m.bank||m.licsch cari,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon from odb.licschkre t,odb.creditinfoguarantee g, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,(select x.pogprospro_ish, x.nacpro_ish-x.pogpro_ish  from view_nacpogprokre_all x where t.licschpkre=x.licschppkre and t.subschkre=x.subschkre) test_faiz,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
                                                                    OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and  t.licschkre ='" + txtsuda.Text + "' and t.subschkre ='" + txtsubkod.Text + "' and t.licschkre=k.licschkre(+) and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn,max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6) ) z4 where z1.qn=z2.qn(+)  and z1.sk=z2.sk(+) and z1.qn=z3.qn(+)  and z1.qn=z4.qn(+) and z3.son_tar=z4.son_tar(+) group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);
//OracleCommand Orcom = new OracleCommand("select z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,sum(z2.faiz) oden_fz ,sum(z3.oden) oden_um,z4.son_tar from (select substr(t.licschkre,10,6) qn,t.licschkre,t.subschkre sk,k.licschpkre,k.licschppkre,sum(t.summa_pog_kre) ced_esas,sum(t.summa_pog_pro) ced_faiz, sum(t.summa_pog_kre+t.summa_pog_pro) ayliq from odb.graphpogkre t,odb.licschkre k where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and t.licschkre ='" + txtsuda.Text + "' and t.subschkre ='" + txtsubkod.Text + "' and k.licschkre=t.licschkre(+) and k.date_close is null group by substr(t.licschkre,10,6),t.licschkre,t.subschkre,k.licschpkre,k.licschppkre) z1,(select substr(z.kredit,10,6) qn, z.ssk sk, sum(z.summa_v_nacval) faiz from odb.arh_dd z, odb.licschkre kr where z.kredit in ( kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.ssk order by substr(z.kredit,10,6)) z2,(select substr(z.kredit,10,6) qn, sum(z.summa_v_nacval) oden,z.date_oper son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6),z.date_oper order by substr(z.kredit,10,6)) z3,(select substr(z.kredit,10,6) qn, max(z.date_oper) son_tar from odb.arh_dd z, odb.licschkre kr where z.kredit in (kr.licschkre,kr.licsch_19, kr.licschppkre,kr.licschpkre) and substr(z.debet,1,1)='4' and z.ssk=kr.subschkre  and kr.date_close is null group by substr(z.kredit,10,6) order by substr(z.kredit,10,6)) z4 where z1.qn=z2.qn(+)  and z1.sk=z2.sk(+) and z1.qn=z3.qn(+)  and z1.qn=z4.qn(+) and z4.son_tar=z3.son_tar(+)group by z1.licschkre,z1.sk,z1.ced_esas,z1.ced_faiz,z3.oden,z3.son_tar,z4.son_tar", Orcon);

            //to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') and
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa_19 from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where substr(t.licschkre, 10, 6)='" + txbqedno.Text + "' and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and t.date_close is null", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView4.DataSource = Ordt;
            Orcon.Close();
            //dataGridView3.Columns[0].HeaderText = "Adı";
            //dataGridView3.Columns[0].Width = 350;

            //dataGridView3.Columns[1].HeaderText = "KS";
            //dataGridView3.Columns[1].Width = 55;

            //dataGridView3.Columns[2].HeaderText = "Sub-qeyd";
            //dataGridView3.Columns[2].Width = 55;

            //dataGridView3.Columns[3].HeaderText = "Hesab No";
            //dataGridView3.Columns[3].Width = 200;

            //dataGridView3.Columns[4].HeaderText = "ver.tarixi";
            //dataGridView3.Columns[4].Width = 120;

            //dataGridView3.Columns[5].HeaderText = "Təyinat";
            //dataGridView3.Columns[5].Width = 120;

            //dataGridView3.Columns[6].HeaderText = "Məbləğ";
            //dataGridView3.Columns[6].Width = 120;

            //dataGridView3.Columns[7].HeaderText = "Məbləğ AZN";
            //dataGridView3.Columns[7].Width = 120;

            //dataGridView3.Columns[8].HeaderText = "Aylıq";
            //dataGridView3.Columns[8].Width = 100;

            //dataGridView3.Columns[9].HeaderText = "Fifd";
            //dataGridView3.Columns[9].Width = 100;

            //dataGridView3.Columns[10].HeaderText = "Faiz";
            //dataGridView3.Columns[10].Width = 100;

            //dataGridView3.Columns[11].HeaderText = "V/K %";
            //dataGridView3.Columns[11].Width = 110;

            //dataGridView3.Columns[12].HeaderText = "Ehtiyyat %";
            //dataGridView3.Columns[12].Width = 120;

            //dataGridView3.Columns[13].HeaderText = "Müddət";
            //dataGridView3.Columns[13].Width = 120;

            //dataGridView3.Columns[14].HeaderText = "Seriya No";
            //dataGridView3.Columns[14].Width = 120;

            //dataGridView3.Columns[15].HeaderText = "Verən orqan";
            //dataGridView3.Columns[15].Width = 200;

            //dataGridView3.Columns[16].HeaderText = "Verilmə tarixi";
            //dataGridView3.Columns[16].Width = 300;

            //dataGridView3.Columns[17].HeaderText = "Mobil";
            //dataGridView3.Columns[17].Width = 300;

            //dataGridView3.Columns[18].HeaderText = "Ünvan";
            //dataGridView3.Columns[18].Width = 100;

            //dataGridView3.Columns[19].HeaderText = "Cari hesab";
            //dataGridView3.Columns[19].Width = 100;

            //dataGridView3.Columns[33].HeaderText = "Faiz";
            //dataGridView3.Columns[33].Width = 300;

            //dataGridView3.Columns[34].HeaderText = "VK faiz";
            //dataGridView3.Columns[34].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "Zamin Ünvan";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[23].HeaderText = "Zamin telefon";
            //dataGridView1.Columns[23].Width = 200;

            //dataGridView1.Columns[30].HeaderText = "tam hesab";
            //dataGridView1.Columns[30].Width = 200;


        }
        DateTime dt = DateTime.Now;

        string tarix_soz;
        string tarix_sozilkintarix;
        string tarix_sozcaritarix;
        public void tarixitapsoz()
        {
            int ay = dt.Month;
            //textBox1.Text = dt.Day.ToString();
            //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
            //int ay = 6;

            if (ay == 1)
            {
                tarix_soz = dt.Day.ToString() + " Yanvar " + dt.Year.ToString();
            }
            else if (ay == 2)
            {
                tarix_soz = dt.Day.ToString() + " Fevral " + dt.Year.ToString();
            }
            else if (ay == 3)
            {
                tarix_soz = dt.Day.ToString() + " Mart " + dt.Year.ToString();
            }
            else if (ay == 4)
            {
                tarix_soz = dt.Day.ToString() + " Aprel " + dt.Year.ToString();
            }
            else if (ay == 5)
            {
                tarix_soz = dt.Day.ToString() + " May " + dt.Year.ToString();
            }
            else if (ay == 6)
            {
                tarix_soz = dt.Day.ToString() + " Iyun " + dt.Year.ToString();
            }
            else if (ay == 7)
            {
                tarix_soz = dt.Day.ToString() + " Iyul " + dt.Year.ToString();
            }
            else if (ay == 8)
            {
                tarix_soz = dt.Day.ToString() + " Avqust " + dt.Year.ToString();
            }
            else if (ay == 9)
            {
                tarix_soz = dt.Day.ToString() + " Sentyabr " + dt.Year.ToString();
            }
            else if (ay == 10)
            {
                tarix_soz = dt.Day.ToString() + " Oktyabr " + dt.Year.ToString();
            }
            else if (ay == 11)
            {
                tarix_soz = dt.Day.ToString() + " Noyabr " + dt.Year.ToString();
            }
            else if (ay == 12)
            {
                tarix_soz = dt.Day.ToString() + " Dekabr " + dt.Year.ToString();
            }
        }
        public void tarixitapsozilkintarix()
        {
            int ay = Convert.ToInt16(dateTimePicker1.Text.Substring(3, 2));
            //textBox1.Text = dt.Day.ToString();
            //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
            //int ay = 6;

            if (ay == 1)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Yanvar " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 2)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Fevral " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 3)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Mart " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 4)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Aprel " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 5)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " May " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 6)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " İyun " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 7)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " İyul " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 8)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Avqust " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 9)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Sentyabr " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 10)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Oktyabr " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 11)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Noyabr " + dateTimePicker1.Text.Substring(6, 4);
            }
            else if (ay == 12)
            {
                tarix_sozilkintarix = dateTimePicker1.Text.Substring(0, 2) + " Dekabr " + dateTimePicker1.Text.Substring(6, 4);
            }
        }
        public void tarixitapsozcaritarix()
        {
            int ay = Convert.ToInt16(dateiddiatar.Text.Substring(3, 2));
            //textBox1.Text = dt.Day.ToString();
            //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
            //int ay = 6;

            if (ay == 1)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Yanvar " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 2)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Fevral " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 3)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Mart " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 4)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Aprel " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 5)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " May " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 6)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " İyun " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 7)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " İyul " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 8)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Avqust " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 9)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Sentyabr " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 10)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Oktyabr " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 11)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Noyabr " + dateiddiatar.Text.Substring(6, 4);
            }
            else if (ay == 12)
            {
                tarix_sozcaritarix = dateiddiatar.Text.Substring(0, 2) + " Dekabr " + dateiddiatar.Text.Substring(6, 4);
            }
        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {
           //testvklar();
            kataloqgetir_();
            kataloqgetir_zamin_ayliq();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            doldur();
            ayitap();
            rusum();
            kataloqgetir_cedveluzre();
            kataloqgetir_cedveluzre_topfaiz();
            decimal odfaiz;
            decimal odesas=Convert.ToDecimal(txtodenilenesas.Text);
            if (dataGridView3.CurrentRow.Cells[4].Value.ToString()=="")
            {
                odfaiz = 0;
            }
            else
            {
                odfaiz=Convert.ToDecimal(dataGridView3.CurrentRow.Cells[4].Value.ToString());

            }
            decimal topod = odesas + odfaiz;
            txtodcem.Text = Decimal.Round(topod, 2).ToString();
            txtodenilmis.Text = txtodcem.Text;
            decimal cedesas = Convert.ToDecimal(dataGridView3.CurrentRow.Cells[2].Value.ToString());
            decimal cedfaiz = Convert.ToDecimal(dataGridView3.CurrentRow.Cells[3].Value.ToString());
            decimal ceduzreodenilmelifaiz = Convert.ToDecimal(dataGridView4.CurrentRow.Cells[3].Value.ToString());
            decimal topayced = cedesas + cedfaiz;
            txtfaizmebleg.Text = Decimal.Round(ceduzreodenilmelifaiz, 2).ToString();
            decimal toplamcedvel = Convert.ToDecimal(textEdit17.Text) + ceduzreodenilmelifaiz;
            txttxttoplamcedvel.Text = Decimal.Round(toplamcedvel, 2).ToString();
            txtayesas.Text = Decimal.Round(cedesas, 2).ToString();
            txtayfaiz.Text = Decimal.Round(cedfaiz, 2).ToString();
            txtaytoplam.Text = Decimal.Round(topayced, 2).ToString();
            txtodenilmeli.Text = Decimal.Round(topayced, 2).ToString();
            if (dataGridView3.CurrentRow.Cells[4].Value.ToString()=="")
            {
                txtodfaiz.Text = "0.00";
            }
            else
            {

            
            txtodfaiz.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
            }
            string testtarix = dataGridView3.CurrentRow.Cells[6].Value.ToString();
            //if (testtarix == "")
            //{
            //    txtpastar.Text = "";
            //}

            //else
            //{
            //    txtsontarix.Text = dataGridView3.CurrentRow.Cells[6].Value.ToString().Substring(0, 10);
            //}
            if (dataGridView3.CurrentRow.Cells[5].Value.ToString()=="")
            {
                txtsonmebleg.Text = "0.00";
            }
            else
            {

            
            txtsonmebleg.Text = dataGridView3.CurrentRow.Cells[5].Value.ToString();
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            KlasorOlusturmaIslemi();

            meknoal();
            meknoalmaliye();
            muracietsayartimsiz();
            muracietsayartimsizmaliyyə();
            if (comboBoxEdit1.Text== "Bir nəfərin zəmanəti")
            {
                word_at();
            }
           else if (comboBoxEdit1.Text == "İki nəfərin zəmanəti")
            {
                word_at_ikizamin();
            }
           else if (comboBoxEdit1.Text == "Üç nəfərin zəmanəti")
            {
                word_at_uczamin();
            }

        }

        private void txtkrmeb_EditValueChanged(object sender, EventArgs e)
        {

        }
    }
}
