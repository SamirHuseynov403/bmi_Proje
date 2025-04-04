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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class Qizil : Form
    {
        
        public Qizil()
        {
            InitializeComponent();
        }
        public string girovnovucombo { get; set; }
        public string muqtipi { get; set; }
        public string umumitar { get; set; }
        public string test { get; set; }
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
        public string girovnovucomboqizil = string.Empty;
        public string qzlgirovu { get; set; }

        public string secilmistarix { get; set; }

        public string icraciadizaminlikde { get; set; }
        public string icraci_adi { get; set; }

        //public string icraci_adi { get; set; }


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

        public static string gedenbilgi = "", kr1 = "", kr2 = "", kr3 = "", kr4 = "", kr5 = "", kr6 = "", kr7 = "";
        public static string db1 = "", db2 = "", db3 = "", db4 = "", db5 = "", db6 = "", db7 = "", c1 = "", sub1 = "";
        public static string meb1 = "", meb2 = "", meb3 = "", meb4 = "", meb5 = "", meb6 = "", meb7 = "";
        public static string teyinat1 = "", teyinat2 = "", teyinat3 = "", teyinat4 = "", teyinat5 = "", teyinat6 = "", teyinat7 = "";

        
        

        

        private void olkeadi()
        {
            if (olke == "AZ")
            {
                cboxqizilOlke.Text = "Azərbaycan Respublikası";
            }
            else if (olke == "IRN")
            {
                cboxqizilOlke.Text = " İran İslam Respublikası";
            }


        }

        

        private void zamincevirme()
        {
            string kassa_hesazn, kassa_hesusd, kassa_hesavro, x_h_kassaazn, x_h_kassausd, x_h_kassaavro, x_h_kreditazn, x_h_kreditusd, x_h_kreditavro;
            string x_h_kredit_azn_QR, x_h_kredit_usd_QR, x_h_kredit_avro_QR;

            kassa_hesazn = "10010000000000100000";
            kassa_hesusd = "10020010000000100000";
            kassa_hesavro = "10020020000000100000";

            x_h_kassaazn = "67010000000000600000";
            x_h_kassausd = "67020010000000600000";
            x_h_kassaavro = "67020020000000600000";

            x_h_kreditazn = "67034000010000600000";
            x_h_kreditusd = "67044010010000600000";
            x_h_kreditavro = "67044020010000600000";

            x_h_kredit_azn_QR = "67020010000000600000";
            x_h_kredit_usd_QR = "67044010050000600000";
            x_h_kredit_avro_QR = "67044020020000600000";

            teyinat1 = DateTime.Now.ToShortDateString() + " il tarixli kredit müq.əsəasən " + txbQizilMuddet.Text + " ay müddətinə kredit verilir";
            teyinat2 = DateTime.Now.ToShortDateString() + " tar.kr.müq.əs 1% x/h tut ";
            teyinat3 = "Kredit verilməsi ilə əlaqədar";
            teyinat4 = "Kassa 0,5 % x/h tutulur";
            teyinat5 = "kred.ver.əlaqədar 0,5 % x/h";

            //c1=;
            sub1 = subhes;
            db1 = sudahes;
            db2 = kassa_hesazn;
            db3 = carihes;
            db4 = carihes;
            db5 = carihes;

            kr1 = carihes;
            kr2 = carihes;
            kr3 = kassa_hesazn;
            kr4 = x_h_kassaazn;
            kr5 = x_h_kreditazn;

            meb1 = mebleg;
            double krmeb = Convert.ToInt32(mebleg) / 100;
            double krxh = krmeb / 2;
            meb2 = krmeb.ToString();
            meb3 = mebleg;
            meb4 = krxh.ToString();
            meb5 = krxh.ToString();

            if (textBox4.Enabled == true)
            {
                db6 = kassa_hesazn;
                kr6 = carihes;
                meb6 = textBox4.Text;
                teyinat6 = "Əmanət";


            }

        }

        private void excele_at()
        {
            try
            {

                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string desktopFolderPK = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopFolder = desktopFolder + "\\Kredit pr1.xls";
                desktopFolderPK = desktopFolderPK + "\\Kredit pr.xls";
                Microsoft.Office.Interop.Excel.Application objexcel = new Microsoft.Office.Interop.Excel.Application();
                //objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook sheet = objexcel.Workbooks.Open(desktopFolder);
                objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objbook = objexcel.Workbooks.Open(desktopFolder);
                Microsoft.Office.Interop.Excel.Worksheet objshet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range objRange;
                Microsoft.Office.Interop.Excel.Range objRange2;
                Microsoft.Office.Interop.Excel.Range objRange3;
                Microsoft.Office.Interop.Excel.Range objRange4, objRange5, objRange6, objRange7, objRange8, objRange9, objkr1, objkr2, objkr3, objkr4, objkr5, objkr6, objmeb1, objmeb2, objmeb3, objmeb4, objmeb5, objmeb6;
                Microsoft.Office.Interop.Excel.Range objtey1, objtey2, objtey3, objtey4, objtey5, objtey6;

                objRange8 = objshet.get_Range("b2", System.Reflection.Missing.Value);
                objRange8.set_Value(System.Reflection.Missing.Value, "");
                objRange8 = objshet.get_Range("b2", System.Reflection.Missing.Value);
                objRange8.set_Value(System.Reflection.Missing.Value, "0");

                objRange7 = objshet.get_Range("c2", System.Reflection.Missing.Value);
                objRange7.set_Value(System.Reflection.Missing.Value, "");
                objRange7 = objshet.get_Range("c2", System.Reflection.Missing.Value);
                objRange7.set_Value(System.Reflection.Missing.Value, sub1);

                objRange = objshet.get_Range("d2", System.Reflection.Missing.Value);
                objRange.set_Value(System.Reflection.Missing.Value, "");
                objRange = objshet.get_Range("d2", System.Reflection.Missing.Value);
                objRange.set_Value(System.Reflection.Missing.Value, db1);

                objRange2 = objshet.get_Range("d3", System.Reflection.Missing.Value);
                objRange2.set_Value(System.Reflection.Missing.Value, "");
                objRange2 = objshet.get_Range("d3", System.Reflection.Missing.Value);
                objRange2.set_Value(System.Reflection.Missing.Value, db2);

                objRange3 = objshet.get_Range("d4", System.Reflection.Missing.Value);
                objRange3.set_Value(System.Reflection.Missing.Value, "");
                objRange3 = objshet.get_Range("d4", System.Reflection.Missing.Value);
                objRange3.set_Value(System.Reflection.Missing.Value, db3);

                objRange9 = objshet.get_Range("h4", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "");
                objRange9 = objshet.get_Range("h4", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "0");

                objRange4 = objshet.get_Range("d5", System.Reflection.Missing.Value);
                objRange4.set_Value(System.Reflection.Missing.Value, "");
                objRange4 = objshet.get_Range("d5", System.Reflection.Missing.Value);
                objRange4.set_Value(System.Reflection.Missing.Value, db4);

                objRange5 = objshet.get_Range("d6", System.Reflection.Missing.Value);
                objRange5.set_Value(System.Reflection.Missing.Value, "");
                objRange5 = objshet.get_Range("d6", System.Reflection.Missing.Value);
                objRange5.set_Value(System.Reflection.Missing.Value, db5);

                objRange6 = objshet.get_Range("d7", System.Reflection.Missing.Value);
                objRange6.set_Value(System.Reflection.Missing.Value, "");
                objRange6 = objshet.get_Range("d7", System.Reflection.Missing.Value);
                objRange6.set_Value(System.Reflection.Missing.Value, db6);

                objkr1 = objshet.get_Range("f2", System.Reflection.Missing.Value);
                objkr1.set_Value(System.Reflection.Missing.Value, "");
                objkr1 = objshet.get_Range("f2", System.Reflection.Missing.Value);
                objkr1.set_Value(System.Reflection.Missing.Value, kr1);

                objkr2 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr2.set_Value(System.Reflection.Missing.Value, "");
                objkr2 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr2.set_Value(System.Reflection.Missing.Value, kr2);

                objkr3 = objshet.get_Range("f4", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, "");
                objkr3 = objshet.get_Range("f4", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, kr3);

                objkr4 = objshet.get_Range("f5", System.Reflection.Missing.Value);
                objkr4.set_Value(System.Reflection.Missing.Value, "");
                objkr4 = objshet.get_Range("f5", System.Reflection.Missing.Value);
                objkr4.set_Value(System.Reflection.Missing.Value, kr4);

                objkr5 = objshet.get_Range("f6", System.Reflection.Missing.Value);
                objkr5.set_Value(System.Reflection.Missing.Value, "");
                objkr5 = objshet.get_Range("f6", System.Reflection.Missing.Value);
                objkr5.set_Value(System.Reflection.Missing.Value, kr5);

                objkr6 = objshet.get_Range("f7", System.Reflection.Missing.Value);
                objkr6.set_Value(System.Reflection.Missing.Value, "");
                objkr6 = objshet.get_Range("f7", System.Reflection.Missing.Value);
                objkr6.set_Value(System.Reflection.Missing.Value, kr6);

                objmeb1 = objshet.get_Range("g2", System.Reflection.Missing.Value);
                objmeb1.set_Value(System.Reflection.Missing.Value, "");
                objmeb1 = objshet.get_Range("g2", System.Reflection.Missing.Value);
                objmeb1.set_Value(System.Reflection.Missing.Value, meb1);

                objmeb2 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb2.set_Value(System.Reflection.Missing.Value, "");
                objmeb2 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb2.set_Value(System.Reflection.Missing.Value, meb2);

                objmeb3 = objshet.get_Range("g4", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, "");
                objmeb3 = objshet.get_Range("g4", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, meb3);

                objmeb4 = objshet.get_Range("g5", System.Reflection.Missing.Value);
                objmeb4.set_Value(System.Reflection.Missing.Value, "");
                objmeb4 = objshet.get_Range("g5", System.Reflection.Missing.Value);
                objmeb4.set_Value(System.Reflection.Missing.Value, meb4);

                objmeb5 = objshet.get_Range("g6", System.Reflection.Missing.Value);
                objmeb5.set_Value(System.Reflection.Missing.Value, "");
                objmeb5 = objshet.get_Range("g6", System.Reflection.Missing.Value);
                objmeb5.set_Value(System.Reflection.Missing.Value, meb5);

                objmeb6 = objshet.get_Range("g7", System.Reflection.Missing.Value);
                objmeb6.set_Value(System.Reflection.Missing.Value, "");
                objmeb6 = objshet.get_Range("g7", System.Reflection.Missing.Value);
                objmeb6.set_Value(System.Reflection.Missing.Value, meb6);

                objtey1 = objshet.get_Range("j2", System.Reflection.Missing.Value);
                objtey1.set_Value(System.Reflection.Missing.Value, "");
                objtey1 = objshet.get_Range("j2", System.Reflection.Missing.Value);
                objtey1.set_Value(System.Reflection.Missing.Value, teyinat1);

                objtey2 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey2.set_Value(System.Reflection.Missing.Value, "");
                objtey2 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey2.set_Value(System.Reflection.Missing.Value, teyinat2);

                objtey3 = objshet.get_Range("j4", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, "");
                objtey3 = objshet.get_Range("j4", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, teyinat3);

                objtey4 = objshet.get_Range("j5", System.Reflection.Missing.Value);
                objtey4.set_Value(System.Reflection.Missing.Value, "");
                objtey4 = objshet.get_Range("j5", System.Reflection.Missing.Value);
                objtey4.set_Value(System.Reflection.Missing.Value, teyinat4);

                objtey5 = objshet.get_Range("j6", System.Reflection.Missing.Value);
                objtey5.set_Value(System.Reflection.Missing.Value, "");
                objtey5 = objshet.get_Range("j6", System.Reflection.Missing.Value);
                objtey5.set_Value(System.Reflection.Missing.Value, teyinat5);

                objtey6 = objshet.get_Range("j7", System.Reflection.Missing.Value);
                objtey6.set_Value(System.Reflection.Missing.Value, "");
                objtey6 = objshet.get_Range("j7", System.Reflection.Missing.Value);
                objtey6.set_Value(System.Reflection.Missing.Value, teyinat6);

                //objbook.Save();
                objbook.SaveAs(desktopFolderPK, Type.Missing);
                objexcel.Quit();

               


            }
            catch (Exception)
            {
            }

        }

        private void excele_at_x_hsiz()
        {
            try
            {

                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string desktopFolderPK = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopFolder = desktopFolder + "\\Kredit pr1.xls";
                desktopFolderPK = desktopFolderPK + "\\Kredit pr.xls";
                Microsoft.Office.Interop.Excel.Application objexcel = new Microsoft.Office.Interop.Excel.Application();
                //objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook sheet = objexcel.Workbooks.Open(desktopFolder);
                objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objbook = objexcel.Workbooks.Open(desktopFolder);
                Microsoft.Office.Interop.Excel.Worksheet objshet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range objRange;
                Microsoft.Office.Interop.Excel.Range objRange2;
                Microsoft.Office.Interop.Excel.Range objRange3;
                Microsoft.Office.Interop.Excel.Range objRange4, objRange5, objRange6, objRange7, objRange8, objRange9, objkr1, objkr2, objkr3, objkr4, objkr5, objkr6, objmeb1, objmeb2, objmeb3, objmeb4, objmeb5, objmeb6;
                Microsoft.Office.Interop.Excel.Range objtey1, objtey2, objtey3, objtey4, objtey5, objtey6;

                objRange8 = objshet.get_Range("b2", System.Reflection.Missing.Value);
                objRange8.set_Value(System.Reflection.Missing.Value, "");
                objRange8 = objshet.get_Range("b2", System.Reflection.Missing.Value);
                objRange8.set_Value(System.Reflection.Missing.Value, "0");

                objRange7 = objshet.get_Range("c2", System.Reflection.Missing.Value);
                objRange7.set_Value(System.Reflection.Missing.Value, "");
                objRange7 = objshet.get_Range("c2", System.Reflection.Missing.Value);
                objRange7.set_Value(System.Reflection.Missing.Value, sub1);

                objRange = objshet.get_Range("d2", System.Reflection.Missing.Value);
                objRange.set_Value(System.Reflection.Missing.Value, "");
                objRange = objshet.get_Range("d2", System.Reflection.Missing.Value);
                objRange.set_Value(System.Reflection.Missing.Value, db1);

                objkr1 = objshet.get_Range("f2", System.Reflection.Missing.Value);
                objkr1.set_Value(System.Reflection.Missing.Value, "");
                objkr1 = objshet.get_Range("f2", System.Reflection.Missing.Value);
                objkr1.set_Value(System.Reflection.Missing.Value, kr1);

                objmeb1 = objshet.get_Range("g2", System.Reflection.Missing.Value);
                objmeb1.set_Value(System.Reflection.Missing.Value, "");
                objmeb1 = objshet.get_Range("g2", System.Reflection.Missing.Value);
                objmeb1.set_Value(System.Reflection.Missing.Value, meb1);

                objtey1 = objshet.get_Range("j2", System.Reflection.Missing.Value);
                objtey1.set_Value(System.Reflection.Missing.Value, "");
                objtey1 = objshet.get_Range("j2", System.Reflection.Missing.Value);
                objtey1.set_Value(System.Reflection.Missing.Value, teyinat1);

                objRange3 = objshet.get_Range("d3", System.Reflection.Missing.Value);
                objRange3.set_Value(System.Reflection.Missing.Value, "");
                objRange3 = objshet.get_Range("d3", System.Reflection.Missing.Value);
                objRange3.set_Value(System.Reflection.Missing.Value, db3);

                objRange9 = objshet.get_Range("h3", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "");
                objRange9 = objshet.get_Range("h3", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "0");



                objkr3 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, "");
                objkr3 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, kr3);



                objmeb3 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, "");
                objmeb3 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, meb3);



                objtey3 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, "");
                objtey3 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, teyinat3);



                //objbook.Save();
                objbook.SaveAs(desktopFolderPK, Type.Missing);
                objexcel.Quit();

                //objexcel.Visible = false;


            }
            catch (Exception)
            {
            }

        }


        private readonly string templatefilname = @"C:\BMI_\BMI\bin\Debug\qiziltablozaminsiz.docx";
        private readonly string qizilzamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_Birzamin.docx";
        private readonly string qizilikizamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_ikizamin.docx";
        private readonly string qiziluczamin = @"C:\BMI_\BMI\bin\Debug\qiziltablo_Uczamin.docx";
        private readonly string qizilpapka = @"C:‪\\192.168.0.5\kred_sob\Qizil_muqavile";

        

        private string IlkHarfleriBuyut(string metin)
        {
            System.Globalization.CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Globalization.TextInfo textInfo = cultureInfo.TextInfo;
            return textInfo.ToTitleCase(metin);
        }
        private void word_at()
        {

            var adi = txbQizilBorcalan.Text;
            var passport = txbQizilPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = comboBox1.Text;
            var Unvan = txbQizilUnvan.Text;
            var CariHesab = txbQizilCarihesab.Text;
            var Valyuta = cboxqizilValyuta.Text;
            //var Serencam = txbQizilSerencam.Text;
            var MuqNo = txbQizilMuqNo.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cboxqizilTeyinat.Text;
            var Olke = cboxqizilOlke.Text;
            var mebleg = txbQizilMebleg.Text;
            var muddet = txbQizilMuddet.Text;
            var faiz = txbQizilFaiz.Text;
            var vkfaiz = txbQizilVKFaiz.Text;
            var ayliq = txbQizilAyliq.Text;
            var fifd = txbQizilFİFD.Text;

            // TODO: Word Export
            var wordapp = new Word.Application();
            wordapp.Visible = false;

            try
            {

                var wordDocument = wordapp.Documents.Open(templatefilname);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                ReplaceWordStub("{CariHesab}", CariHesab, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                //ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg, wordDocument);
                ReplaceWordStub("{muddet}", muddet, wordDocument);
                ReplaceWordStub("{faiz}", faiz, wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz, wordDocument);
                //ReplaceWordStub("{vkfaiz}", vkfaiz, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq, wordDocument);
                ReplaceWordStub("{fifd}", fifd, wordDocument);
                //qizil melumatlari


            }
            catch
            {
                
            }

        }

        private void girov_meb_topla()
        {
            double toplam = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
            }
            label29.Text = toplam.ToString();
        }

        private void teyadi()
        {
            if (teyinat == "2001")
            {
                cboxqizilTeyinat.Text = "mənzil təmiri";
            }
            else if (teyinat == "2002")
            {
                cboxqizilTeyinat.Text = " Avtomobil alınması";
            }
            else if (teyinat == "2003")
            {
                cboxqizilTeyinat.Text = " məişət əşyalarının alınması";
            }

        }
        
        int kr_say;
        int ser_say;
        public int zam_say;
        public int qizil_sayoxu;
        public int qizil_say;
        public void muqavile_nom()
        {

            System.DateTime moment = new System.DateTime(
                                1999, 1, 13, 3, 57, 32, 11);
            // Year gets 1999.
            int year = moment.Year;
            string tarixIl = DateTime.Now.Date.Year.ToString();


            ////string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand komut = new OracleCommand();
            komut.Connection = con;
            komut.CommandText = "select KR_QIZIL,kr_serencam,kr_zaminler,il from odb.muqavile_nomreleri where il='" + tarixIl + "'";
            OracleDataReader dr = komut.ExecuteReader();


            while (dr.Read())
            {
                if (dr["KR_QIZIL"].ToString()=="")
                {
                    txbQizilMuqNo.Text = "1";
                    txbqizilgirNo.Text = "1";
                }
                //qizil_sayoxu = Convert.ToInt32(dr["KR_QIZIL"].ToString());
                //if (qizil_sayoxu.ToString()=="")
                //{
                //    txbQizilMuqNo.Text = "1";
                //    txbqizilgirNo.Text = "1";
                //}
                else
                {
                    zam_say = Convert.ToInt32(dr["kr_zaminler"].ToString()) + 1;
                    kr_say = Convert.ToInt32(dr["KR_QIZIL"].ToString()) + 1;
                    txbQizilMuqNo.Text = kr_say.ToString();
                    txbqizilgirNo.Text = kr_say.ToString();
                    ser_say = Convert.ToInt32(dr["kr_serencam"].ToString()) + 1;
                }

                
                //txbQizilSerencam.Text = ser_say.ToString();


            }


            con.Close();

           
        }
        
        
        private void Qizil_Load(object sender, EventArgs e)
        {
            int zaminlarsay = zam_say;

            Zaminler zmnlar = new Zaminler();
            zmnlar.comboadi = comboBox1.Text;
            zmnlar.qeydnozamin = subkod_qeyd;
            zmnlar.groupBox1.Visible = false;
            zmnlar.groupBox2.Visible = false;
            zmnlar.groupBox3.Visible = false;

            //zmnlar.kataloqtarixi = secilmistarix;
            if (comboBox1.Text == "Bir nəfərin zəmanəti")
            {
                zmnlar.groupBox1.Visible = true;

                zmnlar.txb_zam1No.Text = zaminlarsay.ToString();

                zmnlar.ShowDialog();

                zmnlar.txbzam1adi.Text = zam1adi;
                zmnlar.txbzam1Pasport.Text = zam1pass;
                zmnlar.txbzam1Telefon.Text = zam1tel;
                zmnlar.txbzam1Unvan.Text = zam1unvan;

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

        private void txbQizilAyliq_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    txbQizilAyliq.Text = txbQizilAyliq.Text.ToString().Replace('.', ',');
            //    if (txbQizilAyliq.Text != null)
            //    {
            //        btnqizilYazdir.Enabled = true;
            //    }
            //    else if (txbQizilAyliq.Text == "")
            //    {
            //        btnqizilYazdir.Enabled = false;
            //    }
            //}
            //catch (Exception)
            //{

            //    btnqizilYazdir.Enabled = false;
            //}
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            //e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            //e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            //textBox9.Text = textBox9.Text.ToString().Replace('.', ',');
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Zaminler zmnlar = new Zaminler();
            zmnlar.ShowDialog();

        }

    
        string VAL_AD;
        private void valyuta_adi()
        {
            if (valyuta == "00")
            {
                textBox2.Text = "AZN";
                textBox3.Text = "AZN";
                VAL_AD = "AZN";
            }
            else if (valyuta == "01")
            {
                textBox2.Text = "USD";
                textBox3.Text = "USD";
                VAL_AD = "USD";
            }
            else if (valyuta == "02")
            {
                textBox2.Text = "AVRO";
                textBox3.Text = "AVRO";
                VAL_AD = "AVRO";
            }

        }

        private void Qizil_Load_1(object sender, EventArgs e)
        {

            radioButton1.Checked = true;
            label31.Text = test;
            Menzil_sahibi msah = new Menzil_sahibi();
            //if (cboxqizilGirovsahibi.Text == "Fərqli şəxsə məxsus")
            //{
            //    msah.ShowDialog();
            //}
            if (comboadi == "Progress")
            {
                string adhazir;
                string unvanhazir;
                //DateTime bugun = DateTime.Now.ToShortDateString();
                //label19.Text = DateTime.Now.ToShortDateString();
                teyadi();
                olkeadi();
                valyuta_adi();
                //

                ///
                muqavile_nom();
                //muqavile_nom1();

                adhazir = adi.ToLower();

                txbQizilBorcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);

                //txbzBorcalan.Text = txbzBorcalan.Text.ToUpper();
                txbQizilMebleg.Text = mebleg;
                txbQizilPasport.Text = seriyano;
                txbQizilTelefon.Text = mobil;
                unvanhazir = unvan.ToLower();
                txbQizilUnvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);
                txbQizilVKFaiz.Text = vkfaiz;
                txbQizilFaiz.Text = faiz;
                comboBox1.Text = ver_orqan;
                txbQizilCarihesab.Text = carihes;
                txbQizilAyliq.Text = ayliq;

                txbQizilFİFD.Text = fifd;
                dateTimePicker1.Text = ver_tar;
                int gun = Convert.ToInt32(muddet);
                int gunsay = gun / 30;
                txbQizilMuddet.Text = gunsay.ToString();
                //cboxkzTeyinat.Text = teyinat; 
                if (valyuta == "00")
                {
                    cboxqizilValyuta.Text = "AZN";
                }
                else if (valyuta == "01")
                {
                    cboxqizilValyuta.Text = "USD";
                }
                else if (valyuta == "02")
                {
                    cboxqizilValyuta.Text = "AVRO";
                }
            }
            

        }
 
        private void worda_at_qizil_zaminsiz()
    {
            string template="";

        if (comboBox3.Text=="")
        {
             template = templatefilname;
        }
        else if (comboBox3.Text == "Bir nəfərin zəmanəti")
        {
             template = qizilzamin;
        }
        else if (comboBox3.Text == "İki nəfərin zəmanəti")
        {
             template = qizilikizamin;
        }
        else if (comboBox3.Text == "Üç nəfərin zəmanəti")
        {
             template = qiziluczamin;
        }

        int satirdatagsayi = Convert.ToInt32(dataGridView1.RowCount.ToString());
        int satirsay = dataGridView1.Rows.Count;
        int sutunsay = 8;
        string girov_meb_yaziile = yaziyaCevir(Convert.ToDecimal(label29.Text));
        string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
        string mebleg_yazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
        string ayyazile = yaziyaCevir(Convert.ToDecimal(txbQizilAyliq.Text));
        string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMuddet.Text));
        string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilFaiz.Text));
        string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilVKFaiz.Text));

        var adi = txbQizilBorcalan.Text;
        var passport = txbQizilPasport.Text;
        var pasvertarix = dateTimePicker1.Text;
        var Orqan = comboBox1.Text;
        var Unvan = txbQizilUnvan.Text;
        var tel = txbQizilTelefon.Text;
        var tamcari = tamhesab;
        var CariHesab = txbQizilCarihesab.Text;
        var Valyuta = cboxqizilValyuta.Text;
        //var Serencam = txbQizilSerencam.Text;
        var MuqNo = txbQizilMuqNo.Text;
        var Muqtarix = tamtarix;
        var Teyinat = cboxqizilTeyinat.Text;
        var Olke = cboxqizilOlke.Text;
        var mebleg = txbQizilMebleg.Text + " " + cboxqizilValyuta.Text;
        var muddet = txbQizilMuddet.Text;
        var faiz = txbQizilFaiz.Text;
        var vkfaiz = txbQizilVKFaiz.Text;
        var ayliq = txbQizilAyliq.Text;
        var fifd = txbQizilFİFD.Text+"%";
        var girNo = txbqizilgirNo.Text;
        var mebyaziile = ayyazile;
        
        var girmebleg = girov_meb_yaziile;
        var qizil_meb = label29.Text +" manat";

        var zam1 = zam1adi;
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
        var mevayyazi = ayyazile;
        
        var icraci = icraciadizaminlikde;
       

        
        var wordapp = new Word.Application();
        wordapp.Visible = false;
        var document = wordapp.Documents.Open(template);
        document.Activate();       
        Word.Table table = document.Tables[1];
        Word.Table table1 = document.Tables[3];
               
        for (int i = 0; i < satirdatagsayi; i++)
        {
            table.Rows.Add(table.Rows[1]);
            table1.Rows.Add(table1.Rows[1]);
        }

        table.Rows[1].Cells[1].Range.Text = "Sıra No";
        table.Rows[1].Cells[2].Range.Text = "Qiymətli metal";
        table.Rows[1].Cells[3].Range.Text = "Növü";
        table.Rows[1].Cells[4].Range.Text = "Ədəd sayı";
        table.Rows[1].Cells[5].Range.Text = "Əyarı";
        table.Rows[1].Cells[6].Range.Text = "Çəkisi qr";
        table.Rows[1].Cells[7].Range.Text = "Qiyməti 1qr AZN";
        table.Rows[1].Cells[8].Range.Text = "Dəyəri AZN";
        table.Rows[1].Cells[9].Range.Text = "Qeyd";

        table1.Rows[1].Cells[1].Range.Text = "Sıra No";
        table1.Rows[1].Cells[2].Range.Text = "Qiymətli metal";
        table1.Rows[1].Cells[3].Range.Text = "Növü";
        table1.Rows[1].Cells[4].Range.Text = "Ədəd sayı";
        table1.Rows[1].Cells[5].Range.Text = "Əyarı";
        table1.Rows[1].Cells[6].Range.Text = "Çəkisi qr";
        table1.Rows[1].Cells[7].Range.Text = "Qiyməti 1qr AZN";
        table1.Rows[1].Cells[8].Range.Text = "Dəyəri AZN";
        table1.Rows[1].Cells[9].Range.Text = "Qeyd"; 
       
        for (int i = 0; i < dataGridView1.Rows.Count-1 ; i++)
        {
            
            for (int j = 0; j < 9; j++)
            {

                table.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

            }
        }

        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
        {

            for (int j = 0; j < 9; j++)
            {
                table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
            }
        }

        ReplaceWordStub("{adi}", adi, document);
        ReplaceWordStub("{passport}", passport, document);
        ReplaceWordStub("{pasvertarix}", pasvertarix, document);
        ReplaceWordStub("{Orqan}", Orqan, document);
        ReplaceWordStub("{Unvan}", Unvan, document);
        ReplaceWordStub("{CariHesab}", tamcari, document);
        ReplaceWordStub("{Valyuta}", Valyuta, document);
        //ReplaceWordStub("{Serencam}", Serencam, document);
        ReplaceWordStub("{MuqNo}", MuqNo, document);
        ReplaceWordStub("{Muqtarix}", Muqtarix, document);
        ReplaceWordStub("{Teyinat}", Teyinat, document);
        ReplaceWordStub("{Olke}", Olke, document);
        ReplaceWordStub("{mebleg}", mebleg + " (" + mebleg_yazi + ")", document);
        ReplaceWordStub("{muddet}", muddet + " ay " + "(" + muddetyazi + ")", document);
        ReplaceWordStub("{faiz}", faiz + "% " + " (" + faizyazi + ")", document);
        ReplaceWordStub("{vkfaiz}", vkfaiz + "%" + " (" + vkfaizyazi + ")", document);
        //ReplaceWordStub("{vkfaiz}", vkfaiz, document);
        ReplaceWordStub("{ayliq}", ayliq + VAL_AD + " (" + mebyaziile + ")", document);
        ReplaceWordStub("{fifd}", fifd, document);
        ReplaceWordStub("{tel}", tel, document);
        ReplaceWordStub("{girovNo}", girNo, document);
        ReplaceWordStub("{qizil_mebleg}", qizil_meb, document);
        ReplaceWordStub("{top mebleg}", qizil_meb, document);

        ReplaceWordStub("{zam1ad}", zam1, document);
        ReplaceWordStub("{zam1pas}", zam1pas, document);
        ReplaceWordStub("{zam1telef}", zam1tel, document);
        ReplaceWordStub("{zam1unvani}", zam1unvan, document);
        ReplaceWordStub("{zam1ptarix}", zam1pastarixi, document);
        ReplaceWordStub("{zam1porqan}", zam1pasorqani, document);
        ReplaceWordStub("{zam1olke}", zam1olke, document);

        ReplaceWordStub("{zam2ad}", zam2, document);
        ReplaceWordStub("{zam2pas}", zam2pas, document);
        ReplaceWordStub("{zam2telef}", zam2tel, document);
        ReplaceWordStub("{zam2unvani}", zam2unvan, document);
        ReplaceWordStub("{zam2ptarix}", zam2pastarixi, document);
        ReplaceWordStub("{zam2porqan}", zam2pasorqani, document);
        ReplaceWordStub("{zam2olke}", zam2olke, document);

        ReplaceWordStub("{zam3ad}", zam3, document);
        ReplaceWordStub("{zam3pas}", zam3pas, document);
        ReplaceWordStub("{zam3telef}", zam3tel, document);
        ReplaceWordStub("{zam3unvani}", zam3unvan, document);
        ReplaceWordStub("{zam3ptarix}", zam3pastarixi, document);
        ReplaceWordStub("{zam3porqan}", zam3pasorqani, document);
        ReplaceWordStub("{zam3olke}", zam3olke, document);

        ReplaceWordStub("{zammuq1}", zammuqno1, document);
        ReplaceWordStub("{zammuq2}", zammuqno2, document);
        ReplaceWordStub("{zammuq3}", zammuqno3, document);

        wordapp.Visible = true;
            string muqadi = MuqNo + " " + adi + " " + tamtarix;
            //string muqadi = "yoxlanis";
            //string muqadi = adi;
            // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            document.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Qizil_muqavile\" + muqadi + ".docx");
            //document.SaveAs(qizilpapka + " qızıl" + muqadi);

        }

        private void worda_at_qizil_zamin1()
        {
            int satirsay = dataGridView1.Rows.Count;
            int sutunsay = 8;
            string girov_meb_yaziile = yaziyaCevir(Convert.ToDecimal(label29.Text));
            string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
            string mebleg_yazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
            string ayyazile = yaziyaCevir(Convert.ToDecimal(txbQizilAyliq.Text));
            string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMuddet.Text));
            string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilFaiz.Text));
            string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilVKFaiz.Text));

            var adi = txbQizilBorcalan.Text;
            var passport = txbQizilPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = comboBox1.Text;
            var Unvan = txbQizilUnvan.Text;
            var tel = txbQizilTelefon.Text;
            var tamcari = tamhesab;
            var Valyuta = cboxqizilValyuta.Text;
            //var Serencam = txbQizilSerencam.Text;
            var MuqNo = txbQizilMuqNo.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cboxqizilTeyinat.Text;
            var Olke = cboxqizilOlke.Text;
            var mebleg = txbQizilMebleg.Text + " " + cboxqizilValyuta.Text;
            var muddet = txbQizilMuddet.Text;
            var faiz = txbQizilFaiz.Text;
            var vkfaiz = txbQizilVKFaiz.Text;
            var ayliq = txbQizilAyliq.Text;
            var fifd = txbQizilFİFD.Text + "%";
            var girNo = txbqizilgirNo.Text;
            var mebyaziile = ayyazile;

            var girmebleg = girov_meb_yaziile;
            var qizil_meb = label29.Text + " manat";
            //var Unvan = txbQizilUnvan.Text;

            string template = qizilzamin;
            var wordapp = new Word.Application();
            wordapp.Visible = false;
            //Word.Document document = wordapp.Documents.OpenNoRepairDialog(template);
            var document = wordapp.Documents.Open(qizilzamin);
            document.Activate();

            //Word.Range tablelocation = this.Range(ref satirsay, ref sutunsay);

            Word.Table table = document.Tables[1];
            //Word.Table table1 = document.Tables[2];
            //this.table.Add(tablelocation, 3, 4);
            //table.Cell(1, 1).Range.Text = "Samir";
            //table.Cell(1, 2).Range.Text = "huseynov";
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    table.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    //table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

                }


            }

            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    for (int j = 0; j < 8; j++)
            //    {
            //        table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //        //table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

            //    }
            //}

            ReplaceWordStub("{adi}", adi, document);
            ReplaceWordStub("{passport}", passport, document);
            ReplaceWordStub("{pasvertarix}", pasvertarix, document);
            ReplaceWordStub("{Orqan}", Orqan, document);
            ReplaceWordStub("{Unvan}", Unvan, document);
            ReplaceWordStub("{CariHesab}", tamcari, document);
            ReplaceWordStub("{Valyuta}", Valyuta, document);
            //ReplaceWordStub("{Serencam}", Serencam, document);
            ReplaceWordStub("{MuqNo}", MuqNo, document);
            ReplaceWordStub("{Muqtarix}", Muqtarix, document);
            ReplaceWordStub("{Teyinat}", Teyinat, document);
            ReplaceWordStub("{Olke}", Olke, document);
            ReplaceWordStub("{mebleg}", mebleg + " (" + mebleg_yazi + ")", document);
            ReplaceWordStub("{muddet}", muddet + " ay " + "(" + muddetyazi + ")", document);
            ReplaceWordStub("{faiz}", faiz + "% " + " (" + faizyazi + ")", document);
            ReplaceWordStub("{vkfaiz}", vkfaiz + "%" + " (" + vkfaizyazi + ")", document);
            //ReplaceWordStub("{vkfaiz}", vkfaiz, document);
            ReplaceWordStub("{ayliq}", ayliq + VAL_AD + " (" + mebyaziile + ")", document);
            ReplaceWordStub("{fifd}", fifd, document);
            ReplaceWordStub("{tel}", tel, document);
            ReplaceWordStub("{girovNo}", girNo, document);
            ReplaceWordStub("{qizil_mebleg}", qizil_meb, document);
            ReplaceWordStub("{top mebleg}", qizil_meb, document);

            wordapp.Visible = true;
            string muqadi = MuqNo + " " + adi + " " + tamtarix;
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            document.SaveAs(Application.StartupPath + " qızıl" + muqadi);
            //document.SaveAs(qizilpapka + " qızıl" + muqadi);

        }

        private void worda_at_qizil_zamin2()
        {
            int satirsay = dataGridView1.Rows.Count;
            int sutunsay = 8;
            string girov_meb_yaziile = yaziyaCevir(Convert.ToDecimal(label29.Text));
            string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
            string mebleg_yazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMebleg.Text));
            string ayyazile = yaziyaCevir(Convert.ToDecimal(txbQizilAyliq.Text));
            string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilMuddet.Text));
            string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilFaiz.Text));
            string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbQizilVKFaiz.Text));

            var adi = txbQizilBorcalan.Text;
            var passport = txbQizilPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = comboBox1.Text;
            var Unvan = txbQizilUnvan.Text;
            var tel = txbQizilTelefon.Text;
            var tamcari = tamhesab;
            var Valyuta = cboxqizilValyuta.Text;
            //var Serencam = txbQizilSerencam.Text;
            var MuqNo = txbQizilMuqNo.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cboxqizilTeyinat.Text;
            var Olke = cboxqizilOlke.Text;
            var mebleg = txbQizilMebleg.Text + " " + cboxqizilValyuta.Text;
            var muddet = txbQizilMuddet.Text;
            var faiz = txbQizilFaiz.Text;
            var vkfaiz = txbQizilVKFaiz.Text;
            var ayliq = txbQizilAyliq.Text;
            var fifd = txbQizilFİFD.Text + "%";
            var girNo = txbqizilgirNo.Text;
            var mebyaziile = ayyazile;

            var girmebleg = girov_meb_yaziile;
            var qizil_meb = label29.Text + " manat";
            //var Unvan = txbQizilUnvan.Text;
            var zam1 = zam1adi;
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
            //var mebyaziile = yaziile;
            var mevayyazi = ayyazile;
            //var tamcari = tamhesab;
            //var icraci = icraci_adi;
            var icraci = icraciadizaminlikde;

            string template = qizilikizamin;
            var wordapp = new Word.Application();
            wordapp.Visible = false;
            //Word.Document document = wordapp.Documents.OpenNoRepairDialog(template);
            var document = wordapp.Documents.Open(qizilikizamin);
            document.Activate();

            //Word.Range tablelocation = this.Range(ref satirsay, ref sutunsay);

            Word.Table table = document.Tables[1];
            //Word.Table table1 = document.Tables[2];
            //this.table.Add(tablelocation, 3, 4);
            //table.Cell(1, 1).Range.Text = "Samir";
            //table.Cell(1, 2).Range.Text = "huseynov";
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    table.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    //table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

                }


            }

            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    for (int j = 0; j < 8; j++)
            //    {
            //        table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //        //table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

            //    }
            //}

            ReplaceWordStub("{adi}", adi, document);
            ReplaceWordStub("{passport}", passport, document);
            ReplaceWordStub("{pasvertarix}", pasvertarix, document);
            ReplaceWordStub("{Orqan}", Orqan, document);
            ReplaceWordStub("{Unvan}", Unvan, document);
            ReplaceWordStub("{CariHesab}", tamcari, document);
            ReplaceWordStub("{Valyuta}", Valyuta, document);
            //ReplaceWordStub("{Serencam}", Serencam, document);
            ReplaceWordStub("{MuqNo}", MuqNo, document);
            ReplaceWordStub("{Muqtarix}", Muqtarix, document);
            ReplaceWordStub("{Teyinat}", Teyinat, document);
            ReplaceWordStub("{Olke}", Olke, document);
            ReplaceWordStub("{mebleg}", mebleg + " (" + mebleg_yazi + ")", document);
            ReplaceWordStub("{muddet}", muddet + " ay " + "(" + muddetyazi + ")", document);
            ReplaceWordStub("{faiz}", faiz + "% " + " (" + faizyazi + ")", document);
            ReplaceWordStub("{vkfaiz}", vkfaiz + "%" + " (" + vkfaizyazi + ")", document);
            //ReplaceWordStub("{vkfaiz}", vkfaiz, document);
            ReplaceWordStub("{ayliq}", ayliq + VAL_AD + " (" + mebyaziile + ")", document);
            ReplaceWordStub("{fifd}", fifd, document);
            ReplaceWordStub("{tel}", tel, document);
            ReplaceWordStub("{girovNo}", girNo, document);
            ReplaceWordStub("{qizil_mebleg}", qizil_meb, document);
            ReplaceWordStub("{top mebleg}", qizil_meb, document);

            ReplaceWordStub("{zam1ad}", zam1, document);
            ReplaceWordStub("{zam1pas}", zam1pas, document);
            ReplaceWordStub("{zam1telef}", zam1tel, document);
            ReplaceWordStub("{zam1unvani}", zam1unvan, document);
            ReplaceWordStub("{zam1ptarix}", zam1pastarixi, document);
            ReplaceWordStub("{zam1porqan}", zam1pasorqani, document);
            ReplaceWordStub("{zam1olke}", zam1olke, document);

            ReplaceWordStub("{zam2ad}", zam2, document);
            ReplaceWordStub("{zam2pas}", zam2pas, document);
            ReplaceWordStub("{zam2telef}", zam2tel, document);
            ReplaceWordStub("{zam2unvani}", zam2unvan, document);
            ReplaceWordStub("{zam2ptarix}", zam2pastarixi, document);
            ReplaceWordStub("{zam2porqan}", zam2pasorqani, document);
            ReplaceWordStub("{zam2olke}", zam2olke, document);

            ReplaceWordStub("{zam3ad}", zam3, document);
            ReplaceWordStub("{zam3pas}", zam3pas, document);
            ReplaceWordStub("{zam3telef}", zam3tel, document);
            ReplaceWordStub("{zam3unvani}", zam3unvan, document);
            ReplaceWordStub("{zam3ptarix}", zam3pastarixi, document);
            ReplaceWordStub("{zam3porqan}", zam3pasorqani, document);
            ReplaceWordStub("{zam3olke}", zam3olke, document);

            ReplaceWordStub("{zammuq1}", zammuqno1, document);
            ReplaceWordStub("{zammuq2}", zammuqno2, document);
            ReplaceWordStub("{zammuq3}", zammuqno3, document);

            wordapp.Visible = true;
            string muqadi = MuqNo + " " + adi + " " + tamtarix;
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            document.SaveAs(Application.StartupPath + " qızıl" + muqadi);
            //document.SaveAs(qizilpapka + " qızıl" + muqadi);

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
        private void wordetabloat()
        {
            int satirsay = dataGridView1.Rows.Count;
            int sutunsay = 8;

            string template = @Path.GetDirectoryName(Application.ExecutablePath).Trim() + "\\Test.docx";
            Word.Application wordapp = new Word.Application();
            wordapp.Visible = true;
            Word.Document document = wordapp.Documents.OpenNoRepairDialog(template);
            document.Activate();

            //Word.Range tablelocation = this.Range(ref satirsay, ref sutunsay);

            Word.Table table = document.Tables[1];
            //Word.Table table1 = document.Tables[2];
            //this.table.Add(tablelocation, 3, 4);
            //table.Cell(1, 1).Range.Text = "Samir";
            //table.Cell(1, 2).Range.Text = "huseynov";
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < 8; j++)
                {
                    table.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    //table1.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

                }
            }
            var MuqNo = txbQizilMuqNo.Text;
            wordapp.Visible = true;
            string muqadi = MuqNo + " " + adi + " " + tamtarix;
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            document.SaveAs(Application.StartupPath + " qızıl" + muqadi);
            //document.SaveAs(qizilpapka + " qızıl" + muqadi);

         
        }
        private Word.Range Range(ref int satirsay, ref int sutunsay)
        {
            throw new NotImplementedException();
        }
        private void tabloat()
        {

            object objmissing = System.Reflection.Missing.Value;
            object oEndOfdoc = "C:\\BMI_\\BMI\\bin\\Debug\\testqizil.docx";
            Microsoft.Office.Interop.Word.Application appob;
            Microsoft.Office.Interop.Word.Document docob;
            appob = new Microsoft.Office.Interop.Word.Application();
            appob.Visible = true;
            docob = appob.Documents.Add(ref objmissing, ref objmissing, ref objmissing);
            int i = 0;
            int j = 0;
            Microsoft.Office.Interop.Word.Table tableodb;
            Word.Range wrdrng = docob.Bookmarks.get_Item(ref oEndOfdoc).Range;
            tableodb = docob.Tables.Add(wrdrng, 3, 4, ref objmissing, ref objmissing);
            tableodb.Range.ParagraphFormat.SpaceAfter = 8;
            string str;
            for (i = 0; i <= 3; i++)
            {
                for (j = 0; j <= 4; j++)
                {
                    str = "Row" + i + "Column";
                    tableodb.Cell(i, j).Range.Text = str;

                }
                tableodb.Rows[1].Range.Font.Bold = 1;
                this.Close();
            }



        }
        private void worda_at_qizil_BIRzaminsiz()
        {
            int satirsay = dataGridView1.Rows.Count;
            int sutunsay = 9;

            string girov_meb_yaziile = yaziyaCevir(Convert.ToDecimal(label29.Text));

            var adi = txbQizilBorcalan.Text;
            var passport = txbQizilPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = comboBox1.Text;
            var Unvan = txbQizilUnvan.Text;
            var CariHesab = txbQizilCarihesab.Text;
            var Valyuta = cboxqizilValyuta.Text;
            //var Serencam = txbQizilSerencam.Text;
            var MuqNo = txbQizilMuqNo.Text;
            var Muqtarix = tamtarix;
            var Teyinat = cboxqizilTeyinat.Text;
            var Olke = cboxqizilOlke.Text;
            var mebleg = txbQizilMebleg.Text;
            var muddet = txbQizilMuddet.Text;
            var faiz = txbQizilFaiz.Text;
            var vkfaiz = txbQizilVKFaiz.Text;
            var ayliq = txbQizilAyliq.Text;
            var fifd = txbQizilFİFD.Text;
            var girNo = txbqizilgirNo.Text;
            var girmebleg = girov_meb_yaziile;
            var qizil_meb = label29.Text;
            var tamcari = tamhesab;

            var zam1 = zam1adi;
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
            //var Unvan = txbQizilUnvan.Text;

            string template = qizilzamin;
            var wordapp = new Word.Application();
            wordapp.Visible = false;
            //Word.Document document = wordapp.Documents.OpenNoRepairDialog(template);
            var document = wordapp.Documents.Open(qizilzamin);
            document.Activate();

            ReplaceWordStub("{adi}", adi, document);
            ReplaceWordStub("{passport}", passport, document);
            ReplaceWordStub("{pasvertarix}", pasvertarix, document);
            ReplaceWordStub("{Orqan}", Orqan, document);
            ReplaceWordStub("{Unvan}", Unvan, document);
            ReplaceWordStub("{CariHesab}", tamcari, document);
            ReplaceWordStub("{Valyuta}", Valyuta, document);
            //ReplaceWordStub("{Serencam}", Serencam, document);
            ReplaceWordStub("{MuqNo}", MuqNo, document);
            ReplaceWordStub("{Muqtarix}", Muqtarix, document);
            ReplaceWordStub("{Teyinat}", Teyinat, document);
            ReplaceWordStub("{Olke}", Olke, document);
            ReplaceWordStub("{mebleg}", mebleg, document);
            ReplaceWordStub("{muddet}", muddet, document);
            ReplaceWordStub("{faiz}", faiz, document);
            ReplaceWordStub("{vkfaiz}", vkfaiz, document);
            ReplaceWordStub("{vkfaiz}", vkfaiz, document);
            ReplaceWordStub("{ayliq}", ayliq, document);
            ReplaceWordStub("{fifd}", fifd, document);
            ReplaceWordStub("{girovNo}", girNo, document);
            ReplaceWordStub("{qizil_mebleg}", girmebleg, document);
            ReplaceWordStub("{top mebleg}", girmebleg, document);

            ReplaceWordStub("{zam1ad}", zam1, document);
            ReplaceWordStub("{zam1pas}", zam1pas, document);
            ReplaceWordStub("{zam1telef}", zam1tel, document);
            ReplaceWordStub("{zam1unvani}", zam1unvan, document);
            ReplaceWordStub("{zam1ptarix}", zam1pastarixi, document);
            ReplaceWordStub("{zam1porqan}", zam1pasorqani, document);
            ReplaceWordStub("{zam1olke}", zam1olke, document);

            ReplaceWordStub("{zam2ad}", zam2, document);
            ReplaceWordStub("{zam2pas}", zam2pas, document);
            ReplaceWordStub("{zam2telef}", zam2tel, document);
            ReplaceWordStub("{zam2unvani}", zam2unvan, document);
            ReplaceWordStub("{zam2ptarix}", zam2pastarixi, document);
            ReplaceWordStub("{zam2porqan}", zam2pasorqani, document);
            ReplaceWordStub("{zam2olke}", zam2olke, document);

            ReplaceWordStub("{zam3ad}", zam3, document);
            ReplaceWordStub("{zam3pas}", zam3pas, document);
            ReplaceWordStub("{zam3telef}", zam3tel, document);
            ReplaceWordStub("{zam3unvani}", zam3unvan, document);
            ReplaceWordStub("{zam3ptarix}", zam3pastarixi, document);
            ReplaceWordStub("{zam3porqan}", zam3pasorqani, document);
            ReplaceWordStub("{zam3olke}", zam3olke, document);

            ReplaceWordStub("{zammuq1}", zammuqno1, document);
            ReplaceWordStub("{zammuq2}", zammuqno2, document);
            ReplaceWordStub("{zammuq3}", zammuqno3, document);

            wordapp.Visible = true;
            //string muqadi = MuqNo + " " + adi + " " + tamtarix;
            ////wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
            ////wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            //document.SaveAs(Application.StartupPath + " qızıl" + muqadi);
            //document.SaveAs(qizilpapka + " qızıl" + muqadi);

            //wordapp.Visible = true;
            string muqadi = MuqNo + " " + adi + " " + tamtarix;
            //string muqadi = "yoxlanis";
            //string muqadi = adi;
            // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
            //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
            document.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Qizil_muqavile\" + muqadi + ".docx");

            //\\192.168.0.5\kred_sob\Qizil muqavile
        }

        int yencavab = 0;
        private void muq_no_at()
        {
            System.DateTime moment = new System.DateTime(
                                1999, 1, 13, 3, 57, 32, 11);
            // Year gets 1999.
            int year = moment.Year;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand OCOM = new OracleCommand("Update odb.muqavile_nomreleri set kr_serencam='" + ser_say + "',KR_QIZIL='" + txbQizilMuqNo.Text + "',kr_zaminler='" + zam_say + "' where IL='" + tarixIl + "'", con);
            yencavab = OCOM.ExecuteNonQuery();
            if (yencavab > 0)
            {
                //MessageBox.Show("Məlumat yeniləndi...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // MessageBox.Show("Məlumat yenilənmədi...!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            con.Close();


            


        }

        private void btnqizilYazdir_Click(object sender, EventArgs e)
        {
           

            Zaminler zmnlar = new Zaminler();
            if (comboBox3.Text == "")
            {
                //muqavile_nom();
                muq_no_at();
                worda_at_qizil_zaminsiz();
                zamincevirme();

                if (radioButton2.Checked == true)
                {
                    zamincevirme();
                    excele_at_x_hsiz();
                }
                else if (radioButton1.Checked == true)
                {
                    zamincevirme();
                    excele_at();
                }

            }


            else if (comboBox3.Text == "Bir nəfərin zəmanəti")
            {
                //muqavile_nom();
                muq_no_at();
                worda_at_qizil_zaminsiz();
                zamincevirme();

                if (radioButton2.Checked == true)
                {
                    zamincevirme();
                    excele_at_x_hsiz();
                }
                else if (radioButton1.Checked == true)
                {
                    zamincevirme();
                    excele_at();
                }



            }
            else if (comboBox3.Text == "İki nəfərin zəmanəti")
            {
                //muqavile_nom();
               muq_no_at();
                worda_at_qizil_zaminsiz();
                zamincevirme();

                if (radioButton2.Checked == true)
                {
                    zamincevirme();
                    excele_at_x_hsiz();
                }
                else if (radioButton1.Checked == true)
                {
                    zamincevirme();
                    excele_at();
                }

            }
            else if (comboBox3.Text == "Üç nəfərin zəmanəti")
            {
                //muqavile_nom();
                muq_no_at();
                worda_at_qizil_zaminsiz();
                zamincevirme();

                if (radioButton2.Checked == true)
                {
                    zamincevirme();
                    excele_at_x_hsiz();
                }
                else if (radioButton1.Checked == true)
                {
                    zamincevirme();
                    excele_at();
                }


            }
            else if (txbQizilAyliq.Text == "")
            {
                MessageBox.Show("Məlumatlar tam doldurulmayıb.");
            }
            else if (textBox5.Text == "")
            {
                MessageBox.Show("Məlumatlar tam doldurulmayıb.");
            }
            else if (txbQizilCarihesab.Text == "")
            {
                MessageBox.Show("Məlumatlar tam doldurulmayıb.");
            }
            

        }

        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document document)
        {
            try
            {
                var range = document.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
            }
            catch (Exception)
            {
                
               
            }
            
        }

               
        private void button1_Click(object sender, EventArgs e)
        {
            if (txbQizilAyliq.Text == "")
            {
                MessageBox.Show("Aylıq ödəniş qeyd edilməyib!!!", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {


                yazdir();
                girov_meb_topla();
            }
        }
        string metal, novu, eyar,qeyd;
        double sayi, ceki, qiymet, deyeri;
        int say = 0;

        private void yazdir()
        {
            try
            {
                if (textBox6.Text == "" || textBox8.Text == "" || textBox9.Text == "")
                {
                    MessageBox.Show("Melumatlar doldurulmayıb");
                }
               
                else
                {
                    metal = textBox1.Text;
                    novu = comboBox2.Text;
                    eyar = textBox7.Text;
                    qeyd = textBox10.Text;

                    sayi = Convert.ToDouble(textBox6.Text);
                    ceki = Convert.ToDouble(textBox8.Text);
                    qiymet = Convert.ToDouble(textBox9.Text);
                    deyeri = ceki * sayi * qiymet;

                    say = say + 1;
                    dataGridView1.Rows.Add(say, metal, novu, sayi, eyar, ceki, qiymet, deyeri, qeyd);
                    textBox8.Text = "";
                }
            }
            catch (Exception)
            {

                
            }

            finally { }

            //dataGridView1.DataSource = tablo;
        }

        private void hesabla()
        {
            decimal cem;
            decimal qram;
            decimal qiymet;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                qram = decimal.Parse(dataGridView1.Rows[i].Cells[6].Value.ToString());
                qiymet = decimal.Parse(dataGridView1.Rows[i].Cells[7].Value.ToString());

                cem = qram * qiymet;
                dataGridView1.Rows[i].Cells[8].Value = cem.ToString() + "AZN";
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            int zaminlarsay = zam_say;


            Zaminler zmnlar = new Zaminler();

            zmnlar.zmnlartarix = umumitar;
            zmnlar.muqtipi = label31.Text;
            //zmnlar.muqtipi = girovnovucombo;
            zmnlar.muqtipi = muqtipi;
            zmnlar.comboadi = comboBox3.Text;
            zmnlar.muqtipi = label31.Text;
            zmnlar.qeydnozamin = subkod_qeyd;
            zmnlar.groupBox1.Visible = false;
            zmnlar.groupBox2.Visible = false;
            zmnlar.groupBox3.Visible = false;
            zmnlar.label23.Text = label31.Text;

            //zmnlar.kataloqtarixi = secilmistarix;
            if (comboBox3.Text == "Bir nəfərin zəmanəti")
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
            else if (comboBox3.Text == "İki nəfərin zəmanəti")
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
            else if (comboBox3.Text == "Üç nəfərin zəmanəti")
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
                //zmnlar.muqtipi = test;
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

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void baglaac()
        {
            if (Convert.ToInt16(textBox8.Text) > 0)
            {
                button1.Enabled = true;
            }
            else if (textBox8.Text == "")
            {
                button1.Enabled = false;
            }
            else
            {
                button1.Enabled = false;
            }
            
        }

        //txbzAyliq.Text = txbzAyliq.Text.ToString().Replace('.', ',');
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            //textBox8.Text = textBox8.Text.ToString().Replace('.', ',');
            try
            {
                baglaac();
            }
            catch (Exception)
            {
                
                
            }
           
        }
    }
}
