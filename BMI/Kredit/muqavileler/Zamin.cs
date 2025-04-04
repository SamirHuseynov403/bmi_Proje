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
//using Word = Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Globalization;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class Zamin : Form
    {

        OracleCommand komutmuqavileraz;

        public Zamin()
        {
            InitializeComponent();
        }
        public string adi { get; set;}
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
        public string girovnovucombo { get; set; }
        public string krtarixi { get; set; }
        public string secilmistarix { get; set; }
        public string fin { get; set; }
        public string icraciadizaminlikde { get; set; }
        public string icraci_adi { get; set; }

        public string muqtipi { get; set; }
        public string test { get; set; }
        public string umumitar { get; set; }

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
        public static string db1 = "", db2 = "", db3 = "", db4 = "", db5 = "", db6 = "", db7 = "",c1="",sub1="";
        public static string meb1 = "", meb2 = "", meb3 = "", meb4 = "", meb5 = "", meb6 = "", meb7 = "";
        public static string teyinat1 = "", teyinat2 = "", teyinat3 = "", teyinat4 = "", teyinat5 = "", teyinat6 = "", teyinat7 = "";

        private void kataloqgetir_zaminsiz()
        {

            try
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select t.date_pog,t.summa_pog_kre ced_esas,t.summa_pog_pro ced_faiz, t.summa_pog_kre+t.summa_pog_pro ayliq from odb.graphpogkre t, odb.licschkre k where length(t.licschkre) = 20  and t.licschkre = '" + sudahes+ "' and t.subschkre = '"+subhes+"' and t.licschkre = k.licschkre and k.date_close is null", Orcon);

                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();

                dataGridView1.Columns[0].HeaderText = "Tarix";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "Qaliq";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Esasdan";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Faiz";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "Ayliq";
                dataGridView1.Columns[4].Width = 120;

            }
            catch (Exception)
            {

            }

            finally
            { }
        }
        
        Decimal krdedvelmebleg = 0;
        int say = 0;
        private void exceleattelebetest()
        {
            //try
            //{
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = desktopFolder + "\\Cedvel sablonhazir.xls";

            if (dataGridView1.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Cedvel sablon.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Cedvel sablon.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "Cedvel sablon.xls";

                    worksheet.Cells[1, 2] = adi;
                    worksheet.Cells[2, 2] = "("+subhes.ToString()+")"+sudahes;
                    worksheet.Cells[3, 2] = krtarixi.Substring(0,10)+" il tarixli "+ txbzMuqNo.Text+" saylı";
                    worksheet.Cells[8, 3] = mebleg;
                    worksheet.Cells[8, 4] = cboxkzValyuta.Text;
                    worksheet.Cells[9, 3] = faiz;
                    worksheet.Cells[10, 3] = txbzMuddet.Text;

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        try
                        {
                            
                            worksheet.Cells[i+13, 3] = Convert.ToDecimal(mebleg) - krdedvelmebleg;
                            if (say<=Convert.ToInt32(txbzMuddet.Text))
                            {
                                say = say + 1;
                                worksheet.Cells[i + 13, 1] = say;
                            }
                            
                            worksheet.Cells[i + 13, 2] = dataGridView1.Rows[i].Cells[0].Value;
                            worksheet.Cells[i + 13, 4] = dataGridView1.Rows[i].Cells[1].Value;

                            worksheet.Cells[i + 13, 5] = dataGridView1.Rows[i].Cells[2].Value;
                            worksheet.Cells[i + 13, 6] = dataGridView1.Rows[i].Cells[3].Value;
                            decimal testcedqal = Convert.ToDecimal(dataGridView1.Rows[i].Cells[1].Value.ToString());
                            if (testcedqal.ToString() == "")
                            {

                            }
                            else
                            {
                                krdedvelmebleg = krdedvelmebleg + testcedqal;
                            }
                        }
                        catch (Exception)
                        {

                           
                        }
                           
                    }
                    int settopsay=13 + Convert.ToInt32(txbzMuddet.Text);
                    decimal topesas = 0;
                    decimal topfaiz = 0;
                    decimal topayliq = 0;

                    for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                    {
                        topesas += Convert.ToDecimal(dataGridView1.Rows[i].Cells[1].Value);
                        topfaiz += Convert.ToDecimal(dataGridView1.Rows[i].Cells[2].Value);
                        topayliq += Convert.ToDecimal(dataGridView1.Rows[i].Cells[3].Value);
                    }
                    worksheet.Cells[settopsay, 3] = "Yekun:";
                    worksheet.Cells[settopsay, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    worksheet.Cells[settopsay, 4] = topesas;
                    worksheet.Cells[settopsay, 5] = topfaiz;
                    worksheet.Cells[settopsay, 6] = topayliq;

                    int setirsay = 14 + Convert.ToInt32(txbzMuddet.Text);
                     worksheet.Range[worksheet.Cells[setirsay, 2], worksheet.Cells[setirsay, 5]].Merge();
                    worksheet.Range[worksheet.Cells[setirsay+1, 2], worksheet.Cells[setirsay+1, 5]].Merge();
                    worksheet.Cells[setirsay, 2] = "Диггят: Кредит цзря юдянишляринизи банка эялмядян Пасщапай (Миллиюн) вя Е-МАНАТ терминаллары";
                    worksheet.Cells[setirsay+1, 2] = "иля наьд шякилдя,Е-МАНАТ вя Портманат васитяси иля наьдсыз(онлайн) гайдада иъра едя билярсиниз.";
                    worksheet.Cells[setirsay+2, 2] = "FİN (PİN) kodunuz "+" " +fin;
                    worksheet.Cells[setirsay + 4, 2] = "Tел.: (012)598 90 05, (050)260 54 10";
                    worksheet.Cells[setirsay + 4, 5] = "Borcalan";
                    worksheet.Cells[setirsay + 5, 2] = "Е-mail: bank@bmibaku.az";
                    worksheet.Cells[setirsay + 7, 2] = "Verilmiş kredit üzrə xidmət haqqı 1%";
                    worksheet.Cells[setirsay + 9, 2] = "Bank";
                    worksheet.Cells[setirsay + 10, 2] = "Bank Melli İran Bakı filialı";
                    worksheet.Cells[setirsay + 12, 2] = "_______________";
                    worksheet.Cells[setirsay + 12, 5] = "_______________";
                    worksheet.Range[worksheet.Cells[setirsay, 2], worksheet.Cells[setirsay+12, 5]].Font.name = "A3 Arial Azlat";

                    app.Visible = true;
                    workbook.SaveAs(desktopFolder, Type.Missing);
                    //app.Quit();
                }

            }
            //}
            //catch (Exception)
            //{


            //}
        }
        string VAL_AD;
        private void valyuta_adi()
        {
            if (valyuta=="00")
            {
                textBox2.Text = "AZN";
                textBox3.Text = "AZN";
                VAL_AD = "AZN";
            }
            else if (valyuta=="01")
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

        private void olkeadi()
        {
            if (olke == "AZ")
            {
                cboxkzOlke.Text = "Azərbaycan Respublikası";
            }
            else if (olke == "IRN")
            {
                cboxkzOlke.Text = " İran İslam Respublikası";
            }
            

        }

        private void zamincevirme()
        {
            string kassa_hesazn, kassa_hesusd, kassa_hesavro, x_h_kassaazn, x_h_kassausd, x_h_kassaavro, x_h_kreditazn, x_h_kreditusd, x_h_kreditavro;
            string x_h_kredit_azn_QR,
                 x_h_kredit_avro_QR;//, x_h_kredit_usd_QR;
           
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
            //x_h_kredit_usd_QR = "67044010050000600000";
            x_h_kredit_avro_QR = "67044020020000600000";

            teyinat1 = DateTime.Now.ToShortDateString()+" il tarixli kredit müq.əsəasən "+txbzMuddet.Text+" ay müddətinə kredit verilir";
            teyinat2=DateTime.Now.ToShortDateString()+" tar.kr.müq.əs 1% x/h tut ";
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
            kr2=carihes;
            kr3 = kassa_hesazn;
            kr4 = x_h_kassaazn;
            kr5 = x_h_kreditazn;

            meb1 = mebleg;
            double krmeb = Convert.ToInt32(mebleg)/100;
            double krxh=krmeb/2;
            meb2 = krmeb.ToString();
            meb3 = mebleg;
            meb4 = krxh.ToString();
            meb5 = krxh.ToString();

            if (textBox4.Enabled==true)
            {
                db6=kassa_hesazn;
                kr6 = carihes;
                meb6=textBox4.Text;
                teyinat6 = "Əmanət";


            }

        }
        private void zamincevirme_x_siz()
        {
            string kassa_hesazn, kassa_hesusd, kassa_hesavro, x_h_kassaazn, x_h_kassausd, x_h_kassaavro, x_h_kreditazn, x_h_kreditusd, x_h_kreditavro;
            string x_h_kredit_azn_QR, x_h_kredit_avro_QR;// x_h_kredit_usd_QR, 

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
            //x_h_kredit_usd_QR = "67044010050000600000";
            x_h_kredit_avro_QR = "67044020020000600000";

            teyinat1 = DateTime.Now.ToShortDateString() + " il tarixli kredit müq.əsəasən " + txbzMuddet.Text + " ay müddətinə kredit verilir";
            teyinat2 = DateTime.Now.ToShortDateString() + " tar.kr.müq.əs 1% x/h tut ";
            teyinat3 = "Kredit verilməsi ilə əlaqədar";
            teyinat4 = "Kassa 0,5 % x/h tutulur";
            teyinat5 = "kred.ver.əlaqədar 0,5 % x/h";

            //c1=;
            sub1 = subhes;
            db1 = sudahes;
            db2 = carihes;


            kr1 = carihes;
            kr2 = kassa_hesazn;
           

            meb1 = mebleg;
            double krmeb = Convert.ToInt32(mebleg) / 100;
            double krxh = krmeb / 2;
            meb2 = mebleg;
           

            if (textBox4.Enabled == true)
            {
                db3 = kassa_hesazn;
                kr3 = carihes;
                meb3 = textBox4.Text;
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
                Microsoft.Office.Interop.Excel.Range objRange4, objRange5, objRange6,objRange7,objRange8,objRange9, objkr1, objkr2, objkr3, objkr4, objkr5, objkr6, objmeb1, objmeb2, objmeb3, objmeb4, objmeb5, objmeb6;
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

                //objexcel.Visible = false;


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
                Microsoft.Office.Interop.Excel.Range objRange4, objRange5, objRange6, objRange7, objRange8, objRange9, objkr1, objkr3,  objmeb1, objmeb3;
                Microsoft.Office.Interop.Excel.Range objtey1, objtey2, objtey3,  objtey5, objtey6;

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
                objRange3.set_Value(System.Reflection.Missing.Value, db2);

                objRange9 = objshet.get_Range("h3", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "");
                objRange9 = objshet.get_Range("h3", System.Reflection.Missing.Value);
                objRange9.set_Value(System.Reflection.Missing.Value, "0");

               

                objkr3 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, "");
                objkr3 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, kr2);

                

                objmeb3 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, "");
                objmeb3 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, meb2);

                

                objtey3 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, "");
                objtey3 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, teyinat2);

               

                //objbook.Save();
                objbook.SaveAs(desktopFolderPK, Type.Missing);
                objexcel.Quit();

                //objexcel.Visible = false;


            }
            catch (Exception)
            {
            }

        }

        private readonly string templatefilname = @"C:\BMI_\BMI\bin\Debug\KreditZaminlikle3.docx";

        private void txbzAyliq_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //   txbzAyliq.Text = txbzAyliq.Text.ToString().Replace('.', ',');
            //    if (txbzAyliq.Text !=null)
            //    {
            //        btnkzYazdir.Enabled = true;
            //    }
            //    else if (txbzAyliq.Text =="")
            //    {
            //        btnkzYazdir.Enabled = false;
            //    }
            //}
            //catch (Exception)
            //{

            //    btnkzYazdir.Enabled = false;
            //}
            
           
        }

       

        //public string Templatefilname
        //{
        //    get
        //    {
        //        return templatefilname;
        //    }
        //}

        private string IlkHarfleriBuyut(string metin)
        {
            System.Globalization.CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Globalization.TextInfo textInfo = cultureInfo.TextInfo;
            return textInfo.ToTitleCase(metin);
        }

        private void word_at()
        {
            try
            {

                if (!System.IO.File.Exists(Application.StartupPath + "\\KreditZaminlikle3.docx"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzmebleg.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txbzAyliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzMuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzFaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzVKfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var adi = txbzBorcalan.Text;
                var passport = txbzPasport.Text;
                var pasvertarix = datezPasvertarix.Text;
                var Orqan = cboxkzOrqan.Text;
                var Unvan = txbzUnvan.Text;
                var CariHesab = txbzCarihesab.Text;
                var Valyuta = cboxkzValyuta.Text;
                var Serencam = txbzSerencam.Text;
                var MuqNo = txbzMuqNo.Text;
                var Muqtarix = tamtarix;
                var Teyinat = cboxkzTeyinat.Text;
                var Olke = cboxkzOlke.Text;
                var mebleg = txbzmebleg.Text + " " + cboxkzValyuta.Text;
                var muddet = txbzMuddet.Text;
                var faiz = txbzFaiz.Text;
                var vkfaiz = txbzVKfaiz.Text;
                var telf = txbzTelefon.Text;
                var ayliq = txbzAyliq.Text;
                var fifd = txbzFIFD.Text + " " + textBox5.Text;
                var vsudahes = sudahes;
                var vfaizhes = faizhes;
                var vvkhes = vkhes;
                var vvkhesfaiz = vkfaizhes;
                var vodgun = odgunu;
                var eht_faiz = ehtfaiz;
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
                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;
                var tamcari = tamhesab;
                //var icraci = icraci_adi;
                var icraciadi = icraciadizaminlikde;



                // TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\KreditZaminlikle3.docx");
                //var wordDocument = wordapp.Documents.Open(zamin1);
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
                ReplaceWordStub("{icraci}", icraciadi, wordDocument);

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
                //string muqadi = "yoxlanis";
                //string muqadi = adi;
                // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".docx");


            }
            catch (Exception)
            {

                
            }
            finally { }
        }

        private void word_at_zam2()
        {
            try
            {

                if (!System.IO.File.Exists(Application.StartupPath + "\\KreditZaminlikle2.docx"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzmebleg.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txbzAyliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzMuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzFaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzVKfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var adi = txbzBorcalan.Text;
                var passport = txbzPasport.Text;
                var pasvertarix = datezPasvertarix.Text;
                var Orqan = cboxkzOrqan.Text;
                var Unvan = txbzUnvan.Text;
                var CariHesab = txbzCarihesab.Text;
                var Valyuta = cboxkzValyuta.Text;
                var Serencam = txbzSerencam.Text;
                var MuqNo = txbzMuqNo.Text;
                var Muqtarix = tamtarix;
                var Teyinat = cboxkzTeyinat.Text;
                var Olke = cboxkzOlke.Text;
                var mebleg = txbzmebleg.Text + " " + cboxkzValyuta.Text;
                var muddet = txbzMuddet.Text;
                var faiz = txbzFaiz.Text;
                var vkfaiz = txbzVKfaiz.Text;
                var telf = txbzTelefon.Text;
                var ayliq = txbzAyliq.Text;
                var fifd = txbzFIFD.Text;
                var vsudahes = sudahes;
                var vfaizhes = faizhes;
                var vvkhes = vkhes;
                var vvkhesfaiz = vkfaizhes;
                var vodgun = odgunu;
                var eht_faiz = ehtfaiz;
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
                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;
                var tamcari = tamhesab;
                //var icraci = icraci_adi;
                var icraci = icraciadizaminlikde;



                // TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\KreditZaminlikle2.docx");
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
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                //wordDocument.SaveAs(@"‪\\fs\KRED_SOB\Zaminlik muqavileleri\" + muqadi + ".doc");
                //wordDocument.SaveAs(Application.StartupPath + muqadi);
                wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".doc");
            }
            catch (Exception)
            {
            }
            finally { }

        }
        private void word_at_zam3()
        {

            try
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\KreditZaminlikle3duz.docx"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                Zaminler zmnlar = new Zaminler();
                string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzmebleg.Text));
                string ayyazile = yaziyaCevir(Convert.ToDecimal(txbzAyliq.Text));
                string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzMuddet.Text));
                string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzFaiz.Text));
                string vkfaizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzVKfaiz.Text));
                //string adbalaca = txbzBorcalan.Text;
                var adi = txbzBorcalan.Text;
                var passport = txbzPasport.Text;
                var pasvertarix = datezPasvertarix.Text;
                var Orqan = cboxkzOrqan.Text;
                var Unvan = txbzUnvan.Text;
                var CariHesab = txbzCarihesab.Text;
                var Valyuta = cboxkzValyuta.Text;
                var Serencam = txbzSerencam.Text;
                var MuqNo = txbzMuqNo.Text;
                var Muqtarix = tamtarix;
                var Teyinat = cboxkzTeyinat.Text;
                var Olke = cboxkzOlke.Text;
                var mebleg = txbzmebleg.Text + " " + cboxkzValyuta.Text;
                var muddet = txbzMuddet.Text;
                var faiz = txbzFaiz.Text;
                var vkfaiz = txbzVKfaiz.Text;
                var telf = txbzTelefon.Text;
                var ayliq = txbzAyliq.Text;
                var fifd = txbzFIFD.Text;
                var vsudahes = sudahes;
                var vfaizhes = faizhes;
                var vvkhes = vkhes;
                var vvkhesfaiz = vkfaizhes;
                var vodgun = odgunu;
                var eht_faiz = ehtfaiz;
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
                var mebyaziile = yaziile;
                var mevayyazi = ayyazile;
                var tamcari = tamhesab;
                //var icraci = icraci_adi;
                var icraci = icraciadizaminlikde;



                // TODO: Word Export
                var wordapp = new Microsoft.Office.Interop.Word.Application();
                wordapp.Visible = false;



                var wordDocument = wordapp.Documents.Open(Application.StartupPath + "\\KreditZaminlikle3duz.docx");
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
                wordDocument.SaveAs(@"\\fs\KRED_SOB\Muqavileler\Kredit zaminlik\" + muqadi + ".doc");
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\Erizeler\New folder"+muqadi+".docx");
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);

                //wordDocument.SaveAs(@"‪\\fs\KRED_SOB\Zaminlik muqavileleri\" + muqadi + ".doc");
                //wordDocument.SaveAs(Application.StartupPath + muqadi);
                //‪\\192.168.0.5\kred_sob\Kredit zaminlik
                //this.Application.Documents.Open(Application.StartupPath + "adi.docx", ReadOnly: true);

                //wordDocument.Close(Type.Missing, Type.Missing, Type.Missing);
                ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler\Erize1.docx
                ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler





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

        private void button2_Click(object sender, EventArgs e)
        {
            
            Zaminler zmnlar = new Zaminler();
            //zmnlar.groupBox1.Visible = false;
            //zmnlar.groupBox2.Visible = false;
            //zmnlar.groupBox3.Visible = false;
            //if (comboBox1.Text=="Bir nəfərin zəmanəti")
            //{
            //    zmnlar.groupBox1.Visible = true;
            //}
            //else if (comboBox1.Text=="İki nəfərin zəmanəti")
            //{
            //    zmnlar.groupBox1.Visible = true;
            //    zmnlar.groupBox2.Visible = true;
            //}
            //else if (comboBox1.Text == "Üç nəfərin zəmanəti")
            //{
            //    zmnlar.groupBox1.Visible = true;
            //    zmnlar.groupBox2.Visible = true;
            //    zmnlar.groupBox3.Visible = true;
            //}
            //zmnlar.txbzam1adi.Text = zam1adi;
            //zmnlar.txbzam1Pasport.Text = zam1pass;
            //zmnlar.txbzam1Telefon.Text = zam1tel;
            //zmnlar.txbzam1Unvan.Text = zam1unvan;

            //zmnlar.ShowDialog();
           // word_at();
            
           


        }
        private void teyadi()
        {
            if (teyinat == "2001")
            {
                cboxkzTeyinat.Text = "mənzil təmiri";
            }
            else if (teyinat == "2002")
            {
                cboxkzTeyinat.Text = " Avtomobil alınması";
            }
            else if (teyinat == "2003")
            {
                cboxkzTeyinat.Text = " məişət əşyalarının alınması";
            }

        }
        
        private void Zamin_Load(object sender, EventArgs e)
        {
            label24.Text = test;

            radioButton1.Checked = true;

            //if (txbzAyliq.Text=="")
            //{
            //    btnkzYazdir.Enabled = false;
            //}

            
            if (comboadi=="Progress")
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

                txbzBorcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);
                
                //txbzBorcalan.Text = txbzBorcalan.Text.ToUpper();
                txbzmebleg.Text = mebleg;
                txbzPasport.Text = seriyano;
                txbzTelefon.Text = mobil;
                unvanhazir = unvan.ToLower();
                txbzUnvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);
                txbzVKfaiz.Text = vkfaiz ;
                txbzFaiz.Text = faiz ;
                cboxkzOrqan.Text = ver_orqan;
                txbzCarihesab.Text = carihes;
                txbzAyliq.Text = ayliq;
                
                txbzFIFD.Text = fifd;
                datezPasvertarix.Text = ver_tar;
                int gun = Convert.ToInt32(muddet);
                int gunsay = gun / 30;
                txbzMuddet.Text = gunsay.ToString();
                //cboxkzTeyinat.Text = teyinat; 
                if (valyuta=="00")
                {
                    cboxkzValyuta.Text = "AZN";
                }
                else if (valyuta=="01")
                {
                    cboxkzValyuta.Text = "USD";
                }
                else if (valyuta == "02")
                {
                    cboxkzValyuta.Text = "AVRO";
                }
            }

            kataloqgetir_zaminsiz();
            
            
            
        }
        private void hey_sig()
        {
            
        }
        int kr_say;
        int ser_say;
        public int zam_say;
        public void muqavile_nom()

        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            label23.Text = tarixIl;
            

            ////string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand komut = new OracleCommand();
            komut.Connection = con;
            komut.CommandText = "select kr_zaminlik,kr_serencam,kr_zaminler,il from odb.muqavile_nomreleri where il='" + tarixIl + "'";
            OracleDataReader dr = komut.ExecuteReader();
            
            //txbzMuqNo.Text = "1";
            //zam_say = 1;
            while (dr.Read())
            {

                    zam_say = Convert.ToInt32(dr["kr_zaminler"].ToString()) + 1;
                    kr_say = Convert.ToInt32(dr["kr_zaminlik"].ToString()) + 1; 
                 
                    txbzMuqNo.Text = kr_say.ToString();
                    ser_say = Convert.ToInt32(dr["kr_serencam"].ToString()) + 1;
                    txbzSerencam.Text = ser_say.ToString();
                    
             
            }
            
            
            con.Close();

            //string sql = "select il from odb.muqavile_nomreleri  where il=@il";
            //(sql,baglanti);
            //sorgu.Parameters.AddWithValue("@il", satirsayi);

            ////OracleConnection connection = new OracleConnection(connectrionString);
            ////OracleCommand orcmd = new OracleCommand("select from odb.muqavile_nomreleri where il='2021'", connection);
            //baglanti.Open();
            //OracleDataReader sonuc = sorgu.ExecuteReader();
            //while (sonuc.Read())
            //{
            //    txbzMuqNo.Text=sonuc["kr_zaminlik"].ToString();
            //}
            //baglanti.Close();
        }

        public void muqavile_nom1()
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
            komut.CommandText = "select kr_zaminler from odb.muqavile_nomreleri where il='"+tarixIl+"'";
            OracleDataReader dr = komut.ExecuteReader();


            while (dr.Read())
            {



                zam_say = Convert.ToInt32(dr["kr_zaminler"].ToString()) + 1;
                
                


            }


            con.Close();

            //string sql = "select il from odb.muqavile_nomreleri  where il=@il";
            //(sql,baglanti);
            //sorgu.Parameters.AddWithValue("@il", satirsayi);

            ////OracleConnection connection = new OracleConnection(connectrionString);
            ////OracleCommand orcmd = new OracleCommand("select from odb.muqavile_nomreleri where il='2021'", connection);
            //baglanti.Open();
            //OracleDataReader sonuc = sorgu.ExecuteReader();
            //while (sonuc.Read())
            //{
            //    txbzMuqNo.Text=sonuc["kr_zaminlik"].ToString();
            //}
            //baglanti.Close();
        }

        private void btnkzYazdir_Click(object sender, EventArgs e)
        {
            if (txbzAyliq.Text == "")
            {
                MessageBox.Show("Aylıq ödəniş qeyd edilməyib!!!", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {



                Zaminler zmnlar = new Zaminler();

                if (comboBox1.Text == "")
                {
                    MessageBox.Show("Zamin məlumatları seçilməyib", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }



                else if (comboBox1.Text == "Bir nəfərin zəmanəti")
                {
                    muqavile_nom();
                    muq_no_at();
                    word_at();

                    if (radioButton2.Checked == true)
                    {
                        zamincevirme_x_siz();
                        excele_at_x_hsiz();
                    }
                    else if (radioButton1.Checked == true)
                    {
                        zamincevirme();
                        excele_at();
                    }

                }
                else if (comboBox1.Text == "İki nəfərin zəmanəti")
                {

                    muqavile_nom();
                    muq_no_at();
                    word_at_zam2();
                    if (radioButton2.Checked == true)
                    {
                        zamincevirme_x_siz();
                        excele_at_x_hsiz();
                    }
                    else if (radioButton1.Checked == true)
                    {
                        zamincevirme();
                        excele_at();
                    }

                }
                else if (comboBox1.Text == "Üç nəfərin zəmanəti")
                {
                    muqavile_nom();
                    muq_no_at();
                    word_at_zam3();
                    if (radioButton2.Checked == true)
                    {
                        zamincevirme_x_siz();
                        excele_at_x_hsiz();
                    }
                    else if (radioButton1.Checked == true)
                    {
                        zamincevirme();
                        excele_at();
                    }
                }
                else if (txbzAyliq.Text == "")
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb.");
                }
                else if (textBox5.Text == "")
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb.");
                }
                else if (txbzCarihesab.Text == "")
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb.");
                }
            }
            //exceleattelebetest();
            
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox4.Enabled = true;
                label20.Enabled = true;
            }
            else if (checkBox1.Checked == false)
            {
                textBox4.Enabled = false;
                label20.Enabled = false;
            }
        }
        int yencavab = 0;
        private void muq_no_at()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand OCOM = new OracleCommand("Update odb.muqavile_nomreleri set kr_serencam='" + ser_say + "',kr_zaminlik='" + kr_say + "',kr_zaminler='" + zam_say + "' where IL='" + tarixIl + "'", con);
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


            //try
            //{
               // OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
               // int test = 2021;
               // //kr_zaminlik,kr_serencam,kr_zaminler,il
               // string sorgu5 = "Update odb.muqavile_nomreleri set kr_serencam='" + ser_say + "',kr_zaminlik='" + kr_say + "',kr_zaminler='" + zam_say + "' where IL='2021'";
               // komutmuqavileraz = new OracleCommand(sorgu5, con);
               //komutmuqavileraz.Parameters.AddWithValue("@1", ser_say);
               // komutmuqavileraz.Parameters.AddWithValue("@2", 948);
               //komutmuqavileraz.Parameters.AddWithValue("@3", zam_say);
               //     con.Open();
               //     komutmuqavileraz.ExecuteNonQuery();
               //     con.Close();
                

                

                //string sorgu5 = "Update muqavile_nomreleri set Adi=@Adi,Valyuta=@Valyuta,Muraciet_mebleg=@Muraciet_mebleg,Raziliq_mebleg=@Raziliq_mebleg,Teminat=@Teminat,Etiraz_sebebi=@Etiraz_sebebi,Qeyd=@Qeyd,Komite_qerar=@Komite_qerar,Qerar_qeyd=@Qerar_qeyd,Danisiq=@Danisiq,Telefon=@Telefon,Melumat=@Melumat where Sira=" + txbmursira.Text + "";
                //komutmurraz = new OracleCommand(sorgu5,con );
                ////komutmur.Parameters.AddWithValue("@sira", txbmursira.Text);
                ////komutmur.Parameters.AddWithValue("@tarix", dtpmur);
                //komutmurraz.Parameters.AddWithValue("@Adi", txbmurad.Text);
                //komutmurraz.Parameters.AddWithValue("@Valyuta", cmbmurvaly.Text);
                //komutmurraz.Parameters.AddWithValue("@Muraciet_mebleg", txbmurmeb.Text);
                //komutmurraz.Parameters.AddWithValue("@Raziliq_mebleg", textBox14.Text);
                //komutmurraz.Parameters.AddWithValue("@Teminat", cmbmurtem.Text);
                //komutmurraz.Parameters.AddWithValue("@Etiraz_sebebi", cmbmurseb.Text);
                //komutmurraz.Parameters.AddWithValue("@Qeyd", txbmurqeyd.Text);
                //komutmurraz.Parameters.AddWithValue("@Komite_qerar", cmbmurqerar.Text);
                //komutmurraz.Parameters.AddWithValue("@Qerar_qeyd", txbmurqerqeyd.Text);
                //komutmurraz.Parameters.AddWithValue("@Danisiq", cmbmurdan.Text);
                //komutmurraz.Parameters.AddWithValue("@Telefon", txbmurtel.Text);
                //komutmurraz.Parameters.AddWithValue("@Melumat", cmbmurmel.Text);

                
            }
           
        private void acbagla()
        {
            
        }

        private void cboxkzMuqNo_CheckedChanged(object sender, EventArgs e)
        {
            muqavile_nom();
        }

        private void btn_NOM_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txbzAyliq.Text == "")
            {
                MessageBox.Show("Aylıq ödəniş qeyd edilməyib!!!", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {


                int zaminlarsay = zam_say;


                Zaminler zmnlar = new Zaminler();
                zmnlar.zmnlartarix = umumitar;
                zmnlar.muqtipi = label24.Text;
                //zmnlar.muqtipi = girovnovucombo;
                zmnlar.muqtipi = muqtipi;
                zmnlar.comboadi = comboBox1.Text;
                zmnlar.muqtipi = label24.Text;
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

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void exceleattelebe()
        {

            //saveFileDialog1.FileName = "Personeller (" + DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year + ")";
            //saveFileDialog1.Filter = "XLS Dosyaları (*.xls)|*.xls";

            //saveFileDialog1.InitialDirectory = "c:";

            ////eğer saveFileDiaolog1 açıldığında Evet’e tıklanırsa

            //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            //{

            //    DevExpress.XtraPrinting.XlsExportOptions _Options = new DevExpress.XtraPrinting.XlsExportOptions();

            //    _Options.SheetName = "Ümumi kataloq(" + DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year + ")";

            //    dataGridView1.ExportToXls(saveFileDialog1.FileName, _Options);

            //    if (MessageBox.Show("Yüklədiyiniz excel faylını açmaq istəyirsiniz?", "Excel dosyası", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //    {

            //        //Kaydedilen Excel Dosyasını açar.

            //        System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            //    }

            //}
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //exceleattelebetest();
            if (radioButton2.Checked == true)
            {
                zamincevirme_x_siz();
                excele_at_x_hsiz();
            }
            else if (radioButton1.Checked == true)
            {
                zamincevirme();
                excele_at();
            }

        }
    }
}
