using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;
using BMI.Muhasibat;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Word;

namespace BMI
{
    public partial class Pul_Kocurmesi : Form
    {
        
        public Form1 frman;
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();

        private readonly string templatefilname = @"‪‪C:\\BMI_\\BMI\\bin\\Debug\\Erizeler\\Erize1.docx";
        //C:\FormPK\FormPK\bin\Debug\Erize.docx
        public static string gedenbilgi = "", kr1 = "", kr2 = "", kr3 = "", kr4 = "", kr5 = "", kr6 = "", kr7 = "";
        public static string db1 = "", db2 = "", db3 = "", db4 = "", db5 = "", db6 = "", db7 = "";
        public static string meb1 = "", meb2 = "", meb3 = "", meb4 = "", meb5 = "", meb6 = "", meb7 = "";
        public static string teyinat1 = "", teyinat2 = "", teyinat3 = "", teyinat4 = "", teyinat5 = "", teyinat6 = "", teyinat7 = "";
        public string Xaricmektbtarix = string.Empty;
        public string icraci_kod { get; set; }
        string tamad = "";
        public Pul_Kocurmesi()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }


        string tarixIl = DateTime.Now.Date.Year.ToString().Substring(2, 2);
        //Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\bin\Debug\Irana_kocurmeler.mdb
        OleDbConnection baglanmaq = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\User\Desktop\PK.xlsx;Extended Properties=Excel 12.0 Xml; HDR=YES;");
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.5\12345\Personal\Iran kocurmesi\EL VURMA\Irana_kocurmeler.mdb");

        System.Data.DataTable cedvel = new System.Data.DataTable();

        public string gadi, gsoyadi, gataadi, gpassport, gtelefon, elave, meqsed, aadi, asoyadi, aataadi,  apassport, atelefon, bankadi, filial, ahesab;
        private void button1_Click(object sender, EventArgs e)
        {
            temizle(this);
        }

        private void listele()
        {
            baglanti.Open();
            OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Siyahilar", baglanti);
            isleme.Fill(cedvel);
            dataGridView1.DataSource = cedvel;
            baglanti.Close();
        }
        private void excele_esas_ilk()
        {
            try
            {
                
                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string desktopFolderPK = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopFolder = desktopFolder + "\\Kocur.xls";
                desktopFolderPK = desktopFolderPK + "\\PK_Kocur.xls";
                Microsoft.Office.Interop.Excel.Application objexcel = new Microsoft.Office.Interop.Excel.Application();
                //objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook sheet = objexcel.Workbooks.Open(desktopFolder);
                objexcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook objbook = objexcel.Workbooks.Open(desktopFolder);
                Microsoft.Office.Interop.Excel.Worksheet objshet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range objRange;
                Microsoft.Office.Interop.Excel.Range objRange2;
                Microsoft.Office.Interop.Excel.Range objRange3;
                Microsoft.Office.Interop.Excel.Range objRange4, objRange5, objRange6, objkr1, objkr2, objkr3, objkr4, objkr5, objkr6, objmeb1, objmeb2, objmeb3, objmeb4, objmeb5, objmeb6;
                Microsoft.Office.Interop.Excel.Range objtey1, objtey2, objtey3, objtey4, objtey5, objtey6;
                objRange = objshet.get_Range("d2", System.Reflection.Missing.Value);
                objRange.set_Value(System.Reflection.Missing.Value, db1);
                objRange2 = objshet.get_Range("d3", System.Reflection.Missing.Value);
                objRange2.set_Value(System.Reflection.Missing.Value, db2);
                objRange3 = objshet.get_Range("d4", System.Reflection.Missing.Value);
                objRange3.set_Value(System.Reflection.Missing.Value, db3);
                objRange4 = objshet.get_Range("d5", System.Reflection.Missing.Value);
                objRange4.set_Value(System.Reflection.Missing.Value, db4);
                objRange5 = objshet.get_Range("d6", System.Reflection.Missing.Value);
                objRange5.set_Value(System.Reflection.Missing.Value, db5);
                objRange6 = objshet.get_Range("d7", System.Reflection.Missing.Value);
                objRange6.set_Value(System.Reflection.Missing.Value, db6);
                objkr1 = objshet.get_Range("f2", System.Reflection.Missing.Value);
                objkr1.set_Value(System.Reflection.Missing.Value, kr1);
                objkr2 = objshet.get_Range("f3", System.Reflection.Missing.Value);
                objkr2.set_Value(System.Reflection.Missing.Value, kr2);
                objkr3 = objshet.get_Range("f4", System.Reflection.Missing.Value);
                objkr3.set_Value(System.Reflection.Missing.Value, kr3);
                objkr4 = objshet.get_Range("f5", System.Reflection.Missing.Value);
                objkr4.set_Value(System.Reflection.Missing.Value, kr4);
                objkr5 = objshet.get_Range("f6", System.Reflection.Missing.Value);
                objkr5.set_Value(System.Reflection.Missing.Value, kr5);
                objkr6 = objshet.get_Range("f7", System.Reflection.Missing.Value);
                objkr6.set_Value(System.Reflection.Missing.Value, kr6);
                objmeb1 = objshet.get_Range("g2", System.Reflection.Missing.Value);
                objmeb1.set_Value(System.Reflection.Missing.Value, meb1);
                objmeb2 = objshet.get_Range("g3", System.Reflection.Missing.Value);
                objmeb2.set_Value(System.Reflection.Missing.Value, meb2);
                objmeb3 = objshet.get_Range("g4", System.Reflection.Missing.Value);
                objmeb3.set_Value(System.Reflection.Missing.Value, meb3);
                objmeb4 = objshet.get_Range("g5", System.Reflection.Missing.Value);
                objmeb4.set_Value(System.Reflection.Missing.Value, meb4);
                objmeb5 = objshet.get_Range("g6", System.Reflection.Missing.Value);
                objmeb5.set_Value(System.Reflection.Missing.Value, meb5);
                objmeb6 = objshet.get_Range("g7", System.Reflection.Missing.Value);
                objmeb6.set_Value(System.Reflection.Missing.Value, meb6);
                objtey1 = objshet.get_Range("j2", System.Reflection.Missing.Value);
                objtey1.set_Value(System.Reflection.Missing.Value, teyinat1);
                objtey2 = objshet.get_Range("j3", System.Reflection.Missing.Value);
                objtey2.set_Value(System.Reflection.Missing.Value, teyinat2);
                objtey3 = objshet.get_Range("j4", System.Reflection.Missing.Value);
                objtey3.set_Value(System.Reflection.Missing.Value, teyinat3);
                objtey4 = objshet.get_Range("j5", System.Reflection.Missing.Value);
                objtey4.set_Value(System.Reflection.Missing.Value, teyinat4);
                objtey5 = objshet.get_Range("j6", System.Reflection.Missing.Value);
                objtey5.set_Value(System.Reflection.Missing.Value, teyinat5);
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
        private void excele_at()
        {
            try
            {
                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string sourceFile = Path.Combine(desktopFolder, "Kocur.xls");
                string destinationFile = Path.Combine(desktopFolder, "PK_Kocur.xls");

                var excelApp = new Microsoft.Office.Interop.Excel.Application { Visible = true };
                var workbook = excelApp.Workbooks.Open(sourceFile);
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                // Data to be written into cells
                object[] dbValues = { db1, db2, db3, db4, db5, db6 };
                object[] krValues = { kr1, kr2, kr3, kr4, kr5, kr6 };
                object[] mebValues = { meb1, meb2, meb3, meb4, meb5, meb6 };
                object[] teyinatValues = { teyinat1, teyinat2, teyinat3, teyinat4, teyinat5, teyinat6 };

                // Write values to ranges using a loop
                WriteValuesToRange(worksheet, "D", dbValues);
                WriteValuesToRange(worksheet, "F", krValues);
                WriteValuesToRange(worksheet, "G", mebValues);
                WriteValuesToRange(worksheet, "J", teyinatValues);

                workbook.SaveAs(destinationFile);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                // Handle exception
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        private void WriteValuesToRange(Microsoft.Office.Interop.Excel.Worksheet worksheet, string column, object[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                string cell = column + (i + 2); // Start from row 2
                worksheet.Range[cell].Value = values[i];
            }
        }
        private void Pul_Kocurmesi_Load_1(object sender, EventArgs e)
        {
            hevnom_al();
            Form1 frm111 = new Form1();
            //string adi = frm111.lblTamad.Text;

            // TODO: This line of code loads data into the 'irana_kocurmelerDataSet5.Siyahilar' table. You can move, or remove it, as needed.
            this.siyahilarTableAdapter.Fill(this.irana_kocurmelerDataSet5.Siyahilar);
            satirsayisinitap();
            listele();
            radioButton1.Checked = true;
            rdb_hesab_acmadan.Checked = true;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
            radioButton7.Enabled = false;
            //hevnom_al();
            gedenhevaleyeni frmyged = new gedenhevaleyeni(this);
            //frmyged.hevnom_al();
            //txbhevale.Text = frmyged.txbgedHNo.Text;
            label31.Text = DateTime.Now.Date.ToShortDateString();
            label32.Text = DateTime.Now.Year.ToString();
            cekboxac();

            rdb();
            listele();
            sifaris_mebleg();
            //meblegitap();
            valyutaadi();
            hesab_tap();
        }

        private void rdb_hesabdan_CheckedChanged_1(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = true;
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void rdb_hesab_ve_medaxil_CheckedChanged_1(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = true;
        }

        private void rdb_hesab_acmadan_CheckedChanged_1(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = false;
        }

        private void radioButton2_CheckedChanged_2(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
        }
        private void hesabi_tap()
        {
            //if (radioButton1.Checked==true)
            //{
            //    txb_musteri_hesabi.Text = "45023010010000400000";
            //}
            //else if (radioButton2.Checked == true)
            //{
            //    txb_musteri_hesabi.Text = "45023020010000400000";
            //}
            //else if (radioButton3.Checked == true)
            //{
            //    txb_musteri_hesabi.Text = "45023040010000400000";
            //}
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged_4(object sender, EventArgs e)
        {
            //txb_musteri_hesabi.Text = "45023010010000400000";
        }

        private void radioButton2_CheckedChanged_4(object sender, EventArgs e)
        {
            //txb_musteri_hesabi.Text = "45023020010000400000";
        }

        private void radioButton3_CheckedChanged_2(object sender, EventArgs e)
        {
            //txb_musteri_hesabi.Text = "45023040010000400000";
        }

        private void radioButton2_CheckedChanged_5(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
            txb_Rial_cbar.Text = "";
            txb_valyuta_cbar.Text = "";
            txb_iran_rial.Text = "";
            txb_Rial_cbar.Enabled = false;
            txb_valyuta_cbar.Enabled = false;
            txb_iran_rial.Enabled = false;
            if (rdb_hesab_acmadan.Checked == true && radioButton2.Checked == true)
            {
                txb_musteri_hesabi.Text = "45023020010000400000";
            }
            
        }

        private void radioButton1_CheckedChanged_5(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
            txb_Rial_cbar.Text = "";
            txb_valyuta_cbar.Text = "";
            txb_iran_rial.Text = "";
            txb_Rial_cbar.Enabled = false;
            txb_valyuta_cbar.Enabled = false;
            txb_iran_rial.Enabled = false;
            if (rdb_hesab_acmadan.Checked == true && radioButton1.Checked == true)
            {
            txb_musteri_hesabi.Text = "45023010010000400000";
            }
            
        }

        private void radioButton3_CheckedChanged_3(object sender, EventArgs e)
        {
            radioButton5.Enabled = true;
            radioButton6.Enabled = true;
            radioButton7.Enabled = true;
            txb_Rial_cbar.Enabled = true;
            txb_valyuta_cbar.Enabled = true;
            txb_iran_rial.Enabled = true;
            txb_musteri_hesabi.Text = "45023040010000400000";
            sifaris_mebleg();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            meb_tap();
            valyutaadi();
            //alinanv();
            ged_hev gdn = new ged_hev(this);
            gdn.Show();
            gdn.txbgedmebleg.Text = txb_mebleg.Text;
            gdn.cmbgedvalyuta.Text = label24.Text;
            gdn.textBox1.Text = labelreqemile.Text;
            gdn.comboBox1.Text = label15.Text;
            gdn.txbgedSAA.Text = txb_g_soyadı.Text + " " + txb_g_adi.Text + " " + txb_g_ataadi.Text;
            gdn.txbgedalanb.Text = txbbankadi.Text + " " + txbfilial.Text;
            gdn.txbgedHesNo.Text = txb_musteri_hesabi.Text;
            if (radioButton1.Checked == true)
            {
                gdn.cmbgedvalyuta.Text = "USD";
            }
            else if (radioButton2.Checked == true)
            {
                gdn.cmbgedvalyuta.Text = "AVRO";
            }
            //alinanv();
            //satilanv();
            //valyutaadi();
        }

        private void radioButton1_CheckedChanged_2(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
        }

        private void rdb_hesab_acmadan_CheckedChanged(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = false;
            rdb();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //textBox1.Text = dataGridView1.CurrentRow.Cells["say"].Value.ToString();
        }

        private void btncerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
           // textBox2.Text = dataGridView1.CurrentRow.Cells["say"].Value.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
            uyg.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                myrange.Value2 = dataGridView1.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myrange.Value2 = dataGridView1[i, j].Value;
                }
            }




        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void ıconButton1_Click_1(object sender, EventArgs e)
        {
            temizle(this);
            //gedenhevale frmg = new gedenhevale();
            //frmg.Show();
        }
        private void txb_mebleg_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            txb_musteri_hesabi.Text = "45023020010000400000";
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {
            radioButton5.Checked = false;
            radioButton6.Checked = false;
            txb_musteri_hesabi.Text = "45023010010000400000";
        }
        private void sifaris_mebleg()
        {
            if (radioButton3.Checked == true)
            {
                label39.Text = txb_mebleg.Text;
            }
            else label39.Text = "";

        }
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            txb_g_adi.Text = dataGridView1.CurrentRow.Cells["adi"].Value.ToString();
            txb_g_soyadı.Text = dataGridView1.CurrentRow.Cells["soyadi"].Value.ToString();
            txb_g_ataadi.Text = dataGridView1.CurrentRow.Cells["ata_adi"].Value.ToString();
            txb_gond_passport.Text = dataGridView1.CurrentRow.Cells["gond_passport"].Value.ToString();
            txbtelefon.Text = dataGridView1.CurrentRow.Cells["gond_telefon"].Value.ToString();
            comboelave.Text = dataGridView1.CurrentRow.Cells["ELAVE"].Value.ToString();
            combomeqsed.Text = dataGridView1.CurrentRow.Cells["MEQSED"].Value.ToString();
            txbalanadi.Text = dataGridView1.CurrentRow.Cells["alan_adi"].Value.ToString();
            txbalanataadi.Text = dataGridView1.CurrentRow.Cells["alan_ata_adi"].Value.ToString();
            txbalansoyadi.Text = dataGridView1.CurrentRow.Cells["alan_soyadi"].Value.ToString();
            txbalanpassport.Text = dataGridView1.CurrentRow.Cells["alan_passport"].Value.ToString();
            txbalantelefon.Text = dataGridView1.CurrentRow.Cells["alan_telefon"].Value.ToString();
            txbbankadi.Text = dataGridView1.CurrentRow.Cells["bankin_adi"].Value.ToString();
            txbfilial.Text = dataGridView1.CurrentRow.Cells["filial"].Value.ToString();
            txbalanhesab.Text = dataGridView1.CurrentRow.Cells["alanin_hesabi"].Value.ToString();
            textBox1.Text = dataGridView1.CurrentRow.Cells["say"].Value.ToString();
        }

        private void meblegitap()
        {
            try
            {
                double labelmbl, mbl1, imbl;
                labelmbl = 0;
                if (txb_mebleg.Text == "")
                {
                    //MessageBox.Show("Məbləğ daxil edilməyib");
                    MessageBox.Show("Məbləğ daxil edilməyib...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //txb_mebleg.Text.
                }
                else
                {
                    mbl1 = Convert.ToDouble(txb_mebleg.Text);
                    //imbl = Convert.ToSingle(txb_iran_rial.Text);
                    if (txb_iran_rial.Text == "")
                    {
                        labelyaziile.Text = yaziyaCevir(Convert.ToDecimal(txb_mebleg.Text));
                        labelreqemile.Text = txb_mebleg.Text;
                    }

                    else
                    {
                        labelmbl = mbl1 * Convert.ToDouble(txb_iran_rial.Text);
                        labelyaziile.Text = yaziyaCevir(Convert.ToDecimal(labelmbl.ToString()));
                        labelreqemile.Text = labelmbl.ToString();
                        labelrialmeblegi.Text = labelmbl.ToString();
                    }
                }

            }
            catch (Exception)
            {


            }
        }
        private void cekboxac()
        {
            if (radioButton3.Checked == true)
            {
                radioButton5.Enabled = true;
                radioButton6.Enabled = true;
            }
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void alinanvaltap()
        { 
        
        }
  
        private void button2_Click(object sender, EventArgs e)
        {
            tamad = txb_g_soyadı.Text + " " + txb_g_adi.Text + " " + txb_g_ataadi.Text;
            //meblegitap();
            //hevnom_al();
            
            
            ////try
            ////{
            if (checkBox2.Checked == false)
            {


                if (txb_mebleg.Text == "")
                {
                    //MessageBox.Show("Məbləğ daxil edilməyib", "Qeyd");
                    MessageBox.Show("Məbləğ daxil edilməyib...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //txb_mebleg.Text.
                }
                if (txb_musteri_hesabi.Text == "")
                {
                    //MessageBox.Show("Hesab daxil edilməyib", "Qeyd");
                    MessageBox.Show("Məbləğ daxil edilməyib...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            //if (radioButton3.Checked = true)
                //{
                //    if (txb_Rial_cbar.Text == "" || txb_valyuta_cbar.Text == "" || txb_iran_rial.Text == "")
                //    {
                //        txb_Rial_cbar.BackColor = Color.Yellow;
                //        txb_valyuta_cbar.BackColor = Color.Yellow;
                //        txb_iran_rial.BackColor = Color.Yellow;

            //        MessageBox.Show("Valyuta məlumatlarını tam daxil etməmisiniz");
                //    }


                else
                {
                    txb_Rial_cbar.BackColor = Color.White;
                    txb_valyuta_cbar.BackColor = Color.White;
                    txb_iran_rial.BackColor = Color.White;

                    baglanti.Open();
                    OleDbCommand komut = new OleDbCommand("insert into Siyahilar(Tarix,Adi,Soyadi,Ata_adi,Gond_passport,Mebleg,Rial_Cbar,Valyuta_Cbar,Iran_Rial,HevaleNo,Gond_telefon,Elave,Meqsed,Alan_adi,Alan_ata_adi,Alan_soyadi,Alan_passport,Alan_Telefon,Bankin_adi,Filial,Alanin_hesabi,say) values('" + label31.Text + "','" + txb_g_adi.Text + "','" + txb_g_soyadı.Text + "','" + txb_g_ataadi.Text + "','" + txb_gond_passport.Text + "','" + txb_mebleg.Text + "','" + txb_Rial_cbar.Text + "','" + txb_valyuta_cbar.Text + "','" + txb_iran_rial.Text + "','" + txbhevale.Text + "','" + txbtelefon.Text + "','" + comboelave.Text + "','" + combomeqsed.Text + "','" + txbalanadi.Text + "','" + txbalansoyadi.Text + "','" + txbalanataadi.Text + "','" + txbalanpassport.Text + "','" + txbalantelefon.Text + "','" + txbbankadi.Text + "','" + txbfilial.Text + "','" + txbalanhesab.Text + "','" + textBox1.Text + "')", baglanti);
                    komut.ExecuteNonQuery();
                    baglanti.Close();
                   // MessageBox.Show("Məlumatlar daxil edildi", "Qeyd");
                    MessageBox.Show("Məlumatlar daxil edildi...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cevirme();
                    valyutaadi();
                    alinanv();
                    satilanv();
                    cedvel.Clear();
                    listele();
                    meblegitap();
                    excele_at();
                    word_at();
                    satirsayisinitap();
                }
            }

            else if (checkBox2.Checked == true)
            {
                if (txb_mebleg.Text == "")
                {
                    //MessageBox.Show("Məbləğ daxil edilməyib");
                    MessageBox.Show("Məbləğ daxil edilməyib...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //txb_mebleg.Text.
                }
                else
                {
                    cevirme();
                    valyutaadi();
                    alinanv();
                    satilanv();
                    cedvel.Clear();
                    listele();
                    meblegitap();
                    excele_at();
                    word_at();
                    checkBox2.Checked = false;
                }
            }

            //}
            // catch (Exception)
            // {

            //    MessageBox.Show("Səhv oldu");
            //}
            hevnoat();
            temizle(this);

        }
        private void alinanv()
        {
            label24.Text = "";
            if (radioButton3.Checked == true)
            {
                valyutaadi();
            }

            if (radioButton5.Checked == true)
            {
                label24.Text = "ABŞ dolları";
            }
            else if (radioButton6.Checked == true)
            {
                label24.Text = "AVRO";
            }
            else if (radioButton7.Checked == true)
            {
                label24.Text = "AZN";
            }
        }
        private void satilanv()
        {
            label28.Text = "";
            label30.Text = "";
            if (radioButton3.Checked == true)
            {
                label28.Text = "İran rialı";
                label30.Text = txb_iran_rial.Text;
            }
        }
        private void valyutaadi()
        {
            if (radioButton1.Checked == true)
            {
                label15.Text = "ABŞ dolları";
            }
            else if (radioButton2.Checked == true)
            {
                label15.Text = "AVRO";
            }
            else if (radioButton3.Checked == true)
            {
                label15.Text = "İran Rialı";
            }
        }
        private void rdb_hesabdan_CheckedChanged(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = true;
        }

        private void rdb_hesab_ve_medaxil_CheckedChanged(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Visible = true;
        }
        private void rdb()
        {
            if (rdb_hesab_acmadan.Checked == true && radioButton3.Checked == true)
            {
                radioButton5.Checked = false;
                radioButton6.Checked = false;
                radioButton5.Enabled = true;
                radioButton6.Enabled = true;
            }
            else
            {
                radioButton5.Checked = false;
                radioButton6.Checked = false;
                radioButton5.Enabled = false;
                radioButton6.Enabled = false;
            }


        }

        private void cevirme()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString().Substring(2,2);
            float mbl = 0f, irial, vcbar = 0f, rcbar = 0f, cvb, gedrial = 0f, vcbar_bol_rcbar = 0f, olke_rial_gedenrial = 0f, zerer_gelir = 0f, olke_rial_gedenrial_ferq = 0f;
            int test;
            //olke_rial_gedenrial o durki olke daxili hesablanan realnan irandan verilen kurs arasindaki ferqdir
            float x_h = 10f;
            float x_h_cem = 0f;
            string kusd, keuro,kazn, hacusd, haceuro,haazn, xvalqisatqi, hacIIR, kochesab, mushesolanda, xhusd, xheuro,xhazn, zerer, gelir, husd, heuro;
            string tey1, tey2, tey3, tey4, tey5, tey6, tey7, tey8, tey9, tey10;
            kusd = "10020010000000100000";
            keuro = "10020020000000100000";
            kazn = "10010000000000100000";
            hacusd = "45023010010000400000";
            haceuro = "45023020020000400000";
            haazn = "45013000000000400001";
            xvalqisatqi = "45011000010000400000";
            hacIIR = "45023040020000400000";
            kochesab = "15025040001451500000";
            mushesolanda = "45021040000000400000";
            xhusd = "67023010000000600000";
            xheuro = "67023020000000600000";
            xhazn = "67013000000000600000";
            zerer = "88010000000001200000";
            gelir = "68010000000001200000";
            husd = "35025010000001600000";
            heuro = "35025020000001600000";

            tey1 = "köçürmə və x/h üçün mədaxil";
            tey2 = txbbankadi.Text+" "+txbfilial.Text;
            tey3 = txbbankadi.Text + " " + txbfilial.Text;
            tey4 = "(" + txb_g_adi.Text + " " + txb_g_soyadı.Text + " " + txb_g_ataadi.Text + ")";
            tey5 = txbhevale.Text;
            tey6 = combomeqsed.Text;

            try
            {
                mbl = Convert.ToSingle(txb_mebleg.Text);

                if (rdb_hesab_acmadan.Checked == true && radioButton1.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    db1 = kusd;
                    db2 = hacusd;
                    db3 = hacusd;
                    kr1 = hacusd;
                    kr2 = husd;
                    kr3 = xhusd;
                    meb1 = x_h_cem.ToString();
                    meb2 = mbl.ToString();
                    meb3 = x_h.ToString();
                    teyinat1 = tey1 + " " + tey4;
                    teyinat2 = tey5 + " " + tey2 + " " + tey3 + " " + tey4;
                    teyinat3 = tey5 + " " + tey2 + " " + tey3 + " " + tey4 + " x/h";
                }

                else if (rdb_hesab_acmadan.Checked == true && radioButton2.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    db1 = keuro;
                    db2 = haceuro;
                    db3 = haceuro;
                    kr1 = haceuro;
                    kr2 = heuro;
                    kr3 = xheuro;
                    meb1 = x_h_cem.ToString();
                    meb2 = mbl.ToString();
                    meb3 = x_h.ToString();
                    teyinat1 = tey1 + " " + tey4;
                    teyinat2 = tey5 + " " + tey2 + " " + tey3 + " " + tey4;
                    teyinat3 = tey5 + " " + tey2 + " " + tey3 + " " + tey4 + " x/h";
                }


                else if (rdb_hesab_acmadan.Checked == true && radioButton3.Checked == true && radioButton5.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = kusd;
                    db2 = hacusd;
                    db3 = xvalqisatqi;
                    db4 = hacIIR;
                    db5 = hacIIR;
                    kr1 = hacusd;
                    kr2 = xvalqisatqi;
                    kr3 = hacIIR;
                    kr4 = kochesab;
                    kr5 = xhusd;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();
                        teyinat2 = "1" + " " + radioButton5.Text + "= " +txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();

                        teyinat2 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;
                    }

                }
                    //AZN
                else if (rdb_hesab_acmadan.Checked == true && radioButton3.Checked == true && radioButton7.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = kazn;
                    db2 = haazn;
                    db3 = xvalqisatqi;
                    db4 = hacIIR;
                    db5 = haazn;
                    kr1 = haazn;
                    kr2 = xvalqisatqi;
                    kr3 = hacIIR;
                    kr4 = kochesab;
                    kr5 = xhazn;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb6 = olke_rial_gedenrial_ferq.ToString();
                        teyinat2 = "1" + " " + radioButton7.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton7.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " +txbfilial.Text+" "+ tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + txbfilial.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "AZN-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " AZN)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();

                        teyinat2 = "1" + " " + radioButton7.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton7.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "AZN-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " AZN)" + " " + tey5 + " " + tey4;
                    }

                }
                    //AZN


                else if (rdb_hesab_acmadan.Checked == true && radioButton3.Checked == true && radioButton6.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = keuro;
                    db2 = haceuro;
                    db3 = xvalqisatqi;
                    db4 = hacIIR;
                    db5 = hacIIR;
                    kr1 = haceuro;
                    kr2 = xvalqisatqi;
                    kr3 = hacIIR;
                    kr4 = kochesab;
                    kr5 = xheuro;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "EURO-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " EURO)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();

                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "Avro-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " Avro)" + " " + tey5 + " " + tey4;
                    }
                }
                else if (rdb_hesab_ve_medaxil.Checked == true && radioButton1.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    db1 = kusd;
                    db2 = txb_musteri_hesabi.Text;
                    db3 = txb_musteri_hesabi.Text;
                    kr1 = txb_musteri_hesabi.Text;
                    kr2 = husd;
                    kr3 = xhusd;
                    meb1 = x_h_cem.ToString();
                    meb2 = mbl.ToString();
                    meb3 = x_h.ToString();
                    teyinat1 = tey1 + " " + tey4;
                    teyinat2 = tey5 + " " + tey2 + "  " + tey4;
                    teyinat3 = tey5 + " " + tey2 + "  " + tey4 + " x/h";
                }

                else if (rdb_hesab_ve_medaxil.Checked == true && radioButton2.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    db1 = keuro;
                    db2 = txb_musteri_hesabi.Text;
                    db3 = txb_musteri_hesabi.Text;
                    kr1 = txb_musteri_hesabi.Text;
                    kr2 = heuro;
                    kr3 = xheuro;
                    meb1 = x_h_cem.ToString();
                    meb2 = mbl.ToString();
                    meb3 = x_h.ToString();
                    teyinat1 = tey1 + " " + tey4;
                    teyinat2 = tey5 + " " + tey2 + "  " + tey4;
                    teyinat3 = tey5 + " " + tey2 + "  " + tey4 + " x/h";
                }

                else if (rdb_hesab_ve_medaxil.Checked == true && radioButton3.Checked == true && radioButton5.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = kusd;
                    db2 = txb_musteri_hesabi.Text;
                    db3 = xvalqisatqi;
                    db4 = mushesolanda;
                    db5 = txb_musteri_hesabi.Text;
                    kr1 = txb_musteri_hesabi.Text;
                    kr2 = xvalqisatqi;
                    kr3 = mushesolanda;
                    kr4 = kochesab;
                    kr5 = xhusd;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();
                        teyinat2 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat2 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;
                    }
                    
                }

                    // AZN HESAB VE MEDAXIL

                else if (rdb_hesab_ve_medaxil.Checked == true && radioButton3.Checked == true && radioButton7.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = kazn;
                    db2 = txb_musteri_hesabi.Text;
                    db3 = xvalqisatqi;
                    db4 = mushesolanda;
                    db5 = txb_musteri_hesabi.Text;
                    kr1 = txb_musteri_hesabi.Text;
                    kr2 = xvalqisatqi;
                    kr3 = mushesolanda;
                    kr4 = kochesab;
                    kr5 = xhazn;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();
                        teyinat2 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "AZN dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " AZN dolları)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat2 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;
                    }
                }
                    //AZN HESAB VE MEDAXIL
                else if (rdb_hesab_ve_medaxil.Checked == true && radioButton3.Checked == true && radioButton6.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    db1 = keuro;
                    db2 = txb_musteri_hesabi.Text;
                    db3 = xvalqisatqi;
                    db4 = mushesolanda;
                    db5 = txb_musteri_hesabi.Text;
                    kr1 = txb_musteri_hesabi.Text;
                    kr2 = xvalqisatqi;
                    kr3 = mushesolanda;
                    kr4 = kochesab;
                    kr5 = xheuro;
                    meb1 = x_h_cem.ToString();
                    meb2 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(gedrial).ToString();
                    meb5 = Math.Round(x_h).ToString();
                    teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db6 = zerer;
                        kr6 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb6 = olke_rial_gedenrial_ferq.ToString();
                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "Avro-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " Avro)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db6 = xvalqisatqi;
                        kr6 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat5 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat6 = "Avro-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " Avro)" + " " + tey5 + " " + tey4;
                    }



                }



                ////////////////////hesabdan

                else if (rdb_hesabdan.Checked == true && radioButton1.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    //db1 = kusd;
                    db1 = txb_musteri_hesabi.Text;
                    db2 = txb_musteri_hesabi.Text;
                    //kr1 = txb_musteri_hesabi.Text;
                    kr1 = husd;
                    kr2 = xhusd;
                    //meb1 = x_h_cem.ToString();
                    meb1 = mbl.ToString();
                    meb2 = x_h.ToString();
                    //teyinat1 = tey1 + " " + tey4;
                    teyinat1 = tey5 + " " + tey2 + "  " + tey4;
                    teyinat2 = tey5 + " " + tey2 + "  " + tey4 + " x/h";
                }

                else if (rdb_hesabdan.Checked == true && radioButton2.Checked == true)
                {
                    x_h_cem = mbl + x_h;
                    //db1 = keuro;
                    db1 = txb_musteri_hesabi.Text;
                    db2 = txb_musteri_hesabi.Text;
                    //kr1 = txb_musteri_hesabi.Text;
                    kr1 = heuro;
                    kr2 = xheuro;
                   // meb1 = x_h_cem.ToString();
                    meb1 = mbl.ToString();
                    meb2 = x_h.ToString();
                    //teyinat1 = tey1 + " " + tey4;
                    teyinat1 = tey5 + " " + tey2 + " " + tey4;
                    teyinat2 = tey5 + " " + tey2 + " " + tey4 + " x/h";
                }

                else if (rdb_hesabdan.Checked == true && radioButton3.Checked == true && radioButton5.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    //db1 = kusd;
                    db1 = txb_musteri_hesabi.Text;
                    db2 = xvalqisatqi;
                    db3 = mushesolanda;
                    db4 = txb_musteri_hesabi.Text;
                    //kr1 = txb_musteri_hesabi.Text;
                    kr1 = xvalqisatqi;
                    kr2 = mushesolanda;
                    kr3 = kochesab;
                    kr4 = xhusd;
                    //meb1 = x_h_cem.ToString();
                    meb1 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb2 = Math.Round(gedrial).ToString();
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(x_h).ToString();
                    //teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db5 = zerer;
                        kr5 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb5 = olke_rial_gedenrial_ferq.ToString();
                        teyinat1 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db5 = xvalqisatqi;
                        kr5 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb5 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat1 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;
                    }

                }

                    

                else if (rdb_hesabdan.Checked == true && radioButton3.Checked == true && radioButton7.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    //db1 = kazn;
                    db1 = txb_musteri_hesabi.Text;
                    db2 = xvalqisatqi;
                    db3 = mushesolanda;
                    db4 = txb_musteri_hesabi.Text;
                    //kr1 = txb_musteri_hesabi.Text;
                    kr1 = xvalqisatqi;
                    kr2 = mushesolanda;
                    kr3 = kochesab;
                    kr4 = xhazn;
                   // meb1 = x_h_cem.ToString();
                    meb1 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb2 = Math.Round(gedrial).ToString();
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(x_h).ToString();
                    //teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db5 = zerer;
                        kr5 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb5 = olke_rial_gedenrial_ferq.ToString();
                        teyinat1 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1" + " " + radioButton5.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "AZN dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " AZN dolları)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db5 = xvalqisatqi;
                        kr5 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb5 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat1 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1=" + " " + radioButton5.Text + " " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "ABŞ dolları-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " ABŞ dolları)" + " " + tey5 + " " + tey4;
                    }
                }
                
                else if (rdb_hesabdan.Checked == true && radioButton3.Checked == true && radioButton6.Checked == true)
                {

                    rcbar = Convert.ToSingle(txb_Rial_cbar.Text);
                    vcbar = Convert.ToSingle(txb_valyuta_cbar.Text);
                    irial = Convert.ToSingle(txb_iran_rial.Text);
                    vcbar_bol_rcbar = vcbar / rcbar;

                    x_h_cem = mbl + x_h;
                    //db1 = keuro;
                    db1 = txb_musteri_hesabi.Text;
                    db2 = xvalqisatqi;
                    db3 = mushesolanda;
                    db4 = txb_musteri_hesabi.Text;
                    //kr1 = txb_musteri_hesabi.Text;
                    kr1 = xvalqisatqi;
                    kr2 = mushesolanda;
                    kr3 = kochesab;
                    kr4 = xheuro;
                   // meb1 = x_h_cem.ToString();
                    meb1 = Math.Round(mbl).ToString();
                    gedrial = mbl * irial;
                    meb2 = Math.Round(gedrial).ToString();
                    meb3 = Math.Round(gedrial).ToString();
                    meb4 = Math.Round(x_h).ToString();
                    //teyinat1 = tey1 + " " + tey4;



                    if (vcbar / rcbar < irial)
                    {
                        db5 = zerer;
                        kr5 = xvalqisatqi;
                        olke_rial_gedenrial = irial - vcbar_bol_rcbar;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        //meb6 = Math.Round(olke_rial_gedenrial_ferq).ToString();
                        meb5 = olke_rial_gedenrial_ferq.ToString();
                        teyinat1 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "Avro-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =-" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " Avro)" + " " + tey5 + " " + tey4;

                    }
                    else if (vcbar / rcbar > irial)
                    {
                        db5 = xvalqisatqi;
                        kr5 = gelir;
                        olke_rial_gedenrial = vcbar_bol_rcbar - irial;
                        olke_rial_gedenrial_ferq = olke_rial_gedenrial * mbl * rcbar;
                        meb5 = Math.Round(olke_rial_gedenrial_ferq).ToString();


                        teyinat1 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat2 = "1" + " " + radioButton6.Text + "= " + txb_iran_rial.Text + " " + "İran Rialı" + " " + tey5 + " " + tey4;
                        teyinat3 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs";
                        teyinat4 = tey5 + " " + txbbankadi.Text + " " + tey4 + " " + combomeqsed.Text + " " + "vəsaiti alan şəxs x/h";
                        teyinat5 = "Avro-İran Rialı dilinq fərqi" + " " + vcbar_bol_rcbar.ToString() + " -" + " " + irial.ToString() + " =" + Math.Round(olke_rial_gedenrial).ToString() + " (" + mbl.ToString() + " Avro)" + " " + tey5 + " " + tey4;
                    }



                }


                ///////////////////hesabdan

                //365-54-84
                //cvb = Convert.ToInt32(txb_mebleg.Text);
                // gedenbilgi = cvb.ToString();


            }
            catch (Exception)
            {
            }




        }
        private void word_at()
        {
            try
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Erize1.docx"))
                {
                    MessageBox.Show("Word file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                string g_hesab = rdb_hesab_acmadan.Checked ? "" : txb_musteri_hesabi.Text;

                var replacements = new Dictionary<string, string>
        {
            { "{g_hes}", g_hesab },
            { "{g_adi}", txb_g_adi.Text },
            { "{g_soyadi}", txb_g_soyadı.Text },
            { "{g_ataadi}", txb_g_ataadi.Text },
            { "{G_passport}", txb_gond_passport.Text },
            { "{g_telefon}", txbtelefon.Text },
            { "{bank_adi}", txbbankadi.Text },
            { "{filial}", txbfilial.Text },
            { "{a_adi}", txbalanadi.Text },
            { "{a_soyadi}", txbalansoyadi.Text },
            { "{a_ataadi}", txbalanataadi.Text },
            { "{a_hesab}", txbalanhesab.Text },
            { "{a_passport}", txbalanpassport.Text },
            { "{a_tel}", txbalantelefon.Text },
            { "{T}", txbhevale.Text },
            { "{mebleg}", labelreqemile.Text },
            { "{m_yaziile}", labelyaziile.Text },
            { "{valyuta_novu}", label15.Text },
            { "{meqsed}", combomeqsed.Text },
            { "{elave}", comboelave.Text },
            { "{qeyd}", txb_qeyd.Text },
            { "{alinan_valyuta}", label39.Text + " " + label24.Text },
            { "{satilan_valyuta}", labelrialmeblegi.Text + " " + label28.Text },
            { "{mezenne}", label30.Text },
            { "{tarix}", label31.Text },
            { "{il}", txbhevale.Text },
            { "{ilsoniki}", txbhevale.Text }
        };

                string Erize = System.IO.Path.Combine(qovluqyolu, "Fayllar", "MXD", "Wordler", "Erize1.docx");
                string outputPath = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis wordler", txb_g_soyadı.Text+" "+ txb_g_adi.Text + ".docx");
                                                                                                                                       // Word sənədini yaradın
                aletler.CreateWordDocument(Erize, outputPath, replacements);


                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            radioButton5.Enabled = true;
            radioButton6.Enabled = true;
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
        }
        private void gizle()
        {
            radioButton5.Enabled = false;
            radioButton6.Enabled = false;
        }
        private void temizle_kohne()
        {
            txb_g_adi.Text = "";
            txb_g_soyadı.Text = "";
            txb_g_ataadi.Text = "";
            txb_gond_passport.Text = "";
            txbtelefon.Text = "";
            comboelave.Text = "";
            combomeqsed.Text = "";
            txbalanadi.Text = "";
            txbalanataadi.Text = "";
            txbalansoyadi.Text = "";
            txbalanpassport.Text = "";
            txbalantelefon.Text = "";
            txbbankadi.Text = "";
            txbfilial.Text = "";
            txbalanhesab.Text = "";
            txb_mebleg.Text = "";
            txb_Rial_cbar.Text = "";
            txb_valyuta_cbar.Text = "";
            txb_iran_rial.Text = "";
            txbhevale.Text = "";
        }
        private void temizle(Control parentControl)
        {
            foreach (Control control in parentControl.Controls)
            {
                if (control is TextBox textBox)
                {
                    textBox.Text = "";
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.Text = "";
                }
                else if (control.HasChildren)
                {
                    temizle(control); // Rekursiv olaraq iç kontrolleri yoxla
                }
            }
        }

        private string yaziyaCevir(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için            
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";

            string[] birler = { "", " bir ", " iki ", " üç ", " dört ", " beş ", " altı ", " yeddi ", " səkkiz ", " doqquz " };
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
                yazi += "  ";

            int yaziUzunlugu = yazi.Length;

            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];

            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];

            //if (yazi.Length > yaziUzunlugu)
            //    yazi += " Kr.";
            //else
            //    yazi += "SIFIR Kr.";

            return yazi;

        }
        private void txb_mebleg_TextChanged(object sender, EventArgs e)
        {
            meb_tap();
        }

        private void txb_Rial_cbar_TextChanged(object sender, EventArgs e)
        {
            meb_tap();
        }

        private void txb_valyuta_cbar_TextChanged(object sender, EventArgs e)
        {
            meb_tap();
        }

        private void txb_iran_rial_TextChanged(object sender, EventArgs e)
        {
            meb_tap();
        }
        private void radioButton1_TextChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged_3(object sender, EventArgs e)
        {
            valyutaadi();
            hesab_tap();
        }

        private void radioButton2_CheckedChanged_3(object sender, EventArgs e)
        {
            hesab_tap();
        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {
            hesab_tap();
        }
        private void meb_tap()
        {
            double labelmbl, mbl1, imbl;
            labelmbl = 0;
            cevirme();
            if (txb_mebleg.Text == "")
            {
                //MessageBox.Show("Məbləğ daxil edilməyib");
                MessageBox.Show("Məbləğ daxil edilməyib...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //txb_mebleg.Text.
            }
            else
            {
                mbl1 = Convert.ToDouble(txb_mebleg.Text);
                //imbl = Convert.ToSingle(txb_iran_rial.Text);
                if (txb_iran_rial.Text == "")
                {
                    labelyaziile.Text = yaziyaCevir(Convert.ToDecimal(txb_mebleg.Text));
                    labelreqemile.Text = txb_mebleg.Text;
                }
                else
                {
                    labelmbl = mbl1 * Convert.ToDouble(txb_iran_rial.Text);
                    labelyaziile.Text = yaziyaCevir(Convert.ToDecimal(labelmbl.ToString()));
                    labelreqemile.Text = labelmbl.ToString();
                }
            }
        }
        private void hesab_tap()
        {
            //txb_musteri_hesabi.Text = "";
            if (rdb_hesab_acmadan.Checked == true && radioButton1.Checked == true)
            {

                txb_musteri_hesabi.Text = "45023010010000400000";
            }
            else if (rdb_hesab_acmadan.Checked == true && radioButton2.Checked == true)
            {
                txb_musteri_hesabi.Text = "45023020020000400000";
            }
            else if (rdb_hesab_acmadan.Checked == true && radioButton3.Checked == true)
            {
                txb_musteri_hesabi.Text = "45023040020000400000";
            }
            //else if (rdb_hesab_ve_medaxil.Checked == true)
            //{
            //    txb_musteri_hesabi.Text = "";
            //}


            //else if (rdb_hesab_ve_medaxil.Checked == true)
            //{
            //    txb_musteri_hesabi.Text = "";
            //}
            //else if (rdb_hesabdan.Checked == true)
            //{
            //    txb_musteri_hesabi.Text = "";
            //}
            txb_musteri_hesabi.Text = "";

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //if (textBox2.Text == "")
            //{
            //    MessageBox.Show("Silinəcək sətir seçilməyib", "Qeyd");
            //}
            //else
            //{
                baglanti.Open();
                OleDbCommand komut = new OleDbCommand("delete  from Siyahilar where ID=" + dataGridView1.CurrentRow.Cells["ID"].Value.ToString() + "",baglanti);
                komut.ExecuteNonQuery();
                baglanti.Close();
                //MessageBox.Show("Məlumatlar silindi", "Qeyd");
                MessageBox.Show("Məlumatlar silindi...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cedvel.Clear();
                listele();
                satirsayisinitap();
                textBox13.Text = "";
           // }
        }
        private void satirsayisinitap()
        {
            int satirsayi = -1;

            OleDbCommand orcmd = new OleDbCommand("select count(*) from Siyahilar", baglanti);
            baglanti.Open();
            satirsayi = Convert.ToInt32(orcmd.ExecuteScalar());
            baglanti.Close();
            //string cariil = DateTime.Now.Date.Year.ToString();
            int birartir = satirsayi + 1;
            string nom = birartir.ToString();
            textBox1.Text = nom;
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Siyahilar where Adi like '%" + textBox13.Text + "%' or Soyadi like '%" + textBox13.Text + "%'or Ata_adi like '%" + textBox13.Text + "%'or HevaleNo like '%" + textBox13.Text + "%'", baglanti);
                DataTable tablo2 = new DataTable();
                isleme.Fill(tablo2);
                dataGridView1.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {
            }

        }

        private void rdb_hesab_acmadan_CheckedChanged_2(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
        }

        private void rdb_hesab_ve_medaxil_CheckedChanged_2(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Text = "";
        }

        private void rdb_hesabdan_CheckedChanged_2(object sender, EventArgs e)
        {
            txb_musteri_hesabi.Text = "";
        }

        private void dataGridView1_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txb_g_adi.Text = dataGridView1.CurrentRow.Cells["adi"].Value.ToString();
                txb_g_soyadı.Text = dataGridView1.CurrentRow.Cells["soyadi"].Value.ToString();
                txb_g_ataadi.Text = dataGridView1.CurrentRow.Cells["ata_adi"].Value.ToString();
                txb_gond_passport.Text = dataGridView1.CurrentRow.Cells["gond_passport"].Value.ToString();
                txbtelefon.Text = dataGridView1.CurrentRow.Cells["gond_telefon"].Value.ToString();
                comboelave.Text = dataGridView1.CurrentRow.Cells["ELAVE"].Value.ToString();
                combomeqsed.Text = dataGridView1.CurrentRow.Cells["MEQSED"].Value.ToString();
                txbalanadi.Text = dataGridView1.CurrentRow.Cells["alan_adi"].Value.ToString();
                txbalanataadi.Text = dataGridView1.CurrentRow.Cells["alan_ata_adi"].Value.ToString();
                txbalansoyadi.Text = dataGridView1.CurrentRow.Cells["alan_soyadi"].Value.ToString();
                txbalanpassport.Text = dataGridView1.CurrentRow.Cells["alan_passport"].Value.ToString();
                txbalantelefon.Text = dataGridView1.CurrentRow.Cells["alan_telefon"].Value.ToString();
                txbbankadi.Text = dataGridView1.CurrentRow.Cells["bankin_adi"].Value.ToString();
                txbfilial.Text = dataGridView1.CurrentRow.Cells["filial"].Value.ToString();
                txbalanhesab.Text = dataGridView1.CurrentRow.Cells["alanin_hesabi"].Value.ToString();

                // Seçimin dəyişməsinin qarşısını almaq üçün fokus vəziyyətini saxlayın
                dataGridView1.ClearSelection(); // Seçimi silir
                dataGridView1.Rows[e.RowIndex].Selected = true; // Yalnız mövcud seçimi saxlayır
                // this.Close();
            }
            catch (Exception)
            {
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        

        private void dataGridView1_CellClick_2(object sender, DataGridViewCellEventArgs e)
        {
            //textBox2.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            //// Seçimin dəyişməsinin qarşısını almaq üçün fokus vəziyyətini saxlayın
            //dataGridView1.ClearSelection(); // Seçimi silir
            //dataGridView1.Rows[e.RowIndex].Selected = true; // Yalnız mövcud seçimi saxlayır
        }

        private void txb_mebleg_TextChanged_1(object sender, EventArgs e)
        {
            sifaris_mebleg();
        }
        OracleConnection OrConnect = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass;");
        private void hevnom_al()
        {
            string heval = "";
            string tarixIl = DateTime.Now.Date.Year.ToString();
            string boyukil = "";
            string sonuncuil = "";
            OracleCommand cmd = new OracleCommand();
            OracleCommand cmd1 = new OracleCommand();
            OracleDataReader SR = null;
            OracleDataReader SR1 = null;
            cmd.Connection = OrConnect;
            cmd1.Connection = OrConnect;
            cmd.CommandText = "Select max(to_number(substr(hev_nom,6,4))) from odb.GEDEN_HEVALe where to_char( tarix,'yyyy')='" + tarixIl + "' ";
            cmd1.CommandText = "Select max(to_char( tarix,'yyyy')) from odb.GEDEN_HEVALe ";
            OrConnect.Open();
            SR = cmd.ExecuteReader();
            SR1 = cmd1.ExecuteReader();
            //max(to_number(substr(hev_nom,7,4)))
            if (SR.Read())
            {
                boyukil = SR.GetValue(0).ToString();
                if (boyukil == "")
                {

                }
                else
                {
                    int boyukilsay = Convert.ToInt16(boyukil.ToString());
                }

            }
            if (SR1.Read())
            {
                sonuncuil = SR1.GetValue(0).ToString();
            }

            if (Convert.ToInt16(sonuncuil) < Convert.ToInt16(tarixIl))
            {
                string cariil = "";
                txbhevale.Text = cariil = DateTime.Now.Date.Year.ToString().Substring(2) + "-T-1";

            }
            else
            {
                heval = boyukil.ToString();
                int hevno = Convert.ToInt32(heval.ToString());
                hevno = hevno + 1;
                string cariil = DateTime.Now.Date.Year.ToString().Substring(2);
                txbhevale.Text = cariil + "-T-" + hevno.ToString();
                Pul_Kocurmesi pkh = new Pul_Kocurmesi();
                txbhevale.Text = cariil + "-T-" + hevno.ToString();

            }


            //string hevsecal = heval.Substring(5);

            



        }

        private void hevnoat()
        {
            try
            {
                Xaricmektbtarix = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OrConnect.Close();

                OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Ocon.Open();
                OracleCommand Ocom = new OracleCommand("INSERT INTO ODB.geden_hevale(HEV_NOM,HES_NOM,SAA,Mebleg,TIP_RES,Val_tip,tarix,gon_tip,al_bank,icra) values ('" + txbhevale.Text + "','" + txb_musteri_hesabi.Text + "','" + tamad + "','" + labelreqemile.Text + "','" + "fiziki şəxs" + "','" + label28.Text + "',TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'),'" + "fiziki şəxs" + "','" + txbbankadi.Text + "','" + icraci_kod + "')", Ocon);
                //hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank
                Ocom.ExecuteNonQuery();
                Ocon.Close();
            }
            catch (Exception)
            {
                
                
            }
            
        }

    }
}
