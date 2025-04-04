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
using System.Globalization;
//using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Reflection;


namespace BMI
{
    public partial class qizilmelumatlari : Form
    {
        public qizilmelumatlari()
        {
            InitializeComponent();
        }

        System.Data.DataTable tablo = new System.Data.DataTable();

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        string metal, novu, eyar;
        double sayi, ceki, qiymet, deyeri;
        int say = 0;
        private void yazdir()
        {
            try
            {
                metal = textBox1.Text;
                novu = comboBox1.Text;
                eyar = textBox4.Text;

                sayi = Convert.ToDouble(textBox3.Text);
                ceki = Convert.ToDouble(textBox5.Text);
                qiymet = Convert.ToDouble(textBox6.Text);
                deyeri = ceki * qiymet;

                say = say + 1;
                dataGridView1.Rows.Add(say, metal, novu, sayi, eyar, ceki, qiymet,deyeri);
                textBox5.Text = "";
            }
            catch (Exception)
            {

                throw;
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

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            yazdir();
        }

        //Microsoft.Office.Interop.Word.Application app;
        //Microsoft.Office.Interop.Word.Document doc;
        //object objMiss = Missing.Value;
        //object Tmpfile = System.IO.Path.GetTempPath() + "\\Test.pdf";
        //object FileLocation = @"‪D:\Test\Test.docx";


        private void qizilmelumatlari_Load(object sender, EventArgs e)
        {
            
        }

        //private void FindAndReplace(object FindText, object ReplaceText)
        //{
        //    this.app.Selection.Find.Execute(ref FindText,true,true,false,false,false,true,false, 1 ,ref ReplaceText, 2 ,false,false,false,false);
        //}
        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document WordDocument)
        {
            //var range = WordDocument.Content;
            //range.Find.ClearFormatting();
            //range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            testtab();

            //int satirsay = dataGridView1.Rows.Count;
            //int sutunsay = 8;

            //string template = @Path.GetDirectoryName(Application.ExecutablePath).Trim() + "\\Test.docx";
            //Word.Application wordapp = new Word.Application();
            //wordapp.Visible = true;
            //Word.Document document = wordapp.Documents.OpenNoRepairDialog(template);
            //document.Activate();

            ////Word.Range tablelocation = this.Range(ref satirsay, ref sutunsay);

            //Word.Table table = document.Tables[1];
            ////this.table.Add(tablelocation, 3, 4);
            ////table.Cell(1, 1).Range.Text = "Samir";
            ////table.Cell(1, 2).Range.Text = "huseynov";
            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    for (int j = 0; j < 8; j++)
            //    {
            //        table.Rows[i + 1].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();

            //    }
            //}

           

         
        }

        private Word.Range Range(ref int satirsay, ref int sutunsay)
        {
            throw new NotImplementedException();
        }

        private void tabloat()
        {
        
        object objmissing=System.Reflection.Missing.Value;
        object oEndOfdoc = "C:\\BMI_\\BMI\\bin\\Debug\\testqizil.docx";
        Microsoft.Office.Interop.Word.Application appob;
        Microsoft.Office.Interop.Word.Document docob;
            appob=new Microsoft.Office.Interop.Word.Application();
            appob.Visible = true;
            docob = appob.Documents.Add(ref objmissing, ref objmissing, ref objmissing);
            int i = 0;
            int j=0;
            Microsoft.Office.Interop.Word.Table tableodb;
                Word.Range wrdrng = docob.Bookmarks.get_Item(ref oEndOfdoc).Range;
                tableodb = docob.Tables.Add(wrdrng, 3, 4, ref objmissing, ref objmissing);
            tableodb.Range.ParagraphFormat.SpaceAfter = 8;
            string str;
            for ( i = 0; i <= 3; i++)
            {
                for ( j = 0; j <=4 ; j++)
                {
                    str = "Row" + i + "Column";
                    tableodb.Cell(i, j).Range.Text = str;

                }
                tableodb.Rows[1].Range.Font.Bold = 1;
                this.Close();
            }



        }

        private void testtab()
        {
            int satirsayi = dataGridView1.Rows.Count;
            int sutunsay = 9;
            object oMissing = System.Reflection.Missing.Value;
            object oendofdoc = "\\endofdoc";
            Word.Application wordapp = new Word.Application();
            Word.Document wordoc = wordapp.Documents.Add(ref oMissing,ref oMissing,ref oMissing,ref oMissing);
            wordapp.Visible = true;
            object orng = wordoc.Bookmarks.get_Item(ref oendofdoc).Range;
            orng = wordoc.Bookmarks.get_Item(ref oendofdoc).Range;
            Word.Range wrdrng = wordoc.Bookmarks.get_Item(ref oendofdoc).Range;
            Word.Table tablo = wordoc.Tables.Add(wrdrng,satirsayi,sutunsay, ref oMissing,ref oMissing);
            tablo.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tablo.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //tablo.Rows[1].Cells[1].Range.Text = "Sira No";
            //tablo.Rows[1].Cells[2].Range.Text = "Qiymətli metal";
            //tablo.Rows[1].Cells[3].Range.Text = "Növü";
            //tablo.Rows[1].Cells[4].Range.Text = "Ədəd sayı";
            //tablo.Rows[1].Cells[5].Range.Text = "Əyarı";
            //tablo.Rows[1].Cells[6].Range.Text = "Çəkisi qr";
            //tablo.Rows[1].Cells[7].Range.Text = "qiyməti qr man.";
            //tablo.Rows[1].Cells[8].Range.Text = "Dəyəri man";
            //tablo.Rows[1].Cells[9].Range.Text = "Qeyd";

            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    tablo.Rows[i + 2].Cells[j + 1].Range.Text = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
        }

        //private Word.Range Range(ref int satirsay, ref int sutunsay)
        //{
        //    throw new NotImplementedException();
        //}
        string wordtabat;

        //public Document wordDocument { get; set; }

        public string adi { get; set; }

        
    }
}
