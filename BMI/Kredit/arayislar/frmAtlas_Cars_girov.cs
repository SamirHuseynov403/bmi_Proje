using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Data.OleDb;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Diagnostics;

namespace BMI
{
    public partial class frmAtlas_Cars_girov : Form
    {
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public frmAtlas_Cars_girov()
        {
            InitializeComponent();
        }

        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI_\BMI\bin\Debug\AtlasCars.accdb;Persist Security Info=True");
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI_\BMI\bin\Debug\AtlasCars.accdb
        OracleCommand komyazdir;
        OracleCommand komutmurraz;
        public string icraci_kod = string.Empty;
        DataTable carsbaza = new DataTable();
        DataTable muracietler = new DataTable();
        string tarixIl = DateTime.Now.Date.Year.ToString();
        string tarixay = DateTime.Now.Date.Month.ToString();
        string tarixgun = DateTime.Now.Date.Day.ToString();
        public void carsgirovadusenler()
        {
                carsbaza.Clear();
                //string tarixIl = txbil.Text;
                //int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }

                OracleDataAdapter isleme = new OracleDataAdapter("select otvet_ispoln,date_oper,vbsd,vbsk,licsch,summa,primechanie from vbarh_dd where licsch='40090000001522200000' and vbsk='99789' ", con);
                isleme.Fill(carsbaza);
                dataGridView1.DataSource = carsbaza;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

            dataGridView1.Columns[0].HeaderText = "İcraçı";
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].HeaderText = "Tarix";
            dataGridView1.Columns[1].Width = 110;
            dataGridView1.Columns[2].HeaderText = "Debet";
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].HeaderText = "Kredit";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "Hesab";
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].HeaderText = "Məbləğ";
            dataGridView1.Columns[5].Width = 100;
            dataGridView1.Columns[6].HeaderText = "Təyinat";
            dataGridView1.Columns[6].Width = 850;
        }
        void yazilmis_mektublar()
        {

                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                baglanti.Open();
                muracietler.Clear();
                OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Cars_mektublar ", baglanti);
                isleme.Fill(muracietler);
                dataGridView1.DataSource = muracietler;

                baglanti.Close();
            dataGridView1.Columns[0].HeaderText = "Sira";
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].HeaderText = "Avto No";
            dataGridView1.Columns[1].Width = 100;
            dataGridView1.Columns[2].HeaderText = "Məktub No";
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].HeaderText = "Mühərrik";
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].HeaderText = "BAN";
            dataGridView1.Columns[4].Width = 200;
            dataGridView1.Columns[5].HeaderText = "Tarix";
            dataGridView1.Columns[5].Width = 100;
            dataGridView1.Columns[6].HeaderText = "İran mək No";
            dataGridView1.Columns[6].Width = 100;
        }
        public void carsgirovdancixanlar()
        {

           
            carsbaza.Clear();
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            OracleDataAdapter isleme = new OracleDataAdapter("select otvet_ispoln,date_oper,vbsd,vbsk,licsch,summa,primechanie from vbarh_dd where licsch='40090000001522200000' and vbsd='99789' ", con);
            isleme.Fill(carsbaza);
            dataGridView1.DataSource = carsbaza;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dataGridView1.Columns[0].HeaderText = "İcraçı";
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].HeaderText = "Tarix";
            dataGridView1.Columns[1].Width = 110;
            dataGridView1.Columns[2].HeaderText = "Debet";
            dataGridView1.Columns[2].Width = 100;
            dataGridView1.Columns[3].HeaderText = "Kredit";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "Hesab";
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].HeaderText = "Məbləğ";
            dataGridView1.Columns[5].Width = 100;
            dataGridView1.Columns[6].HeaderText = "Təyinat";
            dataGridView1.Columns[6].Width = 850;
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            
        }
        
        FileInfo[] Files;
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                DirectoryInfo d = new DirectoryInfo(fbd.SelectedPath);
                Files = d.GetFiles("*.*");
                string str = "";
                foreach (FileInfo file in Files)
                {
                    listBox1.Items.Add(file.Name);

                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;
            listBox2.Items.Clear();
            listBox2.Items.Add("Length : " + Files[index].Length);
            listBox2.Items.Add("LastWriteTime : " + Files[index].LastWriteTime);
            listBox2.Items.Add("Extension : " + Files[index].Extension);
            listBox2.Items.Add("LastAccessTime : " + Files[index].LastAccessTime);
            textBox1.Text = File.ReadAllText(Files[index].FullName);
        }

        

       
        private void button2_Click(object sender, EventArgs e)
        {
            //try
            //{

                string pdfFolderPath = @"C:\\BMI_\\2023-132"; // PDF dosyalarının bulunduğu klasör
                string searchText = txtaxtar.Text; // PDF içinde aranacak metin

                string foundPdfPath = FindPdfContainingText(pdfFolderPath, searchText);

                if (foundPdfPath != null)
                {
                    Console.WriteLine($"Aranan metni içeren PDF bulundu: {foundPdfPath}");
                    OpenPdfFile(foundPdfPath);
                }
                else
                {
                    Console.WriteLine("Aranan metni içeren PDF bulunamadı.");
                }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"PDF dosyasını açarken bir hata oluştu: {ex.Message}");
            //    Console.WriteLine($"Hata Detayları: {ex.StackTrace}");
            //}
        }

        static string FindPdfContainingText(string folderPath, string searchText)
        {
            string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.AllDirectories);

            foreach (string pdfFilePath in pdfFiles)
            {
                using (PdfReader pdfReader = new PdfReader(pdfFilePath))
                {
                    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                    {
                        string pageText = PdfTextExtractor.GetTextFromPage(pdfReader, page);

                        if (pageText.Contains(searchText))
                        {
                            return pdfFilePath;
                        }
                    }
                }
            }

            return null;
        }

        static void OpenPdfFile(string pdfFilePath)
        {
            try
            {
                System.Diagnostics.Process.Start(pdfFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF dosyasını açarken bir hata oluştu: {ex.Message}");
                Console.WriteLine($"Hata Detayları: {ex.StackTrace}");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            carsgirovadusenler();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            yazilmis_mektublar();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            carsgirovdancixanlar();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmcarsmektubcix rdmca = new frmcarsmektubcix();
            rdmca.icraci_kod = icraci_kod;
            rdmca.Show();
        }

        private void repositoryItemButtonEdit1_Click(object sender, EventArgs e)
        {

        }

        private void txtaxtar_EditValueChanged(object sender, EventArgs e)
        {

        }
        //projenin ilk oldugu yer
        private void button7_Click(object sender, EventArgs e)
        {
           
            // Faylın tam yolunu qurmaq
            string faylyolu = System.IO.Path.Combine(qovluqyolu, "Fayllar", "AML_Əlavə 2.xlsx");

            // Faylın mövcudluğunu yoxlayırıq
            if (File.Exists(faylyolu))
            {
                Console.WriteLine("Fayl tapıldı: " + faylyolu);
            }
            else
            {
                Console.WriteLine("Fayl tapılmadı: " + faylyolu);
            }

        }
    }
}
