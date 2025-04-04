using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.Xml.Linq;
using System.Xml;

namespace BMI
{
    public partial class AMLexcel : Form
    {
        public AMLexcel()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
    void xmldandatagrideyukle()
        {
            XmlDocument i = new XmlDocument();
            DataSet ds = new DataSet();
            //xml dosyamızı okumak için bir reader oluşturuyoruz.
            XmlReader xmlFile;
            xmlFile = XmlReader.Create(@"DOMESTIC - Copy.xml", new XmlReaderSettings());
            //içeriği Dataset e aktarıyoruz.
            ds.ReadXml(xmlFile);
            //datagridviewin kaynağı olarak dataseti gösteriyoruz.
            dataGridView3.DataSource = ds.Tables[0];
            xmlFile.Close();
        }
        private void gonder()
        {
            

            //try
            //{
            //    string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //    desktopFolder = desktopFolder + "\\AMLexcel.xlsx";
            //    OleDbConnection bağlantıexcel = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+desktopFolder+"; Extended Properties='Excel 12.0 Xml;HDR=YES'");
            //bağlantıexcel.Open();
            //int kayitsay = 0;
            //OleDbCommand komut = new OleDbCommand("Select * From [" + "Axtarilanlar" + "$]", bağlantıexcel);
            //OleDbDataAdapter da = new OleDbDataAdapter(komut);
            //DataTable data = new DataTable();
            //da.Fill(data);
            //dataGridView1.DataSource = data;
            //bağlantıexcel.Close();
            String movzu = "Boş məlumatlar";
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = desktopFolder + "\\AMLexcel.xlsx";

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                string selectQuery = "SELECT * FROM [" + "Axtarilanlar" + "$]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, connection);
                DataTable data = new DataTable();
                adapter.Fill(data);

                // Boş satırları kaldır
                for (int i = data.Rows.Count - 1; i >= 0; i--)
                {
                    bool isRowEmpty = true;
                    for (int j = 0; j < 4; j++) // A'dan E'ye kadar olan sütunları kontrol edin
                    {
                        if (!string.IsNullOrWhiteSpace(data.Rows[i][j].ToString()))
                        {
                            isRowEmpty = false;
                            break;
                        }
                    }
                    if (isRowEmpty)
                    {
                        data.Rows.RemoveAt(i);
                    }
                }

                dataGridView1.DataSource = data;
            }
        
            // DataGridView'e verileri aktar

            //Bu kod, Excel dosyasından verileri alır ve eğer bir setirde A'dan E'ye kadar olan sütunlardan herhangi biri dolu değilse işlemi sonlandırır ve bu setirin tüm boş sütunlarına "bos melumatlar movcuddur" yazar.Eğer işlemi sonlandırmanız gerekiyorsa ve bir sonraki setir tamamen boşsa işlemi sonlandırır. Sonuçları DataGridView'e aktarır.






            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel ilə bağlı xəta başverdi");
            //}
        }
            private void button1_Click(object sender, EventArgs e)
        {
            bazayaat();
            MessageBox.Show("Məlumatlar bazaya yükləndi...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            testyoxla();
        }
        private void bazayaat()
        {
            try
            {
                silme();
                OracleConnection baglanti = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                baglanti.Open();
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                string val1 = "ssssss"; // 1. kolon
                string val2 = "ssssss";
                string val3 = "ssssss";
                string val4 = "";
                string val5 = "ssssss";

                if (!string.IsNullOrWhiteSpace(dataGridView1.Rows[i].Cells[0].Value as string))
                {
                    val1 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                }
                if (!string.IsNullOrWhiteSpace(dataGridView1.Rows[i].Cells[1].Value as string))
                {
                    val2 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                }
                if (!string.IsNullOrWhiteSpace(dataGridView1.Rows[i].Cells[2].Value as string))
                {
                    val3 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                }
                
                //if (!string.IsNullOrWhiteSpace(dataGridView1.Rows[i].Cells[4].Value as string))
                //{
                //    val5 = dataGridView1.Rows[i].Cells[4].Value.ToString();
                //}
                
                    OracleCommand komut = new OracleCommand("insert into odb.AML_YOXLAMA (A_S_A,VOEN,FIN,NOV,TEL) values ('" + val1 + "','" + val2 + "','" + val3 + "','" + val4 + "','" + val5 + "')", baglanti);
                    komut.ExecuteNonQuery();
               
            }
            
                baglanti.Close();
        }
            catch (Exception)
            {
            }
            //finally { };
        }
        private void silme()
        {
            OracleConnection baglanti = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            OracleCommand komut = new OracleCommand("DELETE FROM odb.AML_YOXLAMA", baglanti);

            try
            {
                baglanti.Open();
                komut.ExecuteNonQuery();
                Console.WriteLine("Mevcut veriler başarıyla silindi.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Veri silme işleminde hata oluştu: " + ex.Message);
            }
            finally
            {
                baglanti.Close();
            }
            //Yukarıdaki kod parçası, belirtilen tablodaki mevcut verileri siler.Hata oluşması durumunda bir hata mesajı görüntülenir. "baglanti" nesnesi doğru bir şekilde yapılandırılmış ve bağlantı bilgileriniz doğruysa, bu kod başarılı bir şekilde çalışmalıdır. Yine de dikkatli olmalısınız, çünkü veriler silindiğinde geri alınamazlar.






            //OracleCommand komut = new OracleCommand("DELETE FROM odb.AML_YOXLAMA", baglanti);
            //baglanti.Open();
            //komut.ExecuteNonQuery();
            //baglanti.Close();
        }
        private void AMLexcel_Load(object sender, EventArgs e)
        {
            testyoxla();
            //exceldencek();
            gonder();
        }
        private void testyoxla()
        {
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleCommand Orcom = new OracleCommand("Select * From odb.AML_YOXLAMA", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            for (int row = 0; row < Ordt.Rows.Count; row++)
            {
                for (int col = 0; col < Ordt.Columns.Count; col++)
                {
                    if (Ordt.Rows[row][col].ToString() == "ssssss")
                    {
                        Ordt.Rows[row][col] = DBNull.Value; // Hücreyi boş yap
                    }
                }
            }
            dataGridView2.DataSource = Ordt;
            Orcon.Close();
            dataGridView2.Columns[0].HeaderText = "ADI";
            dataGridView2.Columns[0].Width = 150;
            dataGridView2.Columns[1].HeaderText = "VÖEN";
            dataGridView2.Columns[1].Width = 120;
            dataGridView2.Columns[2].HeaderText = "FİN";
            dataGridView2.Columns[2].Width = 120;
            dataGridView2.Columns[3].HeaderText = "Növü";
            dataGridView2.Columns[3].Width = 50;
            //dataGridView2.Columns[4].HeaderText = "Tel";
            //dataGridView2.Columns[4].Width = 100;
        }
        private void LoadExcelToDataGridView(string excelFilePath)
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string selectQuery = "SELECT * FROM [Sheet1$]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, connection);
                DataTable data = new DataTable();
                adapter.Fill(data);
                dataGridView1.DataSource = data;
            }
        }
        void exceldencek()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = desktopFolder + "\\AMLexcel.xlsx";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string selectQuery = "SELECT * FROM [" + "Axtarilanlar" + "$]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, connection);
                DataTable data = new DataTable();
                adapter.Fill(data);
                dataGridView1.DataSource = data;
            }
        }
    }
}
