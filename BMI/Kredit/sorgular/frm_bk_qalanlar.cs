using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMI.Kredit
{
    public partial class frm_bk_qalanlar : Form
    {
        public frm_bk_qalanlar()
        {
            InitializeComponent();
        }
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        DataTable muracietler = new DataTable();
        private DataTable dataTable;
        public void kr_bagli_olub_bk_qal()
        {
            try
            {
                muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                OracleDataAdapter isleme = new OracleDataAdapter("select i.licsch,i.ssls,i.ostatok_ish,i.vbs,l.licschkre," +
                    "l.subschkre from odb.vblicsch i, odb.licschkre l where substr(l.licschkre,10,6)||l.subschkre = substr(i.licsch,10,6)||i.ssls " +
                    "and i.ostatok_ish>0 and l.date_close is not null and i.vbs IN (99749, 99743, 99742,99740)", con);
                isleme.Fill(muracietler);
                dataGridView1.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                dataGridView1.Columns[0].HeaderText = "BK hesab";
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[1].HeaderText = "BK sk";
                dataGridView1.Columns[1].Width = 60;
                dataGridView1.Columns[2].HeaderText = "BK məbləğ";
                dataGridView1.Columns[2].Width = 150;
                dataGridView1.Columns[3].HeaderText = "BK";
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].HeaderText = "Kataloq hesab";
                dataGridView1.Columns[4].Width = 200;
                dataGridView1.Columns[5].HeaderText = "Kataloq sk";
                dataGridView1.Columns[5].Width = 70;
            }
            catch (Exception)
            {
            }
            finally { };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            kr_bagli_olub_bk_qal();
        }

        private void ExportToExcel()
        {
            // Yeni bir Excel paketi oluştur
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Yeni bir Excel çalışma sayfası oluştur
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // DataGridView'deki verileri Excel sayfasına aktar
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        // DataGridView hücresinin değeri null değilse aktar
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            // Eğer hücre değeri null ise, boş bir değer yaz
                            worksheet.Cells[i + 2, j + 1].Value = "";
                        }
                    }
                }

                // Excel dosyasını kaydet
                FileInfo excelFile = new FileInfo(@"C:\path\to\excel\file.xlsx");
                excelPackage.SaveAs(excelFile);
            }

            MessageBox.Show("Excel dosyası oluşturuldu.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }
    }
}
