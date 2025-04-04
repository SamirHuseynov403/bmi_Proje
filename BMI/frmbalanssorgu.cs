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
using Excel = Microsoft.Office.Interop.Excel;

namespace BMI
{
    public partial class frmbalanssorgu : Form
    {
        public frmbalanssorgu()
        {
            InitializeComponent();
            textBox1.KeyDown += textBox1_KeyDown;
        }
        void axtar()
        {
            try
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                OracleCommand Orcom = new OracleCommand("select r.regnom Qey_nomresi,r.name_regnom ad,d.debet debet ,d.kredit kredit ," +
                "d.summa_v_inval x_valyuta,d.summa_v_nacval azn,d.primechanie " +
                "from docdna d, regnom r where substr(d.kredit, 10, 6) = '" + textBox1.Text + "' and substr(d.kredit, 10, 6) = r.regnom", Orcon);

                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Qeyd No";
                dataGridView1.Columns[0].Width = 70;
                dataGridView1.Columns[1].HeaderText = "Adı";
                dataGridView1.Columns[1].Width = 200;
                dataGridView1.Columns[2].HeaderText = "Debet";
                dataGridView1.Columns[2].Width = 150;
                dataGridView1.Columns[3].HeaderText = "Kredit";
                dataGridView1.Columns[3].Width = 150;
                dataGridView1.Columns[4].HeaderText = "Val_məbl";
                dataGridView1.Columns[4].Width = 70;
                dataGridView1.Columns[5].HeaderText = "Məbləğ";
                dataGridView1.Columns[5].Width = 70;
                dataGridView1.Columns[6].HeaderText = "Təyinat";
                dataGridView1.Columns[6].Width = 200;
            }
            catch (Exception)
            {
            }
            finally { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            axtar();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                axtar();
                textBox1.BackColor = Color.Yellow;
                textBox1.SelectAll();
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.BackColor = Color.White;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                double yekunfaiz = 0.00;
                double yekunesas = 0.00;
                double yekunayliq = 0.00;
                DateTime girilenTarih = DateTime.ParseExact(dateTimePicker1.Text, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                double kr_mebleg = Convert.ToInt32(txtmebleg.Text);
                double faiz =Convert.ToDouble(txtfaiz.Text)/100;
                int muddet = Convert.ToInt32(txtmuddet.Text);

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel yüklü değil veya izinler eksik olabilir.");
                    return;
                }

                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Cells[10, 1] = "Ay";
                worksheet.Cells[10, 2] = "Tarix";
                worksheet.Cells[10, 3] = "Kredit məbləği";
                worksheet.Cells[10, 4] = "Aylıq ödəmə";
                worksheet.Cells[10, 5] = "Faiz";
                worksheet.Cells[10, 6] = "Aylıq ödəniş";

                double esas_qaliq= 0.00;
                    esas_qaliq = kr_mebleg;

                double genislik = 12;

                // Tüm sütunların genişliğini belirli bir sayıda birimle ayarlayın
                
                    
                


                int startRow = 11; // İlk veri satırı
                for (int ay = 1; ay <= muddet; ay++)
                {

                    double onceki_silinme = 0.00;
                    esas_qaliq -= onceki_silinme;
                    DateTime ayinBaslangicTarihi = girilenTarih.AddMonths(ay - 1);
                    DateTime ayinSonuTarihi = girilenTarih.AddMonths(ay);
                    double ayliq_oodenis = 0.00;
                        ayliq_oodenis = Math.Round(CalculateMonthlyPayment(kr_mebleg, faiz, muddet), 2);
                    double faiz_odenis = 0.00;
                        faiz_odenis = Math.Round(esas_qaliq * faiz / 12, 2);
                    double esasdan_silinme = 0.00;
                        esasdan_silinme = Math.Round(ayliq_oodenis - faiz_odenis, 2);

                    worksheet.Cells[1, 2] = "Huseynov Samir";
                    worksheet.Cells[1, 2].EntireColumn.AutoFit();
                    worksheet.Cells[2, 2] = "212100000001474100000";
                    worksheet.Cells[3, 2] = "kredit müqaviləsinə";
                    worksheet.Cells[4 + 3, 2] = "ƏLAVƏ";
                    worksheet.Cells[5 + 4, 2] = "Kredit və ona hesablanmış faizlərin";
                    worksheet.Cells[6, 2] = "ödəniş сədvəli";
                    worksheet.Cells[7, 2] = "Kreditin məbləği :";
                    worksheet.Cells[8, 2] = "İllik faiz :";
                    worksheet.Cells[9, 2] = "Müddət :";
                    worksheet.Cells[7, 3] = kr_mebleg;
                    worksheet.Cells[8, 3] = faiz;
                    worksheet.Cells[9, 3] = muddet;

                    worksheet.Cells[startRow, 1] = ay;
                    worksheet.Cells[startRow, 2] = ayinBaslangicTarihi.ToString("dd-MM-yyyy");
                    Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[startRow, 2];
                    cell.NumberFormat = "dd-MM-yyyy";
                    worksheet.Cells[startRow, 3] =esas_qaliq.ToString("N2");
                    worksheet.Cells[startRow, 4] = esasdan_silinme.ToString("N2");
                    worksheet.Cells[startRow, 5] = Math.Round( faiz_odenis,2).ToString("N2");
                    yekunfaiz = yekunfaiz + faiz_odenis;
                    if (ay == muddet)
                    {
                        worksheet.Cells[startRow, 4] = esas_qaliq.ToString("N2");
                        worksheet.Cells[startRow, 6] = (esas_qaliq + faiz_odenis).ToString("N2");
                        yekunesas = yekunesas + esas_qaliq;
                        yekunayliq = yekunayliq + esas_qaliq + faiz_odenis;

                        worksheet.Cells[startRow + 1, 3] = "Yekun:";
                        worksheet.Cells[startRow+1, 4] = yekunesas.ToString("N2");
                        worksheet.Cells[startRow + 1, 5] = yekunfaiz.ToString("N2");
                        worksheet.Cells[startRow + 1, 6] = yekunayliq.ToString("N2");
                    }
                    else
                    {
                        esas_qaliq -= (esasdan_silinme - onceki_silinme);
                        worksheet.Cells[startRow, 6] = ayliq_oodenis.ToString("N2");
                        yekunayliq = yekunayliq + ayliq_oodenis;
                        yekunesas = yekunesas + esasdan_silinme;
                    }
                    startRow++;
                    worksheet.get_Range("A1", "F1").EntireColumn.ColumnWidth = genislik;
                }
                
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluştu: " + ex.Message);
            }
        }

        double CalculateMonthlyPayment(double principal, double faiz, int muddet)
        {
            double monthlyInterestRate = faiz / 12;
            double payment = principal * (monthlyInterestRate * Math.Pow(1 + monthlyInterestRate, muddet)) / (Math.Pow(1 + monthlyInterestRate, muddet) - 1);
            return Math.Round(payment, 2);
        }

    }
}
