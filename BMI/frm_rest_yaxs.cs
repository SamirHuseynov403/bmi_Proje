using OfficeOpenXml;
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
    public partial class frm_rest_yaxs : Form
    {
        public frm_rest_yaxs()
        {
            InitializeComponent();
        }
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        void excel()
        {
            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            string rest_yaxsi = " select distinct y.hes,r.name_regnom adi,y.sk," +
                "e.summa+e.summa_19 rest_tarixindeki_qaliq,y.qal indiki_qaliq,CASE "+
               " WHEN y.qal <= (e.summa + e.summa_19) / 2 THEN 'Bəli' " +
               " ELSE 'Xeyr' " +
               " END AS sonuc," +
                "e.procstavrez rest_tarixindeki_eht,y.r_tar," +
                "y.r_say,y.eh indiki_eht,y.tip,ss.g_gun rest_taren_boy_gec_gun "+
                  "from arh_licschkre e,regnom r,(select m.licschkre hes, m.subschkre sk, m.summa + m.summa_19 qal, " +
                  "m.procstavrez eh, m.date_restructure r_tar, m.kolic_restructure r_say, odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun," +
                  "g.name tip " +
                  "from arh_licschkre m,view_nacpogprokre_all x, tipkre g " +
                  " where m.tipkredita = g.code and m.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and x.date_oper = m.date_oper " +
                  " and x.licschpkre = m.licschpkre and x.subschkre = m.subschkre and m.date_restructure is not null " +
                  " and m.date_close is null ) y, " +
                  "(select x.licschpkre qeyd_no, x.subschkre sub, max(odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper))) g_gun " +
                  "   from view_nacpogprokre_all x, odb.licschkre s " +
                  "  where x.date_oper between to_date(s.date_restructure, 'dd/mm/yyyy') and TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') " +
                  "and x.subschkre = s.subschkre and x.licschppkre = s.licschppkre " +
                  "  and s.date_close is null and s.date_restructure is not null " +
                  "group by x.licschpkre,x.subschkre) ss " +
                  "where e.date_oper = y.r_tar and e.date_close is null " +
                  " and e.licschkre = y.hes and e.subschkre = y.sk and e.licschpkre = ss.qeyd_no and e.subschkre = ss.sub and substr(e.licschpkre,10,6)=r.regnom order by y.r_tar";
        
        DataTable _dt_rest_yaxsi = new DataTable();

        using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(rest_yaxsi, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
            adapter.Fill(_dt_rest_yaxsi);
                }

                
            }
            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "Restruktruzasiya olunmuş kreditlər" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\Restr_yaxsi.xlsx";
            string filePath = Path.Combine(dosyayolu, fileName);

            if (File.Exists(Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet wsL1 = package.Workbook.Worksheets["Siyahı"];
                int startRow = 2;
                int startColumn = 1;

                // DataTable'deki verileri Excel sayfasına yaz
                for (int row = 0; row < _dt_rest_yaxsi.Rows.Count; row++)
                {
                    for (int col = 0; col < _dt_rest_yaxsi.Columns.Count; col++)
                    {
                        wsL1.Cells[startRow + row, startColumn + col].Value = _dt_rest_yaxsi.Rows[row][col];
                    }
                }
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }

         }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = textBox2.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    textBox2.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel();
        }
    }
}
