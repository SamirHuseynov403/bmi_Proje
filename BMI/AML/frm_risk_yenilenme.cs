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

namespace BMI.AML
{
    public partial class frm_risk_yenilenme : Form
    {
        public frm_risk_yenilenme()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        cl_aletler cl = new cl_aletler();
        string selectedCode="";
        private Dictionary<string, string> riskler = new Dictionary<string, string>();
        private void comboya_at()
        {
            using (OracleConnection connection = new OracleConnection(cl.baglan))
            {
                // SQL sorgusu
                string query = "select r.code,r.name from riskler r";

                // Veritabanı bağlantısını aç
                connection.Open();

                // Komut oluştur ve bağlantıyı belirt
                using (OracleCommand command = new OracleCommand(query, connection))
                {
                    using (OracleDataReader reader = command.ExecuteReader())
                    {
                        //comboBox1.Items.Add("Hamısı");
                        // Okuyucu ile her bir satırı dolaş
                        while (reader.Read())
                        {
                            // Her bir riskin adını ve kodunu dictionary'e ekle
                            string code = reader["code"].ToString();
                            string name = reader["name"].ToString();
                            comboBox1.Items.Add(name);
                            riskler.Add(name, code);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ComboBox'tan seçilen riskin adını al
            string selectedName = comboBox1.SelectedItem.ToString();

            // Dictionary kullanarak seçilen riske ait kodu al
             selectedCode = riskler[selectedName];
            

            // Seçilen kodu kullanarak başka bir işlem yapabilirsiniz
        }
        private void axtar()
        {
            string daily_report;
            if (checkBox1.Checked==true)
            {
                comboBox1.Text = "";
                daily_report = "SELECT distinct m.qeyd,m.ad,m.risk, " +
         " CASE WHEN m.risk = 'AŞAĞI' AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 1095 THEN 'yenilənməli' " +
         " WHEN m.risk = 'ORTA'  AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 730 THEN 'yenilənməli' " +
         " WHEN m.risk = 'YÜKSƏK' AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 365 THEN 'yenilənməli' " +
         " END AS status,m.ac_tar,m.duzelis ,t.tar FROM licsch l," +
         " (SELECT r.regnom AS qeyd, r.name_regnom AS ad," +
         " CASE WHEN rs.name = '%%%%%%%%%%%%%%%%' THEN 'boş' " +
         " ELSE rs.name END AS risk, rs.name," +
         " MAX(l.date_open_licsch) AS ac_tar, r.note AS duzelis " +
         " FROM regnom r, licsch l, riskler rs " +
         " where l.registrac_nomer = r.regnom and " +
         " rs.code = r.riskler and l.date_close_licsch IS NULL" +
         " AND SUBSTR(l.licsch, 1, 5) <> '99999' GROUP BY r.regnom, r.name_regnom," +
         " CASE WHEN rs.name = '%%%%%%%%%%%%%%%%' THEN 'boş'" +
         " ELSE rs.name END, rs.name, r.note)m," +
         " (select distinct r.regnom qeyd," +
         " case when r.fizik = 1 or r.svazanniy = 1 then f.elave_melumatlarin_tarixi else h.customerrenewaldate end tar " +
         "    from regnom r, fiziki_shexs f,huquqi_shexs h, licsch l " +
         " where((r.regnom = f.regnom) or(r.regnom = h.regnom)) and l.registrac_nomer = r.regnom and l.date_close_licsch is null ) t " +
         "    where m.qeyd = t.qeyd and((substr(l.licsch, 1, 1) in (2, 3, 4))" +
         " and(l.registrac_nomer  in (select distinct registrac_nomer qeyd from licsch, balschkli b " +
         " where substr(licsch, 1, 5) = b.balsch and registrac_nomer = l.registrac_nomer))) " +
         " and m.qeyd = l.registrac_nomer order by m.qeyd asc ";
            }
            else
            {
                daily_report = "SELECT distinct m.qeyd,m.ad,m.risk, " +
         " CASE WHEN m.risk = 'AŞAĞI' AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 1095 THEN 'yenilənməli' " +
         " WHEN m.risk = 'ORTA'  AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 730 THEN 'yenilənməli' " +
         " WHEN m.risk = 'YÜKSƏK' AND TO_DATE('15-04-2024', 'DD-MM-YYYY')-ac_tar >= 365 THEN 'yenilənməli' " +
         " END AS status,m.ac_tar,m.duzelis ,t.tar FROM licsch l," +
         " (SELECT r.regnom AS qeyd, r.name_regnom AS ad," +
         " CASE WHEN rs.name = '%%%%%%%%%%%%%%%%' THEN 'boş' " +
         " ELSE rs.name END AS risk, rs.name," +
         " MAX(l.date_open_licsch) AS ac_tar, r.note AS duzelis " +
         " FROM regnom r, licsch l, riskler rs " +
         " where l.registrac_nomer = r.regnom and rs.code ='"+selectedCode+"' and " +
         " rs.code = r.riskler and l.date_close_licsch IS NULL" +
         " AND SUBSTR(l.licsch, 1, 5) <> '99999' GROUP BY r.regnom, r.name_regnom," +
         " CASE WHEN rs.name = '%%%%%%%%%%%%%%%%' THEN 'boş'" +
         " ELSE rs.name END, rs.name, r.note)m," +
         " (select distinct r.regnom qeyd," +
         " case when r.fizik = 1 or r.svazanniy = 1 then f.elave_melumatlarin_tarixi else h.customerrenewaldate end tar " +
         "    from regnom r, fiziki_shexs f,huquqi_shexs h, licsch l " +
         " where((r.regnom = f.regnom) or(r.regnom = h.regnom)) and l.registrac_nomer = r.regnom and l.date_close_licsch is null ) t " +
         "    where m.qeyd = t.qeyd and((substr(l.licsch, 1, 1) in (2, 3, 4))" +
         " and(l.registrac_nomer  in (select distinct registrac_nomer qeyd from licsch, balschkli b " +
         " where substr(licsch, 1, 5) = b.balsch and registrac_nomer = l.registrac_nomer))) " +
         " and m.qeyd = l.registrac_nomer order by m.qeyd asc ";
            }
            //try
            //{
            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
                cl.dtsorgu.Clear();
                
                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    using (OracleCommand command = new OracleCommand(daily_report, connection))
                    {
                        connection.Open();
                        OracleDataAdapter adapter = new OracleDataAdapter(command);
                        adapter.Fill(cl.dtsorgu);
                        if (rdb_umumi.Checked==true)
                        {
                            dataGridView1.DataSource = cl.dtsorgu;
                        }
                        else
                        {
                            DataView dataView = new DataView(cl.dtsorgu);
                            dataView.RowFilter = "status = 'yenilənməli'";
                            // Yenilenmeli olanları içeren satırları bir yeni DataTable'a kopyalayalım
                            DataTable yenilenmeliDataTable = dataView.ToTable();
                            // DataGridView'e yenilenmeli olanları içeren DataTable'ı atayalım
                            dataGridView1.DataSource = yenilenmeliDataTable;
                        }
                    }
                    connection.Close();
                    dataGridView1.Columns[0].HeaderText = "Qeyd No";
                    dataGridView1.Columns[0].Width = 100;
                    dataGridView1.Columns[1].HeaderText = "Hesab adı";
                    dataGridView1.Columns[1].Width = 300;
                    dataGridView1.Columns[2].HeaderText = "Risklər";
                    dataGridView1.Columns[2].Width = 150;
                    dataGridView1.Columns[3].HeaderText = "Status";
                    dataGridView1.Columns[3].Width = 170;   
                    dataGridView1.Columns[4].HeaderText = "Son hesabın açılma tarixi";
                    dataGridView1.Columns[4].Width = 160;
                    dataGridView1.Columns[5].HeaderText = "Note";
                    dataGridView1.Columns[5].Width = 300;
                    dataGridView1.Columns[6].HeaderText = "Yenilənmə tarixi";
                    dataGridView1.Columns[6].Width = 160;
            }

            //}
            //catch (Exception)
            //{
            //}
            //finally { }
        }
        private void frm_risk_yenilenme_Load(object sender, EventArgs e)
        {
            axtar();
            comboya_at();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            axtar();
            topla();
            checkBox1.Checked = false;
        }
        private void topla()
        {
            try
            {
                int dataRowCount = dataGridView1.RowCount; // Başlık satırı hariç toplam satır sayısı
                lbl_say.Text = dataRowCount.ToString();
            }
            catch (Exception)
            {
            }
            finally { }
            
        }
        private void excel()
        {
            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            //string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            //string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "Risk qruplari üzrə sorğu"; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\Risk qruplari üzrə sorğu.xlsx";
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
            using (ExcelPackage excelPackage = new ExcelPackage(templateFile))
            {
                // Excel dosyasına yeni bir çalışma kitabı ekle
                
                ExcelWorksheet wsL1 = excelPackage.Workbook.Worksheets["Siyahı"];

                // DataGridView'deki verileri Excel sayfasına aktar
                int rowCount = dataGridView1.Rows.Count;
                int colCount = dataGridView1.Columns.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        wsL1.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                // Excel dosyasını kaydet
                filePath = Path.Combine(dosyayolu, fileName);
                excelPackage.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excel();
        }

        private void checkBox1_MouseHover(object sender, EventArgs e)
        {
            ToolTip toolTip1 = new ToolTip();
            toolTip1.SetToolTip(checkBox1, "Bütün risk qrupları");
        }
    }
}
