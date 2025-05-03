using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using BMI.Muhasibat;
using OfficeOpenXml;


namespace BMI.Sorgular.Kredit
{
    public partial class frm_kredit_portfeli_aylar_uzre_qaliqalr : Form
    {
        public frm_kredit_portfeli_aylar_uzre_qaliqalr()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();

        void excel()
        {
            DataTable _dt_aylar_uzre_qaliqlar = new DataTable();
            DataTable _dt_aylar_uzre_qaliqlar_normal = new DataTable();
            DataTable _dt_aylar_uzre_qaliqlar_vk = new DataTable();

            string kredit_portfel_aylar_qaliq = @"SELECT 
                l.date_oper,
                SUM(l.summa * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6)) +
                SUM(l.summa_19 * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6)) AS qaliq
            FROM arh_licschkre l 
            WHERE l.date_oper IN (
                SELECT MAX(c.date_oper)
                FROM odb.calendar c
                WHERE 
                    EXTRACT(YEAR FROM c.date_oper) IN (:il_1, :il_2, :il_3)
                    AND (c.space_or_star IS NULL or c.space_or_star='d')
                    AND c.date_oper <= SYSDATE
                GROUP BY TO_CHAR(c.date_oper, 'YYYYMM')
            )
            AND l.date_close IS NULL
            GROUP BY l.date_oper
            ORDER BY l.date_oper";

            string kredit_portfel_aylar_normal_qaliq = @"SELECT 
                l.date_oper,
                SUM(l.summa * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6))
                 AS qaliq
            FROM arh_licschkre l 
            WHERE l.date_oper IN (
                SELECT MAX(c.date_oper)
                FROM odb.calendar c
                WHERE 
                    EXTRACT(YEAR FROM c.date_oper) IN (:il_1, :il_2, :il_3)
                    AND (c.space_or_star IS NULL or c.space_or_star='d')
                    AND c.date_oper <= SYSDATE
                GROUP BY TO_CHAR(c.date_oper, 'YYYYMM')
            )
            AND l.date_close IS NULL
            GROUP BY l.date_oper
            ORDER BY l.date_oper";

            string kredit_portfel_aylar_vk_qaliq = @"SELECT 
                l.date_oper,
                SUM(l.summa_19 * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6))
                 AS qaliq
            FROM arh_licschkre l 
            WHERE l.date_oper IN (
                SELECT MAX(c.date_oper)
                FROM odb.calendar c
                WHERE 
                    EXTRACT(YEAR FROM c.date_oper) IN (:il_1, :il_2, :il_3)
                    AND (c.space_or_star IS NULL or c.space_or_star='d')
                    AND c.date_oper <= SYSDATE
                GROUP BY TO_CHAR(c.date_oper, 'YYYYMM')
            )
            AND l.date_close IS NULL
            GROUP BY l.date_oper
            ORDER BY l.date_oper";

            string il1 = string.IsNullOrWhiteSpace(txt_il1.Text) ? "0" : txt_il1.Text.Trim();
            string il2 = string.IsNullOrWhiteSpace(txt_il2.Text) ? "0" : txt_il2.Text.Trim();
            string il3 = string.IsNullOrWhiteSpace(txt_il3.Text) ? "0" : txt_il3.Text.Trim();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(kredit_portfel_aylar_qaliq, connection))
                {
                    // Parametrləri əlavə et
                    command.Parameters.Add(":il_1", OracleDbType.Varchar2).Value = il1;
                    command.Parameters.Add(":il_2", OracleDbType.Varchar2).Value = il2;
                    command.Parameters.Add(":il_3", OracleDbType.Varchar2).Value = il3;

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_aylar_uzre_qaliqlar);
                }

                using (OracleCommand command = new OracleCommand(kredit_portfel_aylar_normal_qaliq, connection))
                {
                    // Parametrləri əlavə et
                    command.Parameters.Add(":il_1", OracleDbType.Varchar2).Value = il1;
                    command.Parameters.Add(":il_2", OracleDbType.Varchar2).Value = il2;
                    command.Parameters.Add(":il_3", OracleDbType.Varchar2).Value = il3;

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_aylar_uzre_qaliqlar_normal);
                }

                using (OracleCommand command = new OracleCommand(kredit_portfel_aylar_vk_qaliq, connection))
                {
                    // Parametrləri əlavə et
                    command.Parameters.Add(":il_1", OracleDbType.Varchar2).Value = il1;
                    command.Parameters.Add(":il_2", OracleDbType.Varchar2).Value = il2;
                    command.Parameters.Add(":il_3", OracleDbType.Varchar2).Value = il3;

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_aylar_uzre_qaliqlar_vk);
                }
                connection.Close();
                
            }
            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txt_il1.Text + "-" + txt_il2.Text + "-" + txt_il3.Text; // TextBox'tan alınan metni sakla
                                                                                         //string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = textBoxText + " qaliqlar"; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Exceller", "Kredit qaliqlar.xlsx");
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);

            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);
            FileInfo newFile = new FileInfo(filePath);
            templateFile.CopyTo(newFile.FullName, true);

            using (ExcelPackage package = new ExcelPackage(newFile, true))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets["Qaliqlar"];

                // İl başlıqları B5, C5, D5
                ws.Cells["B5"].Value = il1;
                ws.Cells["C5"].Value = il2;
                ws.Cells["D5"].Value = il3;

                // Aylara görə datanı yaz
                foreach (DataRow row in _dt_aylar_uzre_qaliqlar.Rows)
                {
                    DateTime dateOper = Convert.ToDateTime(row["date_oper"]);
                    int year = dateOper.Year;
                    int month = dateOper.Month;

                    double toplam = row["qaliq"] == DBNull.Value ? 0 : Convert.ToDouble(row["qaliq"]);

                    int yearColOffset = (year.ToString() == il1) ? 0 :
                                        (year.ToString() == il2) ? 1 :
                                        (year.ToString() == il3) ? 2 : -1;

                    if (yearColOffset != -1)
                    {
                        ws.Cells[6 + (month - 1), 2 + yearColOffset].Value = toplam;      
                    }
                }

                foreach (DataRow row in _dt_aylar_uzre_qaliqlar_normal.Rows)
                {
                    DateTime dateOper = Convert.ToDateTime(row["date_oper"]);
                    int year = dateOper.Year;
                    int month = dateOper.Month;

                    double toplam = row["qaliq"] == DBNull.Value ? 0 : Convert.ToDouble(row["qaliq"]);

                    int yearColOffset = (year.ToString() == il1) ? 0 :
                                        (year.ToString() == il2) ? 1 :
                                        (year.ToString() == il3) ? 2 : -1;

                    if (yearColOffset != -1)
                    {
                        ws.Cells[21 + (month - 1), 2 + yearColOffset].Value = toplam;
                    }
                }

                foreach (DataRow row in _dt_aylar_uzre_qaliqlar_vk.Rows)
                {
                    DateTime dateOper = Convert.ToDateTime(row["date_oper"]);
                    int year = dateOper.Year;
                    int month = dateOper.Month;

                    double toplam = row["qaliq"] == DBNull.Value ? 0 : Convert.ToDouble(row["qaliq"]);

                    int yearColOffset = (year.ToString() == il1) ? 0 :
                                        (year.ToString() == il2) ? 1 :
                                        (year.ToString() == il3) ? 2 : -1;

                    if (yearColOffset != -1)
                    {
                        ws.Cells[36 + (month - 1), 2 + yearColOffset].Value = toplam;
                    }
                }

                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }
        private void SetupYearTextBox(TextBox txt)
        {
            txt.MaxLength = 4; // 4 rəqəm limiti

            txt.KeyPress += (s, e) =>
            {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                {
                    e.Handled = true; // Rəqəm deyilsə blokla
                }
            };

            txt.TextChanged += (s, e) =>
            {
                if (txt.Text.Length > 4)
                {
                    txt.Text = txt.Text.Substring(0, 4);
                    txt.SelectionStart = txt.Text.Length;
                }
            };

            txt.Leave += (s, e) =>
            {
                if (txt.Text.Length != 4)
                {
                    MessageBox.Show("Zəhmət olmasa 4 rəqəmli il daxil edin.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txt.Focus();
                }
            };
        }

        void dtg_at()
        {
            DataTable _dt_aylar_uzre_qaliqlar = new DataTable();
            DataTable _dt_aylar_uzre_qaliqlar_normal = new DataTable();
            DataTable _dt_aylar_uzre_qaliqlar_vk = new DataTable();

            string kredit_portfel_aylar_qaliq = @"SELECT 
                l.date_oper,
                SUM(l.summa * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6)) as esas,
                SUM(l.summa_19 * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6)) AS vk
            FROM arh_licschkre l 
            WHERE l.date_oper IN (
                SELECT MAX(c.date_oper)
                FROM odb.calendar c
                WHERE 
                    EXTRACT(YEAR FROM c.date_oper) IN (:il_1, :il_2, :il_3)
                    AND c.space_or_star IS NULL
                    AND c.date_oper <= SYSDATE
                GROUP BY TO_CHAR(c.date_oper, 'YYYYMM')
            )
            AND l.date_close IS NULL
            GROUP BY l.date_oper
            ORDER BY l.date_oper";

           

            string il1 = string.IsNullOrWhiteSpace(txt_il1.Text) ? "0" : txt_il1.Text.Trim();
            string il2 = string.IsNullOrWhiteSpace(txt_il2.Text) ? "0" : txt_il2.Text.Trim();
            string il3 = string.IsNullOrWhiteSpace(txt_il3.Text) ? "0" : txt_il3.Text.Trim();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(kredit_portfel_aylar_qaliq, connection))
                {
                    // Parametrləri əlavə et
                    command.Parameters.Add(":il_1", OracleDbType.Varchar2).Value = il1;
                    command.Parameters.Add(":il_2", OracleDbType.Varchar2).Value = il2;
                    command.Parameters.Add(":il_3", OracleDbType.Varchar2).Value = il3;

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_aylar_uzre_qaliqlar);
                }

               
                connection.Close();

            }
            dtg_sorgu.DataSource = _dt_aylar_uzre_qaliqlar;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel();
        }

        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            dtg_at();
        }

        private void frm_kredit_portfeli_aylar_uzre_qaliqalr_Load(object sender, EventArgs e)
        {
            SetupYearTextBox(txt_il1);
            SetupYearTextBox(txt_il2);
            SetupYearTextBox(txt_il3);
        }
    }
}
