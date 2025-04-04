using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.CodeParser;
using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using OfficeOpenXml.Style;
using System.Data.SqlClient;


namespace BMI.Muhasibat
{
    public partial class frm_reziden_ve_qeyri_rezident : Form
    {
        public frm_reziden_ve_qeyri_rezident()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        private string[] hes_kod_1;

        public class Hesabat
        {
            public string Hes { get; set; }
            public string Ad { get; set; }
            public string ExtractedNumber { get; set; }
            public string Tip { get; set; }
            public decimal SaldoVhdNacvalAbs { get; set; }
            public decimal SaldoIshNacvalAbs { get; set; }
        }

        private void hesablari_excel_cek()
        {
            // EPPlus lisenziya kontekstini təyin edirik
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Excel faylının yolunu təyin edirik
            string filePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Yanasmalar", "rezident_qeyri_rezident _yeni_Excelde_cekilenler_ucun.xlsx");

            // Excel faylını oxuyuruq
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet8"]; // İlk səhifəni seçirik

                // Əgər iş səhifəsi boşdursa və ya kifayət qədər sətir yoxdursa, metodu bitiririk
                if (worksheet.Dimension == null || worksheet.Dimension.End.Row < 6)
                    return;

                // A6-dan A1000-ə qədər olan hüceyrələri oxuyuruq
                hes_kod_1 = worksheet.Cells["A6:A1000"]
                   .Where(cell => cell.Value != null) // Boş hüceyrələri nəzərə almırıq
                   .Select(cell => cell.Text) // Hüceyrə mətnini əldə edirik
                   .ToArray(); // Nəticələri massivə çeviririk

                // Nəticələri istifadə etmək üçün burada əlavə əməliyyatlar edə bilərsiniz
                // Məsələn, nəticələri ekrana yazdırmaq:
                foreach (var kod in hes_kod_1)
                {
                    Console.WriteLine(kod);
                }
            }
        }


        private void frm_CorresPondent_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;


            // Keçmiş günləri seçməyə icazə vermək üçün heç bir məhdudiyyət qoymuruq
            // Amma bu məhdudiyyətlərə riayət etmək üçün əlavə nəzarət kodu yazaq.
        }
        private void CreateCustomCalendar(int year, int month)
        {
            // Əvvəlcə köhnə düymələri təmizləyirik
            panelCalendar.Controls.Clear();

            // Ayın ilk gününü və gün sayını tapırıq
            DateTime firstDay = new DateTime(year, month, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Təqvim üçün başlanğıc koordinatları
            int startX = 10, startY = 10;
            int buttonSize = 40;
            int margin = 5;

            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime currentDay = new DateTime(year, month, day);

                Button btnDay = new Button
                {
                    Text = day.ToString(),
                    Width = buttonSize,
                    Height = buttonSize,
                    Left = startX + ((day - 1) % 7) * (buttonSize + margin),
                    Top = startY + ((day - 1) / 7) * (buttonSize + margin),
                    Tag = currentDay
                };

                // Bugünkü və gələcək günlər üçün rəngləmə
                if (currentDay >= DateTime.Now.Date)
                {
                    btnDay.BackColor = Color.LightGray; // Gələcək günlər boz rəngdə
                    btnDay.Enabled = false; // Gələcək günləri deaktiv edirik
                }
                else
                {
                    btnDay.BackColor = Color.LightGreen; // Keçmiş günlər yaşıl rəngdə
                }

                // Düyməni panellə əlavə edirik
                panelCalendar.Controls.Add(btnDay);
            }
        }
        private void monthCalendar1_DateSelected_1(object sender, DateRangeEventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            btn_sorgu.Enabled = true;
            // Bugünkü və gələcək tarixlər seçilibsə:
            if (e.Start >= today)
            {
                MessageBox.Show("Bugünkü və gələcək tarixləri seçmək mümkün deyil!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                // Seçimi geri qaytarırıq (istifadəçiyə xəbərdarlıq veririk və seçimi keçmiş tarixə qaytarırıq).
                monthCalendar1.SetSelectionRange(today.AddDays(-1), today.AddDays(-1)); // Dünənki tarixi seçili edir
                btn_sorgu.Enabled = false;
            }
            else
            {
            // MonthCalendar-dan seçilən tarixi al
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy"); // Oracle üçün uyğun tarix formatı

            // Oracle sorğusu
            string query = @"
        SELECT c.space_or_star AS gun
        FROM odb.calendar c
        WHERE c.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')";

                using (OracleConnection connection = new OracleConnection(cl.con))
                {
                    try
                    {

                        connection.Open();
                        using (OracleCommand command = new OracleCommand(query, connection))
                        {
                            // Tarixi parametr olaraq ötür
                            command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                            // Sorğunu icra et
                            object result = command.ExecuteScalar();

                            if (result != null && result.ToString() == "*")
                            {
                                btn_sorgu.Enabled = false;
                                // Əgər nəticə "*" isə, qeyri-iş günü mesajını göstər
                                MessageBox.Show("Seçdiyiniz tarix qeyri-iş günüdür!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Xəta baş verdikdə mesaj göstər
                        MessageBox.Show($"Xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        List<Hesabat> hesabatlar = new List<Hesabat>();
        private void excel()
        {
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy");

            string hesab_qaliqlari = @"SELECT SUBSTR(s.licsch, 0, 5) AS hes,l.name_licsch AS ad,
                case when (substr(s.licsch, 0, 3)='409' and substr(REGEXP_SUBSTR(l.name_licsch, '\(([^()]*)\)\s*$',1,1,NULL,1),5,1)='5') or SUBSTR(s.licsch, 0, 5)='45029'
                then 'qr' else 'r' end tip,ABS(s.saldo_vhd_nacval) AS saldo_vhd_nacval_abs, ABS(s.saldo_ish_nacval) AS qaliq
                FROM odb.arh_saldo_ls s,licsch l where l.licsch = s.licsch and s.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY') 
                and (l.date_close_licsch is null or l.date_close_licsch>= TO_DATE(:selectedDate, 'DD-MM-YYYY'))
                order by SUBSTR(s.licsch, 0, 5) asc";

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(hesab_qaliqlari, connection))
                {
                    // Parametri əlavə edin
                    command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));
                    using (OracleDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            hesabatlar.Add(new Hesabat
                            {
                                Hes = reader["hes"].ToString(),
                                Ad = reader["ad"].ToString(),
                                Tip = reader["tip"].ToString(),
                                SaldoVhdNacvalAbs = Convert.ToDecimal(reader["qaliq"]),
                                SaldoIshNacvalAbs = Convert.ToDecimal(reader["qaliq"])
                            });
                        }
                    }
                }
            }

            // Excel faylının yolu
            string filePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "rezident, qeyri-rezident yeni.xlsx");
            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            cl.fileName = formattedDate + " rezident, qeyri-rezident yeni" + ".xlsx";
            if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.fileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                cl.fileName = $"{cl.fileName} - {fileCounter}.xlsx";
            }
            // Excel faylını açın
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet8"]; // İlk səhifəni seçirik

                // Ay və ilin son gününü tapırıq
                int year = selectedDate.Year;
                int month = selectedDate.Month;
                int lastDay = DateTime.DaysInMonth(year, month); // Ayın son gününü tapırıq

                // Yeni tarix formatını yaradırıq (ayın son günü)
                DateTime lastDayOfMonth = new DateTime(year, month, lastDay);
                string formattedLastDate = lastDayOfMonth.ToString("dd-MM-yyyy");

                // Excel-ə yazırıq
                worksheet.Cells[1, 4].Value = formattedLastDate;

                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 1; row <= rowCount; row++)
                {
                    string excelHes = worksheet.Cells[row, 1].Text.Trim();
                    if (string.IsNullOrEmpty(excelHes))
                        continue;

                    var matchingRecords = hesabatlar.Where(h => h.Hes.Equals(excelHes, StringComparison.OrdinalIgnoreCase)).ToList();

                    if (matchingRecords.Any())
                    {
                        decimal totalR = matchingRecords.Where(h => h.Tip == "r").Sum(h => h.SaldoIshNacvalAbs);
                        decimal totalQr = matchingRecords.Where(h => h.Tip == "qr").Sum(h => h.SaldoIshNacvalAbs);

                        worksheet.Cells[row, 4].Value = totalR > 0 ? (decimal?)totalR : null;
                        worksheet.Cells[row, 5].Value = totalQr > 0 ? (decimal?)totalQr : null;
                        worksheet.Cells[row, 3].Value = (totalR + totalQr) > 0 ? (decimal?)(totalR + totalQr) : null;
                    }
                }

                // Dəyişiklikləri yadda saxlayın
                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(cl.filePath);
            }
        }
        private void excel_()
        {
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy");

            string hesab_qaliqlari = @"SELECT SUBSTR(s.licsch, 0, 5) AS hes,l.name_licsch AS ad,
        case when (substr(s.licsch, 0, 3)='409' and substr(REGEXP_SUBSTR(l.name_licsch, '\(([^()]*)\)\s*$',1,1,NULL,1),5,1)='5') or SUBSTR(s.licsch, 0, 5)='45029'
        then 'qr' else 'r' end tip,ABS(s.saldo_vhd_nacval) AS saldo_vhd_nacval_abs, ABS(s.saldo_ish_nacval) AS qaliq
        FROM odb.arh_saldo_ls s,licsch l where l.licsch = s.licsch and s.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY') 
        and (l.date_close_licsch is null or l.date_close_licsch>= TO_DATE(:selectedDate, 'DD-MM-YYYY'))
        order by SUBSTR(s.licsch, 0, 5) asc";

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(hesab_qaliqlari, connection))
                {
                    command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));
                    using (OracleDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            hesabatlar.Add(new Hesabat
                            {
                                Hes = reader["hes"].ToString(),
                                Ad = reader["ad"].ToString(),
                                Tip = reader["tip"].ToString(),
                                SaldoVhdNacvalAbs = Convert.ToDecimal(reader["qaliq"]),
                                SaldoIshNacvalAbs = Convert.ToDecimal(reader["qaliq"])
                            });
                        }
                    }
                }
            }

            string filePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "rezident_qeyri-rezident_yeni.xlsx");
            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            cl.fileName = formattedDate + " rezident_qeyri-rezident_yeni.xlsx";

            if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.fileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                cl.fileName = $"{cl.fileName} - {fileCounter}.xlsx";
            }

            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet8");

                if (worksheet == null)
                {
                    MessageBox.Show("Excel faylında 'Sheet8' səhifəsi tapılmadı!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                worksheet.Cells[1, 4].Value = formattedDate;
                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 1; row <= rowCount; row++)
                {
                    string excelHes = worksheet.Cells[row, 1].Text.Trim();
                    if (string.IsNullOrEmpty(excelHes))
                        continue;

                    var matchingRecords = hesabatlar.Where(h => h.Hes.Equals(excelHes, StringComparison.OrdinalIgnoreCase)).ToList();

                    if (matchingRecords.Any())
                    {
                        decimal totalR = matchingRecords.Where(h => h.Tip == "r").Sum(h => h.SaldoIshNacvalAbs);
                        decimal totalQr = matchingRecords.Where(h => h.Tip == "qr").Sum(h => h.SaldoIshNacvalAbs);

                        worksheet.Cells[row, 4].Value = totalR > 0 ? (decimal?)totalR : null;
                        worksheet.Cells[row, 5].Value = totalQr > 0 ? (decimal?)totalQr : null;
                        worksheet.Cells[row, 3].Value = (totalR + totalQr) > 0 ? (decimal?)(totalR + totalQr) : null;
                    }
                }

                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath));

                if (File.Exists(cl.filePath))
                {
                    System.Diagnostics.Process.Start(cl.filePath);
                }
                else
                {
                    MessageBox.Show("Excel faylı uğurla saxlanmadı!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void digerMetod()// kodun islediyini yoxlamaq ucun lisboxa atiriq
        {
            //// 'hes_kod_1' massivindən istifadə edirik
            //if (hes_kod_1 != null)
            //{
            //    // ListBox-u təmizləyirik
            //    listBox1.Items.Clear();

            //    // 'hes_kod_1' massivinin elementlərini ListBox-a əlavə edirik
            //    foreach (var kod in hes_kod_1)
            //    {
            //        listBox1.Items.Add(kod);
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("hes_kod_1 massivi hələ təyin edilməyib.");
            //}
        }
        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            hesablari_excel_cek();
            excel();
        }
    }
}
