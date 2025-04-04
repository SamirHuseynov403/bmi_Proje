using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using BMI.Muhasibat;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DocumentFormat.OpenXml.Wordprocessing;
using Npgsql;
using Oracle.ManagedDataAccess.Client;
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;

namespace BMI.UTimeMaster
{
    public partial class frm_InputOutput : Form
    {
        cl_yanasmalar cl = new cl_yanasmalar();
        cl_Database db = new cl_Database();
        Aletler aletler = new Aletler();
        public frm_InputOutput()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        #region tek yoxlama
//        SELECT
//        pe.first_name,
//        pe.last_name,
//        it.emp_code,
//    TO_CHAR(MIN(it.punch_time AT TIME ZONE 'UTC'), 'HH24:MI') AS giris,
//    TO_CHAR(MAX(it.punch_time AT TIME ZONE 'UTC'), 'HH24:MI') AS cixis
//FROM
//    iclock_transaction it
//JOIN
//    personnel_employee pe ON it.emp_code = pe.emp_code
//WHERE
//    it.emp_code = '2'
//    AND it.punch_time::date = '2025-03-03'
//GROUP BY
//    pe.first_name, pe.last_name, it.emp_code;
        #endregion
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            //btn_sorgu.Enabled = true;
            // Bugünkü və gələcək tarixlər seçilibsə:
            if (e.Start.Date > today.Date)
            {
                MessageBox.Show("Gələcək tarixləri seçmək mümkün deyil!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Seçimi geri qaytarırıq (istifadəçiyə xəbərdarlıq veririk və seçimi keçmiş tarixə qaytarırıq).
                monthCalendar1.SetSelectionRange(today.AddDays(-1), today.AddDays(-1)); // Dünənki tarixi seçili edir
                //btn_sorgu.Enabled = false;
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
                                //btn_sorgu.Enabled = false;
                                // Əgər nəticə "*" isə, qeyri-iş günü mesajını göstər
                                MessageBox.Show("Seçdiyiniz tarix qeyri-iş günüdür!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }
                            else
                            {
                                bazadancek(monthCalendar1.SelectionStart.Date);
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
        private void bazadancek(DateTime selectedDate)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(cl.connStringUTimemaster))
            {
                try
                {
                    conn.Open();
                    string sql = @"SELECT 
    pe.first_name,
    @selectedDate AS work_date,
    
    COALESCE((
        SELECT MIN(TO_CHAR(t.punch_time AT TIME ZONE 'UTC', 'HH24:MI'))
        FROM iclock_transaction t
        WHERE t.emp_code = pe.emp_code
          AND DATE(t.punch_time AT TIME ZONE 'UTC+8' AT TIME ZONE 'Asia/Baku') = @selectedDate
    ), '') AS giris,

   
    COALESCE((
        SELECT MAX(TO_CHAR(t.punch_time AT TIME ZONE 'UTC', 'HH24:MI'))
        FROM iclock_transaction t
        WHERE t.emp_code = pe.emp_code
          AND DATE(t.punch_time AT TIME ZONE 'UTC+8' AT TIME ZONE 'Asia/Baku') = @selectedDate
    ), '') AS cixis,

    COALESCE((
        SELECT CASE 
            WHEN MIN((t.punch_time AT TIME ZONE 'UTC')::time) > TIME '09:01:00' THEN 
                ROUND(EXTRACT(EPOCH FROM (MIN((t.punch_time AT TIME ZONE 'UTC')::time) - TIME '09:00:00')) / 60)::int || ' dəqiqə'
            ELSE ''
        END
        FROM iclock_transaction t
        WHERE t.emp_code = pe.emp_code
          AND DATE(t.punch_time AT TIME ZONE 'UTC+8' AT TIME ZONE 'Asia/Baku') = @selectedDate
    ), '') AS gecikme

FROM 
    personnel_employee pe
LEFT JOIN 
    iclock_transaction it ON pe.emp_code = it.emp_code
    AND DATE(it.punch_time AT TIME ZONE 'UTC+8' AT TIME ZONE 'Asia/Baku') = @selectedDate
WHERE pe.emp_code NOT IN ('23')
GROUP BY 
    pe.first_name, pe.emp_code
ORDER BY 
    pe.emp_code;


";

                    using (NpgsqlCommand cmd = new NpgsqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@selectedDate", selectedDate);

                        NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        dtg_InputOutput.Rows.Clear(); // Cədvəli təmizləyirik

                        foreach (DataRow row in dt.Rows)
                        {
                            string girisSaat = row["giris"].ToString();
                            string cixisSaat = row["cixis"].ToString();  // Çıxış saatını oxuyuruq
                            string gecikme = "";

                            // Cixis saatı giris saatı ilə bərabərdirsə boş qoyulacaq
                            if (girisSaat == cixisSaat)
                            {
                                cixisSaat = "";  // Çıxış saatını boş edirik
                            }

                            if (TimeSpan.TryParse(girisSaat, out TimeSpan girisTime))
                            {
                                TimeSpan standartSaat = new TimeSpan(9, 0, 0); // 09:00:00
                                if (girisTime > standartSaat)
                                {
                                    TimeSpan gecikmeTime = girisTime - standartSaat;
                                    gecikme = gecikmeTime.ToString(@"hh\:mm"); // 00:00 formatında
                                }
                            }

                            dtg_InputOutput.Rows.Add(
                                row["first_name"].ToString(),
                                Convert.ToDateTime(row["work_date"]).ToString("dd-MM-yyyy"),
                                girisSaat,
                                cixisSaat,  // Boş və ya real çıxış saatını göstəririk
                                gecikme
                            );
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message);
                }
            }
        }
        void ExportToExcel()
        {
            DateTime selectedDate = monthCalendar1.SelectionStart.Date;

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            // Başlıq
            worksheet.Cells[1, 1] = "Total Time Card Report";
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 1].Font.Size = 14;

            // Tarix
            worksheet.Cells[2, 1] = $"From {selectedDate.ToString("MMMM dd yyyy")} To {selectedDate.ToString("MMMM dd yyyy")}";

            // 3 və 4-cü sətirlər boş qalır

            // Başlıqlar (5-ci sətir)
            worksheet.Cells[5, 1] = "First Name";
            worksheet.Cells[5, 2] = "Date";
            worksheet.Cells[5, 3] = "Clock In";
            worksheet.Cells[5, 4] = "Clock Out";
            worksheet.Cells[5, 5] = "Late";

            for (int i = 1; i <= 5; i++)
            {
                worksheet.Cells[5, i].Font.Bold = true;
                worksheet.Cells[5, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray); // Başlıq fon rəngi
            }

            int rowCount = dtg_InputOutput.Rows.Count - 1; // son boş sətiri nəzərə alma
            int totalRows = rowCount + 5; // ümumi dolu sətir sayı

            // Məlumatları yazırıq (6-cı sətrdən başlayır)
            for (int i = 0; i < rowCount; i++)
            {
                worksheet.Cells[i + 6, 1] = dtg_InputOutput.Rows[i].Cells[0].Value?.ToString();
                worksheet.Cells[i + 6, 2] = dtg_InputOutput.Rows[i].Cells[1].Value?.ToString();
                worksheet.Cells[i + 6, 3] = dtg_InputOutput.Rows[i].Cells[2].Value?.ToString();
                worksheet.Cells[i + 6, 4] = dtg_InputOutput.Rows[i].Cells[3].Value?.ToString();
                worksheet.Cells[i + 6, 5] = dtg_InputOutput.Rows[i].Cells[4].Value?.ToString();
            }

            // Sütun ölçülərini uyğunlaşdırırıq
            worksheet.Columns.AutoFit();

            // Cədvələ ramka əlavə edirik
            Microsoft.Office.Interop.Excel.Range borderRange = worksheet.Range["A5", $"E{totalRows}"];
            borderRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borderRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            excelApp.Visible = true;
            // Fayl avtomatik saxlanılmır, sadəcə açılır
            //MessageBox.Show("Excel faylı yaradıldı və ekranda açıldı!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btn_Elave_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }
    }
}
