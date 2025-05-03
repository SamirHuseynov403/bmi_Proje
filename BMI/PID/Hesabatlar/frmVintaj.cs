using BMI.Muhasibat;
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
using OfficeOpenXml;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMI.PID.Hesabatlar
{
    public partial class frmVintaj : Form
    {
        public frmVintaj()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        DataTable dtexcelden = new DataTable();

        void excel()
        {
            string vintaj = @"WITH raw_tt AS (
                        SELECT
                            x.licschpkre AS hesab,
                            x.subschkre AS sk,
                            s.krnovlar AS kr_novu,
                            s.krtip AS kr_tipi,
                            s.indexsler AS kr_indexi,
                            LAST_DAY(s.date_open) AS vert,
                            s.meb AS kr,
                            odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper)) AS gecikme_gun,
                            ROW_NUMBER() OVER (PARTITION BY x.licschpkre, x.subschkre ORDER BY NULL) AS rn
                        FROM view_nacpogprokre_all x
                        JOIN (
                            SELECT
                                ar.licschpkre AS hes,
                                ar.subschkre AS sub,
                                ar.summakre AS meb,
                                CASE WHEN tp.code IN (1,3) THEN 'biznes' ELSE tp.name END AS krnovlar,
                                ar.index_otrasli AS indexsler,
                                t.name AS krtip,
                                ar.date_open
                            FROM arh_licschkre ar
                            JOIN tipkre tp ON ar.tipkredita = tp.code
                            JOIN tipzal t ON ar.tipzaloga = t.code
                            WHERE ar.date_open BETWEEN TO_DATE(:dovrevvel, 'DD-MM-YYYY') AND TO_DATE(:dovrson, 'DD-MM-YYYY')
                              AND ar.date_oper = TO_DATE(:dateoper, 'DD-MM-YYYY')
                              AND (ar.date_close IS NULL OR ar.date_close >= TO_DATE(:dovrevvel, 'DD-MM-YYYY'))
                        ) s ON x.licschpkre = s.hes AND x.subschkre = s.sub
                        WHERE x.date_oper = TO_DATE(:dateoper, 'DD-MM-YYYY')
                    ),

                    tt AS (
                        SELECT
                            hesab,
                            sk,
                            kr_tipi,
                            kr_novu,
                            kr_indexi,
                            vert,
                            kr,
                            CASE 
                                WHEN gecikme_gun BETWEEN 90 AND 120 THEN 4
                                WHEN gecikme_gun BETWEEN 121 AND 150 THEN 5
                                WHEN gecikme_gun BETWEEN 151 AND 180 THEN 6
                                WHEN gecikme_gun BETWEEN 181 AND 210 THEN 7
                                WHEN gecikme_gun BETWEEN 211 AND 240 THEN 8
                                WHEN gecikme_gun BETWEEN 241 AND 270 THEN 9
                                WHEN gecikme_gun BETWEEN 271 AND 300 THEN 10
                                WHEN gecikme_gun BETWEEN 301 AND 330 THEN 11
                                WHEN gecikme_gun BETWEEN 331 AND 360 THEN 12
                                WHEN gecikme_gun BETWEEN 361 AND 390 THEN 13
                                WHEN gecikme_gun BETWEEN 391 AND 420 THEN 14
                                WHEN gecikme_gun BETWEEN 421 AND 450 THEN 15
                                WHEN gecikme_gun BETWEEN 451 AND 480 THEN 16
                                WHEN gecikme_gun BETWEEN 481 AND 510 THEN 17
                                WHEN gecikme_gun BETWEEN 511 AND 540 THEN 18
                                WHEN gecikme_gun BETWEEN 541 AND 570 THEN 19
                                WHEN gecikme_gun BETWEEN 571 AND 600 THEN 20
                                WHEN gecikme_gun BETWEEN 601 AND 630 THEN 21
                                WHEN gecikme_gun BETWEEN 631 AND 660 THEN 22
                                WHEN gecikme_gun BETWEEN 661 AND 690 THEN 23
                                WHEN gecikme_gun BETWEEN 691 AND 720 THEN 24
                                WHEN gecikme_gun BETWEEN 721 AND 750 THEN 25
                                WHEN gecikme_gun BETWEEN 751 AND 780 THEN 26
                                WHEN gecikme_gun BETWEEN 781 AND 810 THEN 27
                                WHEN gecikme_gun BETWEEN 811 AND 840 THEN 28
                                WHEN gecikme_gun BETWEEN 841 AND 870 THEN 29
                                WHEN gecikme_gun BETWEEN 871 AND 900 THEN 30
                                WHEN gecikme_gun BETWEEN 901 AND 930 THEN 31
                                WHEN gecikme_gun BETWEEN 931 AND 960 THEN 32
                                WHEN gecikme_gun BETWEEN 961 AND 990 THEN 33
                                WHEN gecikme_gun BETWEEN 991 AND 1020 THEN 34
                                WHEN gecikme_gun BETWEEN 1021 AND 1050 THEN 35
                                WHEN gecikme_gun BETWEEN 1051 AND 1080 THEN 36
                                ELSE 0
                            END AS gecikme_qrupu
                        FROM raw_tt
                        WHERE rn = 1
                    ),

                    total_kr AS (
                        SELECT
                            vert,
                            SUM(kr) AS total
                        FROM tt
                        GROUP BY vert
                    )

                    SELECT
                        TO_CHAR(tt.vert, 'DD-MM-YYYY') AS vert,
                        tt.kr_tipi,
                        tt.kr_novu,
                        tt.kr_indexi,
                        SUM(tt.kr) AS kr,
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 4  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""4"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 5  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""5"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 6  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""6"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 7  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""7"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 8  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""8"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 9  THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""9"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 10 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""10"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 11 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""11"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 12 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""12"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 13 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""13"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 14 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""14"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 15 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""15"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 16 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""16"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 17 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""17"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 18 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""18"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 19 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""19"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 20 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""20"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 21 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""21"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 22 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""22"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 23 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""23"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 24 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""24"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 25 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""25"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 26 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""26"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 27 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""27"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 28 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""28"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 29 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""29"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 30 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""30"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 31 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""31"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 32 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""32"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 33 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""33"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 34 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""34"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 35 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""35"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 36 THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""36""
                    FROM tt
                    JOIN total_kr tk ON tk.vert = tt.vert
                    GROUP BY tt.vert, tt.kr_tipi, tt.kr_novu, tt.kr_indexi, tk.total
                    ORDER BY tt.vert
                    ";
            System.Data.DataTable _dt_vintaj = new System.Data.DataTable();

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(vintaj, connection))
                {
                    command.Parameters.Add("dovrevvel", OracleDbType.Varchar2).Value = dtDovrİlk.Value.ToString("dd-MM-yyyy");
                    command.Parameters.Add("dovrson", OracleDbType.Varchar2).Value = dtDovrSon.Value.ToString("dd-MM-yyyy");
                    command.Parameters.Add("dateoper", OracleDbType.Varchar2).Value = dtHesabatTarixi.Value.ToString("dd-MM-yyyy");
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_vintaj);
                }
                connection.Close();
            }

            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = dtHesabatTarixi.Text; // textBoxText dəyəriniz
            string monthYear = textBoxText.Replace("-", "");
            cl.baseFileName = "Vintaj PAR90 " + monthYear; // Temel dosya adı
            cl.fileName = cl.baseFileName + ".xlsx";
            cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "PID", "Exceller", "Vintaj PAR90.xlsx");
            cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);

            if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                cl.fileName = $"{cl.baseFileName} - {fileCounter}.xlsx";
            }
            FileInfo templateFile = new FileInfo(cl.templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Cəmi"];

                // Başlanğıc və son tarixləri al
                DateTime startDate = dtDovrİlk.Value;
                DateTime endDate = dtDovrSon.Value;

                // İndi aylıq olaraq get
                DateTime currentDate = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                #region Toplam Portfel


                // B5-dən başla
                int startRow = 5;


                while (currentDate <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, DateTime.DaysInMonth(currentDate.Year, currentDate.Month));
                    worksheet1.Cells[startRow, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRow, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate)  // yalnız tarixlər uyğun gələndə
                                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRow, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRow, localStartColumnExcel].Value = totalgun/100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRow++;
                    currentDate = currentDate.AddMonths(1);
                }
                #endregion

                #region Biznes Portfel


                // B5-dən başla
                int startRowBiznes = 47;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateBiznes = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateBiznes <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateBiznes.Year, currentDateBiznes.Month, DateTime.DaysInMonth(currentDateBiznes.Year, currentDateBiznes.Month));
                    worksheet1.Cells[startRowBiznes, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowBiznes, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate &&  row.Field<string>("kr_novu") == "biznes") 
                                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowBiznes, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowBiznes, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowBiznes++;
                    currentDateBiznes = currentDateBiznes.AddMonths(1);
                }
                #endregion

                #region Istehlak Portfel


                // B5-dən başla
                int startRowIstehlak = 212;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlak = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlak <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlak.Year, currentDateIstehlak.Month, DateTime.DaysInMonth(currentDateIstehlak.Year, currentDateIstehlak.Month));
                    worksheet1.Cells[startRowIstehlak, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlak, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate && row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR")
                                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlak, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlak, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlak++;
                    currentDateIstehlak = currentDateIstehlak.AddMonths(1);
                }
                #endregion

                #region Avtomobil Portfel


                // B5-dən başla
                int startRowIstehlakAvtomobil = 254;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakAvtomobil = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakAvtomobil <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakAvtomobil.Year, currentDateIstehlakAvtomobil.Month, DateTime.DaysInMonth(currentDateIstehlakAvtomobil.Year, currentDateIstehlakAvtomobil.Month));
                    worksheet1.Cells[startRowIstehlakAvtomobil, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakAvtomobil, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1905)                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakAvtomobil, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakAvtomobil, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakAvtomobil++;
                    currentDateIstehlakAvtomobil = currentDateIstehlakAvtomobil.AddMonths(1);
                }
                #endregion

                #region Kart Portfel


                // B5-dən başla
                int startRowIstehlakKart = 296;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakKart = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakKart <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakKart.Year, currentDateIstehlakKart.Month, DateTime.DaysInMonth(currentDateIstehlakKart.Year, currentDateIstehlakKart.Month));
                    worksheet1.Cells[startRowIstehlakKart, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakKart, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1907)                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakKart, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakKart, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakKart++;
                    currentDateIstehlakKart = currentDateIstehlakKart.AddMonths(1);
                }
                #endregion

                #region Temir Portfel


                // B5-dən başla
                int startRowIstehlakTemir = 422;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakTemir = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakTemir <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakTemir.Year, currentDateIstehlakTemir.Month, DateTime.DaysInMonth(currentDateIstehlakTemir.Year, currentDateIstehlakTemir.Month));
                    worksheet1.Cells[startRowIstehlakTemir, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakTemir, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  (row.Field<decimal>("kr_indexi") == 1903 || row.Field<decimal>("kr_indexi") == 1904))                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakTemir, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakTemir, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakTemir++;
                    currentDateIstehlakTemir = currentDateIstehlakTemir.AddMonths(1);
                }
                #endregion

                #region Diger Portfel


                // B5-dən başla
                int startRowIstehlakDiger = 464;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakDiger = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakDiger <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakDiger.Year, currentDateIstehlakDiger.Month, DateTime.DaysInMonth(currentDateIstehlakDiger.Year, currentDateIstehlakDiger.Month));
                    worksheet1.Cells[startRowIstehlakDiger, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakDiger, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                        .Where(row =>
                            DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                            sqlDate.Date == excelDate &&
                            row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&
                            ((int)row.Field<decimal>("kr_indexi") != 1903) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1904) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1905) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1907) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1901) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1902)
                        ).ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakDiger, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakDiger, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakDiger++;
                    currentDateIstehlakDiger = currentDateIstehlakDiger.AddMonths(1);
                }
                #endregion

                #region Dasinmaz Portfel


                // B5-dən başla
                int startRowIstehlakDasinmaz = 506;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakDasinmaz = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakDasinmaz <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakDasinmaz.Year, currentDateIstehlakDasinmaz.Month, DateTime.DaysInMonth(currentDateIstehlakDasinmaz.Year, currentDateIstehlakDasinmaz.Month));
                    worksheet1.Cells[startRowIstehlakDasinmaz, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakDasinmaz, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  (row.Field<decimal>("kr_indexi") == 1901 || row.Field<decimal>("kr_indexi") == 1902))                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (cemPortfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = cemPortfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakDasinmaz, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 5;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakDasinmaz, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakDasinmaz++;
                    currentDateIstehlakDasinmaz = currentDateIstehlakDasinmaz.AddMonths(1);
                }
                #endregion

                List<string> tarixler = new List<string>();

                for (int row = 5; row <= 41; row++)
                {
                    var cellValue = worksheet1.Cells[row, 2].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        tarixler.Add(cellValue);
                    }
                }

                // İndi bu tarixləri istədiyin sətirlərə yapışdırırıq
                int[] targetRows = { 47, 88, 129, 170, 212, 254, 296, 338, 380, 422, 464, 506, 548, 590 };

                // Hər bir target sətirə tarixləri yerləşdiririk
                for (int i = 0; i < targetRows.Length; i++)
                {
                    for (int j = 0; j < tarixler.Count; j++)
                    {
                        worksheet1.Cells[targetRows[i] + j, 2].Value = tarixler[j];
                    }
                }

                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet

                // Sonra onu .xlsb formatına çevir və aç  (XLSX NI XLSB YA CEVIRMEK)
                ConvertToXlsbAndOpen(cl.filePath);

                void ConvertToXlsbAndOpen(string xlsxPath)
                {
                    string xlsbPath = Path.ChangeExtension(xlsxPath, ".xlsb");

                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.DisplayAlerts = false;

                    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(xlsxPath);
                    workbook.SaveAs(xlsbPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12); // .xlsb formatı
                    workbook.Close(false);

                    excelApp.Quit();

                    // .xlsb faylını aç
                    System.Diagnostics.Process.Start(xlsbPath);
                }
            }
        }

        void excelrest()
        {
            string vintaj = @"WITH raw_tt AS (
                        SELECT
                            x.licschpkre AS hesab,
                            x.subschkre AS sk,
                            s.krnovlar AS kr_novu,
                            s.krtip AS kr_tipi,
                            s.indexsler AS kr_indexi,
                            LAST_DAY(s.date_open) AS vert,
                            s.meb AS kr,
                            s.res AS kr_res,
                            odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper)) AS gecikme_gun,
                            ROW_NUMBER() OVER (PARTITION BY x.licschpkre, x.subschkre ORDER BY NULL) AS rn
                        FROM view_nacpogprokre_all x
                        JOIN (
                            SELECT
                                ar.licschpkre AS hes,
                                ar.subschkre AS sub,
                                ar.summakre AS meb,
                                CASE WHEN tp.code IN (1,3) THEN 'biznes' ELSE tp.name END AS krnovlar,
                                ar.index_otrasli AS indexsler,
                                t.name AS krtip,
                                ar.date_open,
                                ar.date_restructure res
                            FROM arh_licschkre ar
                            JOIN tipkre tp ON ar.tipkredita = tp.code
                            JOIN tipzal t ON ar.tipzaloga = t.code
                            WHERE ar.date_open BETWEEN TO_DATE(:dovrevvel, 'DD-MM-YYYY') AND TO_DATE(:dovrson, 'DD-MM-YYYY')
                              AND ar.date_oper = TO_DATE(:dateoper, 'DD-MM-YYYY') 
                              AND (ar.date_close IS NULL OR ar.date_close >= TO_DATE(:dovrevvel, 'DD-MM-YYYY'))
                        ) s ON x.licschpkre = s.hes AND x.subschkre = s.sub
                        WHERE x.date_oper = TO_DATE(:dateoper, 'DD-MM-YYYY')
                    ),

                    tt AS (
                        SELECT
                            hesab,
                            sk,
                            kr_tipi,
                            kr_novu,
                            kr_indexi,
                            vert,
                            kr,
                            kr_res,
                            CASE 
                                WHEN gecikme_gun BETWEEN 2 AND 30  THEN 1
                                WHEN gecikme_gun BETWEEN 31 AND 60 THEN 2
                                WHEN gecikme_gun BETWEEN 61 AND 90 THEN 3
                                WHEN gecikme_gun BETWEEN 91 AND 120 THEN 4
                                WHEN gecikme_gun BETWEEN 121 AND 150 THEN 5
                                WHEN gecikme_gun BETWEEN 151 AND 180 THEN 6
                                WHEN gecikme_gun BETWEEN 181 AND 210 THEN 7
                                WHEN gecikme_gun BETWEEN 211 AND 240 THEN 8
                                WHEN gecikme_gun BETWEEN 241 AND 270 THEN 9
                                WHEN gecikme_gun BETWEEN 271 AND 300 THEN 10
                                WHEN gecikme_gun BETWEEN 301 AND 330 THEN 11
                                WHEN gecikme_gun BETWEEN 331 AND 360 THEN 12
                                WHEN gecikme_gun BETWEEN 361 AND 390 THEN 13
                                WHEN gecikme_gun BETWEEN 391 AND 420 THEN 14
                                WHEN gecikme_gun BETWEEN 421 AND 450 THEN 15
                                WHEN gecikme_gun BETWEEN 451 AND 480 THEN 16
                                WHEN gecikme_gun BETWEEN 481 AND 510 THEN 17
                                WHEN gecikme_gun BETWEEN 511 AND 540 THEN 18
                                WHEN gecikme_gun BETWEEN 541 AND 570 THEN 19
                                WHEN gecikme_gun BETWEEN 571 AND 600 THEN 20
                                WHEN gecikme_gun BETWEEN 601 AND 630 THEN 21
                                WHEN gecikme_gun BETWEEN 631 AND 660 THEN 22
                                WHEN gecikme_gun BETWEEN 661 AND 690 THEN 23
                                WHEN gecikme_gun BETWEEN 691 AND 720 THEN 24
                                WHEN gecikme_gun BETWEEN 721 AND 750 THEN 25
                                WHEN gecikme_gun BETWEEN 751 AND 780 THEN 26
                                WHEN gecikme_gun BETWEEN 781 AND 810 THEN 27
                                WHEN gecikme_gun BETWEEN 811 AND 840 THEN 28
                                WHEN gecikme_gun BETWEEN 841 AND 870 THEN 29
                                WHEN gecikme_gun BETWEEN 871 AND 900 THEN 30
                                WHEN gecikme_gun BETWEEN 901 AND 930 THEN 31
                                WHEN gecikme_gun BETWEEN 931 AND 960 THEN 32
                                WHEN gecikme_gun BETWEEN 961 AND 990 THEN 33
                                WHEN gecikme_gun BETWEEN 991 AND 1020 THEN 34
                                WHEN gecikme_gun BETWEEN 1021 AND 1050 THEN 35
                                WHEN gecikme_gun BETWEEN 1051 AND 1080 THEN 36
                                ELSE 0
                            END AS gecikme_qrupu
                        FROM raw_tt
                        WHERE rn = 1
                    ),

                    total_kr AS (
                        SELECT
                            vert,
                            SUM(kr) AS total
                        FROM tt
                        GROUP BY vert
                    )

                    SELECT
                        TO_CHAR(tt.vert, 'DD-MM-YYYY') AS vert,
                        tt.kr_tipi,
                        tt.kr_novu,
                        tt.kr_indexi,
                        tt.kr_res,
                        SUM(tt.kr) AS kr,
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 1  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""1"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 2  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""2"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 3  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""3"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 4  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""4"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 5  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""5"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 6  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""6"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 7  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""7"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 8  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""8"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 9  AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""9"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 10 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""10"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 11 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""11"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 12 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""12"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 13 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""13"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 14 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""14"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 15 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""15"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 16 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""16"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 17 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""17"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 18 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""18"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 19 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""19"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 20 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""20"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 21 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""21"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 22 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""22"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 23 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""23"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 24 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""24"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 25 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""25"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 26 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""26"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 27 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""27"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 28 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""28"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 29 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""29"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 30 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""30"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 31 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""31"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 32 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""32"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 33 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""33"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 34 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""34"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 35 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""35"",
                        ROUND(100 * SUM(CASE WHEN tt.gecikme_qrupu = 36 AND tt.kr_res IS NOT NULL THEN tt.kr ELSE 0 END) / tk.total, 2) AS ""36""

                    FROM tt
                    JOIN total_kr tk ON tk.vert = tt.vert
                    GROUP BY tt.vert, tt.kr_tipi, tt.kr_novu, tt.kr_indexi, tk.total,tt.kr_res
                    ORDER BY tt.vert
                    ";
            System.Data.DataTable _dt_vintaj = new System.Data.DataTable();

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(vintaj, connection))
                {
                    command.Parameters.Add("dovrevvel", OracleDbType.Varchar2).Value = dtDovrİlk.Value.ToString("dd-MM-yyyy");
                    command.Parameters.Add("dovrson", OracleDbType.Varchar2).Value = dtDovrSon.Value.ToString("dd-MM-yyyy");
                    command.Parameters.Add("dateoper", OracleDbType.Varchar2).Value = dtHesabatTarixi.Value.ToString("dd-MM-yyyy");
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_vintaj);
                }
                connection.Close();
            }

            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = dtHesabatTarixi.Text; // textBoxText dəyəriniz
            string monthYear = textBoxText.Replace("-", "");
            cl.baseFileName = "Vintaj Rest " + monthYear; // Temel dosya adı
            cl.fileName = cl.baseFileName + ".xlsx";
            cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "PID", "Exceller", "Vintaj Rest.xlsx");
            cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);

            if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                cl.fileName = $"{cl.baseFileName} - {fileCounter}.xlsx";
            }
            FileInfo templateFile = new FileInfo(cl.templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Cəmi"];

                // Başlanğıc və son tarixləri al
                DateTime startDate = dtDovrİlk.Value;
                DateTime endDate = dtDovrSon.Value;

                // İndi aylıq olaraq get
                DateTime currentDate = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                #region Toplam Portfel


                // B5-dən başla
                int startRow = 5;


                while (currentDate <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, DateTime.DaysInMonth(currentDate.Year, currentDate.Month));
                    worksheet1.Cells[startRow, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRow, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate 
                                                  && !row.IsNull("kr_res")) // yalnız tarixlər uyğun gələndə
                                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate) // yalnız tarixlər uyğun gələndə
                                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRow, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRow, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRow++;
                    currentDate = currentDate.AddMonths(1);
                }
                #endregion

                #region Biznes Portfel


                // B5-dən başla
                int startRowBiznes = 47;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateBiznes = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateBiznes <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateBiznes.Year, currentDateBiznes.Month, DateTime.DaysInMonth(currentDateBiznes.Year, currentDateBiznes.Month));
                    worksheet1.Cells[startRowBiznes, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowBiznes, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate && row.Field<string>("kr_novu") == "biznes" 
                                                  && !row.IsNull("kr_res"))
                                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate && row.Field<string>("kr_novu") == "biznes"
                                                  )
                                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowBiznes, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowBiznes, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowBiznes++;
                    currentDateBiznes = currentDateBiznes.AddMonths(1);
                }
                #endregion

                #region Istehlak Portfel


                // B5-dən başla
                int startRowIstehlak = 212;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlak = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlak <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlak.Year, currentDateIstehlak.Month, DateTime.DaysInMonth(currentDateIstehlak.Year, currentDateIstehlak.Month));
                    worksheet1.Cells[startRowIstehlak, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlak, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate && row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR"
                                                  && !row.IsNull("kr_res"))
                                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                                  sqlDate.Date == excelDate && row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR"
                                                  )
                                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlak, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlak, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlak++;
                    currentDateIstehlak = currentDateIstehlak.AddMonths(1);
                }
                #endregion

                #region Avtomobil Portfel


                // B5-dən başla
                int startRowIstehlakAvtomobil = 254;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakAvtomobil = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakAvtomobil <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakAvtomobil.Year, currentDateIstehlakAvtomobil.Month, DateTime.DaysInMonth(currentDateIstehlakAvtomobil.Year, currentDateIstehlakAvtomobil.Month));
                    worksheet1.Cells[startRowIstehlakAvtomobil, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakAvtomobil, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1905
                                  && !row.IsNull("kr_res"))                 // bu rəqəmdir, int-dir
                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1905
                                  )                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakAvtomobil, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakAvtomobil, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakAvtomobil++;
                    currentDateIstehlakAvtomobil = currentDateIstehlakAvtomobil.AddMonths(1);
                }
                #endregion

                #region Kart Portfel


                // B5-dən başla
                int startRowIstehlakKart = 296;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakKart = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakKart <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakKart.Year, currentDateIstehlakKart.Month, DateTime.DaysInMonth(currentDateIstehlakKart.Year, currentDateIstehlakKart.Month));
                    worksheet1.Cells[startRowIstehlakKart, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakKart, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1907
                                  && !row.IsNull("kr_res"))                 // bu rəqəmdir, int-dir
                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  row.Field<decimal>("kr_indexi") == 1907
                                  )                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakKart, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakKart, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakKart++;
                    currentDateIstehlakKart = currentDateIstehlakKart.AddMonths(1);
                }
                #endregion

                #region Temir Portfel


                // B5-dən başla
                int startRowIstehlakTemir = 422;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakTemir = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakTemir <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakTemir.Year, currentDateIstehlakTemir.Month, DateTime.DaysInMonth(currentDateIstehlakTemir.Year, currentDateIstehlakTemir.Month));
                    worksheet1.Cells[startRowIstehlakTemir, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakTemir, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  (row.Field<decimal>("kr_indexi") == 1903 || row.Field<decimal>("kr_indexi") == 1904)
                                  && !row.IsNull("kr_res"))                 // bu rəqəmdir, int-dir
                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  (row.Field<decimal>("kr_indexi") == 1903 || row.Field<decimal>("kr_indexi") == 1904)
                                  )                 // bu rəqəmdir, int-dir
                    .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakTemir, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakTemir, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakTemir++;
                    currentDateIstehlakTemir = currentDateIstehlakTemir.AddMonths(1);
                }
                #endregion

                #region Diger Portfel


                // B5-dən başla
                int startRowIstehlakDiger = 464;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakDiger = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakDiger <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakDiger.Year, currentDateIstehlakDiger.Month, DateTime.DaysInMonth(currentDateIstehlakDiger.Year, currentDateIstehlakDiger.Month));
                    worksheet1.Cells[startRowIstehlakDiger, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakDiger, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                        .Where(row =>
                            DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                            sqlDate.Date == excelDate &&
                            row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&
                            ((int)row.Field<decimal>("kr_indexi") != 1903) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1904) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1905) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1907) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1901) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1902)
                            && !row.IsNull("kr_res")
                        ).ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                        .Where(row =>
                            DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                            sqlDate.Date == excelDate &&
                            row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&
                            ((int)row.Field<decimal>("kr_indexi") != 1903) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1904) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1905) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1907) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1901) &&
                            ((int)row.Field<decimal>("kr_indexi") != 1902)
                            
                        ).ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakDiger, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakDiger, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakDiger++;
                    currentDateIstehlakDiger = currentDateIstehlakDiger.AddMonths(1);
                }
                #endregion

                #region Dasinmaz Portfel


                // B5-dən başla
                int startRowIstehlakDasinmaz = 506;

                // Başlanğıc və son tarixləri al

                // İndi aylıq olaraq get
                DateTime currentDateIstehlakDasinmaz = new DateTime(startDate.Year, startDate.Month, 1); // ayın əvvəli

                while (currentDateIstehlakDasinmaz <= endDate)
                {
                    DateTime lastDayOfMonth = new DateTime(currentDateIstehlakDasinmaz.Year, currentDateIstehlakDasinmaz.Month, DateTime.DaysInMonth(currentDateIstehlakDasinmaz.Year, currentDateIstehlakDasinmaz.Month));
                    worksheet1.Cells[startRowIstehlakDasinmaz, 2].Value = lastDayOfMonth.ToString("dd-MM-yyyy"); // B sütununa tarix yaz

                    var cellValue = worksheet1.Cells[startRowIstehlakDasinmaz, 2].Value?.ToString();
                    DateTime excelDate = DateTime.ParseExact(cellValue, "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // _dt_vintaj-dan filtrləyirik
                    var cemPortfel = _dt_vintaj.AsEnumerable()
                    .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                  sqlDate.Date == excelDate &&
                                  row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                  (row.Field<decimal>("kr_indexi") == 1901 || row.Field<decimal>("kr_indexi") == 1902)
                                  && !row.IsNull("kr_res"))                 // bu rəqəmdir, int-dir
                    .ToList();

                    var Portfel = _dt_vintaj.AsEnumerable()
                   .Where(row => DateTime.TryParse(row["vert"].ToString(), out DateTime sqlDate) &&
                                 sqlDate.Date == excelDate &&
                                 row.Field<string>("kr_novu") == "FİZİKİ ŞƏXSLƏR" &&  // bu yazıdı, string-dir
                                 (row.Field<decimal>("kr_indexi") == 1901 || row.Field<decimal>("kr_indexi") == 1902)
                                 )                 // bu rəqəmdir, int-dir
                   .ToList();

                    if (Portfel.Count > 0) // Əgər uyğun gələn data varsa
                    {
                        decimal total = Portfel.Sum(row => row.Field<decimal>("kr")) / 1000;
                        worksheet1.Cells[startRowIstehlakDasinmaz, 3].Value = total; // C sütununa ümumi məbləği yaz

                        int localStartColumnData = 6;   // DataTable-da 5-ci sütundan
                        int localStartColumnExcel = 4;  // Excel-də D sütunundan

                        for (int i = 4; i <= 36; i++)  // gecikmə faizləri 4-dən 36-ya qədər
                        {
                            decimal totalgun = cemPortfel.Sum(row => row.Field<decimal>(localStartColumnData));
                            worksheet1.Cells[startRowIstehlakDasinmaz, localStartColumnExcel].Value = totalgun / 100;

                            localStartColumnData++;
                            localStartColumnExcel++;
                        }
                    }
                    else
                    {
                        // Əgər uyğun gələn data yoxdursa, heç nə yazmırıq
                    }

                    startRowIstehlakDasinmaz++;
                    currentDateIstehlakDasinmaz = currentDateIstehlakDasinmaz.AddMonths(1);
                }
                #endregion

                List<string> tarixler = new List<string>();

                for (int row = 5; row <= 41; row++)
                {
                    var cellValue = worksheet1.Cells[row, 2].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        tarixler.Add(cellValue);
                    }
                }

                // İndi bu tarixləri istədiyin sətirlərə yapışdırırıq
                int[] targetRows = { 47, 88, 129, 170, 212, 254, 296, 338, 380, 422, 464, 506, 548, 590 };

                // Hər bir target sətirə tarixləri yerləşdiririk
                for (int i = 0; i < targetRows.Length; i++)
                {
                    for (int j = 0; j < tarixler.Count; j++)
                    {
                        worksheet1.Cells[targetRows[i] + j, 2].Value = tarixler[j];
                    }
                }

                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet

                // Sonra onu .xlsb formatına çevir və aç  (XLSX NI XLSB YA CEVIR)
                ConvertToXlsbAndOpen(cl.filePath);

                void ConvertToXlsbAndOpen(string xlsxPath)
                {
                    string xlsbPath = Path.ChangeExtension(xlsxPath, ".xlsb");

                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.DisplayAlerts = false;

                    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(xlsxPath);
                    workbook.SaveAs(xlsbPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12); // .xlsb formatı
                    workbook.Close(false);

                    excelApp.Quit();

                    // .xlsb faylını aç
                    System.Diagnostics.Process.Start(xlsbPath);
                }
            }
        }

        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            excel();
            excelrest();
        }
    }
}
