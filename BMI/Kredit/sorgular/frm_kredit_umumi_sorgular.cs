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
using BMI.Muhasibat;
//using Microsoft.Office.Interop.Excel;

namespace BMI
{
    public partial class frm_kredit_umumi_sorgular : Form
    {
        public frm_kredit_umumi_sorgular()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;
        
    //    private void exceleat()
    //    {
    //        string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    //        string excelFilePath = @"C:\BMI_\Umumi kredit sorgu.xlsx";

    //        // Veritabanı bağlantısı ve sorguları
    //        string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
    //        string queryVerilmisKrediler = "select l.index_otrasli,count(l.summakre),sum(l.summakre) from licschkre l " +
    //        " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') GROUP BY l.index_otrasli";

    //        string queryOdenilmisKrediler = "SELECT g.name," +
    //               "sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
    //              " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
    //              " sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS vk, " +
    //              " sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
    //        " FROM licschkre l " +
    //        " JOIN tipzal g ON l.tipzaloga = g.code " +
    //        " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
    //        " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
    //          "     AND l.subschkre = ar.ssk and substr(ar.debet,0,1) in (3,4) " +
    //        " GROUP BY g.name";

    //        //string queryHesaplanmisFaizler = "SELECT g.name," +
    //        //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
    //        //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
    //        //" FROM licschkre l " +
    //        //" JOIN tipzal g ON l.tipzaloga = g.code " +
    //        //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('30-10-2023', 'dd/mm/yyyy') " +
    //        //" WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
    //        //  "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
    //        //" GROUP BY g.name";

    //        // Veritabanından verileri çek
    //        System.Data.DataTable verilmisKredilerTable = new System.Data.DataTable();
    //        System.Data.DataTable odenilmisKredilerTable = new System.Data.DataTable();
    //        System.Data.DataTable hesaplanmisFaizlerTable = new System.Data.DataTable();

    //        using (OracleConnection connection = new OracleConnection(connectionString))
    //        {
    //            using (OracleCommand command = new OracleCommand(queryVerilmisKrediler, connection))
    //            {
    //                connection.Open();
    //                OracleDataAdapter adapter = new OracleDataAdapter(command);
    //                adapter.Fill(verilmisKredilerTable);
    //            }

    //            using (OracleCommand command = new OracleCommand(queryOdenilmisKrediler, connection))
    //            {
    //                OracleDataAdapter adapter = new OracleDataAdapter(command);
    //                adapter.Fill(odenilmisKredilerTable);
    //            }

    //            //using (OracleCommand command = new OracleCommand(queryHesaplanmisFaizler, connection))
    //            //{
    //            //    OracleDataAdapter adapter = new OracleDataAdapter(command);
    //            //    adapter.Fill(hesaplanmisFaizlerTable);
    //            //}
    //        }

    //        // Excel uygulamasını başlat
    //         Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
    //        Workbook workbook = excelApp.Workbooks.Add();

    //        // Her bir DataTable'i ilgili sayfaya yaz
    //        AddDataTableToExcelSheet(workbook, verilmisKredilerTable, "Verilmis kreditler");
    //        AddDataTableToExcelSheet(workbook, odenilmisKredilerTable, "Odenilmis kreditler");
    //        AddDataTableToExcelSheet(workbook, hesaplanmisFaizlerTable, "Hesablanmis faizler");

    //        // Excel dosyasını kaydet ve kapat
    //        workbook.SaveAs(excelFilePath);
    //        workbook.Close(false);
    //        excelApp.Quit();

    //        // Excel uygulamasını serbest bırak
    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

    //        MessageBox.Show("Excel dosyası oluşturuldu: " + excelFilePath, "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        
    //}
        private void exceleat2()
        {

            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            //duz olan
            string odenilmis_kr_nov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l,tipzal g,arh_dd ar " +
                         " WHERE l.tipzaloga = g.code and " +
                         " ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') and " +
                         " (ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         "  AND l.subschkre = ar.ssk and " +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) and " +
                         " substr(ar.kredit,11,6)||ar.ssk in (substr(l.licschkre,11,6)||ar.ssk, substr(l.licschpkre,11,6)||ar.ssk, substr(l.licsch_19,11,6)||ar.ssk, substr(l.licschppkre,11,6)||ar.ssk) " +
                         " GROUP BY g.name";

            string odenilmis_kr_girov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l " +
                         " JOIN tipkre g ON l.tipkredita = g.code " +
                         " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
                         " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         " AND l.subschkre = ar.ssk and" +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
                         " GROUP BY g.name";

            //string hesablanmis_faiz_girov = "SELECT g.name," +
            //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //      " FROM licschkre l " +
            //      " JOIN tipzal g ON l.tipzaloga = g.code " +
            //      " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('30-10-2023', 'dd/mm/yyyy') " +
            //      " WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
            //      "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
            //      " GROUP BY g.name";

            // Veritabanından verileri çek
            DataTable odenilmis_kr_NOV = new DataTable();
            DataTable odenilmis_kr_GIROV = new DataTable();
            //DataTable hesablanmis_faiz_GIROV = new DataTable();
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(odenilmis_kr_nov, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_NOV);
                }

                using (OracleCommand command = new OracleCommand(odenilmis_kr_girov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_GIROV);
                }

                //using (OracleCommand command = new OracleCommand(queryHesaplanmisFaizler, connection))
                //{
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(hesaplanmisFaizlerTable);
                //}
            }

                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
                string baseFileName = $"odemeler  "; // Temel dosya adı
                string fileName = baseFileName + ".xlsx";
                string templateFilePath = @"C:\BMI_\Umumi kredit sorgu.xlsx";
                if (File.Exists(Path.Combine(desktopFolder, fileName)))
                {
                    int fileCounter = 1;
                    while (File.Exists(Path.Combine(desktopFolder, $"{baseFileName} - {fileCounter}.xlsx")))
                    {
                        fileCounter++;
                    }
                    fileName = $"{baseFileName} - {fileCounter}.xlsx";
                }
                FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Odenilmis kreditler"];

                int ilksetirKreditnov = 3; // Başlangıç satır numarası
                int ilksutunKreditnov = 1; // Başlangıç sütun numarası
                int lastilksetirKreditnov = ilksetirKreditnov - 1;
                int ilksetirKreditnovnov = 3; // Başlangıç  satır numarası
                int ilksutunKreditnovnov = 6; // Başlangıç sütun numarası
                int lastilksetirKreditnovnov = ilksetirKreditnovnov - 1;

                // Her sütunun toplamını hesaplamak için bir dizi kullanın
                double[] columnTotals = new double[odenilmis_kr_NOV.Columns.Count];//NOVLER UZRE
                double[] columnGIROVTotals = new double[odenilmis_kr_GIROV.Columns.Count];

     //*******************************NOVLER UZRE
                for (int i = 0; i < odenilmis_kr_NOV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_NOV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = odenilmis_kr_NOV.Rows[i][j].ToString();
                        ilksutunKreditnovnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnovnov = 6; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnovnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = columnTotals[j];
                    ilksutunKreditnovnov++; // Sütun numarasını artır
                }
                for (int j = 6; j <= 9; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Size = 14;
                }

                //*******************************GIROVLAR UZRE
                for (int i = 0; i < odenilmis_kr_GIROV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_GIROV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnGIROVTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = odenilmis_kr_GIROV.Rows[i][j].ToString();
                        ilksutunKreditnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnov = 1; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = columnGIROVTotals[j];
                    ilksutunKreditnov++; // Sütun numarasını artır
                }
                for (int j = 1; j <= 4; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Size = 14;
                }

                worksheet.Cells[ilksetirKreditnovnov, 6].Value = "Toplam";
                worksheet.Cells[ilksetirKreditnov, 1].Value = "Toplam";

                string filePath = Path.Combine(desktopFolder, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);

            }
        }
        private void exceleat2test()
        {

            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            //duz olan
            string odenilmis_kr_nov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l,tipzal g,arh_dd ar " +
                         " WHERE l.tipzaloga = g.code and " +
                         " ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') and " +
                         " (ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         "  AND l.subschkre = ar.ssk and " +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) and " +
                         " substr(ar.kredit,11,6)||ar.ssk in (substr(l.licschkre,11,6)||ar.ssk, substr(l.licschpkre,11,6)||ar.ssk, substr(l.licsch_19,11,6)||ar.ssk, substr(l.licschppkre,11,6)||ar.ssk) " +
                         " GROUP BY g.name";

            string odenilmis_kr_girov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l " +
                         " JOIN tipkre g ON l.tipkredita = g.code " +
                         " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
                         " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         " AND l.subschkre = ar.ssk and" +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
                         " GROUP BY g.name";

            string hesablanmis_faiz_girov = "SELECT g.name, " +
        " SUM(CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
        "  SUM(CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz, " +
        " z.ischi, " +
        " z.faiz AS z_faiz, " +
        " z.vk_faiz AS z_vk_faiz " +
        " FROM " +
        "   licschkre l, tipkre g, arh_dd ar, " +
        "  (SELECT 'Isciler' AS ischi, " +
        "      SUM(CASE WHEN ar.debet = l.licschpkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschpkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
        "     SUM(CASE WHEN ar.debet = l.licschppkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschppkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
        " FROM " +
        "    licschkre l, arh_dd ar, regnom r " +
        " WHERE " +
        "    ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
        "  AND SUBSTR(ar.kredit, 1, 1) = 6 " +
        "  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
        "   AND l.subschkre = ar.ssd " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
        "   AND r.regnom = SUBSTR(l.licschpkre, 10, 6) " +
        "   AND r.svazanniy = 1 ) z " +
        " WHERE " +
        "    l.tipkredita = g.code " +
        "   AND ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
        "  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
        "   AND l.subschkre = ar.ssd " +
        "   AND SUBSTR(ar.kredit, 1, 1) IN(6) " +
        "  AND l.tipkredita <> 10 " +
        " GROUP BY " +
        "    g.name, z.ischi, z.faiz, z.vk_faiz";
            //l.tipzaloga = g.code
            string hesablanmis_faiz_nov = "SELECT g.name, " +
        " SUM(CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
        "  SUM(CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz, " +
        " z.ischi, " +
        " z.faiz AS z_faiz, " +
        " z.vk_faiz AS z_vk_faiz " +
        " FROM " +
        "   licschkre l, tipkre g, arh_dd ar, " +
        "  (SELECT 'Isciler' AS ischi, " +
        "      SUM(CASE WHEN ar.debet = l.licschpkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschpkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
        "     SUM(CASE WHEN ar.debet = l.licschppkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschppkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
        " FROM " +
        "    licschkre l, arh_dd ar, regnom r " +
        " WHERE " +
        "    ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
        "  AND SUBSTR(ar.kredit, 1, 1) = 6 " +
        "  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
        "   AND l.subschkre = ar.ssd " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
        "   AND r.regnom = SUBSTR(l.licschpkre, 10, 6) " +
        "   AND r.svazanniy = 1 ) z " +
        " WHERE " +
        "    l.tipkredita = g.code " +
        "   AND ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
        "  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
        "   AND l.subschkre = ar.ssd " +
        "   AND SUBSTR(ar.kredit, 1, 1) IN(6) " +
        "  AND l.tipkredita <> 10 " +
        " GROUP BY " +
        "    g.name, z.ischi, z.faiz, z.vk_faiz ";

            // Veritabanından verileri çek
            DataTable odenilmis_kr_NOV = new DataTable();
            DataTable odenilmis_kr_GIROV = new DataTable();
            DataTable hesablanmis_faiz_GIROV = new DataTable();
            DataTable hesablanmis_faiz_NOV = new DataTable();
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(odenilmis_kr_nov, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_NOV);
                }

                using (OracleCommand command = new OracleCommand(odenilmis_kr_girov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_GIROV);
                }

                using (OracleCommand command = new OracleCommand(hesablanmis_faiz_nov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(hesablanmis_faiz_NOV);
                }
                using (OracleCommand command = new OracleCommand(hesablanmis_faiz_girov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(hesablanmis_faiz_GIROV);
                }
            }

            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            string baseFileName = $"odemeler  "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\Umumi kredit sorgu.xlsx";
            if (File.Exists(Path.Combine(desktopFolder, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(desktopFolder, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Odenilmis kreditler"];
                ExcelWorksheet worksheet_faiz = package.Workbook.Worksheets["Hesablanmis faizler"];//*********
                int ilksetirKreditnov = 3; // Başlangıç satır numarası
                int ilksutunKreditnov = 1; // Başlangıç sütun numarası
                int lastilksetirKreditnov = ilksetirKreditnov - 1;
                int ilksetirKreditnovnov = 3; // Başlangıç  satır numarası
                int ilksutunKreditnovnov = 6; // Başlangıç sütun numarası
                int lastilksetirKreditnovnov = ilksetirKreditnovnov - 1;

                int ilksetirKreditnov_faiz = 3; // Başlangıç satır numarası
                int ilksutunKreditnov_faiz = 1; // Başlangıç sütun numarası
                int lastilksetirKreditnov_faiz = ilksetirKreditnov - 1;
                int ilksetirKreditnovnov_faiz = 3; // Başlangıç  satır numarası
                int ilksutunKreditnovnov_faiz = 5; // Başlangıç sütun numarası
                int lastilksetirKreditnovnov_faiz = ilksetirKreditnovnov - 1;

                // Her sütunun toplamını hesaplamak için bir dizi kullanın
                double[] columnTotals = new double[odenilmis_kr_NOV.Columns.Count];//NOVLER UZRE
                double[] columnGIROVTotals = new double[odenilmis_kr_GIROV.Columns.Count];

                double[] columnTotals_faiz = new double[hesablanmis_faiz_NOV.Columns.Count];//NOVLER UZRE
                double[] columnGIROVTotals_faiz = new double[hesablanmis_faiz_GIROV.Columns.Count];

                //*******************************NOVLER UZRE Odenilmis kreditler****************
                for (int i = 0; i < odenilmis_kr_NOV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_NOV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = odenilmis_kr_NOV.Rows[i][j].ToString();
                        ilksutunKreditnovnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnovnov = 6; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnovnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = columnTotals[j];
                    ilksutunKreditnovnov++; // Sütun numarasını artır
                }
                for (int j = 6; j <= 9; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Size = 14;
                }

                //*******************************GIROVLAR UZRE Odenilmis kreditler**********************
                for (int i = 0; i < odenilmis_kr_GIROV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_GIROV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnGIROVTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = odenilmis_kr_GIROV.Rows[i][j].ToString();
                        ilksutunKreditnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnov = 1; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = columnGIROVTotals[j];
                    ilksutunKreditnov++; // Sütun numarasını artır
                }
                for (int j = 1; j <= 4; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Size = 14;
                }

                //************************************ Hesablanmis faizler Uzre*******************************

                //*******************************NOVLER UZRE faizler
                for (int i = 0; i < hesablanmis_faiz_NOV.Rows.Count; i++)
                {
                    for (int j = 0; j < hesablanmis_faiz_NOV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(hesablanmis_faiz_NOV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnTotals_faiz[j] += cellValue;
                        }

                        worksheet_faiz.Cells[ilksetirKreditnovnov_faiz, ilksutunKreditnovnov_faiz].Value = hesablanmis_faiz_NOV.Rows[i][j].ToString();
                        ilksutunKreditnovnov_faiz++; // Sütun numarasını artır
                    }

                    ilksutunKreditnovnov_faiz = 5; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnovnov_faiz++; // Satır numarasını artır
                }
                for (int j = 0; j < hesablanmis_faiz_NOV.Columns.Count; j++)
                {
                    worksheet_faiz.Cells[ilksetirKreditnovnov_faiz, ilksutunKreditnovnov_faiz].Value = columnTotals_faiz[j];
                    ilksutunKreditnovnov_faiz++; // Sütun numarasını artır
                }
                for (int j = 5; j <= 7; j++)
                {
                    worksheet_faiz.Cells[ilksetirKreditnovnov_faiz, j].Style.Font.Bold = true;
                    worksheet_faiz.Cells[ilksetirKreditnovnov_faiz, j].Style.Font.Size = 14;
                }

                //*******************************GIROVLAR UZRE faizler
                for (int i = 0; i < hesablanmis_faiz_GIROV.Rows.Count; i++)
                {
                    for (int j = 0; j < hesablanmis_faiz_GIROV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(hesablanmis_faiz_GIROV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnGIROVTotals_faiz[j] += cellValue;
                        }

                        worksheet_faiz.Cells[ilksetirKreditnov_faiz, ilksutunKreditnov_faiz].Value = hesablanmis_faiz_GIROV.Rows[i][j].ToString();
                        ilksutunKreditnov_faiz++; // Sütun numarasını artır
                    }

                    ilksutunKreditnov_faiz = 1; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnov_faiz++; // Satır numarasını artır
                }
                for (int j = 0; j < hesablanmis_faiz_GIROV.Columns.Count; j++)
                {
                    worksheet_faiz.Cells[ilksetirKreditnov_faiz, ilksutunKreditnov_faiz].Value = columnGIROVTotals_faiz[j];
                    ilksutunKreditnov_faiz++; // Sütun numarasını artır
                }
                for (int j = 1; j <= 3; j++)
                {
                    worksheet_faiz.Cells[ilksetirKreditnov_faiz, j].Style.Font.Bold = true;
                    worksheet_faiz.Cells[ilksetirKreditnov_faiz, j].Style.Font.Size = 14;
                }


                worksheet.Cells[ilksetirKreditnovnov, 6].Value = "Toplam";
                worksheet.Cells[ilksetirKreditnov, 1].Value = "Toplam";

                worksheet_faiz.Cells[ilksetirKreditnovnov_faiz, 5].Value = "Toplam";
                worksheet_faiz.Cells[ilksetirKreditnov_faiz, 1].Value = "Toplam";



                string filePath = Path.Combine(desktopFolder, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);

            }
        }
        private void exceleat2_adlar()
        {

            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";

            string odenilmis_kr_nov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l " +
                         " JOIN tipzal g ON l.tipzaloga = g.code " +
                         " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
                         " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         "  AND l.subschkre = ar.ssk and " +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) and " +
                         " substr(ar.kredit,11,6)||ar.ssk in (substr(l.licschkre,11,6)||ar.ssk, substr(l.licschpkre,11,6)||ar.ssk, substr(l.licsch_19,11,6)||ar.ssk, substr(l.licschppkre,11,6)||ar.ssk) " +
                         " GROUP BY g.name";


            string odenilmis_kr_girov = "SELECT g.name," +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit,11,6) IN (substr(l.licschkre,11,6), substr(l.licschpkre,11,6), substr(l.licsch_19,11,6), substr(l.licschppkre,11,6)) THEN substr(ar.kredit,11,6) ELSE NULL END)," +
                         " sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                         " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                         " FROM licschkre l " +
                         " JOIN tipkre g ON l.tipkredita = g.code " +
                         " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
                         " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                         " AND l.subschkre = ar.ssk and" +
                         " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                         " OR " +
                         " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
                         " GROUP BY g.name";


            //string odenilmis_kr_girov = "SELECT g.name,rr.name_regnom,ar.ssk," +
            //             //"  CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END + " +
            //             "CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END AS esas, " +
            //             "  CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END + CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END AS faiz " +
            //             " FROM regnom rr,licschkre l " +
            //             " JOIN tipzal g ON l.tipzaloga = g.code " +
            //             " JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('31-10-2023', 'dd/mm/yyyy') " +
            //             " WHERE substr(l.licschkre, 10, 6)=rr.regnom and (ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
            //             " AND ar.ssk =l.subschkre AND " +
            //             " ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
            //             " OR " +
            //             " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)))";

            // Veritabanından verileri çek
            DataTable odenilmis_kr_NOV = new DataTable();
            DataTable odenilmis_kr_GIROV = new DataTable();
            DataTable odenilmis_kr_GIROV_ad = new DataTable();
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(odenilmis_kr_nov, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_NOV);
                }

                using (OracleCommand command = new OracleCommand(odenilmis_kr_girov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr_GIROV);
                }

                //using (OracleCommand command = new OracleCommand(queryHesaplanmisFaizler, connection))
                //{
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(hesaplanmisFaizlerTable);
                //}
            }

            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            string baseFileName = $"odemeler  "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\Umumi kredit sorgu.xlsx";
            if (File.Exists(Path.Combine(desktopFolder, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(desktopFolder, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Odenilmis kreditler"];

                int ilksetirKreditnov = 3; // Başlangıç satır numarası
                int ilksutunKreditnov = 1; // Başlangıç sütun numarası
                int lastilksetirKreditnov = ilksetirKreditnov - 1;
                int ilksetirKreditnovnov = 3; // Başlangıç  satır numarası
                int ilksutunKreditnovnov = 6; // Başlangıç sütun numarası
                int lastilksetirKreditnovnov = ilksetirKreditnovnov - 1;

                // Her sütunun toplamını hesaplamak için bir dizi kullanın
                double[] columnTotals = new double[odenilmis_kr_NOV.Columns.Count];//NOVLER UZRE
                double[] columnGIROVTotals = new double[odenilmis_kr_GIROV.Columns.Count];

                //*******************************NOVLER UZRE
                for (int i = 0; i < odenilmis_kr_NOV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_NOV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = odenilmis_kr_NOV.Rows[i][j].ToString();
                        ilksutunKreditnovnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnovnov = 6; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnovnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_NOV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = columnTotals[j];
                    ilksutunKreditnovnov++; // Sütun numarasını artır
                }
                for (int j = 6; j <= 9; j++)
                {
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnovnov, j].Style.Font.Size = 14;
                }

                //*******************************GIROVLAR UZRE
                for (int i = 0; i < odenilmis_kr_GIROV.Rows.Count; i++)
                {
                    for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                    {
                        double cellValue;
                        if (double.TryParse(odenilmis_kr_GIROV.Rows[i][j].ToString(), out cellValue))
                        {
                            columnGIROVTotals[j] += cellValue;
                        }

                        worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = odenilmis_kr_GIROV.Rows[i][j].ToString();
                        ilksutunKreditnov++; // Sütun numarasını artır
                    }

                    ilksutunKreditnov = 1; // Sütun numarasını sıfırla, bir sonraki satıra geç
                    ilksetirKreditnov++; // Satır numarasını artır
                }
                for (int j = 0; j < odenilmis_kr_GIROV.Columns.Count; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = columnGIROVTotals[j];
                    ilksutunKreditnov++; // Sütun numarasını artır
                }
                for (int j = 1; j <= 4; j++)
                {
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Bold = true;
                    worksheet.Cells[ilksetirKreditnov, j].Style.Font.Size = 14;
                }

                worksheet.Cells[ilksetirKreditnovnov, 6].Value = "Toplam";
                worksheet.Cells[ilksetirKreditnov, 1].Value = "Toplam";

                string filePath = Path.Combine(desktopFolder, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);

            }
        }
        private void testyoxla()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            //(ar.kredit=l.licschpkre or ar.kredit=l.licschppkre) and

            OracleCommand Orcom = new OracleCommand("select * from odb.arh_licschkre lk where lk.date_oper='04-11-2023' and lk.date_close is null ", Orcon);
            //    "arh_dd ar where date_open> '05-01-2023' and (l.index_otrasli='01902'or index_otrasli='01903') " +
            //    "and substr(l.licschkre,11,6)=substr(ar.kredit,11,6) and substr(ar.debet,0,1)='1' and ar.date_oper=l.date_opeN ", Orcon);

            //OracleCommand Orcom = new OracleCommand(" select r.name_regnom,sum(ar.summa_v_nacval), ar.debet,ar.kredit from licschkre l," +
            //    "arh_dd ar,regnom r where ar.date_oper between to_date('01-10-2023', 'dd/mm/yyyy') and to_date('31-10-2023','dd / mm / yyyy') " +
            //    "and ar.debet=l.licschpkre and substr(ar.kredit,0,5)='64012' and l.tipzaloga='10' and r.regnom=substr(l.licschkre,10,6) and r.svazanniy='1' and l.subschkre=ar.ssd GROUP BY r.name_regnom, ar.debet,ar.kredit", Orcon);

            //OracleCommand verilmis_inex_kr = new OracleCommand("select l.index_otrasli,count(l.summakre),sum(l.summakre) from licschkre l " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') GROUP BY l.index_otrasli ", Orcon);

            //OracleCommand verilmis_girov_kr = new OracleCommand("select g.name,count(l.licschkre),sum(l.summakre) from licschkre l,tipzal g " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') and l.tipzaloga = g.code GROUP BY g.name ", Orcon);

            //OracleCommand verilmis_nov_kr = new OracleCommand("select g.name,count(l.licschkre),sum(l.summakre) from licschkre l,tipkre g " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') and l.tipkredita = g.code  GROUP BY g.name ", Orcon);

            //OracleCommand Orcom = new OracleCommand(" select ar.date_oper,ar.debet,ar.kredit,ar.summa_v_nacval from  arh_dd ar where ar.date_oper between to_date('01-10-2023', 'dd/mm/yyyy') and to_date('31-10-2023','dd / mm / yyyy') " +
            //    " and ar.debet= and ar.kredit='64012000000000600000' ", Orcon);

            //            OracleCommand odenilmis_nov_kr = new OracleCommand(" SELECT g.name," +
            //       "sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) AS esas, "+
            //      " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS vk, " +
            //      " sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipkre g ON l.tipkredita = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssk and substr(ar.debet,0,1) in (3,4) " +
            //" GROUP BY g.name ", Orcon);
            //OracleCommand odenilmis_girov_kr = new OracleCommand(" SELECT g.name," +
            //       "sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
            //      " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS vk, " +
            //      " sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipzal g ON l.tipzaloga = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssk and substr(ar.debet,0,1) in (3,4) " +
            //" GROUP BY g.name ", Orcon);

            //OracleCommand hesablanmis_faiz_girov_kr = new OracleCommand(" SELECT g.name," +
            //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipzal g ON l.tipzaloga = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('30-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
            //" GROUP BY g.name ", Orcon);

            //OracleCommand hesablanmis_faiz_nov_kr = new OracleCommand(" SELECT g.name," +
            //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipkre g ON l.tipkredita = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
            //" GROUP BY g.name ", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            System.Data.DataTable Ordt = new System.Data.DataTable();
            Orda.Fill(Ordt);
            //dataGridView2.DataSource = Ordt;
            Orcon.Close();
        }
        private void yoxlama()
        {
           
            Application.DoEvents();
            //string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";


            #region hesablanmis_faiz_girov
            //string hesablanmis_faiz_girov = "SELECT ar.date_oper tarix,ar.ssd subd,ar.debet debet,ar.ssk subk,ar.kredit kredit," +
            //    "ar.summa_v_nacval mebleg,r.regnom qeyd_no,r.svazanniy elaqe,l.licschkre kat,g.name novu,t.name tipi,substr(ar.debet,10,6)||ar.ssd q_no" +
            //" FROM " +
            //"    licschkre l, arh_dd ar, regnom r,tipzal g,tipkre t " +
            //" WHERE " +
            //"    ar.date_oper BETWEEN TO_DATE('" + txt_giristar.Text + "', 'dd/mm/yyyy') AND TO_DATE('" + txt_sontar.Text + "', 'dd/mm/yyyy') " +
            //"  AND SUBSTR(ar.kredit, 1, 1) = 6 and l.tipzaloga = g.code and l.tipkredita = t.code" +
            //"  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
            //"   AND l.subschkre = ar.ssd " +
            //"   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
            //"   AND r.regnom = SUBSTR(l.licschpkre, 10, 6)";
            #endregion
            string odenilmis_kreditler = "SELECT ar.date_oper tarix,ar.ssd subd,ar.debet debet,ar.ssk subk,ar.kredit kredit," +
            " CASE WHEN ar.kredit in( l.licschkre,l.licsch_19) THEN 'esas' WHEN ar.kredit IN (l.licschpkre, l.licschppkre) THEN 'faiz' END as hesabi," +
            " ar.summa_v_nacval mebleg,r.regnom qeyd_no,r.svazanniy elaqe,l.licschkre kat,g.name novu,t.name tipi,substr(ar.kredit,10,6)||ar.ssk q_no" +
            " FROM " +
            "    licschkre l, arh_dd ar, regnom r,tipzal g,tipkre t " +
            " WHERE " +
            "    ar.date_oper BETWEEN TO_DATE('"+txt_giristar.Text+"', 'dd/mm/yyyy') AND TO_DATE('"+txt_sontar.Text+"', 'dd/mm/yyyy') " +
            " and (substr(ar.debet, 0, 1) in (3, 4) OR substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259))" +
            " and l.tipzaloga = g.code and l.tipkredita = t.code" +
            "  and ar.kredit IN(l.licschkre, l.licschpkre,l.licsch_19, l.licschppkre) " +
            "   and l.subschkre = ar.ssk " +
            "   and r.regnom = SUBSTR(ar.kredit, 10, 6) " +
            "   and r.regnom = SUBSTR(l.licschpkre, 10, 6)";

            DataTable hesablanmis_faizler = new DataTable();
            DataTable odenilmis_kr = new DataTable();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                //using (OracleCommand command = new OracleCommand(hesablanmis_faiz_girov, connection))
                //{
                //    connection.Open();
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(hesablanmis_faizler);
                //}
                using (OracleCommand command = new OracleCommand(odenilmis_kreditler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(odenilmis_kr);
                }
            }




            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            cl.baseFileName = $"odemeler  "; // Temel dosya adı
            cl.fileName = cl.baseFileName + ".xlsx";
            cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Exceller", "Umumi kredit sorgu.xlsx");
            cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
            decimal topFsutun=0;
            decimal topFsutuntip = 0;
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
                ExcelWorksheet worksheet_hesfaiz = package.Workbook.Worksheets["Hesablanmis faizler"];
                ExcelWorksheet worksheet_odenilmis = package.Workbook.Worksheets["Odenilmis kreditler"];
                var kredit_novu = hesablanmis_faizler.AsEnumerable()
                    .Select(row => row.Field<string>("novu"))
                    .Distinct();

                var kredit_tipi = hesablanmis_faizler.AsEnumerable()
                    .Select(row => row.Field<string>("tipi"))
                    .Distinct();

                var kredit_novu_od = odenilmis_kr.AsEnumerable()
                    .Select(row => row.Field<string>("novu"))
                    .Distinct();

                var kredit_tipi_od = odenilmis_kr.AsEnumerable()
                    .Select(row => row.Field<string>("tipi"))
                    .Distinct();


                #region odenilmis_yeni


                // Başlangıç sətiri
                int ilksetirTipi = 3; // Tipi üçün başlayacağımız sətir
                int ilksetirKreditnovnovu = 3; // Novu üçün başlayacağımız sətir

                // Tipi üçün sütunlar
                int sutunA = 1; // A sütunu (Tipi)
                int sutunB = 2; // B sütunu (Q_No say)
                int sutunC = 3; // C sütunu (Esas)
                int sutunD = 4; // D sütunu (Faiz)

                // Novu üçün sütunlar
                int sutunF = 6; // F sütunu (Novu)
                int sutunG = 7; // G sütunu (Q_No say)
                int sutunH = 8; // H sütunu (Esas)
                int sutunI = 9; // I sütunu (Faiz)

                // Cəmlər üçün dəyişənlər (Tipi)
                int toplamSayTipi = 0;
                decimal toplamEsasTipi = 0;
                decimal toplamFaizTipi = 0;

                // Tipi üzrə məlumatları işləyirik
                var kreditTipleri = odenilmis_kr.AsEnumerable()
                    .Select(row => row.Field<string>("TIPI"))
                    .Distinct();

                foreach (var tipi in kreditTipleri)
                {
                    // Q_No sayını hesablamaq (unikal)
                    int sayTipi = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("TIPI") == tipi)
                        .Select(row => row.Field<string>("Q_NO"))
                        .Distinct()
                        .Count();

                    // Esas məbləğini hesablamaq
                    decimal toplamEsas = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("TIPI") == tipi && row.Field<string>("HESABI") == "esas")
                        .Sum(row => row.Field<decimal>("MEBLEG"));

                    // Faiz məbləğini hesablamaq
                    decimal toplamFaiz = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("TIPI") == tipi && row.Field<string>("HESABI") == "faiz")
                        .Sum(row => row.Field<decimal>("MEBLEG"));

                    // Məlumatları Excel-ə yaz
                    worksheet_odenilmis.Cells[ilksetirTipi, sutunA].Value = tipi;
                    worksheet_odenilmis.Cells[ilksetirTipi, sutunB].Value = sayTipi;
                    worksheet_odenilmis.Cells[ilksetirTipi, sutunC].Value = toplamEsas;
                    worksheet_odenilmis.Cells[ilksetirTipi, sutunD].Value = toplamFaiz;

                    // Toplamları artır
                    toplamSayTipi += sayTipi;
                    toplamEsasTipi += toplamEsas;
                    toplamFaizTipi += toplamFaiz;

                    // Növbəti sətirə keç
                    ilksetirTipi++;
                }

                // Tipi üçün toplamları yaz
                worksheet_odenilmis.Cells[ilksetirTipi, sutunA].Value = "CƏM";
                worksheet_odenilmis.Cells[ilksetirTipi, sutunB].Value = toplamSayTipi;
                worksheet_odenilmis.Cells[ilksetirTipi, sutunC].Value = toplamEsasTipi;
                worksheet_odenilmis.Cells[ilksetirTipi, sutunD].Value = toplamFaizTipi;
                worksheet_odenilmis.Cells[ilksetirTipi, sutunA, ilksetirTipi, sutunD].Style.Font.Bold = true;

                // Cəmlər üçün dəyişənlər (Novu)
                int toplamSayNovu = 0;
                decimal toplamEsasNovu = 0;
                decimal toplamFaizNovu = 0;

                // Novu üzrə məlumatları işləyirik
                var kreditNovleri = odenilmis_kr.AsEnumerable()
                    .Select(row => row.Field<string>("NOVU"))
                    .Distinct();

                foreach (var novu in kreditNovleri)
                {
                    // Q_No sayını hesablamaq (unikal)
                    int sayNovu = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("NOVU") == novu)
                        .Select(row => row.Field<string>("Q_NO"))
                        .Distinct()
                        .Count();

                    // Esas məbləğini hesablamaq
                    decimal toplamEsas = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("NOVU") == novu && row.Field<string>("HESABI") == "esas")
                        .Sum(row => row.Field<decimal>("MEBLEG"));

                    // Faiz məbləğini hesablamaq
                    decimal toplamFaiz = odenilmis_kr.AsEnumerable()
                        .Where(row => row.Field<string>("NOVU") == novu && row.Field<string>("HESABI") == "faiz")
                        .Sum(row => row.Field<decimal>("MEBLEG"));

                    // Məlumatları Excel-ə yaz
                    worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunF].Value = novu;
                    worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunG].Value = sayNovu;
                    worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunH].Value = toplamEsas;
                    worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunI].Value = toplamFaiz;

                    // Toplamları artır
                    toplamSayNovu += sayNovu;
                    toplamEsasNovu += toplamEsas;
                    toplamFaizNovu += toplamFaiz;

                    // Növbəti sətirə keç
                    ilksetirKreditnovnovu++;
                }

                // Novu üçün toplamları yaz
                worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunF].Value = "CƏM";
                worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunG].Value = toplamSayNovu;
                worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunH].Value = toplamEsasNovu;
                worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunI].Value = toplamFaizNovu;
                worksheet_odenilmis.Cells[ilksetirKreditnovnovu, sutunF, ilksetirKreditnovnovu, sutunI].Style.Font.Bold = true;
                #endregion
                #region hesablanmisFaiz_yeni
                //// Başlangıç sətiri və sütunlar
                //int ilksetirKreditnovnov = 3; // Kredit növləri üçün başlayan sətir
                //int ilksetirKreditnov = 3;    // Kredit tipləri üçün başlayan sətir
                //int ilksutunKreditnovnov = 1; // Kredit növləri üçün başlayan sütun
                //int ilksutunKreditnov = 5;    // Kredit tipləri üçün başlayan sütun

                //// Cəmlər üçün dəyişənlər
                //int topFsutunNov = 0;
                //decimal totalMeblegNov = 0;

                //int topFsutunTip = 0;
                //decimal totalMeblegTip = 0;

                //// *********** Kreditin növləri üzrə *************
                //foreach (var name in kredit_novu)
                //{
                //    // Kredit növü üzrə məbləği hesablamaq
                //    decimal totalForName = hesablanmis_faizler.AsEnumerable()
                //        .Where(row => row.Field<string>("novu") == name )
                //        .Sum(row => row.Field<decimal>("mebleg"));

                //    // Kredit növü üzrə qeyd_no-nu hesablamaq (təkrarlar daxil edilmədən)
                //    int totalSayi = hesablanmis_faizler.AsEnumerable()
                //        .Where(row => row.Field<string>("novu") == name )
                //        .Select(row => row.Field<string>("q_no"))
                //        .Distinct()
                //        .Count();

                //    // Məlumatları Excel-ə yazdır
                //    worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = name;
                //    worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov + 1].Value = totalSayi;
                //    worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov + 2].Value = totalForName;

                //    // Cəmləri artır
                //    topFsutunNov += totalSayi;
                //    totalMeblegNov += totalForName;

                //    // Növbəti sətirə keç
                //    ilksetirKreditnovnov++;
                //}

                //// Kredit növləri üçün cəmləri yazdır
                //worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov].Value = "CƏM";
                //worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov + 1].Value = topFsutunNov;
                //worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov + 2].Value = totalMeblegNov;
                //worksheet_hesfaiz.Cells[ilksetirKreditnovnov, ilksutunKreditnovnov, ilksetirKreditnovnov, ilksutunKreditnovnov + 2].Style.Font.Bold = true;

                //// *********** Kreditin tipləri üzrə *************
                //foreach (var name in kredit_tipi)
                //{
                //    // Kredit tipi üzrə məbləği hesablamaq
                //    decimal totalForName = hesablanmis_faizler.AsEnumerable()
                //        .Where(row => row.Field<string>("tipi") == name)
                //        .Sum(row => row.Field<decimal>("mebleg"));

                //    // Kredit tipi üzrə qeyd_no-nu hesablamaq (təkrarlar daxil edilmədən)
                //    int totalSayi = hesablanmis_faizler.AsEnumerable()
                //        .Where(row => row.Field<string>("tipi") == name)
                //        .Select(row => row.Field<string>("q_no"))
                //        .Distinct()
                //        .Count();

                //    // Məlumatları Excel-ə yazdır
                //    worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = name;
                //    worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov + 1].Value = totalSayi;
                //    worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov + 2].Value = totalForName;

                //    // Cəmləri artır
                //    topFsutunTip += totalSayi;
                //    totalMeblegTip += totalForName;

                //    // Növbəti sətirə keç
                //    ilksetirKreditnov++;
                //}

                //// Kredit tipləri üçün cəmləri yazdır
                //worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov].Value = "CƏM";
                //worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov + 1].Value = topFsutunTip;
                //worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov + 2].Value = totalMeblegTip;
                //worksheet_hesfaiz.Cells[ilksetirKreditnov, ilksutunKreditnov, ilksetirKreditnov, ilksutunKreditnov + 2].Style.Font.Bold = true;

                #endregion
                label1.Text = "Hazırdır.";
                Application.DoEvents();
                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(cl.filePath);
                Application.DoEvents();
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_giristar.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_giristar.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_sontar.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_sontar.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox3.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_sontar.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //exceleat2test();
            yoxlama();
        }
    }
}
