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
using Excel = Microsoft.Office.Interop.Excel;

namespace BMI.Muhasibat
{
    public partial class frm_valyuta_alis_satis : Form
    {
        public frm_valyuta_alis_satis()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;

        string[] val_hes_60 = { "10060" };
        string[] val_hes_50 = { "10050" };
        string[] hes_1 = { "10020" };
        string[] hes_2 = { "41" };
        string[] hes_3 = { "41040", "41045", "41050", "41055" };
        string[] hes_4 = { "45023010010000400000" , "45023020020000400000" };
        string[] hes_6 = { "66220" };
        string[] hes_7 = { "35020", "35025", "35026", "35940" };
        string[] hes_8 = { "86220" };

        int setirsay = 0;

        string dosyayolu  ;
        string textBoxText  ; // TextBox'tan alınan metni sakla
        string yeniMetin   ;
        string baseFileName  ; // Temel dosya adı
        string fileName  ;
        string templateFilePath  ;
        string filePath  ;


        private void makro()
        {
            
            //string excelyolu = filePath; // Excel dosyasının yolu
            string macroName = "btnCalculate_Click"; // Çalıştırmak istediğiniz makro adı
            string macrosil = "btnDeleteSelectedRow_Click"; // Çalıştırmak istediğiniz makro adı
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            // Excel dosyasını aç
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet worksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["Forma-NXVS-1"];
            Microsoft.Office.Interop.Excel.Worksheet worksheet2 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["Forma-NXVS-1a"];
            Microsoft.Office.Interop.Excel.Worksheet worksheet3 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["Forma-NXVS-2"];
            excelApp.Run(macroName);
            Microsoft.Office.Interop.Excel.Range cellRange = null;

            int baslangicSatir = 20+setirsay;
            int bitisSatir = 65;

            string baslangicHucresi = "B" + baslangicSatir.ToString();
            string bitisHucresi = "B" + bitisSatir.ToString();

            string aralik = baslangicHucresi + ":" + bitisHucresi;
            cellRange = worksheet1.Range[aralik];
            excelApp.AlertBeforeOverwriting = false; // Hücre içeriği değiştirilirken uyarı almayı devre dışı bırakır
            excelApp.DisplayAlerts = false; // Uyarıları gösterme
            excelApp.Run(macrosil); // "btnDeleteSelectedRow_Click" makrosunu çalıştır
            excelApp.DisplayAlerts = true;
            excelApp.AlertBeforeOverwriting = true;
            excelApp.Visible = true;
            
        }

        private void Excel_val_alis_satis()
        {
            
            //duz olan
            #region sql_kodlar

           
            string giris_qal = "select sum(l.saldo_vhd_nacval)/1000 giris_qal from odb.arh_saldo_ls l " +
                "where l.date_oper=to_date('" + txtgiris.Text + "', 'DD-MM-YYYY') and substr( l.licsch,1,5) in ('10020','10080')";
            string arx_balans = "select d.date_oper tar,substr(d.debet,1,5) db_ilk_5,d.debet db,substr(d.kredit,1,5) kr_ilk_5, "+
     "d.kredit kr, d.summa_v_inval*d.kurs_valuti mebavl, d.summa_v_nacval mebaz from arh_dd d where " +
     "d.date_oper BETWEEN to_date('" + txtgiris.Text + "', 'DD-MM-YYYY') AND to_date('" + txtcixis.Text + "', 'DD-MM-YYYY')";
            string valyuta_alis_satis = "SELECT COALESCE(alis.tarix, satis.tarix) AS tarix, " +
       "COALESCE(alis.val, satis.val) AS val, " +
       "COALESCE(alis.val_meb_alis, 0) AS val_meb_alis, " +
       "COALESCE(alis.azn_meb, 0) AS azn_meb_alis, " +
       "COALESCE(alis.say, 0) AS say_Alis, " +
       "COALESCE(satis.val_meb_satis, 0) AS val_meb_satis, " +
       "COALESCE(satis.azn_meb, 0) AS azn_meb_satis, " +
       "COALESCE(satis.say, 0) AS say_Satis " +
       "FROM " +
       "(SELECT d.date_oper AS tarix, " +
       "case    when substr(d.debet, 6, 2)='00' then 'AZN' "+
            "when substr(d.debet, 6, 2)= '01' then 'USD' " +
            "when substr(d.debet, 6, 2)= '02' then 'EUR' " +
            "when substr(d.debet, 6, 2)= '03' then 'RUB' " +
            "when substr(d.debet, 6, 2)= '04' then '100_İRR' " +
            "when substr(d.debet, 6, 2)= '05' then 'AED' end " +
            " AS val, " +
       "sum(d.summa_v_inval) AS val_meb_alis, " +
       "sum(d.summa_v_inval * d.kurs_valuti) AS azn_meb, " +
       "count(d.debet) AS say, " +
       "COUNT(DISTINCT CASE WHEN substr(d.debet, 1, 5) IN('10060', '10050') THEN d.debet END) AS tekcut " +
       "FROM arh_dd d " +
       "WHERE d.date_oper BETWEEN to_date('" + txtgiris.Text + "', 'DD-MM-YYYY') AND to_date('" + txtcixis.Text + "', 'DD-MM-YYYY') " +
       "AND substr(d.debet, 1, 5) = '10060' " +
       "AND substr(d.kredit, 1, 5) = '10050' " +
       "GROUP BY d.date_oper, substr(d.debet, 6, 2)) alis " +
       "FULL OUTER JOIN " +
       "(SELECT d.date_oper AS tarix, " +
       "case    when substr(d.kredit, 6, 2)='00' then 'AZN' " +
            "when substr(d.kredit, 6, 2)= '01' then 'USD' " +
            "when substr(d.kredit, 6, 2)= '02' then 'EUR' " +
            "when substr(d.kredit, 6, 2)= '03' then 'RUB' " +
            "when substr(d.kredit, 6, 2)= '04' then '100_İRR' " +
            "when substr(d.kredit, 6, 2)= '05' then 'AED' end " +
            " AS val, " +
       "sum(d.summa_v_inval) AS val_meb_satis, " +
       "sum(d.summa_v_inval * d.kurs_valuti) AS azn_meb, " +
       "count(d.debet) AS say " +
       "FROM arh_dd d " +
       "WHERE d.date_oper BETWEEN to_date('" + txtgiris.Text + "', 'DD-MM-YYYY') AND to_date('" + txtcixis.Text + "', 'DD-MM-YYYY') " +
       "AND substr(d.debet,1,5)= '10050' " +
       "AND substr(d.kredit,1,5)= '10060' " +
       "GROUP BY d.date_oper, substr(d.kredit, 6, 2)) satis " +
       "ON alis.tarix = satis.tarix AND alis.val = satis.val " +
       "ORDER BY COALESCE(alis.tarix, satis.tarix), COALESCE(alis.val, satis.val)";

            #endregion

            DataTable _dt_giris_qal = new DataTable();
            DataTable _dt_val_alis_satis = new DataTable();
            DataTable _dt_arx_balans = new DataTable();

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                using (OracleCommand command = new OracleCommand(giris_qal, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_giris_qal);
                }

                using (OracleCommand command = new OracleCommand(valyuta_alis_satis, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_val_alis_satis);
                }
                using (OracleCommand command = new OracleCommand(arx_balans, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_arx_balans);
                }
                connection.Close();
            }
             dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            textBoxText = txtgiris.Text; // TextBox'tan alınan metni sakla
             yeniMetin = textBoxText.Replace("-", ""); ;
             baseFileName = "NXVS_124_"+yeniMetin.Substring(2,6); // Temel dosya adı
             fileName = baseFileName + ".xlsm";
             templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "Valyuta hesabatai.xlsm");
             filePath = Path.Combine(dosyayolu, fileName);

            if (File.Exists(Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Forma-NXVS-1"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["Forma-NXVS-1a"];
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets["Forma-NXVS-2"];
                
                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    for (int j = 0; j < 3; j++) // Sadece 0, 1 ve 2. sütunları al
                    {
                        worksheet1.Cells[i + 20, j + 2].Value = _dt_val_alis_satis.Rows[i][j];
                    }
                }

                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                        worksheet1.Cells[i + 20, 7].Value = _dt_val_alis_satis.Rows[i][4];
                }
                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    worksheet1.Cells[i + 20, 6].Value = _dt_val_alis_satis.Rows[i][3];
                }

                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    for (int j = 0; j < 2; j++) // Sadece 0, 1 ve 2. sütunları al
                    {
                        worksheet1.Cells[i + 20, j + 8].Value = _dt_val_alis_satis.Rows[i][j];
                    }
                }
                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    worksheet1.Cells[i + 20, 10].Value = _dt_val_alis_satis.Rows[i][5];
                }

                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    worksheet1.Cells[i + 20, 13].Value = _dt_val_alis_satis.Rows[i][7];
                }
                for (int i = 0; i < _dt_val_alis_satis.Rows.Count; i++)
                {
                    worksheet1.Cells[i + 20, 12].Value = _dt_val_alis_satis.Rows[i][6];
                }

                var USD_top = _dt_val_alis_satis.AsEnumerable()
                .Where(row => row.Field<string>(1) == "USD")
                .ToList();
                decimal total_USD_alis = USD_top.Sum(row => row.Field<decimal>(2));
                decimal total_USD_alis_AZN_ile = USD_top.Sum(row => row.Field<decimal>(3));
                decimal total_USD_alis_say = USD_top.Sum(row => row.Field<decimal>(4));

                decimal total_USD_satis = USD_top.Sum(row => row.Field<decimal>(5));
                decimal total_USD_satis_AZN_ile = USD_top.Sum(row => row.Field<decimal>(6));
                decimal total_USD_satis_say = USD_top.Sum(row => row.Field<decimal>(7));

                worksheet2.Cells["F24"].Value = total_USD_alis;
                worksheet2.Cells["H24"].Value = total_USD_alis_AZN_ile;
                worksheet2.Cells["I24"].Value = total_USD_alis_say;

                worksheet2.Cells["K24"].Value = total_USD_satis;
                worksheet2.Cells["M24"].Value = total_USD_satis_AZN_ile;
                worksheet2.Cells["N24"].Value = total_USD_satis_say;

                var EUR_top = _dt_val_alis_satis.AsEnumerable()
                .Where(row => row.Field<string>(1) == "EUR")
                .ToList();
                decimal total_EUR_alis = EUR_top.Sum(row => row.Field<decimal>(2));
                decimal total_EUR_alis_AZN_ile = EUR_top.Sum(row => row.Field<decimal>(3));
                decimal total_EUR_alis_say = EUR_top.Sum(row => row.Field<decimal>(4));

                decimal total_EUR_satis = EUR_top.Sum(row => row.Field<decimal>(5));
                decimal total_EUR_satis_AZN_ile = EUR_top.Sum(row => row.Field<decimal>(6));
                decimal total_EUR_satis_say = EUR_top.Sum(row => row.Field<decimal>(7));

                worksheet2.Cells["P24"].Value = total_EUR_alis;
                worksheet2.Cells["R24"].Value = total_EUR_alis_AZN_ile;
                worksheet2.Cells["S24"].Value = total_EUR_alis_say;

                worksheet2.Cells["U24"].Value = total_EUR_satis;
                worksheet2.Cells["W24"].Value = total_EUR_satis_AZN_ile;
                worksheet2.Cells["X24"].Value = total_EUR_satis_say;

                var NXVS2_e17 = _dt_giris_qal.AsEnumerable().ToList();

                decimal total_NXVS2_e17 = NXVS2_e17.Sum(row => row.Field<decimal>(0));
                worksheet3.Cells["E17"].Value = total_NXVS2_e17;

                var NXVS2_e21 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10060" && row.Field<string>(3) == "10050")
                .ToList();
                decimal total_NXVS2_e21 = NXVS2_e21.Sum(row => row.Field<decimal>(5));

                var NXVS2_e21_mq = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10060" && row.Field<string>(3) == "10050")
                .ToList();
                decimal total_NXVS2_e21_mq = NXVS2_e21_mq.Sum(row => row.Field<decimal>(6));
                decimal ferqi_cix = (total_NXVS2_e21_mq-total_NXVS2_e21 );

                worksheet3.Cells["E21"].Value = total_NXVS2_e21/1000;

                var NXVS2_e22 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && row.Field<string>(3).Substring(0,2) == "41")
                .ToList();
                decimal total_NXVS2_e22 = NXVS2_e22.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E22"].Value = total_NXVS2_e22 / 1000;

                var NXVS2_e23 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && row.Field<string>(3).Substring(0, 2) == "41" && row.Field<string>(3).Substring(4, 1) == "5")
                .ToList();
                decimal total_NXVS2_e23 = NXVS2_e23.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E23"].Value = total_NXVS2_e23 / 1000;

                var NXVS2_e24 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && hes_3.Contains(row.Field<string>(3)))
                .ToList();
                decimal total_NXVS2_e24 = NXVS2_e24.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E24"].Value = total_NXVS2_e24 / 1000;

                var NXVS2_e25 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && (row.Field<string>(4) == "45023010010000400000" || row.Field<string>(4) == "45023020020000400000"))
                .ToList();
                decimal total_NXVS2_e25 = NXVS2_e25.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E25"].Value = total_NXVS2_e25 / 1000;

                var NXVS2_e26 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && !hes_7.Contains(row.Field<string>(3)) && 
                (row.Field<string>(3).Substring(0,2) == "40" || row.Field<string>(3).Substring(0,1) == "3"))
                .ToList();
                decimal total_NXVS2_e26 = NXVS2_e26.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E26"].Value = total_NXVS2_e26 / 1000;

                var NXVS2_e29 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && row.Field<string>(3) == "66220")
                .ToList();
                var NXVS2_e29_1 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && row.Field<string>(4).Substring(0,2) == "45" && !hes_4.Contains(row.Field<string>(4)))
                .ToList();
                var NXVS2_e29_2 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10020" && row.Field<string>(3).Substring(0, 2) == "25")
                .ToList();

                decimal total_NXVS2_e29_1 = NXVS2_e29_1.Sum(row => row.Field<decimal>(6));
                decimal total_NXVS2_e29 = NXVS2_e29.Sum(row => row.Field<decimal>(6));
                decimal total_NXVS2_e29_2 = NXVS2_e29_2.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E29"].Value = (total_NXVS2_e29_1+ total_NXVS2_e29+ ferqi_cix+ total_NXVS2_e29_2) / 1000;

                var NXVS2_e33 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10050" && row.Field<string>(3) == "10060")
                .ToList();
                decimal total_NXVS2_e33 = NXVS2_e33.Sum(row => row.Field<decimal>(5));

                var NXVS2_e33_mq = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(1) == "10050" && row.Field<string>(3) == "10060")
                .ToList();
                decimal total_NXVS2_e33_mq = NXVS2_e33_mq.Sum(row => row.Field<decimal>(6));
                decimal ferqi_cix_ = total_NXVS2_e33_mq-total_NXVS2_e33  ;

                worksheet3.Cells["E33"].Value = total_NXVS2_e33 / 1000;


                var NXVS2_e34 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && row.Field<string>(2).Substring(0, 2) == "41")
                .ToList();
                decimal total_NXVS2_e34 = NXVS2_e34.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E34"].Value = total_NXVS2_e34 / 1000;

                var NXVS2_e35 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && row.Field<string>(1).Substring(0, 2) == "41" && row.Field<string>(1).Substring(4, 1) == "5")
                .ToList();
                decimal total_NXVS2_e35 = NXVS2_e35.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E35"].Value = total_NXVS2_e35 / 1000;

                var NXVS2_e36 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && hes_3.Contains(row.Field<string>(1)))
                .ToList();
                decimal total_NXVS2_e36 = NXVS2_e36.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E36"].Value = total_NXVS2_e36 / 1000;

                var NXVS2_e37 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && (row.Field<string>(2) == "45023010010000400000" || row.Field<string>(2) == "45023020020000400000"))
                .ToList();
                decimal total_NXVS2_e37 = NXVS2_e37.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E37"].Value = total_NXVS2_e37 / 1000;

                var NXVS2_e38 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && !hes_7.Contains(row.Field<string>(1)) &&
                (row.Field<string>(1).Substring(0, 2) == "40" || row.Field<string>(1).Substring(0, 1) == "3"))
                .ToList();
                decimal total_NXVS2_e38 = NXVS2_e38.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E38"].Value = total_NXVS2_e38 / 1000;

                var NXVS2_e41 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && row.Field<string>(1) == "86220")
                .ToList();
                var NXVS2_e41_1 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && row.Field<string>(2).Substring(0, 2) == "45" && !hes_4.Contains(row.Field<string>(2)))
                .ToList();
                var NXVS2_e41_2 = _dt_arx_balans.AsEnumerable()
                .Where(row => row.Field<string>(3) == "10020" && row.Field<string>(1).Substring(0, 2) == "25")
                .ToList();

                decimal total_NXVS2_e41_1 = NXVS2_e41_1.Sum(row => row.Field<decimal>(6));
                decimal total_NXVS2_e41 = NXVS2_e41.Sum(row => row.Field<decimal>(6));
                decimal total_NXVS2_e41_2 = NXVS2_e41_2.Sum(row => row.Field<decimal>(6));
                worksheet3.Cells["E41"].Value = (total_NXVS2_e41_1 + total_NXVS2_e41+ ferqi_cix_+ total_NXVS2_e41_2) / 1000;

                setirsay = _dt_val_alis_satis.Rows.Count;
                #endregion
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private void txtdtbugun_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txtgiris.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txtgiris.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txtdtbugun_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txtcixis.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txtdtdunen_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txtcixis.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txtcixis.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txtdtdunen_KeyDown(object sender, KeyEventArgs e)
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
            Excel_val_alis_satis();
            
        }
    }
}
