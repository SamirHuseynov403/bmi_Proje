using OfficeOpenXml;
using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BMI.Muhasibat
{
    public partial class frm_Dep : Form
    {
        public frm_Dep()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl_yanasma = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;
        private void Excel()
        {
            //duz olan
            #region sqlkodlar

            
            string huquqi_top10 = "select 'huquqi' AS nov, "+
         "b.registrac_nomer qeyd, r.name_regnom,  " +
         "      sum(-round(l.saldo_ish_nacval / 1000, 2))top, " +
         "      SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '00' THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS azn, " +
         "      SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '01' THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS usd, " +
         "      SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) NOT IN('00', '01') THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS diger " +
         "from odb.arh_saldo_ls l,regnom r, licsch b " +
         " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') " +
         "and b.licsch = l.licsch " +
         "and(substr(l.licsch, 1, 2) = '40' or substr(l.licsch, 1, 1) = '3') " +
         "and substr(l.licsch,1,5) not in (35020, 35025, 35026, 35940) and substr(l.licsch,10,6)<> '000004' " +
         "and r.regnom = b.registrac_nomer " +
         "group by b.registrac_nomer,r.name_regnom, 'huquqi' " +
        " order by sum(l.saldo_ish_nacval / 1000) " +
        " FETCH FIRST 10 ROWS ONLY";

            string sefirlik_qal = "select 'sefirlik' nov,b.registrac_nomer qeyd, " +
     " case when substr(b.licsch,6,2)= '00' then 'azn' " +
     " when substr(b.licsch,6,2)= '01' then 'usd' " +
     " when substr(b.licsch,6,2)not in ('00', '01') then 'diger' end val " +
     " ,r.name_regnom ad, round(l.saldo_ish_nacval / 1000, 2) meb from odb.arh_saldo_ls l, regnom r,licsch b " +
     " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and b.registrac_nomer = '008002' " +
     " and b.licsch = l.licsch and r.regnom = b.registrac_nomer";

            string fiziki_top10 = "SELECT 'fiziki' AS nov, "+
       "SUBSTR(l.licsch, 10, 6) AS hes,r.name_regnom, sum(-round(l.saldo_ish_nacval / 1000, 2))top, " +
      "SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '00' THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS azn, " +
      " SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '01' THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS usd, " +
      " SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) NOT IN('00', '01') THEN ROUND(-l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS diger " +
"FROM odb.arh_saldo_ls l, regnom r " +
"WHERE l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') " +
"  AND SUBSTR(l.licsch, 1, 2) = '41' " +
"  AND r.regnom = SUBSTR(l.licsch, 10, 6) " +
" AND SUBSTR(l.licsch, 10, 6) <> '000004' " +
"  AND l.saldo_ish_nacval <> 0 " +
"GROUP BY SUBSTR(l.licsch, 10, 6), 'fiziki',r.name_regnom " +
"order by sum(l.saldo_ish_nacval / 1000) " +
"FETCH FIRST 10 ROWS ONLY";

            string gunun_qaliqlari = "select 'fiziki_qal' qaliq, "+
        " case when substr(l.licsch,6,02)= '00' then 'manat' " +
        " when substr(l.licsch,6,02)= '01' then 'usd' " +
        " when substr(l.licsch,6,02)= '01' then 'usd' else 'diger' end valyuta, count(l.saldo_vhd_nacval) say,  " +
        " -round(sum(l.saldo_vhd_nacval) / 1000, 4)giris,-round(sum(l.saldo_ish_nacval) / 1000, 4) cixis " +
        " from odb.arh_saldo_ls l " +
        " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and(l.saldo_ish_nacval < 0 or l.saldo_vhd_nacval < 0) " +
        " and substr(l.licsch,1,2)= '41' " +
        " group by substr(l.licsch, 6, 02) " +
        " union " +
        " select 'huquqi_qal'qaliq,  " +
        " case when substr(l.licsch,6,02)= '00' then 'manat' " +
        " when substr(l.licsch,6,02)= '01' then 'usd' " +
        " when substr(l.licsch,6,02)= '01' then 'usd' else 'diger' end valyuta, count(l.saldo_vhd_nacval) say,  " +
        " -round(sum(l.saldo_vhd_nacval) / 1000, 4)giris,-round(sum(l.saldo_ish_nacval) / 1000, 4) cixis from odb.arh_saldo_ls l " +
        " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and(l.saldo_ish_nacval < 0 or l.saldo_vhd_nacval < 0) " +
        " and(substr(l.licsch, 1, 2) = '40' or substr(l.licsch, 1, 1) = '3') and substr(l.licsch,1,5) not in (35020, 35025, 35026, 35940) " +
        " group by substr(l.licsch, 6, 02)";

            string medaxiller = "select 'fiziki_med', "+
           " case when substr(d.kredit,6,02)= '00' then 'manat' " +
           " when substr(d.kredit,6,2)= '01' then 'usd' " +
           " when substr(d.kredit,6,02)= '01' then 'usd' else 'diger' end valyuta, count('huquqi') say, " +
           " round(sum(d.summa_v_nacval) / 1000, 2) meb " +
           " from arh_dd d " +
           " where d.date_oper = to_date('"+textBox2.Text+"', 'dd/mm/yyyy') and " +
           " substr(d.kredit, 1, 2) = '41' and substr(d.debet,1,5) not in ('66220', '86220') " +
           " group by substr(d.kredit, 6, 2) " +
           " union " +
           " select 'huquqi_med', " +
           " case when substr(d.kredit,6,02)= '00' then 'manat' " +
           " when substr(d.kredit,6,2)= '01' then 'usd' " +
           " when substr(d.kredit,6,02)= '01' then 'usd' else 'diger' end valyuta, count('huquqi') say, " +
           " round(sum(d.summa_v_nacval) / 1000, 2) meb " +
           " from arh_dd d " +
           " where d.date_oper = to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and " +
           " (substr(d.kredit, 1, 2) = '40' or substr(d.kredit, 1, 1) = '3') and substr(d.kredit,1,5) not in (35020, 35025, 35026, 35940) and substr(d.debet,1,5) not in ('66220', '86220') " +
           " group by substr(d.kredit, 6, 2) " +
           " union " +
           " select 'fiziki_mex', " +
           " case when substr(d.debet,6,02)= '00' then 'manat' " +
           " when substr(d.debet,6,2)= '01' then 'usd' " +
           " when substr(d.debet,6,02)= '01' then 'usd' else 'diger' end valyuta, count('huquqi') say, " +
           " round(sum(d.summa_v_nacval) / 1000, 2) meb " +
           " from arh_dd d " +
           " where d.date_oper = to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and " +
           " substr(d.debet, 1, 2) = '41' and substr(d.kredit,1,5) not in ('66220', '86220') " +
           " group by substr(d.debet, 6, 2) " +
           " union " +
           " select 'huquqi_mex', " +
           " case when substr(d.debet,6,02)= '00' then 'manat' " +
           " when substr(d.debet,6,2)= '01' then 'usd' " +
           " when substr(d.debet,6,02)= '01' then 'usd' else 'diger' end valyuta, count('huquqi') say, " +
           " round(sum(d.summa_v_nacval) / 1000, 2) meb " +
           " from arh_dd d " +
           " where d.date_oper = to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and " +
           " (substr(d.debet, 1, 2) = '40' or substr(d.debet, 1, 1) = '3') and substr(d.debet,1,5) not in (35020, 35025, 35026, 35940) and substr(d.kredit,1,5) not in ('66220', '86220') " +
           " group by substr(d.debet, 6, 2)";

            string umumi_qaliqlar = "select 'fiziki' nov,SUBSTR(l.licsch,10,6),R.NAME_REGNOM ad,-round(sum(l.saldo_ish_nacval/1000),2) AZN " +
         "from odb.arh_saldo_ls l,regnom r "+
         "where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and " +
         "substr(l.licsch, 1, 2) = '41' " +
         "and r.regnom = substr(l.licsch, 10, 6) and substr(l.licsch,10,6)<> '000004' " +
         "GROUP BY  SUBSTR(l.licsch, 10, 6),R.NAME_REGNOM " +
         "order by round(sum(l.saldo_ish_nacval/ 1000),2) " +
         "FETCH FIRST 10 ROWS ONLY";
            string likvidler = "select "+
         "case when substr(l.licsch,1,3)= '100' then 'C5' " +
         "when substr(l.licsch,1,5)= '15025' then 'C7' " +
         "when substr(l.licsch,1,5)= '15020' then 'C8' " +
         "when substr(l.licsch,1,5)in ('11010', '11020') and l.licsch not in ('11010000010000200000', '11010000040000200000', '11020020010000200000') then 'C9' " +
         "when substr(l.licsch,1,5)in ('14010', '14014', '14030', '14034')then 'C10' end as qruplar " +
         ",sum(round(l.saldo_ish_nacval / 1000, 2))top, " +
         "SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '00' THEN ROUND(l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS azn, " +
         "SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) = '01' THEN ROUND(l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS usd, " +
         "SUM(CASE WHEN SUBSTR(l.licsch, 6, 2) NOT IN('00', '01') THEN ROUND(l.saldo_ish_nacval / 1000, 2) ELSE 0 END) AS diger " +
         "from odb.arh_saldo_ls l " +
         "where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') " +
         "and(substr(l.licsch, 1, 5) = '15020' or substr(l.licsch, 1, 3) = '100' or substr(l.licsch, 1, 5) = '15025' or " +
         "substr(l.licsch, 1, 5)in ('11010', '11020')or substr(l.licsch, 1, 5)in ('14010', '14014', '14030', '14034')) " +
         "and l.licsch not in ('11010000010000200000', '11010000040000200000', '11020020010000200000') " +
         "group by case when substr(l.licsch,1,3)= '100' then 'C5' " +
         "when substr(l.licsch,1,5)= '15025' then 'C7' " +
         "when substr(l.licsch,1,5)= '15020' then 'C8' " +
         "when substr(l.licsch,1,5)in ('11010', '11020') and l.licsch not in ('11010000010000200000', '11010000040000200000', '11020020010000200000') then 'C9' " +
         "when substr(l.licsch,1,5)in ('14010', '14014', '14030', '14034')then 'C10' end";
            #endregion
            DataTable _dt_huqtop10 = new DataTable();
            DataTable _dt_sefirlik = new DataTable();
            DataTable _dt_fiztop10 = new DataTable();
            DataTable _dt_medaxiller = new DataTable();
            DataTable _dt_qaliqlar = new DataTable();
            DataTable _dt_likvid = new DataTable();
            using (OracleConnection connection = new OracleConnection(cl_yanasma.con))
            {
                using (OracleCommand command = new OracleCommand(huquqi_top10, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_huqtop10);
                }
                using (OracleCommand command = new OracleCommand(sefirlik_qal, connection))
                {
                    //connection.Open();

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_sefirlik);
                }
                using (OracleCommand command = new OracleCommand(fiziki_top10, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_fiztop10);
                }
                using (OracleCommand command = new OracleCommand(medaxiller, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_medaxiller);
                }
                using (OracleCommand command = new OracleCommand(gunun_qaliqlari, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qaliqlar);
                }
                using (OracleCommand command = new OracleCommand(likvidler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_likvid);
                }
            }
            
            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "DEP.v01.1124d" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "Dep.xlsm");
            string filePath = Path.Combine(dosyayolu, fileName);

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
                #region MyRegion

                
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["ESAS"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["Portfelin strukturu"];
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets["Hüquqi şəxslər"];
                ExcelWorksheet worksheet4 = package.Workbook.Worksheets["Fiziki şəxslər"];
                ExcelWorksheet worksheet5 = package.Workbook.Worksheets["Likvid vəsaitlərə dair"];

                //*****************ESAS
                //*****GUNUN EVVELINE QALIQ
                var esas_c8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal")
                                    .ToList();
                decimal total_esas_c8 = esas_c8.Sum(row => row.Field<decimal>(3));

                var esas_c10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal")
                                    .ToList();
                decimal total_esas_c10 = esas_c10.Sum(row => row.Field<decimal>(3));
                //********
                var esas_d8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_d8 = esas_d8.Sum(row => row.Field<decimal>(3));

                var esas_d10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_d10 = esas_d10.Sum(row => row.Field<decimal>(3));

                //********
                var esas_e8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_e8 = esas_e8.Sum(row => row.Field<decimal>(2));

                var esas_e10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_e10 = esas_e10.Sum(row => row.Field<decimal>(2));
                //********
                var esas_f8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_f8 = esas_f8.Sum(row => row.Field<decimal>(3));

                var esas_f10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_f10 = esas_f10.Sum(row => row.Field<decimal>(3));

                //********
                var esas_g8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_g8 = esas_g8.Sum(row => row.Field<decimal>(2));

                var esas_g10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_g10 = esas_g10.Sum(row => row.Field<decimal>(2));

                //*****GUN ERZINDE MEDAXILLER
                var esas_H8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_med" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_H8 = esas_H8.Sum(row => row.Field<decimal>(3));
                var esas_H10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_med" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_H10 = esas_H10.Sum(row => row.Field<decimal>(3));
                var esas_I8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_med" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_I8 = esas_I8.Sum(row => row.Field<decimal>(2));
                var esas_I10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_med" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_I10 = esas_I10.Sum(row => row.Field<decimal>(2));

                var esas_J8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_med" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_J8 = esas_J8.Sum(row => row.Field<decimal>(3));
                var esas_J10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_med" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_J10 = esas_J10.Sum(row => row.Field<decimal>(3));
                var esas_K8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_med" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_K8 = esas_K8.Sum(row => row.Field<decimal>(2));
                var esas_K10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_med" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_K10 = esas_K10.Sum(row => row.Field<decimal>(2));

                //*****GUN ERZINDE MEXARICLER
                var esas_L8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_mex" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_L8 = esas_L8.Sum(row => row.Field<decimal>(3));
                var esas_L10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_mex" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_L10 = esas_L10.Sum(row => row.Field<decimal>(3));
                var esas_M8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_mex" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_M8 = esas_M8.Sum(row => row.Field<decimal>(2));
                var esas_M10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_mex" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_M10 = esas_M10.Sum(row => row.Field<decimal>(2));

                var esas_N8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_mex" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_N8 = esas_N8.Sum(row => row.Field<decimal>(3));
                var esas_N10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_mex" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_N10 = esas_N10.Sum(row => row.Field<decimal>(3));
                var esas_O8 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_mex" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_O8 = esas_O8.Sum(row => row.Field<decimal>(2));
                var esas_O10 = _dt_medaxiller.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_mex" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_O10 = esas_O10.Sum(row => row.Field<decimal>(2));

                //*****GUNUN SONUNA QALIQ
                var esas_Q8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal")
                                    .ToList();
                decimal total_esas_Q8 = esas_Q8.Sum(row => row.Field<decimal>(4));

                var esas_Q10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal")
                                    .ToList();
                decimal total_esas_Q10 = esas_Q10.Sum(row => row.Field<decimal>(4));
                //********
                var esas_R8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_R8 = esas_R8.Sum(row => row.Field<decimal>(4));

                var esas_R10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_R10 = esas_R10.Sum(row => row.Field<decimal>(4));

                //********
                var esas_S8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_S8 = esas_S8.Sum(row => row.Field<decimal>(2));

                var esas_S10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "manat")
                                    .ToList();
                decimal total_esas_S10 = esas_S10.Sum(row => row.Field<decimal>(2));
                //********
                var esas_T8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_T8 = esas_T8.Sum(row => row.Field<decimal>(4));

                var esas_T10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_T10 = esas_T10.Sum(row => row.Field<decimal>(4));

                //********
                var esas_U8 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_U8 = esas_U8.Sum(row => row.Field<decimal>(2));

                var esas_U10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "usd")
                                    .ToList();
                decimal total_esas_U10 = esas_U10.Sum(row => row.Field<decimal>(2));

                //*****************Portfelin strukturu
                var portfel_F10 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "fiziki_qal" && row.Field<string>(1) == "diger")
                                    .ToList();
                decimal total_portfel_F10 = portfel_F10.Sum(row => row.Field<decimal>(4));

                var portfel_F7 = _dt_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "huquqi_qal" && row.Field<string>(1) == "diger")
                                    .ToList();
                decimal total_portfel_F7 = portfel_F7.Sum(row => row.Field<decimal>(4));
                //*****************Likvid
                var likvid_C5 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C5")
                                    .ToList();
                decimal total_likvid_C5 = likvid_C5.Sum(row => row.Field<decimal>(2));
                var likvid_E5 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C5")
                                    .ToList();
                decimal total_likvid_E5 = likvid_E5.Sum(row => row.Field<decimal>(3));
                var likvid_F5 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C5")
                                    .ToList();
                decimal total_likvid_F5 = likvid_F5.Sum(row => row.Field<decimal>(4));

                var likvid_C7 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C7")
                                    .ToList();
                decimal total_likvid_C7 = likvid_C7.Sum(row => row.Field<decimal>(2));
                var likvid_E7 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C7")
                                    .ToList();
                decimal total_likvid_E7 = likvid_E7.Sum(row => row.Field<decimal>(3));
                var likvid_F7 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C7")
                                    .ToList();
                decimal total_likvid_F7 = likvid_F7.Sum(row => row.Field<decimal>(4));

                var likvid_C8 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C8")
                                    .ToList();
                decimal total_likvid_C8 = likvid_C8.Sum(row => row.Field<decimal>(2));
                var likvid_E8 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C8")
                                    .ToList();
                decimal total_likvid_E8 = likvid_E8.Sum(row => row.Field<decimal>(3));
                var likvid_F8 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C8")
                                    .ToList();
                decimal total_likvid_F8 = likvid_F8.Sum(row => row.Field<decimal>(4));

                var likvid_C9 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C9")
                                    .ToList();
                decimal total_likvid_C9 = likvid_C9.Sum(row => row.Field<decimal>(2));
                var likvid_E9 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C9")
                                    .ToList();
                decimal total_likvid_E9 = likvid_E9.Sum(row => row.Field<decimal>(3));
                var likvid_F9 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C9")
                                    .ToList();
                decimal total_likvid_F9 = likvid_F9.Sum(row => row.Field<decimal>(4));

                var likvid_C10 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C10")
                                    .ToList();
                decimal total_likvid_C10 = likvid_C10.Sum(row => row.Field<decimal>(2));
                var likvid_E10 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C10")
                                    .ToList();
                decimal total_likvid_E10 = likvid_E10.Sum(row => row.Field<decimal>(3));
                var likvid_F10 = _dt_likvid.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "C10")
                                    .ToList();
                decimal total_likvid_F10 = likvid_F10.Sum(row => row.Field<decimal>(4));

                worksheet1.Cells[8, 3].Value = total_esas_c8;
                worksheet1.Cells[10, 3].Value = total_esas_c10;
                worksheet1.Cells[8, 4].Value = total_esas_d8;
                worksheet1.Cells[10, 4].Value = total_esas_d10;
                worksheet1.Cells[8, 5].Value = total_esas_e8;
                worksheet1.Cells[10, 5].Value = total_esas_e10;
                worksheet1.Cells[8, 6].Value = total_esas_f8;
                worksheet1.Cells[10, 6].Value = total_esas_f10;
                worksheet1.Cells[8, 7].Value = total_esas_g8;
                worksheet1.Cells[10, 7].Value = total_esas_g10;
                worksheet1.Cells[8, 8].Value = total_esas_H8;
                worksheet1.Cells[10, 8].Value = total_esas_H10;
                worksheet1.Cells[8, 9].Value = total_esas_I8;
                worksheet1.Cells[10, 9].Value = total_esas_I10;
                worksheet1.Cells[8, 10].Value = total_esas_J8;
                worksheet1.Cells[10, 10].Value = total_esas_J10;

                worksheet1.Cells[8, 11].Value = total_esas_K8;
                worksheet1.Cells[10, 11].Value = total_esas_K10;
                worksheet1.Cells[8, 12].Value = total_esas_L8;
                worksheet1.Cells[10, 12].Value = total_esas_L10;
                worksheet1.Cells[8, 13].Value = total_esas_M8;
                worksheet1.Cells[10, 13].Value = total_esas_M10;
                worksheet1.Cells[8, 14].Value = total_esas_N8;
                worksheet1.Cells[10, 14].Value = total_esas_N10;
                worksheet1.Cells[8, 15].Value = total_esas_O8;
                worksheet1.Cells[10, 15].Value = total_esas_O10;

                worksheet1.Cells[8, 17].Value = total_esas_Q8;
                worksheet1.Cells[10, 17].Value = total_esas_Q10;
                worksheet1.Cells[8, 18].Value = total_esas_R8;
                worksheet1.Cells[10, 18].Value = total_esas_R10;
                worksheet1.Cells[8, 19].Value = total_esas_S8;
                worksheet1.Cells[10, 19].Value = total_esas_S10;
                worksheet1.Cells[8, 20].Value = total_esas_T8;
                worksheet1.Cells[10, 20].Value = total_esas_T10;
                worksheet1.Cells[8, 21].Value = total_esas_U8;
                worksheet1.Cells[10, 21].Value = total_esas_U10;
                //portfel
                worksheet2.Cells[7, 6].Value = total_portfel_F7;
                worksheet2.Cells[10, 6].Value = total_portfel_F10;
                //huquqi qaliq
                worksheet3.Cells[4, 3].Value = _dt_huqtop10.Rows[0][2];
                worksheet3.Cells[5, 3].Value = _dt_huqtop10.Rows[1][2];
                worksheet3.Cells[6, 3].Value = _dt_huqtop10.Rows[2][2];
                worksheet3.Cells[7, 3].Value = _dt_huqtop10.Rows[3][2];
                worksheet3.Cells[8, 3].Value = _dt_huqtop10.Rows[4][2];
                worksheet3.Cells[9, 3].Value = _dt_huqtop10.Rows[5][2];
                worksheet3.Cells[10, 3].Value = _dt_huqtop10.Rows[6][2];
                worksheet3.Cells[11, 3].Value = _dt_huqtop10.Rows[7][2];
                worksheet3.Cells[12, 3].Value = _dt_huqtop10.Rows[8][2];
                worksheet3.Cells[13, 3].Value = _dt_huqtop10.Rows[9][2];

                int rowCount_huqtop10 = _dt_huqtop10.Rows.Count;

                for (int row = 4; row <= Math.Min(13, rowCount_huqtop10 + 3); row++)
                {
                    for (int col = 5; col <= 7; col++)
                    {
                        if (row - 4 < rowCount_huqtop10 && _dt_huqtop10.Rows[row - 4][col - 3].ToString() == "004801")
                        {
                            decimal huqtop10Value = Convert.ToDecimal(_dt_huqtop10.Rows[row - 4][col - 1]);

                            decimal sefirlikValue1 = Convert.ToDecimal(_dt_sefirlik.Rows[2][4]);
                            decimal sefirlikValue2 = Convert.ToDecimal(_dt_sefirlik.Rows[0][4]);
                            decimal sefirlikValue3 = Convert.ToDecimal(_dt_sefirlik.Rows[1][4]);

                            worksheet3.Cells[row, 5].Value = huqtop10Value + sefirlikValue1;
                            worksheet3.Cells[row, 6].Value = huqtop10Value + sefirlikValue2;
                            worksheet3.Cells[row, 7].Value = huqtop10Value + sefirlikValue3;
                        }
                        else if (row - 4 < rowCount_huqtop10)
                        {
                            worksheet3.Cells[row, col].Value = Convert.ToDecimal(_dt_huqtop10.Rows[row - 4][col - 1]);
                        }
                        else
                        {
                            // _dt_huqtop10-da həmin setir mövcud deyilsə, uyğun maneəni təyin edin.
                            //Console.WriteLine("Məlumat mövcud deyil.");
                        }
                    }
                }
                //fiziki qaliq
                worksheet4.Cells[4, 3].Value = _dt_fiztop10.Rows[0][2];
                worksheet4.Cells[5, 3].Value = _dt_fiztop10.Rows[1][2];
                worksheet4.Cells[6, 3].Value = _dt_fiztop10.Rows[2][2];
                worksheet4.Cells[7, 3].Value = _dt_fiztop10.Rows[3][2];
                worksheet4.Cells[8, 3].Value = _dt_fiztop10.Rows[4][2];
                worksheet4.Cells[9, 3].Value = _dt_fiztop10.Rows[5][2];
                worksheet4.Cells[10, 3].Value = _dt_fiztop10.Rows[6][2];
                worksheet4.Cells[11, 3].Value = _dt_fiztop10.Rows[7][2];
                worksheet4.Cells[12, 3].Value = _dt_fiztop10.Rows[8][2];
                worksheet4.Cells[13, 3].Value = _dt_fiztop10.Rows[9][2];

                int rowCount = _dt_fiztop10.Rows.Count;
                for (int row = 4; row <= Math.Min(13, rowCount + 3); row++)
{
                    for (int col = 5; col <= 7; col++)
                    {
                        // _dt_fiztop10 DataTable-indəki dəyəri Decimal formatına çevirib Excel hücrəsinə yazırıq.
                        if (row - 4 < rowCount)
                        {
                            worksheet4.Cells[row, col].Value = Convert.ToDecimal(_dt_fiztop10.Rows[row - 4][col - 1]);
                        }
                        else
                        {
                            // _dt_fiztop10-da həmin setir mövcud deyilsə, uyğun maneəni təyin edin.
                            Console.WriteLine("Məlumat mövcud deyil.");
                        }
                    }
                }
                //Likvid vəsaitlərə dair
                worksheet5.Cells[5, 4].Value = total_likvid_C5;
                worksheet5.Cells[5, 5].Value = total_likvid_E5;
                worksheet5.Cells[5, 6].Value = total_likvid_F5;
                worksheet5.Cells[7, 4].Value = total_likvid_C7;
                worksheet5.Cells[7, 5].Value = total_likvid_E7;
                worksheet5.Cells[7, 6].Value = total_likvid_F7;
                worksheet5.Cells[8, 4].Value = total_likvid_C8;
                worksheet5.Cells[8, 5].Value = total_likvid_E8;
                worksheet5.Cells[8, 6].Value = total_likvid_F8;
                worksheet5.Cells[9, 4].Value = total_likvid_C9;
                worksheet5.Cells[9, 5].Value = total_likvid_E9;
                worksheet5.Cells[9, 6].Value = total_likvid_F9;
                worksheet5.Cells[10, 4].Value = total_likvid_C10;
                worksheet5.Cells[10, 5].Value = total_likvid_E10;
                worksheet5.Cells[10, 6].Value = total_likvid_F10;

                //worksheet1.Cells[2, 1].Value = "Bank Melli İran Bakı filialı";
                worksheet1.Cells[3, 1].Value = "Tarix:" + textBox2.Text;
                #endregion
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
                }
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (button1.Text == "Daily Report")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                    button1.Focus();
                    e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
                }
            }
            else
            {
                if (e.KeyCode == Keys.Enter)
                {
                    // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                    button1.Focus();
                    e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            
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
            Excel();
        }
    }
}
