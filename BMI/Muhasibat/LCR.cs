using DevExpress.XtraGrid;
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
using System.Windows.Controls;
using System.Windows.Forms;


namespace BMI.Muhasibat
{
    public partial class LCR : Form
    {
        public LCR()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_hesabat_tarixi.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_hesabat_tarixi.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }
        
        private void excel()
        {
            

            DataTable dt_L2 = new DataTable();
            DataTable dt_proqnoz = new DataTable();
            DataTable dt_xett = new DataTable();
            DataTable dt_likvid = new DataTable();
            DataTable dt_qiym_kag = new DataTable();
            DataTable dt_akk_qarant = new DataTable();
            DataTable dt_99531 = new DataTable();

            #region sql_kodlar

            
            string LCR = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "SUBSTR(ar.licsch, 6, 2) AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";
            #region Proqnoz kohne
            //      string Proqnoz = "select distinct n.licschpkre,n.val,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez,n.gec_gun," +
            //      "n.meb*ROUND(odb.func_get_kurval(substr(n.licschpkre,6,2),TO_DATE('" + txt_hesabat_tarixi.Text + "', 'DD-MM-YYYY')),6) ekv,n.faiz,n.mud," +
            //          "'' ay,asz.odenis,n.tip from " +
            //"(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, substr(m.licschpkre, 6, 2) val, m.meb, m.faiz, m.mud, m.tip, count(*) kol from " +
            //"(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19, s.setmininterestreserves min_rez, " +
            //"odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, t.summakre meb, t.procstavkre faiz, t.srok mud, t.tipkredita tip " +
            //"from view_nacpogprokre_all x, odb.arh_licschkre t, odb.srokpogprockre s, tipkre g " +
            //"where t.tipkredita = g.code and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null " +
            //"and x.date_oper = to_date('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') and x.date_oper = t.date_oper " +
            //"order by(x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)), x.date_oper asc) m " +
            //"group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, m.meb, m.faiz, m.mud, m.tip " +
            //"order by m.gec_gun) n, (SELECT distinct substr(ar.kredit, 10, 6) qeyd, ar.ssk, sum(ar.summa_v_nacval) odenis " +
            //"                    FROM regnom rr, licschkre l " +
            //"                    JOIN arh_dd ar ON EXTRACT(MONTH FROM ar.date_oper) = EXTRACT(MONTH FROM TO_DATE('" + txt_hesabat_tarixi.Text + "', 'DD-MM-YYYY')) " +
            //"                    AND EXTRACT(YEAR FROM ar.date_oper) = EXTRACT(YEAR FROM TO_DATE('" + txt_hesabat_tarixi.Text + "', 'DD-MM-YYYY')) " +
            //"                    WHERE substr(l.licschkre, 10, 6) = rr.regnom and(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
            //"                    AND ar.ssk = l.subschkre AND " +
            //"                    ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
            //"                     group by substr(ar.kredit, 10, 6),ar.ssk) asz where substr(n.licschpkre, 10, 6) = asz.qeyd(+)and n.sk = asz.ssk(+)";
            #endregion
            #region Prognozyeni
            string Proqnoz = "select distinct n.licschpkre,n.val,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez,n.gec_gun," +
            "n.meb*ROUND(odb.func_get_kurval(substr(n.licschpkre,6,2),TO_DATE('" + txt_hesabat_tarixi.Text + "', 'DD-MM-YYYY')),6) ekv,n.faiz,n.mud," +
                "'' ay,asz.odenis,n.tip from " +
      "(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, substr(m.licschpkre, 6, 2) val, m.meb, m.faiz, m.mud, m.tip, count(*) kol from " +
      "(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19, s.setmininterestreserves min_rez, " +
      "odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, t.summakre meb, t.procstavkre faiz, t.srok mud, t.tipkredita tip " +
      "from view_nacpogprokre_all x, odb.arh_licschkre t, odb.srokpogprockre s, tipkre g " +
      "where t.tipkredita = g.code and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null " +
      "and x.date_oper = to_date('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') and x.date_oper = t.date_oper " +
      "order by(x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)), x.date_oper asc) m " +
      "group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, m.meb, m.faiz, m.mud, m.tip " +
      "order by m.gec_gun) n, (SELECT distinct substr(ar.kredit, 10, 6) qeyd, ar.ssk, round( sum(ar.summa_v_nacval)/6,2) odenis " +
                         " FROM regnom rr, licschkre l, arh_dd ar " +
                          "WHERE substr(l.licschkre, 10, 6) = rr.regnom and(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre)" +
                          "AND ar.ssk = l.subschkre AND ar.date_oper between TO_DATE('" + txt_sonaltiay.Text + "', 'DD-MM-YYYY') and TO_DATE('" + txt_hesabat_tarixi.Text + "', 'DD-MM-YYYY') and " +
                          "((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
                           "group by substr(ar.kredit, 10, 6),ar.ssk) asz where substr(n.licschpkre, 10, 6) = asz.qeyd(+)and n.sk = asz.ssk(+)";
            #endregion
            #region Proqnoz en son deyiiklik
            string Prognoz = @"
SELECT 
    ar.licschkre hes,
    SUBSTR(ar.licschkre, 6, 2) val,
    ar.subschkre sub,
    odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper)) gec_gun,
    ar.tipkredita tip,
    (
        SELECT ROUND(
            SUM(dd.summa_v_nacval) / 
            NULLIF(MONTHS_BETWEEN(
                TO_DATE(:hesabat_tar, 'DD-MM-YYYY'),
                TO_DATE(:six_months_ago, 'DD-MM-YYYY')
            ), 0), 2
        )
        FROM arh_dd dd 
        WHERE SUBSTR(dd.debet, 1, 1) IN ('3','4') 
          AND dd.kredit IN (ar.licschkre, ar.licschpkre, ar.licsch_19, ar.licschppkre)
          AND dd.date_oper BETWEEN TO_DATE(:six_months_ago, 'DD-MM-YYYY') 
                               AND TO_DATE(:hesabat_tar, 'DD-MM-YYYY')
    ) odenis,
    (
        SELECT g.summa_pog_kre + g.summa_pog_pro 
        FROM graphpogkre g
        WHERE g.subschkre = ar.subschkre 
          AND g.licschkre = ar.licschkre
          AND TO_CHAR(g.date_pog, 'MM-YYYY') = TO_CHAR(
              ADD_MONTHS(TO_DATE(:hesabat_tar, 'DD-MM-YYYY'), 1), 'MM-YYYY'
          )
    ) ayliq
FROM arh_licschkre ar, view_nacpogprokre_all x
WHERE 
    (ar.date_close IS NULL OR ar.date_close > TO_DATE(:hesabat_tar, 'DD-MM-YYYY'))
    AND ar.date_oper = TO_DATE(:hesabat_tar, 'DD-MM-YYYY')
    AND x.licschpkre = ar.licschpkre 
    AND x.subschkre = ar.subschkre 
    AND x.date_oper = ar.date_oper";

            #endregion
            string daily_report_bk_xett = "select distinct t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish, "+
                "t.ostatok_ish* ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv, "+
                "ROUND(odb.func_get_kurval(substr(t.licsch, 6, 2), t.date_oper), 6)  kurs ,ar.date_planclose," +
                "odb.tar_ferq360(ar.date_planclose,t.date_oper) gun_ferqi," +
                "ar.tipkredita " +
                "from odb.arh_saldo_vbls t, arh_licschkre ar,tipkre g where ar.tipkredita = g.code " +
                "and t.date_oper = TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99540,99550,99531) and t.ostatok_ish > 0 " +
                "and substr(t.licsch,10,6)= substr(ar.licschkre, 10, 6) and t.ssls = ar.subschkre and t.date_oper=ar.date_oper and t.licsch is not null";

            string likvid = "select l.licsch,substr(l.licsch,6,2),round(sum(l.saldo_ish_nacval/1000),2) from odb.arh_saldo_ls l " +
                   "where l.date_oper=TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') " +
                   " and substr(l.licsch,1,5) in ('15770') " +
                   " group by l.licsch";

            string qiymetli_kag = "select substr(t.licsch_cb,0,5)hes,substr(t.licsch_cb,6,2) val,t.subsch_cb,t.summa_cb," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licsch_cb,6,2),t.date_oper),6) ekv,t.diskont," +
                "odb.tar_ferq360(t.date_planclose,t.date_oper) gun_ferqi from odb.arh_licsch_cb t "+
                "where t.date_oper = TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy') and t.summa > 0 order by t.licsch_cb,t.subsch_cb";

            string akk_qarant = "select t.vbs bk,substr(t.licschgar,6,2)val," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licschgar,6,2),t.date_oper),6) ekv ,"+
               " case when t.date_prolong is null then odb.tar_ferq360(t.date_planclose, t.date_oper) " +
               "else odb.tar_ferq360(t.date_prolong, t.date_oper) end gun_ferqi "+
               " from arh_licschgar t where t.date_oper = TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy')and t.summa > 0";

            string birterefli_99531 = "select * from odb.arh_saldo_vbls t where t.ostatok_ish > 0 and t.vbs='99531' " +
                "and t.date_oper = TO_DATE('" + txt_hesabat_tarixi.Text + "', 'dd/mm/yyyy')";
            #endregion

            using (OracleConnection connection = new OracleConnection(cl.con))

            {
                using (OracleCommand command = new OracleCommand(LCR, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_L2);
                }
                using (OracleCommand command = new OracleCommand(Prognoz, connection))
                {
                    // Tarixləri textbox-lardan oxu
                    string hesabatTar = txt_hesabat_tarixi.Text.Trim();      // Məs: "31-03-2025"
                    string altıAyEvvel = txt_sonaltiay.Text.Trim();      // Məs: "30-09-2024"

                    // Parametrləri əlavə et
                    command.Parameters.Add("hesabat_tar", OracleDbType.Varchar2).Value = hesabatTar;
                    command.Parameters.Add("six_months_ago", OracleDbType.Varchar2).Value = altıAyEvvel;

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_proqnoz);
                }

                //using (OracleCommand command = new OracleCommand(Proqnoz, connection))
                //{
                //    //connection.Open();
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(dt_proqnoz);


                //    //foreach (DataRow row in dt_proqnoz.Rows)  SONRADAN DATATABLEYE AYLIQ ATMAQ UCUN
                //    //{
                //    //    double meb = Convert.ToDouble(row[7]);
                //    //    double faizOranı = Convert.ToDouble(row[8]);
                //    //    int vadeMüddeti = Convert.ToInt32(row[9]) / 30;

                //    //    double aylıkÖdeme = Math.Round(CalculateMonthlyPayment(meb, vadeMüddeti, faizOranı), 2);
                //    //    row[10] = aylıkÖdeme;
                //    //}

                //}
                using (OracleCommand command = new OracleCommand(daily_report_bk_xett, connection))
                {
                    dt_xett.Clear();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_xett);
                    connection.Close();
                    
                    //gridControl1.DataSource = dt_xett;
                }

                using (OracleCommand command = new OracleCommand(likvid, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_likvid);

                }

                using (OracleCommand command = new OracleCommand(qiymetli_kag, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_qiym_kag);
                    
                }

                using (OracleCommand command = new OracleCommand(akk_qarant, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_akk_qarant);
                }
                using (OracleCommand command = new OracleCommand(birterefli_99531, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_99531);
                    connection.Close();
                }

            }

            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txt_hesabat_tarixi.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "LCR_1124m" +yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "LCR_.xlsm");
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

            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet wsL1 = package.Workbook.Worksheets["L1"];
                ExcelWorksheet wsL2 = package.Workbook.Worksheets["L2"];
                ExcelWorksheet wsL3_A = package.Workbook.Worksheets["L3 (A)"];
                ExcelWorksheet wsL3_B = package.Workbook.Worksheets["L3 (B)"];

                string[] L2_c15 = { "100" };
                string[] L2_c16 = { "11010","11110", "11710" };
                string[] L2_d16 = { "11020" };
                string[] L2_c16_ist = { "11010000010000200000"};
                string[] L2_d16_ist = { "11020020010000200000" };
                string[] L2_c17 = { "14010", "14012", "14014", "14030", "14032", "14034" };
                string[] L2_f16 = { "11010000040000200000" };

                string[] L3A_c21 = { "410", "419" };
                string[] L3A_f24 = { "411", "412" };
                string[] L3A_c36 = { "35015", "35020", "35025", "35026" };
                string[] L3A_c37 = { "49025" };
                string[] L3A_c39 = { "35090", "35190" };
                string[] L3A_c42 = { "35938" };
                string[] L3A_c42_qisa = { "38", "39", "40" };

                //L2
                var sh_L2_c15 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c15 = sh_L2_c15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_d15 = dt_L2.AsEnumerable()
                .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                .ToList();
                decimal total_sh_L2_d15 = sh_L2_d15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c16 = dt_L2.AsEnumerable()
                    .Where(row => L2_c16.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00"
                     && !L2_c16_ist.Contains(row.Field<string>(1)))
                    .ToList();
                decimal total_sh_L2_c16 = sh_L2_c16.Sum(row => row.Field<decimal>(3)) / 1000;


                var sh_L2_d16 = dt_L2.AsEnumerable()
                .Where(row => L2_d16.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00"
                 && !L2_d16_ist.Contains(row.Field<string>(1)))
                .ToList();
                decimal total_sh_L2_d16 = sh_L2_d16.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c17 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c17.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c17 = sh_L2_c17.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_f16 = dt_L2.AsEnumerable()
                                    .Where(row => L2_f16.Contains(row.Field<string>(1)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_f16 = sh_L2_f16.Sum(row => row.Field<decimal>(3)) / 1000;

                //L3A

                var sh_L3A_c21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c21 = sh_L3A_c21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d21 = sh_L3A_d21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_f24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_f24 = sh_L3A_f24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d24 = sh_L3A_d24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c36 = sh_L3A_c36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d36 = sh_L3A_d36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c37 = sh_L3A_c37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d37 = sh_L3A_d37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c39 = sh_L3A_c39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d39 = sh_L3A_d39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c42 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                                    .ToList();
                decimal total_sh_L3A_c42 = sh_L3A_c42.Sum(row => row.Field<decimal>(3)) / 1000;

                //DataTable filteredDataTable = dt_L2.Clone(); // İlk tablonun şemasını kopyala
                //foreach (var row in sh_L3A_c42)
                //{
                //    filteredDataTable.Rows.Add(row.ItemArray);
                //}

                //// DataGridView'e yeni DataTable'ı atayarak güncelle
                //gridControl1.DataSource = filteredDataTable;

                var sh_L3A_d42 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                .ToList();

                decimal total_sh_L3A_d42 = sh_L3A_d42.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_c43 = sh_L3A_c43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_d43 = sh_L3A_d43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_c44 = sh_L3A_c44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_d44 = sh_L3A_d44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c45 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                                    .ToList();
                decimal total_sh_L3A_c45 = sh_L3A_c45.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d45 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                .ToList();

                decimal total_sh_L3A_d45 = sh_L3A_d45.Sum(row => row.Field<decimal>(3)) / 1000;

                //L3 B PROQNOZ

                var sh_L3B_c19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(3) <= 90 && row.Field<string>(1) == "00")
                      .ToList();

                decimal total_sh_L3B_c19 = sh_L3B_c19
                    .Where(row => row.Field<decimal?>(6) != null)
                .Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3B_c19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(3) > 90 && row.Field<string>(1) == "00")
                      .ToList();

                //double total_sh_L3B_c19_90_cox = sh_L3B_c19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;
                decimal total_sh_L3B_c19_90_cox = sh_L3B_c19_90_cox                         //sh_L3B_c19_90 bunu deyisidm
                .Where(row => row.Field<decimal?>(5) != null)
                .Sum(row => row.Field<decimal>(5)) / 1000;


                var sh_L3B_d19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(3) <= 90 && row.Field<string>(1) != "00")
                      .ToList();

                decimal total_sh_L3B_d19 = sh_L3B_d19.Sum(row => Convert.ToDecimal(row.Field<string>(6))) / 1000;

                var sh_L3B_d19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(3) > 90 && row.Field<string>(1) != "00")
                      .ToList();

                //double total_sh_L3B_d19_90_cox = sh_L3B_d19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;

                decimal total_sh_L3B_d19_90_cox = sh_L3B_d19
                .Where(row => row.Field<decimal?>(5) != null)
                .Sum(row => row.Field<decimal>(5)) / 1000;

                DataTable filteredDataTableproqnoz = dt_proqnoz.Clone(); // İlk tablonun şemasını kopyala
                foreach (var row in sh_L3B_d19)
                {
                    filteredDataTableproqnoz.Rows.Add(row.ItemArray);
                }

                // DataGridView'e yeni DataTable'ı atayarak güncelle
                dataGridView1.DataSource = filteredDataTableproqnoz;

                //QIYMETLI KAGIZLAR

                var sh_L3B_c33 = dt_likvid.AsEnumerable()
                     .Where(row => row.Field<string>(1) == "00")
                     .ToList();
                decimal total_sh_L3B_c33 = sh_L3B_c33.Sum(row => row.Field<decimal>(2)) ;

                //var sh_L3B_c49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) == "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_c49 = sh_L3B_c49.Sum(row => row.Field<decimal>(4)) / 1000;

                //var sh_L3B_d49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) != "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_d49 = sh_L3B_d49.Sum(row => row.Field<decimal>(4)) / 1000;

                //Akkreditiv 30 gune
                var sh_L3A_c76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_c76_30 = sh_L3A_c76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_d76_30 = sh_L3A_d76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Akkreditiv 30 gunden cox
                var sh_L3A_f76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_f76_90 = sh_L3A_f76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_g76_90 = sh_L3A_g76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gune
                var sh_L3A_c77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_c77_30 = sh_L3A_c77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_d77_30 = sh_L3A_d77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gunden cox
                var sh_L3A_f77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_30 = sh_L3A_f77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_30 = sh_L3A_g77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_f77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_90 = sh_L3A_f77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_90 = sh_L3A_g77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //XETLER 30 79
                var sh_L3A_c79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00" 
                        && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                double total_sh_L3A_c79_30 = sh_L3A_c79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                var sh_L3A_d79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00" 
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                    .ToList();


                double total_sh_L3A_d79_30 = sh_L3A_d79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                //XETLER 30 80
                var sh_L3A_c80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                      &&  row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_c80_30 = sh_L3A_c80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_d80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_d80_30 = sh_L3A_d80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 79
                var sh_L3A_f79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) == "00" 
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_f79_30 = sh_L3A_f79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) != "00" 
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_g79_30 = sh_L3A_g79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 80
                var sh_L3A_f80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_f80_30 = sh_L3A_f80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_g80_30 = sh_L3A_g80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //Birtərəfli ləğv etmə
                var sh_L3A_f84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) == "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_f84 = sh_L3A_f84.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_g84 = sh_L3A_g84.Sum(row => row.Field<decimal>(6)) / 1000;

                wsL1.Cells[9, 3].Value = txt_hesabat_tarixi.Text;

                wsL2.Cells[15, 3].Value = total_sh_L2_c15;
                wsL2.Cells[15, 4].Value = total_sh_L2_d15;

                wsL2.Cells[16, 3].Value = total_sh_L2_c16;
                wsL2.Cells[16, 4].Value = total_sh_L2_d16;
                wsL2.Cells[16, 6].Value = total_sh_L2_f16;

                wsL2.Cells[17, 3].Value = total_sh_L2_c17;

                wsL3_A.Cells[20, 3].Value = -total_sh_L3A_c21;
                wsL3_A.Cells[20, 4].Value = -total_sh_L3A_d21;

                wsL3_A.Cells[24, 3].Value = -total_sh_L3A_f24;
                wsL3_A.Cells[24, 4].Value = -total_sh_L3A_f24;

                wsL3_A.Cells[36, 3].Value = -total_sh_L3A_c36;
                wsL3_A.Cells[36, 4].Value = -total_sh_L3A_d36;

                wsL3_A.Cells[37, 6].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[37, 7].Value = -total_sh_L3A_d37;

                wsL3_A.Cells[39, 3].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[39, 4].Value = -total_sh_L3A_d39;

                wsL3_A.Cells[42, 3].Value = -total_sh_L3A_c42;
                wsL3_A.Cells[42, 4].Value = -total_sh_L3A_d42;

                wsL3_A.Cells[43, 3].Value = -total_sh_L3A_c43;
                wsL3_A.Cells[43, 4].Value = -total_sh_L3A_d43;

                wsL3_A.Cells[44, 3].Value = -total_sh_L3A_c44;
                wsL3_A.Cells[44, 4].Value = -total_sh_L3A_d44;

                wsL3_A.Cells[45, 3].Value = -total_sh_L3A_c45;
                wsL3_A.Cells[45, 4].Value = -total_sh_L3A_d45;

                wsL3_B.Cells[19, 3].Value = total_sh_L3B_c19 + total_sh_L3B_c19_90_cox;
                wsL3_B.Cells[19, 4].Value = total_sh_L3B_d19;

                wsL3_A.Cells[79, 3].Value = total_sh_L3A_c79_30;
                wsL3_A.Cells[79, 4].Value = total_sh_L3A_d79_30;

                wsL3_A.Cells[80, 3].Value = total_sh_L3A_c80_30;
                wsL3_A.Cells[80, 4].Value = total_sh_L3A_d80_30;

                wsL3_A.Cells[79, 6].Value = total_sh_L3A_f79_30;
                wsL3_A.Cells[79, 7].Value = total_sh_L3A_g79_30;

                wsL3_A.Cells[80, 6].Value = total_sh_L3A_f80_30;
                wsL3_A.Cells[80, 7].Value = total_sh_L3A_g80_30;

                wsL3_B.Cells[37, 3].Value = total_sh_L3B_c33;//hazirlamaq
                wsL3_B.Cells[37, 4].Value = 0;

                //wsL3_B.Cells[49, 3].Value = total_sh_L3B_c49;
                //wsL3_B.Cells[49, 4].Value = total_sh_L3B_d49;

                wsL3_A.Cells[76, 3].Value = total_sh_L3A_c76_30;
                wsL3_A.Cells[76, 4].Value = total_sh_L3A_d76_30;

                wsL3_A.Cells[76, 6].Value = total_sh_L3A_f76_90;
                wsL3_A.Cells[76, 7].Value = total_sh_L3A_f76_90;

                wsL3_A.Cells[77, 3].Value = total_sh_L3A_c77_30;
                wsL3_A.Cells[77, 4].Value = total_sh_L3A_d77_30;

                wsL3_A.Cells[77, 6].Value = total_sh_L3A_f77_90;
                wsL3_A.Cells[77, 7].Value = total_sh_L3A_g77_90;

                wsL3_A.Cells[84, 6].Value = total_sh_L3A_f84;
                wsL3_A.Cells[84, 7].Value = total_sh_L3A_g84;

                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }

        public static double CalculateMonthlyPayment(double presentValue, double financingPeriod, double interestRatePerYear)
        {
            double monthlyInterestRate = interestRatePerYear / 1200;
            double x = Math.Pow(1 + monthlyInterestRate, financingPeriod);
            double monthlyPayment = (presentValue * monthlyInterestRate * x) / (x - 1);
            return monthlyPayment;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel();
        }

        private void LCR_Load(object sender, EventArgs e)
        {
            
        }

        private void txt_sonaltiay_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_sonaltiay.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_sonaltiay.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txt_sonaltiay_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_hesabat_tarixi.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }
    }
}
