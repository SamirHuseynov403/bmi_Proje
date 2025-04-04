using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;

namespace BMI
{
    public partial class frmehtiyyatlar : Form
    {
        public frmehtiyyatlar()
        {
            InitializeComponent();
        }
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        OracleCommand komyazdir;
        OracleCommand komutmurraz;
        System.Data.DataTable muracietler = new System.Data.DataTable();
        System.Data.DataTable dt_axtarilan = new System.Data.DataTable();
        System.Data.DataTable dt_axtarilantek = new System.Data.DataTable();

        private System.Data.DataTable dataTable;
        public void listelemuraciet()
        {

            //try
            //{
                muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }

                OracleDataAdapter isleme = new OracleDataAdapter("select distinct n.licschpkre,n.sk,n.procstavrez,n.procstavrez_19," +
                    "n.min_rez,n.gec_gun from(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, " +
                    "count(*) kol from(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19," +
                    " s.setmininterestreserves min_rez,odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate,x.date_oper)) as gec_gun from view_nacpogprokre_all x, odb.licschkre t," +
                    " odb.srokpogprockre s where(x.lastoverduedate_ish is not null)and" +
                    " odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate,x.date_oper)) > 30 and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre " +
                    "and x.subschkre = s.subschkre and t.date_close is null and x.date_oper between to_date('" + txtgiris.Text + "','dd/mm/yyyy') " +
                    "and to_date('" + txtcixis.Text + "','dd/mm/yyyy')order by " +
                    " odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish)," +
                    " x.date_oper))) m group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun order by m.date_oper," +
                    " m.gec_gun) n where n.kol > 1 and n.gec_gun > 60 and n.min_rez <= 15 order by n.licschpkre, n.sk", con);

                //OracleDataAdapter isleme = new OracleDataAdapter("select distinct n.licschpkre,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez from (select m.licschpkre,m.sk,m.procstavrez,m.procstavrez_19, m.min_rez, m.gec_gun,count(*) kol from (select x.date_oper,x.licschpkre,x.subschkre sk,t.procstavrez,t.procstavrez_19,s.setmininterestreserves min_rez, tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)) gec_gun from view_nacpogprokre_all x, odb.licschkre t,odb.srokpogprockre s where (x.lastoverduedate_ish is not null) and (x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish,x.lastoverduedate,x.lodinterest_ish),x.date_oper)) >30 and x.licschpkre=t.licschpkre and x.subschkre=t.subschkre and t.licschkre=s.licschkre and x.subschkre=s.subschkre and t.date_close is null and x.date_oper between to_date('"+txtgiris.Text+ "','dd/mm/yyyy') and to_date('" + txtcixis.Text + "','dd/mm/yyyy') order by (x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish,x.lastoverduedate,x.lodinterest_ish),x.date_oper)),x.date_oper asc) m group by m.licschpkre,m.sk,m.procstavrez,m.procstavrez_19,m.min_rez,m.gec_gun order by m.gec_gun) n where n.kol>1 or n.gec_gun>60 and n.min_rez<15 and n.procstavrez<25", con);
                isleme.Fill(muracietler);
                dtg_siyahi.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

            dtg_siyahi.Columns[0].HeaderText = "Suda hesabı";
            dtg_siyahi.Columns[0].Width = 250;
            dtg_siyahi.Columns[1].HeaderText = "Sub kod";
            dtg_siyahi.Columns[1].Width = 250;
            dtg_siyahi.Columns[2].HeaderText = "Ehtiyyatar";
            dtg_siyahi.Columns[2].Width = 250;
            dtg_siyahi.Columns[3].HeaderText = "VK ehtiyyat";
            dtg_siyahi.Columns[3].Width = 200;
            dtg_siyahi.Columns[4].HeaderText = "Min ehtiyyat";
            dtg_siyahi.Columns[4].Width = 200;


            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        public void listele_30_iki_defe()
        {

            //try
            //{
            muracietler.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            OracleDataAdapter isleme = new OracleDataAdapter("WITH min_gecikme_her_ay AS (SELECT x.licschpkre AS hes,x.subschkre AS sk, " +
            " TO_CHAR(x.date_oper, 'YYYY-MM') AS ay_il, " +
            " MIN(odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper))) AS min_gec_gun, " +
            " MAX(odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper))) AS max_gec_gun, " +
            " MIN(x.date_oper) AS min_gecikme_tarixi " +
            " FROM view_nacpogprokre_all x JOIN odb.licschkre t ON x.licschpkre = t.licschpkre AND x.subschkre = t.subschkre " +
            " WHERE x.date_oper BETWEEN TO_DATE('"+txtgiris.Text+"', 'DD-MM-YYYY') AND TO_DATE('"+txtcixis.Text+"', 'DD-MM-YYYY') " +
            " AND t.date_close IS NULL AND t.procstavrez < 15 " +
            " AND odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper)) > 30 " +
            " GROUP BY x.licschpkre, x.subschkre, TO_CHAR(x.date_oper, 'YYYY-MM') " +
            " ), ay_farki AS(SELECT a.hes, a.sk, a.ay_il AS ay_il_a, a.min_gecikme_tarixi AS tarix_a, b.ay_il AS ay_il_b, " +
            " b.min_gecikme_tarixi AS tarix_b, b.min_gec_gun AS ikinci_ay_min_gecikme, " +
            " b.max_gec_gun AS ikinci_ay_max_gecikme, " +
            " ABS(a.min_gecikme_tarixi - b.min_gecikme_tarixi) AS gun_farki " +
            " FROM min_gecikme_her_ay a " +
            " JOIN min_gecikme_her_ay b ON a.hes = b.hes AND a.sk = b.sk AND a.ay_il < b.ay_il), " +
            " unik_hesablar AS(SELECT hes, sk, MIN(ay_il_a) AS ilk_ay, MIN(ay_il_b) AS ikinci_ay, " +
            " MIN(ikinci_ay_min_gecikme) AS ikinci_ay_min_gecikme, " +
            " MIN(ikinci_ay_max_gecikme) AS ikinci_ay_max_gecikme, " +
            " MIN(gun_farki) AS aradaki_gun " +
            " FROM ay_farki WHERE gun_farki >= 30 GROUP BY hes, sk) " +
            " SELECT hes, sk, ilk_ay, ikinci_ay, ikinci_ay_min_gecikme, " +
            " ikinci_ay_max_gecikme, aradaki_gun FROM unik_hesablar ORDER BY hes, sk", con);

            //OracleDataAdapter isleme = new OracleDataAdapter("select distinct n.licschpkre,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez from (select m.licschpkre,m.sk,m.procstavrez,m.procstavrez_19, m.min_rez, m.gec_gun,count(*) kol from (select x.date_oper,x.licschpkre,x.subschkre sk,t.procstavrez,t.procstavrez_19,s.setmininterestreserves min_rez, tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)) gec_gun from view_nacpogprokre_all x, odb.licschkre t,odb.srokpogprockre s where (x.lastoverduedate_ish is not null) and (x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish,x.lastoverduedate,x.lodinterest_ish),x.date_oper)) >30 and x.licschpkre=t.licschpkre and x.subschkre=t.subschkre and t.licschkre=s.licschkre and x.subschkre=s.subschkre and t.date_close is null and x.date_oper between to_date('"+txtgiris.Text+ "','dd/mm/yyyy') and to_date('" + txtcixis.Text + "','dd/mm/yyyy') order by (x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish,x.lastoverduedate,x.lodinterest_ish),x.date_oper)),x.date_oper asc) m group by m.licschpkre,m.sk,m.procstavrez,m.procstavrez_19,m.min_rez,m.gec_gun order by m.gec_gun) n where n.kol>1 or n.gec_gun>60 and n.min_rez<15 and n.procstavrez<25", con);
            isleme.Fill(muracietler);
            /////////////
    //        var maxOfThirdColumn = muracietler.AsEnumerable()
    //.Where(row => !row.IsNull("gun")) // Null değerleri filtrele
    //.Select(row =>
    //{
    //    if (decimal.TryParse(row.Field<decimal>("gun").ToString(), out decimal value)) // Değeri int'e dönüştürmeyi dene
    //    {
    //        return value; // Dönüşüm başarılıysa değeri döndür
    //    }
    //    else
    //    {
    //        return int.MinValue; // Dönüşüm başarısızsa en küçük int değeri döndür
    //    }
    //})
    //.Max();

    //        // Maksimum değer ile filtreleme yaparak uygun satırları seçme
    //        var largestRows = muracietler.AsEnumerable()
    //.Where(row => !row.IsNull("gun") && row["gun"] != DBNull.Value) // Null veya DBNull olmayanları filtrele
    //.Where(row =>
    //{
    //    // "gun" sütunundaki değeri string olarak alın
    //    string gunStringValue = row.Field<object>("gun").ToString();

    //    // String değeri int türüne dönüştürmeyi dene
    //    if (int.TryParse(gunStringValue, out int gunValue))
    //    {
    //        // Dönüşüm başarılıysa ve değer, maxOfThirdColumn ile eşitse true döndür
    //        return gunValue == maxOfThirdColumn;
    //    }
    //    else
    //    {
    //        // Dönüşüm başarısızsa false döndür
    //        return false;
    //    }
    //});

    //        // Yalnızca 1. sütundaki tekrar eden hesapları kaldırma
    //        var uniqueRows = largestRows.GroupBy(row => row.Field<string>("1")).Select(grp => grp.First());

    //        // GridControl'ü temizleme
    //        dtg_siyahi.DataSource = null;

    //        // Yeni bir DataTable oluşturma ve GridControl'e atama
    //        System.Data.DataTable newDataTable = uniqueRows.CopyToDataTable();
    //        dtg_siyahi.DataSource = newDataTable;

            // GridControl'deki sütunları özelleştirme (isteğe bağlı)
            //GridView gridView = dtg_siyahi.MainView as GridView;
            /////////////
            dtg_siyahi.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dtg_siyahi.Columns[0].HeaderText = "Suda hesabı";
            dtg_siyahi.Columns[0].Width = 250;
            dtg_siyahi.Columns[1].HeaderText = "Sub kod";
            dtg_siyahi.Columns[1].Width = 250;
            dtg_siyahi.Columns[2].HeaderText = "İlk ay";
            dtg_siyahi.Columns[2].Width = 100;
            dtg_siyahi.Columns[3].HeaderText = "İkinci ay";
            dtg_siyahi.Columns[3].Width = 100;
            dtg_siyahi.Columns[4].HeaderText = "Növbəti min 30+ gecikmə";
            dtg_siyahi.Columns[4].Width = 200;
            dtg_siyahi.Columns[5].HeaderText = "Növbəti max 30+ gecikmə";
            dtg_siyahi.Columns[5].Width = 200;
            dtg_siyahi.Columns[6].HeaderText = "Gün fərqi";
            dtg_siyahi.Columns[6].Width = 100;



            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        string hesab = "";
        string sub = "";
        public void listelemuraciet30_bidefe()
        {

            //try
            //{
            muracietler.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }


            OracleDataAdapter isleme = new OracleDataAdapter("WITH min_gecikme_her_ay AS (SELECT x.licschpkre AS hes,x.subschkre AS sk, "+
" TO_CHAR(x.date_oper, 'YYYY-MM') AS ay_il, "+
" MAX(odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper))) AS max_gec_gun " +
" FROM view_nacpogprokre_all x " +
" JOIN odb.licschkre t ON x.licschpkre = t.licschpkre AND x.subschkre = t.subschkre " +
" WHERE x.date_oper BETWEEN TO_DATE('"+txtgiris.Text+"', 'DD-MM-YYYY') AND TO_DATE('"+txtcixis.Text+"', 'DD-MM-YYYY') " +
" AND t.date_close IS NULL AND t.procstavrez = 15 " +
" GROUP BY x.licschpkre, x.subschkre, TO_CHAR(x.date_oper, 'YYYY-MM')), " +
" gecikmesi_olan_aylar AS(SELECT hes, sk, COUNT(*) AS gecikme_sayi " +
" FROM min_gecikme_her_ay WHERE max_gec_gun > 30 GROUP BY hes, sk) " +
" SELECT m.hes,m.sk,MIN(m.ay_il) AS ilk_ay_il, MAX(m.ay_il) AS son_ay_il, " +
" MAX(m.max_gec_gun) AS max_gecikme FROM " +
" min_gecikme_her_ay m LEFT JOIN gecikmesi_olan_aylar g ON m.hes = g.hes AND m.sk = g.sk " +
" WHERE(g.gecikme_sayi IS NULL OR g.gecikme_sayi < 2) " +
" GROUP BY m.hes, m.sk ORDER BY m.hes, m.sk", con);
            isleme.Fill(muracietler);
            dtg_siyahi.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dtg_siyahi.Columns[0].HeaderText = "Suda hesabı";
            dtg_siyahi.Columns[0].Width = 250;
            dtg_siyahi.Columns[1].HeaderText = "Sub kod";
            dtg_siyahi.Columns[1].Width = 250;
            dtg_siyahi.Columns[2].HeaderText = "İlk ay";
            dtg_siyahi.Columns[2].Width = 200;
            dtg_siyahi.Columns[3].HeaderText = "Növbəti ay";
            dtg_siyahi.Columns[3].Width = 200;
            dtg_siyahi.Columns[4].HeaderText = "Max gecikmə";
            dtg_siyahi.Columns[4].Width = 200;


            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        private void axtar()
        {
            dt_axtarilan.Clear();
            dt_axtarilantek.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            
            OracleDataAdapter isleme = new OracleDataAdapter("select x.date_oper,t.licschkre,t.subschkre, odb.tar_ferq360(x.date_oper, " +
               " nvl(x.lastoverduedate, x.date_oper)) gec_gun from view_nacpogprokre_all x, odb.licschkre t "+
               " where x.licschpkre = t.licschpkre and x.subschkre = t.subschkre "+
               " and x.date_oper between to_date('"+txtgiris.Text+"', 'dd/mm/yyyy') and to_date('"+txtcixis.Text+"', 'dd/mm/yyyy') and t.date_close is null "+
               " and substr(t.licschkre, 10, 6) = substr('"+hesab+"', 10, 6) and x.subschkre = '"+sub+"' "+
               " order by odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) desc", con);
            isleme.Fill(dt_axtarilantek);
            dtg_axtar.DataSource = dt_axtarilantek;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dtg_axtar.Columns[0].HeaderText = "Tarix";
            dtg_axtar.Columns[0].Width = 200;
            dtg_axtar.Columns[1].HeaderText = "Hesab";
            dtg_axtar.Columns[1].Width = 250;
            dtg_axtar.Columns[2].HeaderText = "SK";
            dtg_axtar.Columns[2].Width = 150;
            dtg_axtar.Columns[3].HeaderText = "Gecikmə gün";
            dtg_axtar.Columns[3].Width = 200;
            

        }
        private void axtar_elave()
        {
            dt_axtarilan.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            OracleDataAdapter isleme = new OracleDataAdapter("select TO_CHAR(x.date_oper, 'DD-MM-YYYY') AS tarix,LPAD(t.licschkre, 20, '0') AS licschkre_20_rəqəmli,t.subschkre, odb.tar_ferq360(x.date_oper, " +
               " nvl(x.lastoverduedate, x.date_oper)) gec_gun from view_nacpogprokre_all x, odb.licschkre t " +
               " where x.licschpkre = t.licschpkre and x.subschkre = t.subschkre " +
               " and x.date_oper between to_date('" + txtgiris.Text + "', 'dd/mm/yyyy') and to_date('" + txtcixis.Text + "', 'dd/mm/yyyy') and t.date_close is null " +
               " and substr(t.licschkre, 10, 6) = substr('" + txt_hes.Text + "', 10, 6) and x.subschkre = '" + txt_sub.Text + "' " +
               " order by odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) desc", con);
            isleme.Fill(dt_axtarilan);
            dtg_axtar.DataSource = dt_axtarilan;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dtg_axtar.Columns[0].HeaderText = "Tarix";
            dtg_axtar.Columns[0].Width = 200;
            dtg_axtar.Columns[1].HeaderText = "Hesab";
            dtg_axtar.Columns[1].Width = 250;
            dtg_axtar.Columns[2].HeaderText = "SK";
            dtg_axtar.Columns[2].Width = 150;
            dtg_axtar.Columns[3].HeaderText = "Gecikmə gün";
            dtg_axtar.Columns[3].Width = 200;


        }
        public void tamkataloq()
        {

            try
            {
                muracietler.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            //lk.date_close is null and
            //OracleDataAdapter isleme = new OracleDataAdapter("select b.date_oper tarix,b.otvet_ispoln icraci,ssd sub_debet,debet,ssk sub_kredit,kredit,summa_v_nacval mebleg,primechanie qeyd from odb.arh_dd b where b.date_oper>='01-05-2023' and b.date_oper<='11-05-2023'", con);
            //OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre, lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close, lk.date_planclose, lk.kolic_prolong, lk.date_prolong, lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok, lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08, sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18,sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth, r.adress, r.telefon, r.mobilniy,(trunc(months_between(vk.date_oper, nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper))) * 30) + extract(day from vk.date_oper) - extract(day from nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper)) as gec_gun from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr, view_nacpogprokre_all vk, creditinfo ci where substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and lk.date_oper = to_date('" + txtgiris.Text + "', 'dd/mm/yyyy') and lk.date_close is null and lk.licschppkre = vk.licschppkre and lk.subschkre = vk.subschkre and lk.licschpkre = vk.licschpkre order by  lk.hes_18", con);
            OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre," +
                " lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close," +
                " lk.date_planclose, lk.kolic_prolong, lk.date_prolong,lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok," +
                " lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga," +
                "lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08," +
                " sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18," +
                "sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth," +
                " r.adress, r.telefon, r.mobilniy from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr," +
                " creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and " +
                " lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and" +
                " lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and " +
                "lk.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and lk.date_close is null ", con);
            isleme.Fill(muracietler);
                dtg_siyahi.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            dtg_siyahi.Columns[0].HeaderText  = "Adı";
            dtg_siyahi.Columns[0].Width = 250;
            dtg_siyahi.Columns[1].HeaderText  = "Ssuda";
            dtg_siyahi.Columns[1].Width = 250;
            dtg_siyahi.Columns[2].HeaderText  = "Ssuda faiz";
            dtg_siyahi.Columns[2].Width = 250;
            dtg_siyahi.Columns[3].HeaderText  = "Ssuda vk";
            dtg_siyahi.Columns[3].Width = 200;
            dtg_siyahi.Columns[4].HeaderText  = "Ssuda vk faiz";
            dtg_siyahi.Columns[4].Width = 200;
            dtg_siyahi.Columns[5].HeaderText  = "Faiz";
            dtg_siyahi.Columns[5].Width = 100;
            dtg_siyahi.Columns[6].HeaderText  = "Ehtiyyat %";
            dtg_siyahi.Columns[6].Width = 100;
            dtg_siyahi.Columns[7].HeaderText  = "Ehtiyyat vk %";
            dtg_siyahi.Columns[7].Width = 100;
            dtg_siyahi.Columns[8].HeaderText  = "Min rez";
            dtg_siyahi.Columns[8].Width = 100;
            dtg_siyahi.Columns[9].HeaderText  = "Kredit məbləği";
            dtg_siyahi.Columns[9].Width = 100;
            dtg_siyahi.Columns[10].HeaderText  = "Qalıq";
            dtg_siyahi.Columns[10].Width = 100;
            dtg_siyahi.Columns[11].HeaderText  = "Vk qalıq";
            dtg_siyahi.Columns[11].Width = 100;
            dtg_siyahi.Columns[12].HeaderText  = "Verilmə tarixi";
            dtg_siyahi.Columns[12].Width = 100;
            dtg_siyahi.Columns[13].HeaderText  = "Bağlanma tarixi";
            dtg_siyahi.Columns[13].Width = 100;
            dtg_siyahi.Columns[14].HeaderText  = "K.p.b.tarixi";
            dtg_siyahi.Columns[14].Width = 100;
            dtg_siyahi.Columns[15].HeaderText  = "U.S.";
            dtg_siyahi.Columns[15].Width = 100;
            dtg_siyahi.Columns[16].HeaderText  = "U.T";
            dtg_siyahi.Columns[16].Width = 100;
            dtg_siyahi.Columns[17].HeaderText  = "RS";
            dtg_siyahi.Columns[17].Width = 100;
            dtg_siyahi.Columns[18].HeaderText  = "R.T";
            dtg_siyahi.Columns[18].Width = 100;
            dtg_siyahi.Columns[19].HeaderText  = "Ö.G";
            dtg_siyahi.Columns[19].Width = 100;
            dtg_siyahi.Columns[20].HeaderText  = "K.K.";
            dtg_siyahi.Columns[20].Width = 100;
            dtg_siyahi.Columns[21].HeaderText  = "K.M";
            dtg_siyahi.Columns[21].Width = 100;
            dtg_siyahi.Columns[22].HeaderText  = "G.M";
            dtg_siyahi.Columns[22].Width = 100;
            dtg_siyahi.Columns[23].HeaderText  = "Kreditin növü";
            dtg_siyahi.Columns[23].Width = 100;
            dtg_siyahi.Columns[24].HeaderText  = "Girovun növü";
            dtg_siyahi.Columns[24].Width = 200;
            dtg_siyahi.Columns[25].HeaderText  = "Girovun hesabı";
            dtg_siyahi.Columns[25].Width = 200;
            dtg_siyahi.Columns[26].HeaderText  = "Girovun məbləği";
            dtg_siyahi.Columns[26].Width = 100;
            dtg_siyahi.Columns[27].HeaderText  = "Sayı";
            dtg_siyahi.Columns[27].Width = 100;
            dtg_siyahi.Columns[28].HeaderText  = "G.Y.G";
            dtg_siyahi.Columns[28].Width = 100;
            dtg_siyahi.Columns[29].HeaderText  = "G.Y.G.T";
            dtg_siyahi.Columns[29].Width = 100;
            dtg_siyahi.Columns[30].HeaderText  = "İtem1";
            dtg_siyahi.Columns[30].Width = 100;
            dtg_siyahi.Columns[31].HeaderText  = "İtem2";
            dtg_siyahi.Columns[31].Width = 100;
            dtg_siyahi.Columns[32].HeaderText  = "İtem3";
            dtg_siyahi.Columns[32].Width = 100;
            dtg_siyahi.Columns[33].HeaderText  = "İtem4";
            dtg_siyahi.Columns[33].Width = 100;
            dtg_siyahi.Columns[34].HeaderText  = "İtem5";
            dtg_siyahi.Columns[34].Width = 100;
            dtg_siyahi.Columns[35].HeaderText  = "İtem6";
            dtg_siyahi.Columns[35].Width = 100;
            dtg_siyahi.Columns[36].HeaderText  = "İtem7";
            dtg_siyahi.Columns[36].Width = 100;
            dtg_siyahi.Columns[37].HeaderText  = "İtem8";
            dtg_siyahi.Columns[37].Width = 100;
            dtg_siyahi.Columns[38].HeaderText  = "İtem9";
            dtg_siyahi.Columns[38].Width = 100;
            dtg_siyahi.Columns[39].HeaderText  = "İtem10";
            dtg_siyahi.Columns[39].Width = 100;
            dtg_siyahi.Columns[40].HeaderText  = "İtem11";
            dtg_siyahi.Columns[40].Width = 100;
            dtg_siyahi.Columns[41].HeaderText  = "İtem12";
            dtg_siyahi.Columns[41].Width = 100;


            dtg_siyahi.Columns[42].HeaderText  = "İtem13";
            dtg_siyahi.Columns[42].Width = 100;
            dtg_siyahi.Columns[43].HeaderText  = "İtem14";
            dtg_siyahi.Columns[43].Width = 100;
            dtg_siyahi.Columns[44].HeaderText  = "İtem15";
            dtg_siyahi.Columns[44].Width = 100;
            dtg_siyahi.Columns[45].HeaderText  = "İtem16";
            dtg_siyahi.Columns[45].Width = 100;
            dtg_siyahi.Columns[46].HeaderText  = "İtem17";
            dtg_siyahi.Columns[46].Width = 100;
            dtg_siyahi.Columns[47].HeaderText  = "İtem18";
            dtg_siyahi.Columns[47].Width = 100;
            dtg_siyahi.Columns[48].HeaderText  = "İtem19";
            dtg_siyahi.Columns[48].Width = 100;
            dtg_siyahi.Columns[49].HeaderText  = "İtem20";
            dtg_siyahi.Columns[49].Width = 100;
            dtg_siyahi.Columns[50].HeaderText  = "Aylıq borc";
            dtg_siyahi.Columns[50].Width = 100;
            dtg_siyahi.Columns[51].HeaderText  = "Gəlir";
            dtg_siyahi.Columns[51].Width = 100;
            dtg_siyahi.Columns[52].HeaderText  = "FİFD";
            dtg_siyahi.Columns[52].Width = 100;
            dtg_siyahi.Columns[53].HeaderText  = "BGN";
            dtg_siyahi.Columns[53].Width = 100;
            dtg_siyahi.Columns[54].HeaderText  = "Fin";
            dtg_siyahi.Columns[54].Width = 100;
            dtg_siyahi.Columns[55].HeaderText  = "Doğum tarixi";
            dtg_siyahi.Columns[55].Width = 100;

            dtg_siyahi.Columns[56].HeaderText  = "Doğum yeri";
            dtg_siyahi.Columns[56].Width = 200;
            dtg_siyahi.Columns[57].HeaderText  = "Ünvanı";
            dtg_siyahi.Columns[57].Width = 500;
            dtg_siyahi.Columns[58].HeaderText  = "Əlaqə nömrələri";
            dtg_siyahi.Columns[58].Width = 400;
            dtg_siyahi.Columns[59].HeaderText  = "Mobil";
            dtg_siyahi.Columns[59].Width = 100;
            dtg_siyahi.Columns[60].HeaderText  = "Gecikmə günü";
            dtg_siyahi.Columns[60].Width = 100;
                //gridmuracietler.RefreshData();


            }
            catch (Exception)
            {


            }
            //finally { };


        }
        public void testetmek()
        {

            //try
            //{
                muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
            //,odb.licschkre l where substr(l.licschkre,10,6)||l.subschkre = substr(i.licsch,10,6)||i.ssls and i.passive>0 and l.date_close is not null 
            //lk.date_close is null and  trunc(mod(cast(sysdate as date) - cast(n.LASTOVERDUEDATE_ISH as date),30))  trunc(months_between(cast(sysdate as date),cast(n.LASTOVERDUEDATE_ISH as date)))*30
            //OracleDataAdapter isleme = new OracleDataAdapter("select * from odb.iavv i where i.passive>0", con);
            OracleDataAdapter isleme = new OracleDataAdapter("select i.licsch,i.ssls,i.ostatok_ish,i.vbs,l.licschkre,l.subschkre from odb.vblicsch i, odb.licschkre l where substr(l.licschkre,10,6)||l.subschkre = substr(i.licsch,10,6)||i.ssls and i.ostatok_ish>0 and l.date_close is not null and i.vbs IN (99749, 99743, 99742,99740)", con);
            //OracleDataAdapter isleme = new OracleDataAdapter("select l.licschkre,l.subschkre sk,l.tipkredita,r.name_regnom,l.procstavrez,l.procstavrez_19,sr.setmininterestreserves," +
            //    "odb.tar_ferq360(vk.date_oper, nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate," +
            //    " vk.lodinterest_ish), vk.date_oper)) gec_gun from odb.view_nacpogprokre_all vk, licschkre l,regnom r,srokpogprockre sr " +
            //    "where l.date_close is null and substr(l.licschkre,10,6)= r.regnom and l.licschpkre = vk.LICSCHPKRE " +
            //    "and l.subschkre = vk.SUBSCHKRE and substr(l.licschkre, 10, 6) = substr(vk.LICSCHPKRE, 10, 6) and l.licschkre = sr.licschkre and l.subschkre = sr.subschkre " +
            //    "and vk.date_oper = to_date('" + txtgiris.Text + "', 'dd/mm/yyyy') and l.summa_19 > 0 order by  l.licschkre,l.subschkre", con);

            //           OracleDataAdapter isleme = new OracleDataAdapter("select x.date_oper,x.licschpkre,x.subschkre sk,t.procstavrez,t.procstavrez_19,s.setmininterestreserves min_rez,t.date_restructure,round(months_between(sysdate, t.date_restructure), 0) ay, z.max_gun,odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)) as gec_gun from view_nacpogprokre_all x, odb.licschkre t, odb.srokpogprockre s,(select t.licschkre, t.subschkre sk, max(odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper))) max_gun from view_nacpogprokre_all x, odb.licschkre t where  x.licschpkre = t.licschpkre and x.subschkre = t.subschkre  and t.date_close is null and x.date_oper between to_date(t.date_restructure, 'dd/mm/yyyy') and to_date(sysdate, 'dd/mm/yyyy') group by t.licschkre, t.subschkre) z where t.date_restructure is not null and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null and x.date_oper = to_date('12/06/2023', 'dd/mm/yyyy') and t.licschkre = z.licschkre and t.subschkre = z.sk order by odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper))", con);
            //OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre, lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close, lk.date_planclose, lk.kolic_prolong, lk.date_prolong,lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok, lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08, sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18,sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth, r.adress, r.telefon, r.mobilniy from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr, creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and  lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and lk.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and lk.date_close is null ", con);
            isleme.Fill(muracietler);
                dtg_siyahi.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

                //gridmuracietler.Columns[0].Caption = "Adı";
                //gridmuracietler.Columns[0].Width = 250;
                //gridmuracietler.Columns[1].Caption = "Ssuda";
                //gridmuracietler.Columns[1].Width = 250;
                //gridmuracietler.Columns[2].Caption = "Ssuda faiz";
                //gridmuracietler.Columns[2].Width = 250;
                //gridmuracietler.Columns[3].Caption = "Ssuda vk";
                //gridmuracietler.Columns[3].Width = 200;
                //gridmuracietler.Columns[4].Caption = "Ssuda vk faiz";
                //gridmuracietler.Columns[4].Width = 200;
                //gridmuracietler.Columns[5].Caption = "Faiz";
                //gridmuracietler.Columns[5].Width = 100;
                //gridmuracietler.Columns[6].Caption = "Ehtiyyat %";
                //gridmuracietler.Columns[6].Width = 100;
                //gridmuracietler.Columns[7].Caption = "Ehtiyyat vk %";
                //gridmuracietler.Columns[7].Width = 100;
                //gridmuracietler.Columns[8].Caption = "Min rez";
                //gridmuracietler.Columns[8].Width = 100;
                //gridmuracietler.Columns[9].Caption = "Kredit məbləği";
                //gridmuracietler.Columns[9].Width = 100;
                //gridmuracietler.Columns[10].Caption = "Qalıq";
                //gridmuracietler.Columns[10].Width = 100;
                //gridmuracietler.Columns[11].Caption = "Vk qalıq";
                //gridmuracietler.Columns[11].Width = 100;
                //gridmuracietler.Columns[12].Caption = "Verilmə tarixi";
                //gridmuracietler.Columns[12].Width = 100;
                //gridmuracietler.Columns[13].Caption = "Bağlanma tarixi";
                //gridmuracietler.Columns[13].Width = 100;
                //gridmuracietler.Columns[14].Caption = "K.p.b.tarixi";
                //gridmuracietler.Columns[14].Width = 100;
                //gridmuracietler.Columns[15].Caption = "U.S.";
                //gridmuracietler.Columns[15].Width = 100;
                //gridmuracietler.Columns[16].Caption = "U.T";
                //gridmuracietler.Columns[16].Width = 100;
                //gridmuracietler.Columns[17].Caption = "RS";
                //gridmuracietler.Columns[17].Width = 100;
                //gridmuracietler.Columns[18].Caption = "R.T";
                //gridmuracietler.Columns[18].Width = 100;
                //gridmuracietler.Columns[19].Caption = "Ö.G";
                //gridmuracietler.Columns[19].Width = 100;
                //gridmuracietler.Columns[20].Caption = "K.K.";
                //gridmuracietler.Columns[20].Width = 100;
                //gridmuracietler.Columns[21].Caption = "K.M";
                //gridmuracietler.Columns[21].Width = 100;
                //gridmuracietler.Columns[22].Caption = "G.M";
                //gridmuracietler.Columns[22].Width = 100;
                //gridmuracietler.Columns[23].Caption = "Kreditin növü";
                //gridmuracietler.Columns[23].Width = 100;
                //gridmuracietler.Columns[24].Caption = "Girovun növü";
                //gridmuracietler.Columns[24].Width = 200;
                //gridmuracietler.Columns[25].Caption = "Girovun hesabı";
                //gridmuracietler.Columns[25].Width = 200;
                //gridmuracietler.Columns[26].Caption = "Girovun məbləği";
                //gridmuracietler.Columns[26].Width = 100;
                //gridmuracietler.Columns[27].Caption = "Sayı";
                //gridmuracietler.Columns[27].Width = 100;
                //gridmuracietler.Columns[28].Caption = "G.Y.G";
                //gridmuracietler.Columns[28].Width = 100;
                //gridmuracietler.Columns[29].Caption = "G.Y.G.T";
                //gridmuracietler.Columns[29].Width = 100;
                //gridmuracietler.Columns[30].Caption = "İtem1";
                //gridmuracietler.Columns[30].Width = 100;
                //gridmuracietler.Columns[31].Caption = "İtem2";
                //gridmuracietler.Columns[31].Width = 100;
                //gridmuracietler.Columns[32].Caption = "İtem3";
                //gridmuracietler.Columns[32].Width = 100;
                //gridmuracietler.Columns[33].Caption = "İtem4";
                //gridmuracietler.Columns[33].Width = 100;
                //gridmuracietler.Columns[34].Caption = "İtem5";
                //gridmuracietler.Columns[34].Width = 100;
                //gridmuracietler.Columns[35].Caption = "İtem6";
                //gridmuracietler.Columns[35].Width = 100;
                //gridmuracietler.Columns[36].Caption = "İtem7";
                //gridmuracietler.Columns[36].Width = 100;
                //gridmuracietler.Columns[37].Caption = "İtem8";
                //gridmuracietler.Columns[37].Width = 100;
                //gridmuracietler.Columns[38].Caption = "İtem9";
                //gridmuracietler.Columns[38].Width = 100;
                //gridmuracietler.Columns[39].Caption = "İtem10";
                //gridmuracietler.Columns[39].Width = 100;
                //gridmuracietler.Columns[40].Caption = "İtem11";
                //gridmuracietler.Columns[40].Width = 100;
                //gridmuracietler.Columns[41].Caption = "İtem12";
                //gridmuracietler.Columns[41].Width = 100;


                //gridmuracietler.Columns[42].Caption = "İtem13";
                //gridmuracietler.Columns[42].Width = 100;
                //gridmuracietler.Columns[43].Caption = "İtem14";
                //gridmuracietler.Columns[43].Width = 100;
                //gridmuracietler.Columns[44].Caption = "İtem15";
                //gridmuracietler.Columns[44].Width = 100;
                //gridmuracietler.Columns[45].Caption = "İtem16";
                //gridmuracietler.Columns[45].Width = 100;
                //gridmuracietler.Columns[46].Caption = "İtem17";
                //gridmuracietler.Columns[46].Width = 100;
                //gridmuracietler.Columns[47].Caption = "İtem18";
                //gridmuracietler.Columns[47].Width = 100;
                //gridmuracietler.Columns[48].Caption = "İtem19";
                //gridmuracietler.Columns[48].Width = 100;
                //gridmuracietler.Columns[49].Caption = "İtem20";
                //gridmuracietler.Columns[49].Width = 100;
                //gridmuracietler.Columns[50].Caption = "Aylıq borc";
                //gridmuracietler.Columns[50].Width = 100;
                //gridmuracietler.Columns[51].Caption = "Gəlir";
                //gridmuracietler.Columns[51].Width = 100;
                //gridmuracietler.Columns[52].Caption = "FİFD";
                //gridmuracietler.Columns[52].Width = 100;
                //gridmuracietler.Columns[53].Caption = "BGN";
                //gridmuracietler.Columns[53].Width = 100;
                //gridmuracietler.Columns[54].Caption = "Fin";
                //gridmuracietler.Columns[54].Width = 100;
                //gridmuracietler.Columns[55].Caption = "Doğum tarixi";
                //gridmuracietler.Columns[55].Width = 100;

                //gridmuracietler.Columns[56].Caption = "Doğum yeri";
                //gridmuracietler.Columns[56].Width = 200;
                //gridmuracietler.Columns[57].Caption = "Ünvanı";
                //gridmuracietler.Columns[57].Width = 500;
                //gridmuracietler.Columns[58].Caption = "Əlaqə nömrələri";
                //gridmuracietler.Columns[58].Width = 400;
                //gridmuracietler.Columns[59].Caption = "Mobil";
                //gridmuracietler.Columns[59].Width = 100;
                //gridmuracietler.Columns[60].Caption = "Gecikmə günü";
                //gridmuracietler.Columns[60].Width = 100;



            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        public void kr_bagli_olub_bk_qal()
        {
            try
            {
            muracietler.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            OracleDataAdapter isleme = new OracleDataAdapter("select i.licsch,i.ssls,i.ostatok_ish,i.vbs,l.licschkre," +
                "l.subschkre from odb.vblicsch i, odb.licschkre l where substr(l.licschkre,10,6)||l.subschkre = substr(i.licsch,10,6)||i.ssls " +
                "and i.ostatok_ish>0 and l.date_close is not null and i.vbs IN (99749, 99743, 99742,99740)", con);
            isleme.Fill(muracietler);
            dtg_siyahi.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }
                dtg_siyahi.Columns[0].HeaderText  = "BK hesab";
                dtg_siyahi.Columns[0].Width = 250;
                dtg_siyahi.Columns[1].HeaderText  = "BK ssuda";
                dtg_siyahi.Columns[1].Width = 100;
                dtg_siyahi.Columns[2].HeaderText  = "BK məbləğ";
                dtg_siyahi.Columns[2].Width = 200;
                dtg_siyahi.Columns[3].HeaderText  = "BK";
                dtg_siyahi.Columns[3].Width = 150;
                dtg_siyahi.Columns[4].HeaderText  = "Kataloq hesab";
                dtg_siyahi.Columns[4].Width = 250;
                dtg_siyahi.Columns[5].HeaderText  = "Kataloq ssuda";
                dtg_siyahi.Columns[5].Width = 100;
            }
            catch (Exception)
            {
            }
            finally { };
        }
        public void testlerucun()
        {
            string[] ehtiyatlar = {
                "15910000000001001001", "15910000000001001100", "15910000000001002001",
                "15910000000001002100", "15910000000001100001", "15910000000001102001",
                "15910000000001102100", "15910000000001200001", "15910000000001200030",
                "15910000000001200060", "15910000000001200100", "15910000000001202100",
                "15910000000001301001", "15910000000001301100", "15910000000001302001",
                "15910000000001302100", "15910000000001400001", "15910000000001402001",
                "15910000000001402100", "15910000000001500001", "15910000000001500030",
                "15910000000001500060", "15910000000001500100", "15910000000001502100",
                "15911000000001000001", "15911000000001001001", "15911000000001003001",
                "15911000000001002001", "15911000000001004001", "15911000000001002030",
                "15911000000001004030", "15911000000001005030", "15911000000001005001",
                "15911000000001105001", "15911000000001102001", "15911000000001301001",
                "15911000000001302001", "15911000000001402001", "15911000000001405001",
                "15918000000001100001", "15918000000001400001", "15918000000001500001",
                "15918000003001600001", "20910000000001000001", "20910000000001100001",
                "20910000000001100005", "20910000001001100005", "20910000003001100001",
                "20910000003001100005", "20910000000001100030", "20910000000001100060",
                "20910000000001100100", "20910000003001100100", "20910000003001100030",
                "20910000000001200001", "20910000000001200005", "20910000000001200030",
                "20910000003001200001", "20910000003001200030", "20910000000001200060",
                "20910000000001200100", "20910000003001200100", "20910000000001201100",
                "20910000000001400001", "20910000000001400005", "20910000001001400005",
                "20910000003001200005", "20910000003001400001", "20910000005001100001",
                "20910000005001200001", "20910000005001400001", "20910000005001500001",
                "20910000003001400005", "20910000000001400030", "20910000003001200060",
                "20910000003001400030", "20910000003001400100", "20910000000001400100",
                "20910000000001500001", "20910000000001500005", "20910000000001500030",
                "20910000003001500001", "20910000003001500030", "20910000000001500100",
                "20910000003001500100", "20910000005001100002", "20910000005001100025",
                "20910000005001100050", "20910000005001200002", "20910000005001100100",
                "20910000005001200100", "20910000005001400002", "20910000005001400025",
                "20910000005001400050", "20910000005001500002", "20910000005001400100",
                "20910000005001500100", "20910000000001501100", "20910000003001600001",
                "21910000000001000001", "21910000005001100001", "21910000000001100001",
                "21910000000001100002", "21910000000001100005", "21910000000001100030",
                "21910000000001100060", "21910000000001100100", "21910000000001101001",
                "21910000000001100015", "21910000005001200001", "21910000000001200001",
                "21910000000001200002", "21910000000001200005", "21910000000001200015",
                "21910000000001200030", "21910000000001200060", "21910000000001200100",
                "21910000000001201001", "21910000000001201100", "21910000005001400001",
                "21910000000001400001", "21910000000001400002", "21910000000001400005",
                "21910000000001400015", "21910000000001400030", "21910000000001400060",
                "21910000000001400100", "21910000000001401001", "21910000000001401100",
                "21910000005001500001", "21910000000001500001", "21910000000001500002",
                "21910000000001500005", "21910000000001500015", "21910000000001500030",
                "21910000000001500060", "21910000000001500100", "21910000000001501001",
                "21910000000001501100", "21910000000001600001", "21910000000001600005",
                "21910000000001600030", "21910000000001600060", "21910000000001600100",
                "21910000003001600001", "21910000004001600025", "21911000000001000001",
                "21911000000001100001", "21911000000001100002", "21911000000001100005",
                "21911000000001100030", "21911000000001100060", "21911000000001100100",
                "21911000000001101001", "21911000000001101002", "21911000000001101005",
                "21911000000001101030", "21911000000001101060", "21911000000001101100",
                "21911000000001102001", "21911000000001102002", "21911000000001102005",
                "21911000000001102010", "21911000000001102030", "21911000000001104002",
                "21911000000001200001", "21911000000001200002", "21911000000001200005",
                "21911000000001200030", "21911000000001200060", "21911000000001200100",
                "21911000000001201001", "21911000000001201002", "21911000000001201005",
                "21911000000001201030", "21911000000001201060", "21911000000001201100",
                "21911000000001202100", "21911000000001202002", "21911000000001204002",
                "21911000000001400001", "21911000000001400002", "21911000000001400005",
                "21911000000001400030", "21911000000001400060", "21911000000001400100",
                "21911000000001401001", "21911000000001401002", "21911000000001401005",
                "21911000000001401030", "21911000000001401100", "21911000000001402001",
                "21911000000001402002", "21911000000001402005", "21911000000001402010",
                "21911000000001402030", "21911000000001404002", "21911000000001500001",
                "21911000000001500002", "21911000000001500005", "21911000000001500030",
                "21911000000001500060", "21911000000001500100", "21911000000001501001",
                "21911000000001501002", "21911000000001501005", "21911000000001501030",
                "21911000000001501100", "21911000000001502002", "21911000000001502100",
                "21911000000001504002", "21911000000001600001", "21911000000001600030",
                "21911000003001600001", "21911000004001600025", "21911000000001502010",
                "21912000000001000001", "21912000000001100001", "21912000000001100002",
                "21912000000001100005", "21912000000001100030", "21912000000001100060",
                "21912000000001100100", "21912000000001101005", "21912000000001200001",
                "21912000000001200005", "21912000000001200030", "21912000000001200060",
                "21912000000001200100", "21912000000001200105", "21912000000001201005",
                "21912000000001400001",
                "21912000000001400002", "21912000000001400005", "21912000000001400030",
                "21912000000001400100", "21912000000001401005", "21912000000001500001",
                "21912000000001500005", "21912000000001500030", "21912000000001500100",
                "21912000000001501005", "21913000000001100001", "21913000000001101100",
                "21913000000001201100", "21913000000001400001", "21913000000001401100",
                "23910000000001201100", "23910000000001501100", "23911000000001400001",
                "23911000000009300001", "27013000000001000006", "23911000000001100100"
            };
            string[] setir_7 = {
                 "27010", "27011", "27012", "27013",
                 "28020", "28021", "28030", "28031",
                 "28040", "28041", "28050", "28051",
                 "28060", "28061"};
            string[] setir_9 = {
                "13012", "13032", "14012", "14082", "15022", "15022", "15027", "15027", "15028",
                "15212", "15214", "15770", "15222", "15222", "15224", "15224", "15227", "15227",
                "15622", "15624", "15770", "20362", "20392", "20394", "20632", "20634", "20652",
                "20654", "20662", "20664", "20682", "20764", "21072", "21074", "21077", "21079",
                "21112", "21114", "21117", "21119", "21127", "21127", "21129", "21129", "21142",
                "21144", "21212", "21214", "21217", "21219", "21222", "21224", "21227", "21227",
                "21227", "21229", "21229", "21229", "21242", "21244", "21247", "21252", "21254",
                "21257", "21259", "23124", "23302", "24010", "25010", "25011", "25019", "25020",
                "25020", "25021", "25021", "25021", "25021", "25029", "25029", "25052", "25059",
                "25069", "25069", "25069", "25079", "25089", "25089", "25100", "25101", "25103",
                "25109", "25110", "25120", "25121", "25122", "25123", "25129", "25139", "25159",
                "25270", "25280", "25280", "28110", "28111", "28120", "28121", "28130", "28131",
                "15910000000001301100", "15910000000001302100", "15910000000001402100",
                "15910000000001500030", "15910000000001500060", "15910000000001500100",
                "15910000000001502100", "20910000000001400030", "20910000003001200060",
                "20910000003001400030", "20910000003001400100", "20910000000001400100",
                "20910000000001500030", "20910000003001500030", "20910000000001500100",
                "20910000003001500100", "20910000005001100002", "20910000005001100025",
                "20910000005001100050", "20910000005001200002", "20910000005001100100",
                "20910000005001200100", "20910000005001400002", "20910000005001400025",
                "20910000005001400050", "20910000005001500002", "20910000005001400100",
                "20910000005001500100", "21910000000001400030", "21910000000001400060",
                "21910000000001400100", "21910000000001401100", "21910000000001500030",
                "21910000000001500060", "21910000000001500100", "21910000000001501001",
                "21910000000001501100", "21910000000001600001", "21910000000001600005",
                "21911000000001400030", "21911000000001400060", "21911000000001400100",
                "21911000000001401030", "21911000000001401100", "21911000000001402030",
                "21911000000001500030", "21911000000001500060", "21911000000001500100",
                "21911000000001501030", "21911000000001501100", "21911000000001502100",
                "21912000000001400030", "21912000000001400100", "21912000000001500030",
                "21912000000001500100", "21913000000001401100", "21913000000001501100",
                "23910000000001501100", "23911000000001100100"};
            string[] setir_9_a = {
                "13012", "13032", "14012", "14082", "15022", "15022", "15027", "15027", "15028", "15212",
                "15214", "15770", "15222", "15222", "15224", "15224", "15227", "15227", "15622", "15624",
                "20392", "20394", "20632", "20634", "20652", "20654", "20662", "20664", "20682", "20764",
                "21072", "21074", "21077", "21079", "21112", "21114", "21117", "21119", "21127", "21127",
                "21129", "21129", "21142", "21144", "21212", "21214", "21217", "21219", "21222", "21224",
                "21227", "21227", "21227", "21229", "21229", "21229", "21242", "21244", "21247", "21252",
                "21254", "21257", "21259", "23124", "23302", "25069"
            };
            string[] setir_3 = { "13010", "13030", "14010", "14014", "14015", "14030", "14034", "14080", "14913" };
            string[] setir_15_a = { "35772", "35227", "41122", "45029", "45029" };
            string[] setir_16_uzun = {"15910000000001001001",
                "15910000000001002001", "15910000000001100001", "15910000000001102001",
                "15910000000001200001", "15910000000001301001", "15910000000001302001",
                "15910000000001400001", "15910000000001402001", "15910000000001500001",
                "15911000000001000001", "15911000000001001001", "15911000000001003001",
                "15911000000001002001", "15911000000001004001", "15911000000001005001",
                "15911000000001105001", "15911000000001102001", "15911000000001301001",
                "15911000000001302001", "15911000000001402001", "15911000000001405001",
                "15918000000001100001", "15918000000001400001", "15918000000001500001",
                "15918000003001600001", "20910000000001000001", "20910000000001100001",
                "20910000000001100005", "20910000001001100005", "20910000003001100001",
                "20910000003001100005", "20910000000001200001", "20910000000001200005",
                "20910000000001400001", "20910000000001400005", "20910000001001400005",
                "20910000003001200005", "20910000003001400001", "20910000005001100001",
                "20910000005001200001", "20910000005001400001", "20910000005001500001",
                "20910000003001400005", "20910000000001500001", "20910000000001500005",
                "20910000003001500001", "20910000003001600001", "21910000000001000001",
                "21910000005001100001", "21910000000001100001", "21910000000001100002",
                "21910000000001100005", "21910000000001101001", "21910000000001100015",
                "21910000005001200001", "21910000000001200001", "21910000000001200002",
                "21910000000001200005", "21910000000001200015", "21910000005001400001",
                "21910000000001400001", "21910000000001400002", "21910000000001400005",
                "21910000000001400015", "21910000000001401001", "21910000005001500001",
                "21910000000001500001", "21910000000001500002", "21910000000001500005",
                "21910000000001500015", "21910000003001600001", "21911000000001000001",
                "21911000000001100001", "21911000000001100002", "21911000000001100005",
                "21911000000001101001", "21911000000001101002", "21911000000001101005",
                "21911000000001102001", "21911000000001102002", "21911000000001102005",
                "21911000000001102010", "21911000000001104002", "21911000000001200001",
                "21911000000001200002", "21911000000001200005", "21911000000001201001",
                "21911000000001201002", "21911000000001400001", "21911000000001400002",
                "21911000000001400005", "21911000000001401001", "21911000000001401002",
                "21911000000001401005", "21911000000001402001", "21911000000001402002",
                "21911000000001402005", "21911000000001402010", "21911000000001404002",
                "21911000000001500001", "21911000000001500002", "21911000000001500005",
                "21911000000001501001", "21911000000001501002", "21911000000001501005",
                "21911000000001502002", "21911000000001504002", "21911000000001600001",
                "21911000003001600001", "21911000000001502010", "21912000000001000001",
                "21912000000001100001", "21912000000001100002", "21912000000001100005",
                "21912000000001101005","21912000000001200001","21912000000001200005","21912000000001200105","21912000000001201005",
                "21912000000001400001","21912000000001400002","21912000000001400005","21912000000001401005",
                "21912000000001500001","21912000000001500005","","21912000000001501005","21913000000001100001",
                "21913000000001400001","23911000000001400001","23911000000009300001" };
            string[] setir_16 = { "50020", "50060", "50110", "50120", "50130", "50130" };
            string[] setir_10_b = {
                "41010", "41011", "41015", "41016", "41020", "41020", "41021", "41021", "41025",
                "41025", "41025", "41026", "41026", "41040", "41045", "41050", "41050", "41055",
                "41055", "41930", "41931", "41932", "41933", "41940", "41940", "41941", "41941",
                "41942", "41942", "41943"};
            //try
            //{
            muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                OracleDataAdapter isleme = new OracleDataAdapter("select distinct n.licschpkre,n.val,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez,n.gec_gun,n.meb,n.faiz,n.mud,'' ay,asz.odenis,n.tip from "+
      "(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, substr(m.licschpkre, 6, 2) val, m.meb, m.faiz, m.mud, m.tip, count(*) kol from " +
      "(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19, s.setmininterestreserves min_rez, " +
      "odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, t.summakre meb, t.procstavkre faiz, t.srok mud, t.tipkredita tip " +
      "from view_nacpogprokre_all x, odb.arh_licschkre t, odb.srokpogprockre s, tipkre g " +
      "where t.tipkredita = g.code and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null " +
      "and x.date_oper = to_date('30/11/2023', 'dd/mm/yyyy') and x.date_oper = t.date_oper " +
      "order by(x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)), x.date_oper asc) m " +
      "group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, m.meb, m.faiz, m.mud, m.tip " +
      "order by m.gec_gun) n, (SELECT distinct substr(ar.kredit, 10, 6) qeyd, ar.ssk, sum(ar.summa_v_nacval) odenis " +
      "                    FROM regnom rr, licschkre l " +
      "                    JOIN arh_dd ar ON EXTRACT(MONTH FROM ar.date_oper) = EXTRACT(MONTH FROM TO_DATE('30-11-2023', 'DD-MM-YYYY')) " +
      "                    AND EXTRACT(YEAR FROM ar.date_oper) = EXTRACT(YEAR FROM TO_DATE('30-11-2023', 'DD-MM-YYYY')) " +
      "                    WHERE substr(l.licschkre, 10, 6) = rr.regnom and(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
      "                    AND ar.ssk = l.subschkre AND " +
      "                    ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
      "                     group by substr(ar.kredit, 10, 6),ar.ssk) asz where substr(n.licschpkre, 10, 6) = asz.qeyd(+)and n.sk = asz.ssk(+)", con);
                isleme.Fill(muracietler);
                dtg_siyahi.DataSource = muracietler;
               // muracietler.Rows.Add(DBNull.Value);
            foreach (DataRow row in muracietler.Rows)
            {
                double meb = Convert.ToDouble(row[7]);
                double faizOranı = Convert.ToDouble(row[8]);
                int vadeMüddeti = Convert.ToInt32(row[9])/30;

                double aylıkÖdeme =Math.Round( CalculateMonthlyPayment(meb, vadeMüddeti, faizOranı ),2);
                row[10] = aylıkÖdeme;
            }
            if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                //gridmuracietler.Columns[0].Caption = "BK hesab";
                //gridmuracietler.Columns[0].Width = 250;
                //gridmuracietler.Columns[1].Caption = "BK ssuda";
                //gridmuracietler.Columns[1].Width = 100;
                //gridmuracietler.Columns[2].Caption = "BK məbləğ";
                //gridmuracietler.Columns[2].Width = 200;
                //gridmuracietler.Columns[3].Caption = "BK";
                //gridmuracietler.Columns[3].Width = 150;
                //gridmuracietler.Columns[4].Caption = "Kataloq hesab";
                //gridmuracietler.Columns[4].Width = 250;
                //gridmuracietler.Columns[5].Caption = "Kataloq ssuda";
                //gridmuracietler.Columns[5].Width = 100;
            //}
            //catch (Exception)
            //{
            //}
            //finally { };
        }
        public void restruk()
        {

            //try
            //{
            muracietler.Clear();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            int iltarix = Convert.ToInt32(tarixIl);
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            //lk.date_close is null and  trunc(mod(cast(sysdate as date) - cast(n.LASTOVERDUEDATE_ISH as date),30))  trunc(months_between(cast(sysdate as date),cast(n.LASTOVERDUEDATE_ISH as date)))*30
            //OracleDataAdapter isleme = new OracleDataAdapter("select * from odb.view_nacpogprokre_all where substr(LICSCHPKRE,10,6)='012266'", con);
            //OracleDataAdapter isleme = new OracleDataAdapter("select l.licschkre,l.subschkre sk,l.tipkredita,r.name_regnom,l.procstavrez,l.procstavrez_19,sr.setmininterestreserves," +
            //    "odb.tar_ferq360(vk.date_oper, nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate," +
            //    " vk.lodinterest_ish), vk.date_oper)) gec_gun from odb.view_nacpogprokre_all vk, licschkre l,regnom r,srokpogprockre sr " +
            //    "where l.date_close is null and substr(l.licschkre,10,6)= r.regnom and l.licschpkre = vk.LICSCHPKRE " +
            //    "and l.subschkre = vk.SUBSCHKRE and substr(l.licschkre, 10, 6) = substr(vk.LICSCHPKRE, 10, 6) and l.licschkre = sr.licschkre and l.subschkre = sr.subschkre " +
            //    "and vk.date_oper = to_date('" + txtgiris.Text + "', 'dd/mm/yyyy') and l.summa_19 > 0 order by  l.licschkre,l.subschkre", con);

            OracleDataAdapter isleme = new OracleDataAdapter("select x.date_oper,x.licschpkre,x.subschkre sk,t.procstavrez,t.procstavrez_19,s.setmininterestreserves min_rez,t.date_restructure,round(months_between(sysdate, t.date_restructure), 0) ay, z.max_gun,odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)) as gec_gun from view_nacpogprokre_all x, odb.licschkre t, odb.srokpogprockre s,(select t.licschkre, t.subschkre sk, max(odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper))) max_gun from view_nacpogprokre_all x, odb.licschkre t where  x.licschpkre = t.licschpkre and x.subschkre = t.subschkre  and t.date_close is null and x.date_oper between to_date(t.date_restructure, 'dd/mm/yyyy') and to_date(sysdate, 'dd/mm/yyyy') group by t.licschkre, t.subschkre) z where t.date_restructure is not null and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null and x.date_oper = to_date('12/06/2023', 'dd/mm/yyyy') and t.licschkre = z.licschkre and t.subschkre = z.sk order by odb.tar_ferq360(x.date_oper, nvl(odb.func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper))", con);
            //OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre, lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close, lk.date_planclose, lk.kolic_prolong, lk.date_prolong,lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok, lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08, sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18,sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth, r.adress, r.telefon, r.mobilniy from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr, creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and  lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and lk.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and lk.date_close is null ", con);
            isleme.Fill(muracietler);
            dtg_siyahi.DataSource = muracietler;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            //gridmuracietler.Columns[0].Caption = "Adı";
            //gridmuracietler.Columns[0].Width = 250;
            //gridmuracietler.Columns[1].Caption = "Ssuda";
            //gridmuracietler.Columns[1].Width = 250;
            //gridmuracietler.Columns[2].Caption = "Ssuda faiz";
            //gridmuracietler.Columns[2].Width = 250;
            //gridmuracietler.Columns[3].Caption = "Ssuda vk";
            //gridmuracietler.Columns[3].Width = 200;
            //gridmuracietler.Columns[4].Caption = "Ssuda vk faiz";
            //gridmuracietler.Columns[4].Width = 200;
            //gridmuracietler.Columns[5].Caption = "Faiz";
            //gridmuracietler.Columns[5].Width = 100;
            //gridmuracietler.Columns[6].Caption = "Ehtiyyat %";
            //gridmuracietler.Columns[6].Width = 100;
            //gridmuracietler.Columns[7].Caption = "Ehtiyyat vk %";
            //gridmuracietler.Columns[7].Width = 100;
            //gridmuracietler.Columns[8].Caption = "Min rez";
            //gridmuracietler.Columns[8].Width = 100;
            //gridmuracietler.Columns[9].Caption = "Kredit məbləği";
            //gridmuracietler.Columns[9].Width = 100;
            //gridmuracietler.Columns[10].Caption = "Qalıq";
            //gridmuracietler.Columns[10].Width = 100;
            //gridmuracietler.Columns[11].Caption = "Vk qalıq";
            //gridmuracietler.Columns[11].Width = 100;
            //gridmuracietler.Columns[12].Caption = "Verilmə tarixi";
            //gridmuracietler.Columns[12].Width = 100;
            //gridmuracietler.Columns[13].Caption = "Bağlanma tarixi";
            //gridmuracietler.Columns[13].Width = 100;
            //gridmuracietler.Columns[14].Caption = "K.p.b.tarixi";
            //gridmuracietler.Columns[14].Width = 100;
            //gridmuracietler.Columns[15].Caption = "U.S.";
            //gridmuracietler.Columns[15].Width = 100;
            //gridmuracietler.Columns[16].Caption = "U.T";
            //gridmuracietler.Columns[16].Width = 100;
            //gridmuracietler.Columns[17].Caption = "RS";
            //gridmuracietler.Columns[17].Width = 100;
            //gridmuracietler.Columns[18].Caption = "R.T";
            //gridmuracietler.Columns[18].Width = 100;
            //gridmuracietler.Columns[19].Caption = "Ö.G";
            //gridmuracietler.Columns[19].Width = 100;
            //gridmuracietler.Columns[20].Caption = "K.K.";
            //gridmuracietler.Columns[20].Width = 100;
            //gridmuracietler.Columns[21].Caption = "K.M";
            //gridmuracietler.Columns[21].Width = 100;
            //gridmuracietler.Columns[22].Caption = "G.M";
            //gridmuracietler.Columns[22].Width = 100;
            //gridmuracietler.Columns[23].Caption = "Kreditin növü";
            //gridmuracietler.Columns[23].Width = 100;
            //gridmuracietler.Columns[24].Caption = "Girovun növü";
            //gridmuracietler.Columns[24].Width = 200;
            //gridmuracietler.Columns[25].Caption = "Girovun hesabı";
            //gridmuracietler.Columns[25].Width = 200;
            //gridmuracietler.Columns[26].Caption = "Girovun məbləği";
            //gridmuracietler.Columns[26].Width = 100;
            //gridmuracietler.Columns[27].Caption = "Sayı";
            //gridmuracietler.Columns[27].Width = 100;
            //gridmuracietler.Columns[28].Caption = "G.Y.G";
            //gridmuracietler.Columns[28].Width = 100;
            //gridmuracietler.Columns[29].Caption = "G.Y.G.T";
            //gridmuracietler.Columns[29].Width = 100;
            //gridmuracietler.Columns[30].Caption = "İtem1";
            //gridmuracietler.Columns[30].Width = 100;
            //gridmuracietler.Columns[31].Caption = "İtem2";
            //gridmuracietler.Columns[31].Width = 100;
            //gridmuracietler.Columns[32].Caption = "İtem3";
            //gridmuracietler.Columns[32].Width = 100;
            //gridmuracietler.Columns[33].Caption = "İtem4";
            //gridmuracietler.Columns[33].Width = 100;
            //gridmuracietler.Columns[34].Caption = "İtem5";
            //gridmuracietler.Columns[34].Width = 100;
            //gridmuracietler.Columns[35].Caption = "İtem6";
            //gridmuracietler.Columns[35].Width = 100;
            //gridmuracietler.Columns[36].Caption = "İtem7";
            //gridmuracietler.Columns[36].Width = 100;
            //gridmuracietler.Columns[37].Caption = "İtem8";
            //gridmuracietler.Columns[37].Width = 100;
            //gridmuracietler.Columns[38].Caption = "İtem9";
            //gridmuracietler.Columns[38].Width = 100;
            //gridmuracietler.Columns[39].Caption = "İtem10";
            //gridmuracietler.Columns[39].Width = 100;
            //gridmuracietler.Columns[40].Caption = "İtem11";
            //gridmuracietler.Columns[40].Width = 100;
            //gridmuracietler.Columns[41].Caption = "İtem12";
            //gridmuracietler.Columns[41].Width = 100;


            //gridmuracietler.Columns[42].Caption = "İtem13";
            //gridmuracietler.Columns[42].Width = 100;
            //gridmuracietler.Columns[43].Caption = "İtem14";
            //gridmuracietler.Columns[43].Width = 100;
            //gridmuracietler.Columns[44].Caption = "İtem15";
            //gridmuracietler.Columns[44].Width = 100;
            //gridmuracietler.Columns[45].Caption = "İtem16";
            //gridmuracietler.Columns[45].Width = 100;
            //gridmuracietler.Columns[46].Caption = "İtem17";
            //gridmuracietler.Columns[46].Width = 100;
            //gridmuracietler.Columns[47].Caption = "İtem18";
            //gridmuracietler.Columns[47].Width = 100;
            //gridmuracietler.Columns[48].Caption = "İtem19";
            //gridmuracietler.Columns[48].Width = 100;
            //gridmuracietler.Columns[49].Caption = "İtem20";
            //gridmuracietler.Columns[49].Width = 100;
            //gridmuracietler.Columns[50].Caption = "Aylıq borc";
            //gridmuracietler.Columns[50].Width = 100;
            //gridmuracietler.Columns[51].Caption = "Gəlir";
            //gridmuracietler.Columns[51].Width = 100;
            //gridmuracietler.Columns[52].Caption = "FİFD";
            //gridmuracietler.Columns[52].Width = 100;
            //gridmuracietler.Columns[53].Caption = "BGN";
            //gridmuracietler.Columns[53].Width = 100;
            //gridmuracietler.Columns[54].Caption = "Fin";
            //gridmuracietler.Columns[54].Width = 100;
            //gridmuracietler.Columns[55].Caption = "Doğum tarixi";
            //gridmuracietler.Columns[55].Width = 100;

            //gridmuracietler.Columns[56].Caption = "Doğum yeri";
            //gridmuracietler.Columns[56].Width = 200;
            //gridmuracietler.Columns[57].Caption = "Ünvanı";
            //gridmuracietler.Columns[57].Width = 500;
            //gridmuracietler.Columns[58].Caption = "Əlaqə nömrələri";
            //gridmuracietler.Columns[58].Width = 400;
            //gridmuracietler.Columns[59].Caption = "Mobil";
            //gridmuracietler.Columns[59].Width = 100;
            //gridmuracietler.Columns[60].Caption = "Gecikmə günü";
            //gridmuracietler.Columns[60].Width = 100;



            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        public void ehtiyatlarabaxis()
        {

            try
            {
                muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }

                //lk.date_close is null and
                OracleDataAdapter isleme = new OracleDataAdapter("select odb.func_utf8_to_latin(kl.ad) ad,t.licschkre hes_18, kl.fz, kl.fz_19,kl.procstavrez rez_fz,kl.procstavrez_19 rez_fz_19,kl.pul, kl.ver_kred,kl.date_prolong,kl.kolic_prolong,kl.kolic_restructure,kl.date_restructure, kl.summa_pereocen_zaloga, kl.data_pereocen_zaloga, pr.aylig_gelir, pr.aylig_borc_yuku, pr.bgn, kl.hes_18 kredhes, kl.sk, kl.meb_18 esasborc, kl.hes_19 vkhes, kl.meb_19 vkborc, kl.hes_24 fzhes, kl.meb_24 fzborc, kl.fzborcqaz fzqaz, kl.fzborcoden fzoden, kl.vkfzhes, kl.meb19_24 vkfzborc, kl.vk_fz_qaz, kl.vk_fz_oden, kl.gec_gun, round(kl.gec_gun / 30, 0) gec_ay, cer.cerime cer_meb_um, round(cer.cerime / kl.fz_19 * (kl.fz_19 - kl.fz), 2) cer_meb, g1.gr_mab, (kl.ver_kred - kl.meb_18 - meb_19) fakt_od_mab, (kl.ver_kred - kl.meb_18 - meb_19) - g1.gr_mab Vaxt_ev_od, kl.ver_tar, kl.son_tar, kl.zal_mab, (kl.meb_18 + kl.meb_24 + kl.meb_19 + kl.meb19_24) cem_meb, (kl.meb_18 + kl.meb_19) * kl.procstavrez / 100 ef, (kl.meb_24 + kl.meb19_24) * kl.procstavrez_19 / 100 ef_fz from(select r.name_regnom ad, t.procstavkre fz, t.procstav_19 fz_19, t.procstavrez, t.procstavrez_19, t.summa_zaloga zal_mab, t.date_open ver_tar, t.date_planclose son_tar, t.summakre ver_kred, t.date_prolong, t.kolic_prolong, t.kolic_restructure, t.date_restructure, t.summa_pereocen_zaloga, t.data_pereocen_zaloga,  t.subschkre sk, t.summa meb_18, t.licsch_19 hes_19, t.summa_19 meb_19, t.licschppkre vkfzhes, substr(t.licschkre, 6, 2) pul, t.licschpkre hes_24, vk.nacpro_ish fzborcqaz, vk.pogpro_ish fzborcoden, NVL((vk.nacpro_ish - vk.pogpro_ish), 0) meb_24, NVL((vk.nacprospro_ish - vk.pogprospro_ish), 0) meb19_24, vk.nacprospro_ish vk_fz_qaz, vk.pogprospro_ish vk_fz_oden, vk.licschppkre, (trunc(months_between(vk.date_oper, nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper))) * 30) + extract(day from vk.date_oper) - extract(day from nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper)) as gec_gun from odb.arh_licschkre t, odb.balschkre b, odb.arh_saldo_ls l, odb.regnom r, odb.view_nacpogprokre_all vk where t.date_close is null and substr(t.licschkre, 1, 5) = b.balsch and t.licschppkre = vk.licschppkre and t.subschkre = vk.subschkre and t.licschpkre = vk.licschpkre and t.licschpkre = l.licsch and substr(t.licschkre, 10, 6) = r.regnom and t.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and l.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and vk.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy')) kl, odb.srokpogprockre pr, (select z.debet, z.ssd sk, sum(z.summa_v_nacval) cerime from odb.arh_dd z, odb.licschkre kr  where z.debet = kr.licschppkre and z.kredit = kr.trlicsch_19 and z.ssd = kr.subschkre group by z.debet, z.ssd) cer,  (select gr.licschkre,gr.subschkre sk, sum(gr.summa_pog_kre) gr_mab from odb.graphpogkre gr, odb.licschkre kr  where length(gr.licschkre)= 20 and gr.licschkre = kr.licschkre and gr.subschkre = kr.subschkre and gr.date_pog <= to_date('" + txtgiris.Text + "','dd/mm/yyyy') group by gr.licschkre,gr.subschkre ) g1  where  kl.licschppkre = cer.debet(+) and kl.sk = cer.sk(+) and kl.hes_18 = pr.licschkre(+) and kl.sk = pr.subschkre(+) and kl.hes_18 = g1.licschkre(+) and kl.sk = g1.sk(+) and(kl.meb_18 + kl.meb_24 + kl.meb_19 + kl.meb19_24)>= 0 order by  kl.hes_18", con);
                //OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre, lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close, lk.date_planclose, lk.kolic_prolong, lk.date_prolong, lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok, lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08, sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18,sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth, r.adress, r.telefon, r.mobilniy,(trunc(months_between(vk.date_oper, nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper))) * 30) + extract(day from vk.date_oper) - extract(day from nvl(odb.func_get_overdue_min_date(vk.lastoverduedate_ish, vk.lastoverduedate, vk.lodinterest_ish), vk.date_oper)) as gec_gun from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr, view_nacpogprokre_all vk, creditinfo ci where substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and lk.date_oper = to_date('" + txtgiris.Text + "', 'dd/mm/yyyy') and lk.date_close is null and lk.licschppkre = vk.licschppkre and lk.subschkre = vk.subschkre and lk.licschpkre = vk.licschpkre order by  lk.hes_18", con);
                //OracleDataAdapter isleme = new OracleDataAdapter("select distinct r.name_regnom, lk.licschkre,  lk.licschpkre,lk.licsch_19, lk.licschppkre, lk.procstavkre, lk.procstavrez, lk.procstav_19,lk.procstavrez_19, lk.summakre, lk.summa, lk.summa_19, lk.date_open,lk.date_close, lk.date_planclose, lk.kolic_prolong, lk.date_prolong,lk.kolic_restructure, lk.date_restructure, lk.day_uderproc, kk.f_i_o,lk.srok, lk.lgotperiod, tk.name, g.name, lk.licsch_zaloga, lk.summa_zaloga,lk.kolic_pereocen_zaloga, lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,sr.item_01, sr.item_02, sr.item_03, sr.item_04, sr.item_05, sr.item_06,sr.item_07, sr.item_08, sr.item_09, sr.item_10, sr.item_11, sr.item_12,sr.item_13, sr.item_14, sr.item_15, sr.item_16, sr.item_17, sr.item_18,sr.item_19, sr.item_20, sr.aylig_borc_yuku, sr.aylig_gelir, sr.fifd,sr.bgn, ci.pincode, ci.dateofbirth, ci.placeofbirth, r.adress, r.telefon, r.mobilniy from odb.arh_licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr, creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and  lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre and lk.subschkre = ci.subschkre and lk.date_oper = to_date('" + txtgiris.Text + "','dd/mm/yyyy') and lk.date_close is null ", con);
                isleme.Fill(muracietler);
                dtg_siyahi.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

                dtg_siyahi.Columns[0].HeaderText  = "Adı";
                dtg_siyahi.Columns[0].Width = 250;
                dtg_siyahi.Columns[1].HeaderText  = "Ssuda";
                dtg_siyahi.Columns[1].Width = 250;
                dtg_siyahi.Columns[2].HeaderText  = "Ssuda faiz";
                dtg_siyahi.Columns[2].Width = 250;
                dtg_siyahi.Columns[2].HeaderText  = "Ssuda vk faiz";
                dtg_siyahi.Columns[2].Width = 250;
                dtg_siyahi.Columns[3].HeaderText  = "Ssuda vk";
                dtg_siyahi.Columns[3].Width = 200;
                dtg_siyahi.Columns[4].HeaderText  = "Ssuda vk faiz";
                dtg_siyahi.Columns[4].Width = 200;
                dtg_siyahi.Columns[5].HeaderText  = "Faiz";
                dtg_siyahi.Columns[5].Width = 100;
                dtg_siyahi.Columns[6].HeaderText  = "Ehtiyyat %";
                dtg_siyahi.Columns[6].Width = 100;
                dtg_siyahi.Columns[7].HeaderText  = "Ehtiyyat vk %";
                dtg_siyahi.Columns[7].Width = 100;
                dtg_siyahi.Columns[8].HeaderText  = "Min rez";
                dtg_siyahi.Columns[8].Width = 100;
                dtg_siyahi.Columns[9].HeaderText  = "Kredit məbləği";
                dtg_siyahi.Columns[9].Width = 100;
                dtg_siyahi.Columns[10].HeaderText  = "Qalıq";
                dtg_siyahi.Columns[10].Width = 100;
                dtg_siyahi.Columns[11].HeaderText  = "Vk qalıq";
                dtg_siyahi.Columns[11].Width = 100;
                dtg_siyahi.Columns[12].HeaderText  = "Verilmə tarixi";
                dtg_siyahi.Columns[12].Width = 100;
                dtg_siyahi.Columns[13].HeaderText  = "Bağlanma tarixi";
                dtg_siyahi.Columns[13].Width = 100;
                dtg_siyahi.Columns[14].HeaderText  = "K.p.b.tarixi";
                dtg_siyahi.Columns[14].Width = 100;
                dtg_siyahi.Columns[15].HeaderText  = "U.S.";
                dtg_siyahi.Columns[15].Width = 100;
                dtg_siyahi.Columns[16].HeaderText  = "U.T";
                dtg_siyahi.Columns[16].Width = 100;
                dtg_siyahi.Columns[17].HeaderText  = "RS";
                dtg_siyahi.Columns[17].Width = 100;
                dtg_siyahi.Columns[18].HeaderText  = "R.T";
                dtg_siyahi.Columns[18].Width = 100;
                dtg_siyahi.Columns[19].HeaderText  = "Ö.G";
                dtg_siyahi.Columns[19].Width = 100;
                dtg_siyahi.Columns[20].HeaderText  = "K.K.";
                dtg_siyahi.Columns[20].Width = 100;
                dtg_siyahi.Columns[21].HeaderText  = "K.M";
                dtg_siyahi.Columns[21].Width = 100;
                dtg_siyahi.Columns[22].HeaderText  = "G.M";
                dtg_siyahi.Columns[22].Width = 100;
                dtg_siyahi.Columns[23].HeaderText  = "Kreditin növü";
                dtg_siyahi.Columns[23].Width = 100;
                dtg_siyahi.Columns[24].HeaderText  = "Girovun növü";
                dtg_siyahi.Columns[24].Width = 200;
                dtg_siyahi.Columns[25].HeaderText  = "Girovun hesabı";
                dtg_siyahi.Columns[25].Width = 200;
                dtg_siyahi.Columns[26].HeaderText  = "Girovun məbləği";
                dtg_siyahi.Columns[26].Width = 100;
                dtg_siyahi.Columns[27].HeaderText  = "Sayı";
                dtg_siyahi.Columns[27].Width = 100;
                dtg_siyahi.Columns[28].HeaderText  = "G.Y.G";
                dtg_siyahi.Columns[28].Width = 100;
                dtg_siyahi.Columns[29].HeaderText  = "G.Y.G.T";
                dtg_siyahi.Columns[29].Width = 100;
                dtg_siyahi.Columns[30].HeaderText  = "İtem1";
                dtg_siyahi.Columns[30].Width = 100;
                dtg_siyahi.Columns[31].HeaderText  = "İtem2";
                dtg_siyahi.Columns[31].Width = 100;
                dtg_siyahi.Columns[32].HeaderText  = "İtem3";
                dtg_siyahi.Columns[32].Width = 100;
                dtg_siyahi.Columns[33].HeaderText  = "İtem4";
                dtg_siyahi.Columns[33].Width = 100;
                dtg_siyahi.Columns[34].HeaderText  = "İtem5";
                dtg_siyahi.Columns[34].Width = 100;
                dtg_siyahi.Columns[35].HeaderText  = "İtem6";
                dtg_siyahi.Columns[35].Width = 100;
                dtg_siyahi.Columns[36].HeaderText  = "İtem7";
                dtg_siyahi.Columns[36].Width = 100;
                dtg_siyahi.Columns[37].HeaderText  = "İtem8";
                dtg_siyahi.Columns[37].Width = 100;
                dtg_siyahi.Columns[38].HeaderText  = "İtem9";
                dtg_siyahi.Columns[38].Width = 100;
                dtg_siyahi.Columns[39].HeaderText  = "İtem10";
                dtg_siyahi.Columns[39].Width = 100;
                dtg_siyahi.Columns[40].HeaderText  = "İtem11";
                dtg_siyahi.Columns[40].Width = 100;
                dtg_siyahi.Columns[41].HeaderText  = "İtem12";
                dtg_siyahi.Columns[41].Width = 100;


                dtg_siyahi.Columns[42].HeaderText  = "İtem13";
                dtg_siyahi.Columns[42].Width = 100;
                dtg_siyahi.Columns[43].HeaderText  = "İtem14";
                dtg_siyahi.Columns[43].Width = 100;
                dtg_siyahi.Columns[44].HeaderText  = "İtem15";
                dtg_siyahi.Columns[44].Width = 100;
                dtg_siyahi.Columns[45].HeaderText  = "İtem16";
                dtg_siyahi.Columns[45].Width = 100;
                dtg_siyahi.Columns[46].HeaderText  = "İtem17";
                dtg_siyahi.Columns[46].Width = 100;
                dtg_siyahi.Columns[47].HeaderText  = "İtem18";
                dtg_siyahi.Columns[47].Width = 100;
                dtg_siyahi.Columns[48].HeaderText  = "İtem19";
                dtg_siyahi.Columns[48].Width = 100;
                dtg_siyahi.Columns[49].HeaderText  = "İtem20";
                dtg_siyahi.Columns[49].Width = 100;
                dtg_siyahi.Columns[50].HeaderText  = "Aylıq borc";
                dtg_siyahi.Columns[50].Width = 100;
                dtg_siyahi.Columns[51].HeaderText  = "Gəlir";
                dtg_siyahi.Columns[51].Width = 100;
                dtg_siyahi.Columns[52].HeaderText  = "FİFD";
                dtg_siyahi.Columns[52].Width = 100;
                dtg_siyahi.Columns[53].HeaderText  = "BGN";
                dtg_siyahi.Columns[53].Width = 100;
                dtg_siyahi.Columns[54].HeaderText  = "Fin";
                dtg_siyahi.Columns[54].Width = 100;
                dtg_siyahi.Columns[55].HeaderText  = "Doğum tarixi";
                dtg_siyahi.Columns[55].Width = 100;

                dtg_siyahi.Columns[56].HeaderText  = "Doğum yeri";
                dtg_siyahi.Columns[56].Width = 200;
                dtg_siyahi.Columns[57].HeaderText  = "Ünvanı";
                dtg_siyahi.Columns[57].Width = 500;
                dtg_siyahi.Columns[58].HeaderText  = "Əlaqə nömrələri";
                dtg_siyahi.Columns[58].Width = 400;
                dtg_siyahi.Columns[59].HeaderText  = "Mobil";
                dtg_siyahi.Columns[59].Width = 100;
                dtg_siyahi.Columns[60].HeaderText  = "Gecikmə günü";
                dtg_siyahi.Columns[60].Width = 100;



            }
            catch (Exception)
            {


            }
            finally { };


        }
        void ayliq()
        {
            dataTable = new System.Data.DataTable();
            dataTable.Columns.Add("Meblağ", typeof(decimal));
            dataTable.Columns.Add("Faiz", typeof(double));
            dataTable.Columns.Add("Müddet", typeof(int));
            dataTable.Columns.Add("AylıkÖdeme", typeof(decimal));

            // Örnek veriler ekleniyor.
            dataTable.Rows.Add(10000, 0.10, 12, DBNull.Value);

            // DataGridView'e DataTable'ı bağla
            //dataGridView.DataSource = dataTable;

            // DataTable'daki her bir satır için aylık ödemeyi hesapla ve sütuna yaz.
            foreach (DataRow row in dataTable.Rows)
            {
                double meb = Convert.ToDouble(row["Meblağ"]);
                double faizOranı = Convert.ToDouble(row["Faiz"]);
                int vadeMüddeti = Convert.ToInt32(row["Müddet"]);

                double aylıkÖdeme = calcPayment(meb, faizOranı, vadeMüddeti);
                row["AylıkÖdeme"] = aylıkÖdeme;
            }
        }
        private void exceleat()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "AML";
            worksheet.Columns["A"].ColumnWidth = 20; // A sütunu genişliyini artırır
            worksheet.Columns["B"].ColumnWidth = 25;
            // storing header part in Excel  
            for (int i = 1; i < dtg_axtar.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dtg_axtar.Columns[i - 1].HeaderText;
            }
            // Məlumatları Excel-ə əlavə etmək
            for (int i = 0; i < dtg_axtar.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dtg_axtar.Columns.Count; j++)
                {
                    if (j == 1) // 1-ci sütun: 20 rəqəmli format
                    {
                        worksheet.Cells[i + 2, j + 1].NumberFormat = "@"; // Hüceyrəni mətn formatında təyin et
                        worksheet.Cells[i + 2, j + 1] = dtg_axtar.Rows[i].Cells[j].Value.ToString().PadLeft(20, '0'); // Dəyərləri doldur
                    }
                    else
                    {
                        worksheet.Cells[i + 2, j + 1] = dtg_axtar.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
        }
        public static double calcPayment(double presentValue, double financingPeriod, double interestRatePerYear)
        {
            double a, b, x;
            double monthlyPayment;
            a = (1 + interestRatePerYear / 1200);
            b = financingPeriod;
            x = Math.Pow(a, b);
            x = 1 / x;
            x = 1 - x;
            monthlyPayment = (presentValue) * (interestRatePerYear / 1200) / x;
            return (monthlyPayment);
        }
        public static double CalculateMonthlyPayment(double presentValue, double financingPeriod, double interestRatePerYear)
        {
            double monthlyInterestRate = interestRatePerYear / 1200;
            double x = Math.Pow(1 + monthlyInterestRate, financingPeriod);
            double monthlyPayment = (presentValue * monthlyInterestRate * x) / (x - 1);
            return monthlyPayment;
        }

        private void frmehtiyyatlar_Load(object sender, EventArgs e)
        {
            
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBoxEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBoxEdit1.Text== "Ümumi Kataloq"|| comboBoxEdit1.Text == "Gecikmə günü olan müştərilər")
            //{
            //    labelControl1.Visible = true;
            //    labelControl1.Text ="Tarix üzrə";
            //    txtcixis.Visible = false;
            //    txtgiris.Visible = true;
            //}
            //else if (comboBoxEdit1.Text == "Son altı ayda  iki dəfə 30+ gecikmiş müştərilər"|| comboBoxEdit1.Text == "Son altı ayda  iki dəfə 30+ gecikməsi olmayan min res 15% olanlar")
            //{
            //    labelControl1.Visible = true;
            //    labelControl1.Text = "Tarixlər üzrə";
            //    txtgiris.Visible = true;
            //    txtcixis.Visible = true;
            //}
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            testlerucun();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
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

        private void textBox1_TextChanged(object sender, EventArgs e)
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

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            
                if (e.KeyCode == Keys.Enter)
                {
                    // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                    txtcixis.Focus();
                    e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
                }
        }

        private void txtcixis_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                btn_sorgu.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            muracietler.Clear();
            //labelControl1.Visible = true;
            //labelControl1.Text = "Tarixlər üzrə";
            txtgiris.Visible = true;
            txtcixis.Visible = true;
            if (rdb_iki.Checked == true)
            {
                listele_30_iki_defe();
            }
            else if (rdb_bir.Checked == true)
            {
                listelemuraciet30_bidefe();
            }
            else if (radioButton1.Checked==true)
            {
                axtar_elave();
            }
            //if (comboBoxEdit1.Text== "Son altı ayda  iki dəfə 30+ gecikmiş müştərilər")
            //{
            //    listele_30_iki_defe();
            //    //listelemuraciet();
            //}
            //else if (comboBoxEdit1.Text == "Son altı ayda  iki dəfə 30+ gecikməsi olmayan min res 15% olanlar")
            //{
            //    listelemuraciet30_bidefe();
            //}
            //else if (comboBoxEdit1.Text == "BGN üzrə səhv ehtiyyatlar")
            //{

            //}
            //else if (comboBoxEdit1.Text == "Ümumi Kataloq")
            //{
            //    tamkataloq();
            //}
            //else if (comboBoxEdit1.Text == "Gecikmə günü olan müştərilər")
            //{

            //}
            //else if (comboBoxEdit1.Text == "Restruktrurizasiya olunmuş kreditlərdə yaxşılaşma")
            //{

            //}
            //else if (comboBoxEdit1.Text == "Krediti bağlı olub bk qalanlar")
            //{
            //    kr_bagli_olub_bk_qal();
            //}




        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmehtiyatlarin_cedveli frmehtcedvel = new frmehtiyatlarin_cedveli();
            frmehtcedvel.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            exceleat();
        }

        private void dtg_siyahi_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // 0-cı sütundakı dəyəri əldə edirik
                hesab = dtg_siyahi.Rows[e.RowIndex].Cells[0].Value.ToString();
                sub = dtg_siyahi.Rows[e.RowIndex].Cells[1].Value.ToString();
                axtar();

            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                panel1.Visible = true;
            }
            else
            {
                panel1.Visible = false;
            }
            }
    }
}
