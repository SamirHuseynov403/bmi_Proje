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
using DevExpress.CodeParser;
using DevExpress.DataProcessing.InMemoryDataProcessor;
using static DevExpress.Utils.Drawing.Helpers.NativeMethods;

namespace BMI.Muhasibat
{
    public partial class frm_iran_hesabat : Form
    {
        public frm_iran_hesabat()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl_yanasma = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        string tarix_soz;
        DateTime dt = DateTime.Now;
        public void tarixitapsoz()
        {
            int ay = Convert.ToInt16(txt_tarix_evvel.Text.Substring(3, 2));
            //textBox1.Text = dt.Day.ToString();
            //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
            //int ay = 6;

            if (ay == 1)
            {
                tarix_soz = "January";
            }
            else if (ay == 2)
            {
                tarix_soz = "February";
            }
            else if (ay == 3)
            {
                tarix_soz = "March";
            }
            else if (ay == 4)
            {
                tarix_soz = "April";
            }
            else if (ay == 5)
            {
                tarix_soz = "May";
            }
            else if (ay == 6)
            {
                tarix_soz = "June";
            }
            else if (ay == 7)
            {
                tarix_soz = "July";
            }
            else if (ay == 8)
            {
                tarix_soz = "August";
            }
            else if (ay == 9)
            {
                tarix_soz = "September";
            }
            else if (ay == 10)
            {
                tarix_soz = "October";
            }
            else if (ay == 11)
            {
                tarix_soz = "November";
            }
            else if (ay == 12)
            {
                tarix_soz = "December";
            }
        }
        private void excel()
        {
            tarixitapsoz();

            DataTable dt_interaktiv = new DataTable();
            DataTable dt_isci_giris_qaliq = new DataTable();
            DataTable dt_isci_ver_kr = new DataTable();
            DataTable dt_odenilmis = new DataTable();
            DataTable dt_faiz_giris_qaliq = new DataTable();
            //DataTable dt_nagd_vesait_kr = new DataTable();
            DataTable dt_hesablanmis = new DataTable();
            DataTable dt_ehtiyyat = new DataTable();
            DataTable dt_isci_89_2 = new DataTable();
            DataTable dt_umumi_89_2 = new DataTable();
            DataTable dt_isci_2_89 = new DataTable();
            DataTable dt_umumi_2_89 = new DataTable();
            DataTable dt_bk_qaliqlar = new DataTable();

            string ilktar=txt_tarix_evvel.Text;
            string sontar=txt_tarix_son.Text;
            string kecenay = txt_evvelki_ay.Text;

            string formattedIlktar = DateTime.ParseExact(ilktar, "dd-mm-yyyy", CultureInfo.InvariantCulture).ToString("dd-mm-yyyy");
            string formattedSontar = DateTime.ParseExact(sontar, "dd-mm-yyyy", CultureInfo.InvariantCulture).ToString("dd-mm-yyyy");
            string formattedKecenay = DateTime.ParseExact(kecenay, "dd-mm-yyyy", CultureInfo.InvariantCulture).ToString("dd-mm-yyyy");
            #region sql_kodlar

           
            string interaktiv = "select dax.hes,dax.D,dovr.e,dovr.f, "+
"case when substr(dax.hes,5,1) in (2, 4, 7, 9) then 'faiz' " +
 "                else (case when substr(dax.hes,5,1) in (0, 3, 5, 8) then 'esas' end)end esas_faiz, " +
  "                  case when(substr(dax.hes, 1, 2) in (15, 20) or substr(dax.hes, 1, 4) in (2124, 2125, 2114, 2115))  then 'sahibkar' " +
   "              else 'fiziki'end s_f from" +
"(select substr(t.licsch, 1, 5) hes, sum(t.saldo_vhd_nacval) D from odb.arh_saldo_ls t " +
"where substr(t.licsch, 1, 5)between(20000)and(23123) and t.date_oper = odb.ish_gun_cari1(to_date('" + txt_tarix_evvel.Text + "', 'dd/mm/yyyy')) " +
"group by substr(t.licsch, 1, 5) order by substr(t.licsch, 1, 5) asc)dax," +
"(select substr(t.licsch, 1, 5) hes,sum(t.oboroti_debet_nacval) E,sum(t.oboroti_kredit_nacval) F from odb.arh_saldo_ls t " +
"where substr(t.licsch,1,5)between(20000)and(23123) and t.date_oper BETWEEN TO_DATE('"+txt_tarix_evvel.Text+"', 'dd/mm/yyyy') AND TO_DATE('"+txt_tarix_son.Text+"', 'dd/mm/yyyy') " +
"group by substr(t.licsch, 1, 5) order by substr(t.licsch, 1, 5) asc) dovr where dax.hes = dovr.hes and " +
"substr(dax.hes, 1, 3) not in ('159', '209', '219', '239', '259')";

            string isci_ver_kr = "select substr(t.licsch,10,6) hes, "+
            "case when t.licsch in (l.licschkre, l.licsch_19) then 'esas' " + 
            "when t.licsch in(l.licschpkre, l.licschppkre) then 'faiz' end tip, " +
            "sum(t.oboroti_debet_nacval)from odb.arh_saldo_ls t,regnom r, arh_licschkre l " +
            "where t.date_oper BETWEEN TO_DATE('" + txt_tarix_evvel.Text + "', 'dd/mm/yyyy') AND TO_DATE('" + txt_tarix_son.Text + "', 'dd/mm/yyyy')  " +
            "and r.svazanniy = 1 and r.regnom = substr(t.licsch, 10, 6) and t.licsch in (l.licschkre, l.licsch_19, l.licschpkre, l.licschppkre) and t.date_oper = l.date_oper " +
            "and t.oboroti_debet_nacval <> 0 group by substr(t.licsch, 10, 6) , " +
            "case when t.licsch in (l.licschkre, l.licsch_19) then 'esas' " +
            "when t.licsch in(l.licschpkre, l.licschppkre) then 'faiz' end";

            string  interaktiv_isci= "select case when s.licsch in (ar.licschkre, ar.licsch_19) then 'esas' else 'faiz' end tip,"+
            " sum(s.saldo_ish_nacval) qal from odb.arh_saldo_ls s, arh_licschkre ar,regnom r" +
            " where ar.date_oper = TO_DATE('" + txt_tarix_son.Text + "', 'dd/mm/yyyy') and s.licsch in (ar.licschkre, ar.licschpkre, ar.licsch_19, ar.licschppkre)" +
            "  and r.svazanniy = 1 and r.regnom = substr(ar.licschkre, 10, 6) and ar.date_close is null and ar.date_oper = s.date_oper" +
            " group by CASE WHEN s.licsch IN(ar.licschkre, ar.licsch_19) THEN 'esas' ELSE 'faiz' END";

    //        string nagd_vesait_kredit = "select "+
    //"CASE " +
    //"WHEN r.svazanniy = 1 THEN 'isci' " +
    //"WHEN r.fizik = 1 AND r.svazanniy = 0 THEN 'fiziki' " +
    //"WHEN r.fizik = 0 AND r.svazanniy = 0 THEN 'sahibkar' " +
    //"ELSE 'unknown' END AS tip,sum(l.summa + l.summa_19) " +
    //"from arh_licschkre l ,regnom r " +
    //"where l.date_oper = to_date('17/04/2018', 'dd/mm/yyyy') " +
    //"and l.tipzaloga = 3 and l.date_close is null and r.regnom = substr(l.licschkre, 10, 6) " +
    //"group by " +
    //"CASE " +
    //"WHEN r.svazanniy = 1 THEN 'isci' " +
    //"WHEN r.fizik = 1 AND r.svazanniy = 0 THEN 'fiziki' " +
    //"WHEN r.fizik = 0 AND r.svazanniy = 0 THEN 'sahibkar' " +
    //"ELSE 'unknown' END";

            string faiz_giris_qaliqlar = "select TO_CHAR(l.tipkredita) hes,sum((n.nacprospro_ish - n.pogprospro_ish)+(n.nacpro_ish-n.pogpro_ish)) meb from view_nacpogprokre_all n ,arh_licschkre l,regnom r " +
   "where n.date_oper = to_date('" + txt_evvelki_ay.Text + "', 'DD/MM/YYYY') and l.date_oper = n.date_oper and l.date_close is null and n.licschpkre in (l.licschpkre, l.licschppkre) and n.subschkre = l.subschkre " +
   "and r.regnom = substr(l.licschkre, 10, 6) and r.svazanniy <> 1 " +
   "group by l.tipkredita " +
   "union " +
   "select  TO_CHAR('4') as hes,sum((n.nacprospro_ish - n.pogprospro_ish) + (n.nacpro_ish - n.pogpro_ish)) meb from view_nacpogprokre_all n, arh_licschkre l,regnom r " +
   "where n.date_oper = to_date('" + txt_evvelki_ay.Text + "', 'DD/MM/YYYY') and l.date_oper = n.date_oper and l.date_close is null " +
   "and n.licschpkre in (l.licschpkre, l.licschppkre) and n.subschkre = l.subschkre and r.regnom = substr(l.licschkre, 10, 6) and r.svazanniy = 1";

            string odenilmis_kr_nov = " SELECT g.name,'qeyri' zaloq, " +
                         " COUNT(DISTINCT CASE WHEN substr(ar.kredit, 11, 6) IN(substr(l.licschkre, 11, 6), substr(l.licschpkre, 11, 6), substr(l.licsch_19, 11, 6), substr(l.licschppkre, 11, 6)) THEN substr(ar.kredit, 11, 6) ELSE NULL END)say, " +
                          " sum(CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) + sum(CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                          "sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum(CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                          " FROM regnom r,licschkre l " +
                          " JOIN tipkre g ON l.tipkredita = g.code " +
                          " JOIN arh_dd ar ON ar.date_oper BETWEEN to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY')  " +
                          " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                          " AND l.subschkre = ar.ssk and r.regnom = substr(l.licschkre, 10, 6) and r.svazanniy <> 1 and " +
                          "       ((substr(ar.debet, 10, 6) = substr(ar.kredit, 10, 6) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                          " OR " +
                          " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)))  " +
                          " GROUP BY g.name " +
                          " union " +
                          " SELECT g.name,'isci' zaloq, " +
                          " COUNT(DISTINCT CASE WHEN substr(ar.kredit, 11, 6) IN(substr(l.licschkre, 11, 6), substr(l.licschpkre, 11, 6), substr(l.licsch_19, 11, 6), substr(l.licschppkre, 11, 6)) THEN substr(ar.kredit, 11, 6) ELSE NULL END)say, " +
                          " sum(CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) + sum(CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
                          " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END)+sum(CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
                          " FROM regnom r,licschkre l " +
                          " JOIN tipkre g ON l.tipkredita = g.code " +
                          " JOIN arh_dd ar ON ar.date_oper BETWEEN to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY')  " +
                          " WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
                          " AND l.subschkre = ar.ssk and r.regnom = substr(l.licschkre, 10, 6) and r.svazanniy = 1 and " +
                          "       ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)) " +
                          " OR " +
                          " (substr(ar.debet, 0, 3) in (159, 209, 219, 239, 259) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre)))  " +
                          " GROUP BY g.name";

            string hesablanmis_faizler_kr_nov = "SELECT '4' AS ischi, " +
              "SUM(CASE WHEN ar.debet = l.licschpkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschpkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) + " +
             "SUM(CASE WHEN ar.debet = l.licschppkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschppkre, 10, 6) AND r.svazanniy = 1 THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
         "FROM licschkre l, arh_dd ar, regnom r WHERE " +
         "   ar.date_oper BETWEEN to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY')  " +
         "  AND r.regnom = SUBSTR(ar.debet, 10, 6) and substr(ar.kredit,1,5) not in ('66220', '86220') " +
         " AND ar.debet IN(l.licschpkre, l.licschppkre) " +
         "  AND l.subschkre = ar.ssd " +
         "  AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
         "  AND r.regnom = SUBSTR(l.licschpkre, 10, 6) " +
         "  AND r.svazanniy = 1 " +
         "  union " +
         "  SELECT TO_CHAR(l.tipkredita) AS ischi, " +
         "     SUM(CASE WHEN ar.debet = l.licschpkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschpkre, 10, 6)  THEN ar.summa_v_nacval ELSE 0 END)  + " +
         "    SUM(CASE WHEN ar.debet = l.licschppkre AND r.regnom = SUBSTR(ar.debet, 10, 6) AND r.regnom = SUBSTR(l.licschppkre, 10, 6)  THEN ar.summa_v_nacval ELSE 0 END) AS faiz " +
        " FROM licschkre l, arh_dd ar, regnom r WHERE " +
        "    ar.date_oper BETWEEN to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY') " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) and substr(ar.kredit,1,5) not in ('66220', '86220')" +
        "  AND ar.debet IN(l.licschpkre, l.licschppkre) " +
        "   AND l.subschkre = ar.ssd " +
        "   AND r.regnom = SUBSTR(ar.debet, 10, 6) " +
        "   AND r.regnom = SUBSTR(l.licschpkre, 10, 6) " +
        "   AND r.svazanniy <> 1 " +
        "   group by l.tipkredita";
            string icmal_ehtiyatlar = "SELECT t1.eht,t1.tip,t1.qaliq - COALESCE(t2.azalma, 0) AS qaliq, "+
"    t1.hes,t1.sk " +
" FROM("+
" SELECT CASE WHEN l.procstavrez < 25 THEN 'adi' ELSE 'meqsedli' END AS eht, " +
" CASE WHEN r.svazanniy = 1 THEN 'isci' WHEN r.fizik = 1 AND r.svazanniy = 0 THEN 'fiziki' " +
" WHEN r.fizik = 0 AND r.svazanniy = 0 THEN 'sahibkar' " +
" ELSE 'unknown' END AS tip, " +
" SUM(ROUND((l.summa + l.summa_19) * POWER(1 + l.procstavrez / 100, 1), 2) - (l.summa + l.summa_19)) AS qaliq, " +
" l.licschkre AS hes, l.subschkre AS sk " +
" FROM arh_licschkre l JOIN srokpogprockre s ON s.licschkre = l.licschkre AND s.subschkre = l.subschkre " +
" JOIN regnom r ON r.regnom = SUBSTR(l.licschkre, 10, 6) " +
" WHERE l.date_oper = TO_DATE('" + kecenay+"', 'dd/mm/yyyy') "+
" AND l.date_close IS NULL GROUP BY " +
" CASE WHEN l.procstavrez < 25 THEN 'adi' ELSE 'meqsedli' END, " +
" l.licschkre, l.subschkre, CASE " +
" WHEN r.svazanniy = 1 THEN 'isci' " +
" WHEN r.fizik = 1 AND r.svazanniy = 0 THEN 'fiziki' " +
" WHEN r.fizik = 0 AND r.svazanniy = 0 THEN 'sahibkar' " +
" ELSE 'unknown' END " +
" ) t1 LEFT JOIN(" +
" SELECT ar.licschkre AS hesab, ar.subschkre AS sk, " +
" ((ar.summa + ar.summa_19) + " +
" (SELECT(v.nacpro_ish - v.pogpro_ish) + (v.nacprospro_ish - v.pogprospro_ish) " +
" FROM view_nacpogprokre_all v " +
" WHERE v.date_oper = TO_DATE('" + kecenay+"', 'dd/mm/yyyy') "+
" AND v.subschkre = ar.subschkre " +
" AND v.licschpkre = ar.licschpkre)) *s.girovun_likvidlik_faiz_derece / 100 AS azalma " +
" FROM arh_licschkre ar " +
" JOIN girovun_bazar_deyeri g ON ar.licschkre = g.licschkre AND ar.subschkre = g.subschkre " +
" JOIN srokpogprockre s ON s.licschkre = ar.licschkre AND s.subschkre = ar.subschkre " +
" WHERE ar.date_oper = TO_DATE('" + kecenay+"', 'dd/mm/yyyy') "+
" AND ar.date_close IS NULL " +
" AND s.girovun_likvidlik_faiz_derece IS NOT NULL " +
" AND g.tarix <= TO_DATE('" + kecenay+"', 'dd/mm/yyyy') "+
" GROUP BY ar.licschkre,ar.subschkre, s.girovun_likvidlik_faiz_derece, ar.summa, " +
" ar.summa_19,ar.procstavrez, ar.licschpkre) t2 ON t1.hes = t2.hesab AND t1.sk = t2.sk"; 



            string isci_89_2 = "select ar.debet db,ar.kredit kr,ar.summa_v_nacval meb, " +
"case when  TO_NUMBER(REGEXP_SUBSTR(l.name_licsch, '\\d+'))<25 then 'adi' else 'meqsedli' end AS tey, " +
 "        case when(substr(ar.debet, 1, 4) = '8916' or substr(ar.debet, 1, 4) = '8911') then 'sah' else 'fizik' end tey, SUBSTR(ar.primechanie, INSTR(ar.primechanie, '/') + 1) AS tey_adi " +
 "from arh_dd ar,licsch l, regnom r " +
 "  where ar.date_oper between to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY') " +
 "and substr(ar.debet,1,2)= '89' and substr(ar.kredit,1,2) in ('20','21') and ar.kredit = l.licsch and r.svazanniy = 1 and SUBSTR(ar.primechanie, INSTR(ar.primechanie, '/') +1)= r.name_regnom";

            string umumi_89_2 = "select ar.debet db,ar.kredit kr,ar.summa_v_nacval meb, "+
"case when  TO_NUMBER(REGEXP_SUBSTR(l.name_licsch, '\\d+'))<25 then 'adi' else 'meqsedli' end AS tey, " +
"case when (substr(ar.debet,1,4)='8916' or substr(ar.debet,1,4)='8911') then 'sah' else 'fizik' end tey" +
 " from arh_dd ar,licsch l "+
 " where ar.date_oper between to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY') " +
 "and substr(ar.debet,1,2)= '89' and substr(ar.kredit,1,2) in ('20','21') and ar.kredit = l.licsch";

            string isci_2_89 = "select distinct ar.debet db,ar.kredit kr, ar.summa_v_nacval meb, "+
 "case when l.procstavrez < 25 then 'adi' else 'meqsedli' end AS eh, " +
"ar.summa_v_nacval* l.procstavrez / 100 from arh_dd ar,arh_licschkre l, regnom r " +
"   where ar.date_oper between to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY') and " +
"ar.kredit in (l.licschkre, l.licsch_19) and r.regnom = substr(l.licschkre, 10, 6) and r.svazanniy = 1 and l.date_oper = ar.date_oper and l.date_close is null " +
"and substr(ar.debet,1,1)= '4'";

            string umumi_2_89 = "select ar.debet db,ar.kredit kr,ar.summa_v_nacval meb, "+
"case when  TO_NUMBER(REGEXP_SUBSTR(l.name_licsch, '\\d+'))<25 then 'adi' else 'meqsedli' end AS tey, " +
 "        case when(substr(ar.kredit, 1, 4) = '8916' or substr(ar.kredit, 1, 4) = '8911') then 'sah' else 'fizik' end tey " +
 "from arh_dd ar,licsch l " +
 " where ar.date_oper between to_date('" + txt_tarix_evvel.Text + "', 'DD/MM/YYYY') and to_date('" + txt_tarix_son.Text + "', 'DD/MM/YYYY') " +
 "and substr(ar.debet,1,2) in ('20','21') and substr(ar.kredit,1,2)= '89' and ar.debet = l.licsch";

            string bk_qaliqlar = "select t.date_oper tarix,t.vbs, " +
 " CASE WHEN r.svazanniy = 1 THEN 'isci' " +
 "  WHEN r.fizik = 1 AND r.svazanniy = 0 THEN 'fiziki' " +
 "  WHEN r.fizik = 0 AND r.svazanniy = 0 THEN 'sahibkar' " +
 "  ELSE 'unknown' END AS tip, " +
"t.licsch,substr(t.licsch, 6, 2),t.ssls, " +
"  t.ostatok_ish* ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv, " +
"  ROUND(odb.func_get_kurval(substr(t.licsch, 6, 2), t.date_oper), 6)  kurs " +
"      from odb.arh_saldo_vbls t, arh_licschkre l,regnom r " +
"  where t.date_oper = TO_DATE('"+txt_tarix_son.Text+"', 'dd/mm/yyyy') and t.vbs in (99740, 99742, 99743, 99749) and t.ostatok_ish <> 0 and t.date_oper = l.date_oper " +
"  and substr(t.licsch,10,6) = substr(l.licschkre, 10, 6) and t.ssls = l.subschkre and r.regnom = substr(l.licschkre, 10, 6)";
            #endregion
            using (OracleConnection connection = new OracleConnection(cl_yanasma.con))

            {
                using (OracleCommand command = new OracleCommand(interaktiv, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_interaktiv);
                }
                using (OracleCommand command = new OracleCommand(interaktiv_isci, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_isci_giris_qaliq);
                    
                }
                using (OracleCommand command = new OracleCommand(isci_ver_kr, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_isci_ver_kr);

                }
                using (OracleCommand command = new OracleCommand(faiz_giris_qaliqlar, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_faiz_giris_qaliq);

                }
                //using (OracleCommand command = new OracleCommand(nagd_vesait_kredit, connection))
                //{
                //    //connection.Open();
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(dt_nagd_vesait_kr);

                //}
                using (OracleCommand command = new OracleCommand(odenilmis_kr_nov, connection))
                {
                    dt_odenilmis.Clear();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_odenilmis);
                    //    connection.Close();

                    //    //gridControl1.DataSource = dt_xett;
                 }

                using (OracleCommand command = new OracleCommand(hesablanmis_faizler_kr_nov, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_hesablanmis);

                }

                using (OracleCommand command = new OracleCommand(icmal_ehtiyatlar, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_ehtiyyat);  
                }
                using (OracleCommand command = new OracleCommand(isci_89_2, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_isci_89_2);
                    
                }

                using (OracleCommand command = new OracleCommand(umumi_89_2, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_umumi_89_2);
                    
                }
                using (OracleCommand command = new OracleCommand(umumi_2_89, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_umumi_2_89);

                }
                using (OracleCommand command = new OracleCommand(isci_2_89, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_isci_2_89);

                }

                using (OracleCommand command = new OracleCommand(bk_qaliqlar, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_bk_qaliqlar);
                    connection.Close();
                }
            }

            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txt_tarix_evvel.Text.Substring(4,2); // TextBox'tan alınan metni sakla


            string yeniMetin = tarix_soz;
            string baseFileName = "یادداشت تسهیلاتسپرده ها و زیر خط - نهاییJ"+ yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "iran_hesabat.xlsx");
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
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                
                ExcelWorksheet ws = package.Workbook.Worksheets["Note 13"];

                //*********************************
                //F sutun
                var sh_F7 = dt_interaktiv.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "esas" && row.Field<string>(5) == "fiziki")
                                    .ToList();
                decimal total_sh_F7 = sh_F7.Sum(row => row.Field<decimal>(1));

                var sh_F8 = dt_isci_giris_qaliq.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "esas")
                                    .ToList();
                decimal total_sh_F8 = sh_F8.Sum(row => row.Field<decimal>(1));

                var sh_F9 = dt_interaktiv.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "esas" && row.Field<string>(5) == "sahibkar")
                                    .ToList();
                decimal total_sh_F9 = sh_F9.Sum(row => row.Field<decimal>(1));
                //*********************************
                //G sutun
                var sh_G7 = dt_interaktiv.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "esas" && row.Field<string>(5) == "fiziki")
                                    .ToList();
                decimal total_sh_G7 = sh_G7.Sum(row => row.Field<decimal>(2));

                var sh_G8 = dt_isci_ver_kr.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "esas")
                                    .ToList();
                decimal total_sh_G8 = sh_G8.Sum(row => row.Field<decimal>(2));

                var sh_G9 = dt_interaktiv.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "esas" && row.Field<string>(5) == "sahibkar")
                                    .ToList();
                decimal total_sh_G9 = sh_G9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //H sutun
                var sh_H7 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "FİZİKİ ŞƏXSLƏR" && row.Field<string>(1) != "isci")
                                    .ToList();
                decimal total_sh_H7 = sh_H7.Sum(row => row.Field<decimal>(3));

                var sh_H8 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "isci")
                                    .ToList();
                decimal total_sh_H8 = sh_H8.Sum(row => row.Field<decimal>(3));

                var sh_H9 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) != "FİZİKİ ŞƏXSLƏR")
                                    .ToList();
                decimal total_sh_H9 = sh_H9.Sum(row => row.Field<decimal>(3));

                //*********************************
                //K sutun
                var sh_K7 = dt_faiz_giris_qaliq.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "2" )
                                    .ToList();
                decimal total_sh_K7 = sh_K7.Sum(row => row.Field<decimal>(1));

                var sh_K8 = dt_faiz_giris_qaliq.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "4" )
                                    .ToList();
                DataTable filteredDataTable = dt_faiz_giris_qaliq.Clone(); // İlk tablonun şemasını kopyala
                foreach (var row in sh_K8)
                {
                    filteredDataTable.Rows.Add(row.ItemArray);
                }

                // DataGridView'e yeni DataTable'ı atayarak güncelle
                dataGridView1.DataSource = filteredDataTable;

                decimal total_sh_K8 = sh_K8.Sum(row => row.Field<decimal>(1));//---***************************

                var sh_K9 = dt_faiz_giris_qaliq.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "1" || row.Field<string>(0) == "3")
                                    .ToList();
                decimal total_sh_K9 = sh_K9.Sum(row => row.Field<decimal>(1));

                //*********************************
                //L sutun
                var sh_L7 = dt_hesablanmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "2" )
                                    .ToList();
                decimal total_sh_L7 = sh_L7.Sum(row => row.Field<decimal>(1));

                var sh_L8 = dt_hesablanmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "4")
                                    .ToList();
                decimal total_sh_L8 = sh_L8.Sum(row => row.Field<decimal>(1));

                var sh_L9 = dt_hesablanmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "1" || row.Field<string>(0) == "3")
                                    .ToList();
                decimal total_sh_L9 = sh_L9.Sum(row => row.Field<decimal>(1));

                //*********************************
                //M sutun
                var sh_M7 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "FİZİKİ ŞƏXSLƏR" && row.Field<string>(1) != "isci")
                                    .ToList();
                decimal total_sh_M7 = sh_M7.Sum(row => row.Field<decimal>(4));

                var sh_M8 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "isci")
                                    .ToList();
                decimal total_sh_M8 = sh_M8.Sum(row => row.Field<decimal>(4));

                var sh_M9 = dt_odenilmis.AsEnumerable()
                                    .Where(row => row.Field<string>(0) != "FİZİKİ ŞƏXSLƏR")
                                    .ToList();
                decimal total_sh_M9 = sh_M9.Sum(row => row.Field<decimal>(4));

                //*********************************
                //AE sutun
                var sh_AE7 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "adi" && row.Field<string>(1) == "fiziki")
                                    .ToList();
                decimal total_sh_AE7 = sh_AE7.Sum(row => row.Field<decimal>(2));

                var sh_AE8 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "adi" && row.Field<string>(1) == "isci")
                                    .ToList();
                decimal total_sh_AE8 = sh_AE8.Sum(row => row.Field<decimal>(2));

                var sh_AE9 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "adi" && row.Field<string>(1) == "sahibkar")
                                    .ToList();
                decimal total_sh_AE9 = sh_AE9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AF sutun
                var sh_AF7 = dt_umumi_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "fizik" && row.Field<string>(3) == "adi")
                                    .ToList();
                decimal total_sh_AF7 = sh_AF7.Sum(row => row.Field<decimal>(2));

                var sh_AF8 = dt_isci_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(3) == "adi")
                                    .ToList();
                decimal total_sh_AF8 = sh_AF8.Sum(row => row.Field<decimal>(2));

                var sh_AF9 = dt_umumi_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "sah" && row.Field<string>(3) == "adi")
                                    .ToList();
                decimal total_sh_AF9 = sh_AF9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AG sutun

                var sh_AG7 = dt_umumi_2_89.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "fizik" && row.Field<string>(3) == "adi")
                                    .ToList();
                decimal total_sh_AG7 = sh_AG7.Sum(row => row.Field<decimal>(2));
                decimal total_sh_AG8 = 0;
                if (dt_isci_2_89 != null && dt_isci_2_89.Rows.Count > 0)
                {
                    var sh_AG8 = dt_isci_2_89.AsEnumerable()
                                        .Where(row => row.Field<string>(3) == "adi")
                                        .ToList();
                    total_sh_AG8 = sh_AG8.Sum(row => row.Field<decimal>(4));
                    // total_sh_AG8 ile ilgili başka işlemleri gerçekleştirin
                }
                var sh_AG9 = dt_umumi_2_89.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "sah" && row.Field<string>(3) == "adi")
                                    .ToList();
                decimal total_sh_AG9 = sh_AG9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AI sutun
                var sh_AI7 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "meqsedli" && row.Field<string>(1) == "fiziki")
                                    .ToList();
                decimal total_sh_AI7 = sh_AI7.Sum(row => row.Field<decimal>(2));

                var sh_AI8 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "meqsedli" && row.Field<string>(1) == "isci")
                                    .ToList();
                decimal total_sh_AI8 = sh_AI8.Sum(row => row.Field<decimal>(2));

                var sh_AI9 = dt_ehtiyyat.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "meqsedli" && row.Field<string>(1) == "sahibkar")
                                    .ToList();
                decimal total_sh_AI9 = sh_AI9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AJ sutun
                var sh_AJ7 = dt_umumi_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "fizik" && row.Field<string>(3) == "meqsedli")
                                    .ToList();
                decimal total_sh_AJ7 = sh_AJ7.Sum(row => row.Field<decimal>(2));

                var sh_AJ8 = dt_isci_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(3) == "meqsedli")
                                    .ToList();
                decimal total_sh_AJ8 = sh_AJ8.Sum(row => row.Field<decimal>(2));

                var sh_AJ9 = dt_umumi_89_2.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "sah" && row.Field<string>(3) == "meqsedli")
                                    .ToList();
                decimal total_sh_AJ9 = sh_AJ9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AK sutun

                var sh_AK7 = dt_umumi_2_89.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "fizik" && row.Field<string>(3) == "meqsedli")
                                    .ToList();
                decimal total_sh_AK7 = sh_AK7.Sum(row => row.Field<decimal>(2));
                decimal total_sh_AK8 = 0;
                if (dt_isci_2_89 != null && dt_isci_2_89.Rows.Count > 0)
                {
                    var sh_AK8 = dt_isci_2_89.AsEnumerable()
                                        .Where(row => row.Field<string>(3) == "meqsedli")
                                        .ToList();
                    total_sh_AK8 = sh_AK8.Sum(row => row.Field<decimal>(4));
                    // total_sh_AG8 ile ilgili başka işlemleri gerçekleştirin
                }
                var sh_AK9 = dt_umumi_2_89.AsEnumerable()
                                    .Where(row => row.Field<string>(4) == "sah" && row.Field<string>(3) == "meqsedli")
                                    .ToList();
                decimal total_sh_AK9 = sh_AK9.Sum(row => row.Field<decimal>(2));

                //*********************************
                //AV sutun
                var sh_AV7 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99743" && row.Field<string>(2) == "fiziki")
                                    .ToList();
                decimal total_sh_AV7 = sh_AV7.Sum(row => row.Field<decimal>(6));

                var sh_AV8 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99743" && row.Field<string>(2) == "isci")
                                    .ToList();
                decimal total_sh_AV8 = sh_AV8.Sum(row => row.Field<decimal>(6));

                var sh_AV9 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99743" && row.Field<string>(2) == "sahibkar")
                                    .ToList();
                decimal total_sh_AV9 = sh_AV9.Sum(row => row.Field<decimal>(6));

                //*********************************
                //AY sutun
                var sh_AY7 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) != "99743" && row.Field<string>(2) == "fiziki")
                                    .ToList();
                decimal total_sh_AY7 = sh_AY7.Sum(row => row.Field<decimal>(6));

                var sh_AY8 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) != "99743" && row.Field<string>(2) == "isci")
                                    .ToList();
                decimal total_sh_AY8 = sh_AY8.Sum(row => row.Field<decimal>(6));

                var sh_AY9 = dt_bk_qaliqlar.AsEnumerable()
                                    .Where(row => row.Field<string>(1) != "99743" && row.Field<string>(2) == "sahibkar")
                                    .ToList();
                decimal total_sh_AY9 = sh_AY9.Sum(row => row.Field<decimal>(6));

                ws.Cells[7, 6].Value = total_sh_F7- total_sh_F8;
                ws.Cells[8, 6].Value = total_sh_F8;
                ws.Cells[9, 6].Value = total_sh_F9;

                ws.Cells[7, 7].Value = total_sh_G7 ;
                ws.Cells[8, 7].Value = total_sh_G8;
                ws.Cells[9, 7].Value = total_sh_G9;

                ws.Cells[7, 8].Value = total_sh_H7;
                ws.Cells[8, 8].Value = total_sh_H8;
                ws.Cells[9, 8].Value = total_sh_H9;

                ws.Cells[7, 11].Value = total_sh_K7;
                ws.Cells[8, 11].Value = total_sh_K8;
                ws.Cells[9, 11].Value = total_sh_K9;

                ws.Cells[7, 12].Value = total_sh_L7;
                ws.Cells[8, 12].Value = total_sh_L8;
                ws.Cells[9, 12].Value = total_sh_L9;

                ws.Cells[7, 13].Value = total_sh_M7;
                ws.Cells[8, 13].Value = total_sh_M8;
                ws.Cells[9, 13].Value = total_sh_M9;

                ws.Cells[7, 31].Value = total_sh_AE7;
                ws.Cells[8, 31].Value = total_sh_AE8;
                ws.Cells[9, 31].Value = total_sh_AE9;

                ws.Cells[7, 32].Value = total_sh_AF7- total_sh_AF8;
                ws.Cells[8, 32].Value = total_sh_AF8;
                ws.Cells[9, 32].Value = total_sh_AF9;

                ws.Cells[7, 33].Value = total_sh_AG7 - total_sh_AG8;
                ws.Cells[8, 33].Value = total_sh_AG8;
                ws.Cells[9, 33].Value = total_sh_AG9;

                ws.Cells[7, 35].Value = total_sh_AI7;
                ws.Cells[8, 35].Value = total_sh_AI8;
                ws.Cells[9, 35].Value = total_sh_AI9;

                ws.Cells[7, 36].Value = total_sh_AJ7;
                ws.Cells[8, 36].Value = total_sh_AJ8;
                ws.Cells[9, 36].Value = total_sh_AJ9;

                ws.Cells[7, 37].Value = total_sh_AK7;
                ws.Cells[8, 37].Value = total_sh_AK8;
                ws.Cells[9, 37].Value = total_sh_AK9;

                ws.Cells[7, 48].Value = total_sh_AV7;
                ws.Cells[8, 48].Value = total_sh_AV8;
                ws.Cells[9, 48].Value = total_sh_AV9;

                ws.Cells[7, 51].Value = total_sh_AY7;
                ws.Cells[8, 51].Value = total_sh_AY8;
                ws.Cells[9, 51].Value = total_sh_AY9;
                #endregion
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }
        //private void is_tar_getir()
        //{
        //    try
        //    {
        //        string tarix = txt_tarix_evvel.Text;
        //        string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
        //        using (OracleConnection connection = new OracleConnection(connectionString))
        //        {
        //            connection.Open();

        //            using (OracleCommand command = new OracleCommand("SELECT odb.ish_gun_cari1(TO_DATE('05/02/2024', 'dd/mm/yyyy')) FROM dual", connection))
        //            {
        //                // OracleDataReader kullanarak sorguyu çalıştırın
        //                using (OracleDataReader reader = command.ExecuteReader())
        //                {
        //                    // Eğer veri varsa okuyun
        //                    if (reader.Read())
        //                    {
        //                        // Tarihi TextBox'a ata
        //                        txt_tarix_son.Text = reader.GetDateTime(0).ToString("dd/mm/yyyy");
        //                    }
        //                    else
        //                    {
        //                        txt_tarix_evvel.Text = "Sonuç bulunamadı.";
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error: " + ex.Message);
        //    }
        //}
    
        private void button1_Click(object sender, EventArgs e)
        {
            //is_tar_getir();
            excel();
        }

        private void txt_tarix_evvel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_tarix_son.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_tarix_evvel_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_tarix_evvel.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_tarix_evvel.Text = yeniFormatliTarih;
                }
                else
                {
                    
                }
            }
        }

        private void txt_tarix_son_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_tarix_son.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_tarix_son.Text = yeniFormatliTarih;
                }
                else
                {
                    
                }
            }
        }

        private void txt_evvelki_ay_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_evvelki_ay.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_evvelki_ay.Text = yeniFormatliTarih;
                }
                else
                {
                    
                }
            }
        }

        private void txt_tarix_son_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_evvelki_ay.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_evvelki_ay_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void frm_iran_hesabat_Load(object sender, EventArgs e)
        {
            txt_tarix_evvel.Select();
            //textBox1.SelectAll();
        }
    }
}
