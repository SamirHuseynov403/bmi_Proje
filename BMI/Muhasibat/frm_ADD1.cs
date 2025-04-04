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
using Excel.FinancialFunctions;
using DocumentFormat.OpenXml.Presentation;
using DevExpress.CodeParser;
using DevExpress.DataProcessing.InMemoryDataProcessor;

namespace BMI.Muhasibat
{
    public partial class frm_ADD1 : Form
    {
        public frm_ADD1()
        {
            InitializeComponent();
        }
        Aletler aletler =new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        DataTable dtexcelden = new DataTable();
        private void exceldencek()
        {
            string filePath = System.IO.Path.Combine(qovluqyolu,"Fayllar", "Muhasibat", "Exceller", "siyahi.xlsx");  // Fayl yolunu dəyişdirin
            // DataTable yaratmaq
            dtexcelden.Clear();  // DataTable-dən mövcud bütün məlumatları silir
            dtexcelden.Columns.Clear(); // Sütunları silir, əgər lazım olarsa
            dtexcelden.Columns.Add("A sütunu");
            dtexcelden.Columns.Add("B sütunu");

            // Excel faylını oxumaq
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk səhifəni götürürük

                // İlk 1-ci sətirdən başlayaraq A və B sütunlarını oxuyuruq
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    string colA = worksheet.Cells[row, 1].Text;  // A sütunu
                    string colB = worksheet.Cells[row, 2].Text;  // B sütunu

                    // DataTable-ə yeni sətir əlavə edirik
                    dtexcelden.Rows.Add(colA, colB);
                }
            }

        }

        private void FIFD1()
        {
            //double muddet = Convert.ToDouble(Kat_muddet.Text);
            //double xhaq = Convert.ToDouble(Kat_kredit.Text) * 1 / 100;
            //double mebleg = Convert.ToDouble(Kat_kredit.Text) * -1;
            //double kr = mebleg + xhaq;
            //double ay = Convert.ToDouble(Kat_ayliq.Text);

            //// muddet sayına uyğun values massivini yarat
            //double[] values = new double[(int)muddet + 1];
            //values[0] = kr;  // İlk dəyər kreditin dəyəri olur
            //for (int i = 1; i <= muddet; i++)
            //{
            //    values[i] = ay;  // Sonrakı aylıq dəyərlər
            //}

            //// IRR hesablamaq
            //double FIFD = Math.Round(Financial.Irr(values) * 12 * 100, 3);
            //string netice = FIFD.ToString("F2");
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
        private void excel()
        {
            exceldencek();
            label1.Visible = true;
            label1.Text = "Hazırlanır gözləyin...";
            Application.DoEvents();

            #region ADD1
            string ADD1 = "select '\"BANK MELLİ İRAN\" BAKI FİLİALI' bank_C,r.name_regnom ad_D,case when ar.tipkredita=2 then 1 when ar.tipkredita=1 then 2 when ar.tipkredita=3 then 3 end mensubiyyet_E," +
                " c.name olke_F,case when length(ar.subschkre)=1 then '(0'||ar.subschkre||')'||l.registrac_nomer else '('||ar.subschkre||')'||l.registrac_nomer end qeydno_G," +
                " case when length(ar.subschkre)=1 then '(0'||ar.subschkre||')'||ar.licschkre else '('||ar.subschkre||')'||ar.licschkre end hesab_H," +
                " case when ar.tipkredita <>'2' then r.inn_regnom else '' end voen_I,case when ar.tipkredita ='2' then r.pincode else '' end fin_J," +
                " case when ar.tipkredita ='1' then r.rukovod else '' end rehber_K,'' payci_L,'' elaqeli_M,'yox' aidiyyatsexs_N," +
                " case when ar.graphpogkre in ('3','2') then '2' else '1'end emel_xarakter_O,ar.date_open ver_tar_P," +
                " ar.date_planclose qaytar_tar_Q,ar.date_prolong uzadilma_tar_R ,case when substr(ar.licschkre,6,2)='00' then '1'when substr(ar.licschkre,6,2)='01'" +
                "  then '2'when substr(ar.licschkre,6,2)='02' then '3'when substr(ar.licschkre,6,2)='03' then '6' else '8' end val_S,ar.procstavkre faiz_T,s.fifd fifd_U, " +
                "  ar.summakre * odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) ilkinmeb_V," +
                " (ar.summa+ar.summa_19)* odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) qaliq_W," +
                " (select t.ostatok_ish*odb.func_get_kurval(substr(t.licsch,6,2),ar.date_oper)from odb.arh_saldo_vbls t" +
                "  where t.licsch=ar.licsch_19 and t.vbs='99300' and t.ssls=ar.subschkre and t.date_oper=ar.date_oper) bk_X," +
                " ar.summa_19* odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) vk_Y," +
                " (select (v.nacpro_ish-v.pogpro_ish)+(v.nacprospro_ish-v.pogprospro_ish)from view_nacpogprokre_all v " +
                " where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre and v.licschpkre=ar.licschpkre)qaliqfaiz_Z," +
                " (select (v.nacpro_ish-v.pogpro_ish)from view_nacpogprokre_all v " +
                " where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre and v.licschpkre=ar.licschpkre)faiz_AA," +
                " (select (v.nacprospro_ish-v.pogprospro_ish)from view_nacpogprokre_all v " +
                " where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre and v.licschpkre=ar.licschpkre)vkfaiz_AB,'0' sutun_AC," +
                " '30' sutun_AD,'30' sutun_AE,ar.lgotperiod guzest_mud_AF,(select max(z.date_oper) son_tar from odb.arh_dd z" +
                " where z.kredit in(ar.licschkre, ar.licsch_19, ar.licschppkre, ar.licschpkre)and substr(z.debet, 1, 1) = '4'" +
                " and z.ssk = ar.subschkre )son_odenis_tar_AG,(SELECT SUM(z.summa_v_nacval) AS son_tar FROM odb.arh_dd z" +
                " WHERE z.kredit IN (ar.licschkre, ar.licsch_19, ar.licschppkre, ar.licschpkre)" +
                " AND SUBSTR(z.debet, 1, 1) = '4' AND z.ssk = ar.subschkre AND z.date_oper = (SELECT MAX(z1.date_oper) " +
                " FROM odb.arh_dd z1 WHERE z1.kredit IN (ar.licschkre, ar.licsch_19, ar.licschppkre, ar.licschpkre)" +
                " AND SUBSTR(z1.debet, 1, 1) = '4'AND z1.ssk = ar.subschkre))son_odenis_meb_AH," +
               "  (SELECT CASE WHEN x.lastoverduedate IS NULL THEN 0 " +
               "  ELSE odb.tar_ferq360(x.date_oper, x.lastoverduedate)END" +
               "   FROM view_nacpogprokre_all x" +
               "   WHERE ar.date_oper = x.date_oper" +
               "   AND ar.licschpkre = x.licschpkre" +
               "   AND ar.subschkre = x.subschkre) gec_gun_esas_AI," +
               "  (select CASE WHEN x.lodinterest_ish IS NULL THEN 0" +
               "  ELSE odb.tar_ferq360(x.date_oper, x.lodinterest_ish) END" +
               "  from view_nacpogprokre_all x where ar.date_oper = x.date_oper" +
               "  and ar.licschpkre = x.licschpkre and ar.subschkre = x.subschkre)gec_gun_faiz_AJ," +
                " ar.procstavrez esas_eht_AK," +
                "(ar.summa + ar.summa_19) * ar.procstavrez / 100 as esas_eht_meb_AL," +
                " ar.procstavrez_19 faiz_eht_AM," +
                " round((select (v.nacpro_ish-v.pogpro_ish)+(v.nacprospro_ish-v.pogprospro_ish)from view_nacpogprokre_all v " +
                " where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre " +
                " and v.licschpkre=ar.licschpkre)*ar.procstavrez_19/100,2) faiz_eht_meb_AN," +
                " case when ar.date_restructure is null then '0' else '1' end rest__AO," +
                " case when ar.date_restructure is not null then ar.date_restructure end rest__AP," +
                " ar.kolic_restructure rset_say_AQ," +
                " (select count(distinct l.date_restructure) from arh_licschkre l where l.date_restructure is not null " +
                " and l.licschkre=ar.licschkre and l.subschkre=ar.subschkre" +
                " AND l.date_restructure > TO_DATE('30-07-2022', 'DD-MM-YYYY')) rest_say_sertli_AR," +
                " (SELECT (l.summa + l.summa_19)* odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) AS toplam" +
                " FROM arh_licschkre l WHERE " +
                " l.licschkre = ar.licschkre AND l.subschkre = ar.subschkre" +
                " AND l.date_oper = (SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre)) son_rest_qal_AS," +
                " (SELECT (v.nacpro_ish-v.pogpro_ish)+(v.nacprospro_ish-v.pogprospro_ish) AS toplam" +
                " FROM arh_licschkre l,view_nacpogprokre_all v" +
                " WHERE l.date_restructure = (SELECT MAX(l2.date_restructure)" +
                " FROM arh_licschkre l2 WHERE l2.date_restructure IS NOT NULL" +
                " AND l2.licschkre = ar.licschkre AND l2.subschkre = ar.subschkre)" +
                " AND l.licschkre = ar.licschkre AND l.subschkre = ar.subschkre" +
                " AND l.date_oper = (SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre)and v.licschpkre=l.licschpkre and v.subschkre=l.subschkre" +
                "  and v.date_oper=(SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre))son_rest_faiz_qal_AT," +
                " (SELECT case when ar.procstavrez ='1' then '1'" +
                " when (ar.procstavrez='2' and substr(ar.licschkre,6,2)<>'00') then '1' " +
                " when (ar.procstavrez='5' and ar.tipkredita='2' and substr(ar.licschkre,6,2)='00') then '2'" +
                " when (ar.procstavrez='10'and ar.tipkredita='2'and substr(ar.licschkre,6,2)<>'00')then '2'" +
                " when (ar.procstavrez='2'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '2'" +
                " when (ar.procstavrez='3'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)<>'00')then '2'" +
                " when (ar.procstavrez='2'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '2'" +
                " when (ar.procstavrez='2'and ar.index_otrasli='01902')then '2'" +
                " when (ar.procstavrez='15'and ar.tipkredita='2'and substr(ar.licschkre,6,2)='00')then '3'" +
                " when (ar.procstavrez='20'and ar.tipkredita='2'and substr(ar.licschkre,6,2)<>'00')then '3'" +
                " when (ar.procstavrez='10'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '3'" +
                " when (ar.procstavrez='12'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)<>'00')then '3'" +
                " when (ar.procstavrez='10'and ar.index_otrasli='01902')then '3'" +
                " when (ar.procstavrez='25')then '6'when (ar.procstavrez='50')then '5'" +
                " when (ar.procstavrez='100')then '4' end rst_tar_tes_AU   " +
                " FROM arh_licschkre l WHERE l.date_restructure = (SELECT MAX(l2.date_restructure)" +
                " FROM arh_licschkre l2 WHERE l2.date_restructure IS NOT NULL" +
                " AND l2.licschkre = ar.licschkre AND l2.subschkre = ar.subschkre)" +
                " AND l.licschkre = ar.licschkre AND l.subschkre = ar.subschkre" +
                " AND l.date_oper = (SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre)) son_rest_qal_AS," +
                " (SELECT odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate,x.date_oper)) AS toplam" +
                " FROM arh_licschkre l,view_nacpogprokre_all x WHERE l.date_restructure = (SELECT MAX(l2.date_restructure)" +
                " FROM arh_licschkre l2 WHERE l2.date_restructure IS NOT NULL" +
                " AND l2.licschkre = ar.licschkre AND l2.subschkre = ar.subschkre)" +
                " AND l.licschkre = ar.licschkre AND l.subschkre = ar.subschkre" +
                " AND l.date_oper = (SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre) and x.licschpkre=ar.licschpkre and x.subschkre=ar.subschkre" +
                "  and x.date_oper=(SELECT MAX(l3.date_restructure)" +
                " FROM arh_licschkre l3 WHERE l3.licschkre = ar.licschkre" +
                " AND l3.subschkre = ar.subschkre)) son_rest_gecgun_AV," +
                " case when (ar.tipkredita='2'and ar.index_otrasli<>'01902') then '1'" +
                " when (ar.tipkredita<>'2' and ar.index_otrasli<>'01902') then '2'" +
                " when (ar.index_otrasli='01902') then '3' end kre_tip_AW," +
                " case when (ar.tipkredita='2') then s.aylig_gelir end xalis_gelir_AX," +
                " case when (ar.tipkredita='2') then s.aylig_gelir_xv end xalis_val_gelir_AY," +
                " case when (ar.tipkredita='2') then s.aylig_borc_yuku end ist_borc_AZ," +
                " case when ar.graphpogkre not in ('2','3') then" +
                " round((ar.summakre * (ar.procstavkre/12/100)) / (1 - POWER((1 + (ar.procstavkre/12/100)), -(ar.srok / 30))),2)" +
                " else 0 end ayliq_od_BA,'0' yigim_sigorta_BB," +
                " (select (case when ((max( odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate,x.date_oper)))>30)" +
                " or max( odb.tar_ferq360(x.date_oper, nvl(x.lodinterest_ish,x.date_oper)))>30) then '1' else '0' end) sorgu" +
                " from view_nacpogprokre_all x" +
                " where x.date_oper between ADD_MONTHS(to_date('"+txt_hes_tar.Text+"','dd/mm/yyyy'), -6) and to_date('"+txt_hes_tar.Text+"','dd/mm/yyyy')" +
                " and x.licschpkre=ar.licschpkre and x.subschkre=ar.subschkre) son_6_ay_vk_BC," +
                " case when ar.tipkredita in ('1','3') and ar.licschkre<>'20660000001366100000' then '1'" +
                "   when ar.tipkredita='1' and ar.licschkre='20660000001366100000' then '2' else '' end sub_bolgusu_sutun_BD," +
                " case when ar.procstavrez ='1' then '1'when (ar.procstavrez='2' and substr(ar.licschkre,6,2)<>'00') then '1' " +
                " when (ar.procstavrez='5' and ar.tipkredita='2' and substr(ar.licschkre,6,2)='00') then '2'" +
                " when (ar.procstavrez='10'and ar.tipkredita='2'and substr(ar.licschkre,6,2)<>'00')then '2'" +
                " when (ar.procstavrez='2'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '2'" +
                " when (ar.procstavrez='3'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)<>'00')then '2'" +
                " when (ar.procstavrez='2'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '2'" +
                " when (ar.procstavrez='2'and ar.index_otrasli='01902')then '2'" +
                " when (ar.procstavrez='15'and ar.tipkredita='2'" +
                " and substr(ar.licschkre,6,2)='00')then '3'" +
                " when (ar.procstavrez='20'and ar.tipkredita='2'and substr(ar.licschkre,6,2)<>'00')then '3'" +
                " when (ar.procstavrez='10'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)='00')then '3'" +
                " when (ar.procstavrez='12'and ar.tipkredita<>'2'and substr(ar.licschkre,6,2)<>'00')then '3'" +
                " when (ar.procstavrez='10'and ar.index_otrasli='01902')then '3'" +
                " when (ar.procstavrez='25')then '6'" +
                " when (ar.procstavrez='50')then '5'" +
                " when (ar.procstavrez='100')then '4' end key_mey_BE," +
                " case when (ar.tipkredita <> '2' and ar.procstavrez <25) then '1'" +
                " when (ar.tipkredita <> '2' and ar.procstavrez >=25) then '0'" +
                " else '' end maliy_vez_BF," +
                " case when (ar.tipkredita <> '2' and ar.procstavrez <25) then '1'" +
                " when (ar.tipkredita <> '2' and ar.procstavrez >=25) then '0'" +
                " else '' end risk_BG," +
                "   '1' sutun_BH,  '0' sutun_BI," +
                " case when (ar.tipkredita <> '2' and ar.procstavrez <25) then '1'" +
                " when (ar.tipkredita <> '2' and ar.procstavrez >=25) then '0'" +
                " else '' end sutun_BJ,case when (ar.tipkredita <> '2' and ar.procstavrez <25) then '1'" +
                " when (ar.tipkredita <> '2' and ar.procstavrez >=25) then '0'else '' end sutun_BK," +
                " '' sutun_BL,'' sutun_BM,'' sutun_BN,'' sutun_BO,'' sutun_BP,'' sutun_BQ,'' sutun_BR," +
                " case when ( ar.procstavrez <25) then '1'when (ar.tipkredita <> '2' and ar.procstavrez >=25) then '0'" +
                " else '' end sutun_BS,'0' sutun_BT,'ayri selectden atacam' sutun_BU,'ayri selectden atacam' sutun_BV," +
                " case when s.item_01='İCRADA' then '1'when s.item_01='QƏTNAMƏ' then '2' else '3' end mehkeme_BW," +
                " case when (ar.tipzaloga in (1,2,3,6,9,16) and ((ar.summa+ar.summa_19)*150/100)<coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga))" +
                " then '1' when (ar.tipzaloga in (1,2,3,6,9,16) and ((ar.summa+ar.summa_19)*150/100)>=coalesce(ar.summa_pereocen_zaloga," +
                " ar.summa_zaloga))then '2' else '1' end akt_novu_B,case when ar.tipzaloga=3 then coalesce(ar.summa_pereocen_zaloga," +
                " ar.summa_zaloga)* odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga,ar.licschkre),6,2),ar.date_oper) else 0 end qrup_1_BY," +
                "'0' sutun_BZ,'0' sutun_CA,'' sutun_CB," +
                " case when ar.tipzaloga=16 then coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga)* odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga," +
                " ar.licschkre),6,2),ar.date_oper) end qrup_2_CC,'0' sutun_CD,case when s.item_20='q/y' then coalesce(ar.summa_pereocen_zaloga," +
                " ar.summa_zaloga)* odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga,ar.licschkre),6,2),ar.date_oper) else 0 end qrup_3_CE," +
                " case when (ar.tipzaloga=1) then coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga)* odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga," +
                " ar.licschkre),6,2),ar.date_oper) else 0 end-case when s.item_20='q/y' " +
                " then coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga)* odb.func_get_kurval(substr(ar.licschkre,6,2)," +
                " ar.date_oper) else 0 end qrup_3_CF,case when ar.tipzaloga=9 then coalesce(ar.summa_pereocen_zaloga," +
                " ar.summa_zaloga)* odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga,ar.licschkre),6,2),ar.date_oper) else 0 " +
                " end qrup_4_CG,case when ar.tipzaloga not in (1,3,9,16) then coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga)" +
                " * odb.func_get_kurval(substr(coalesce(ar.licsch_zaloga,ar.licschkre),6,2),ar.date_oper) else 0 end qrup_5_CH," +
                " case when ar.tipzaloga not in (1,3,9,16) then t.name else '0' end qrup_5_CI," +
                " case when (ar.tipzaloga in (1,2,3,6,9,16))then ar.date_open end son_qiy_tar_CJ," +
                " case when (ar.tipzaloga in (1,2,3,6,9,16))then ar.data_pereocen_zaloga end son_qiy_tar_CK," +
                " case when (ar.tipzaloga in (1,2,3,6,9,16))then s.asas_qiymatlendirici_adi end son_qiy_sirk_CL," +
                " n.name teyinat_CM," +
                " i.name_index_otrasli saheler_CN ," +
                "'' sutun_CO,'1' maliy_menbeyi_CP,'' sutun_CQ,'' sutun_CR," +
                " '' sutun_CS,'' sutun_CT,'' sutun_CU from arh_licschkre ar,regnom r,licsch l,countrycode c,srokpogprockre s," +
                " tipzal t,naznackredita n,index_otrasli i where ar.licschkre=l.licsch and l.registrac_nomer=r.regnom " +
                " and ar.date_oper = to_date('" + txt_hes_tar.Text + "','dd/mm/yyyy') and t.code=ar.tipzaloga and ar.naznackredita=n.code" +
                " and ar.date_close is null and ar.date_oper=to_date('"+txt_hes_tar.Text+ "','dd/mm/yyyy') and ar.index_otrasli=i.index_otrasli" +
                " and l.countrycode=c.code and ar.subschkre=s.subschkre and ar.licschkre=s.licschkre order by substr(ar.licschkre,2,5) asc";
            #endregion
            #region gecgun6
            string gecgun6 = "WITH max_gecgun_per_month " +
" AS (SELECT case when length(x.subschkre)=1 then '(0'||x.subschkre||')'||substr(x.licschpkre,10,6) else '('||x.subschkre||')'||substr(x.licschpkre,10,6) end reg, " +
" x.licschpkre AS hes,x.subschkre AS sk, TO_CHAR(x.date_oper, 'MM-YYYY') AS tarix,MAX(odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper))) AS max_gecgun FROM view_nacpogprokre_all x, arh_licschkre al " +
" WHERE x.date_oper BETWEEN ADD_MONTHS(TO_DATE('" + txt_hes_tar.Text + "', 'DD-MM-YYYY'), -6) AND TO_DATE('" + txt_hes_tar.Text + "', 'DD-MM-YYYY')" +
" AND x.licschpkre = al.licschpkre AND x.subschkre = al.subschkre AND al.date_oper = TO_DATE('30-09-2024', 'DD-MM-YYYY')" +
" AND al.date_close IS NULL GROUP BY x.licschpkre, x.subschkre, TO_CHAR(x.date_oper, 'MM-YYYY'))" +
" SELECT hes, sk,reg,case when COUNT(CASE WHEN max_gecgun > 0 THEN 1 END)>0 then COUNT(CASE WHEN max_gecgun > 0 THEN 1 END) else 0 end AS gecgun_sayi" +
" FROM max_gecgun_per_month GROUP BY hes,reg, sk ORDER BY hes, sk";
            #endregion
            #region gecgun12
            string gecgun12 = "WITH max_gecgun_per_month " +
" AS (SELECT case when length(x.subschkre)=1 then '(0'||x.subschkre||')'||substr(x.licschpkre,10,6) else '('||x.subschkre||')'||substr(x.licschpkre,10,6) end reg, " +
" x.licschpkre AS hes,x.subschkre AS sk, TO_CHAR(x.date_oper, 'MM-YYYY') AS tarix,MAX(odb.tar_ferq360(x.date_oper, NVL(x.lastoverduedate, x.date_oper))) AS max_gecgun FROM view_nacpogprokre_all x, arh_licschkre al " +
" WHERE x.date_oper BETWEEN ADD_MONTHS(TO_DATE('" + txt_hes_tar.Text + "', 'DD-MM-YYYY'), -12) AND TO_DATE('" + txt_hes_tar.Text + "', 'DD-MM-YYYY')" +
" AND x.licschpkre = al.licschpkre AND x.subschkre = al.subschkre AND al.date_oper = TO_DATE('" + txt_hes_tar.Text + "', 'DD-MM-YYYY')" +
" AND al.date_close IS NULL GROUP BY x.licschpkre, x.subschkre, TO_CHAR(x.date_oper, 'MM-YYYY'))" +
" SELECT hes, sk,reg,case when COUNT(CASE WHEN max_gecgun > 0 THEN 1 END)>0 then COUNT(CASE WHEN max_gecgun > 0 THEN 1 END) else 0 end AS gecgun_sayi" +
" FROM max_gecgun_per_month GROUP BY hes,reg, sk ORDER BY hes, sk";
            #endregion
            #region ADD1_2
             string ADD1_2 = "select '\"BANK MELLİ İRAN\" BAKI FİLİALI' bank_B,r.name_regnom ad_C,'Vuqar' sutun_D," +
                " case when ar.tipkredita in ('1','3') then '3'when ar.tipkredita in ('2') then '4' end sutun_D," +
                " coalesce(r.inn_regnom,r.pincode) sutun_E,case when ar.tipkredita in ('1') then r.rukovod end sutun_F," +
                " (select case when ar.tipkredita in ('1') and i.tesischinin_payi>10 then i.soyadi||' '||i.adi||' '||i.ata_adi||' '||i.tesischinin_payi end " +
                " from imza_huquqi_olan_shexsler i where r.regnom=i.regnom ) sutun_G,'1' sutun_J,'0' sutun_K," +
                " ar.date_open sutun_L,ar.date_planclose sutun_M,ar.date_prolong sutun_N,ar.kolic_prolong sutun_O," +
                " case when substr(ar.licschkre,6,2)='00' then '1'when substr(ar.licschkre,6,2)='01' then '2'" +
                " when substr(ar.licschkre,6,2)='02' then '3'when substr(ar.licschkre,6,2)='03' then '6' else '8' end sutun_P," +
                " ar.summakre*odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) sutun_Q,(ar.summa+ar.summa_19)*odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper) sutun_R," +
                " 'el ile yaz' sutun_S,(select (v.nacpro_ish-v.pogpro_ish)+(v.nacprospro_ish-v.pogprospro_ish)from view_nacpogprokre_all v " +
                " where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre and v.licschpkre=ar.licschpkre)sutun_T," +
                " ar.procstavrez sutun_U,((ar.summa+ar.summa_19)*odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper))*ar.procstavrez/100 sutun_V," +
                " ar.procstavrez_19 sutun_W,(select ((v.nacpro_ish-v.pogpro_ish)+(v.nacprospro_ish-v.pogprospro_ish))*odb.func_get_kurval(substr(ar.licschkre,6,2),ar.date_oper)" +
                " from view_nacpogprokre_all v where v.date_oper=ar.date_oper and v.subschkre=ar.subschkre and v.licschpkre=ar.licschpkre)*ar.procstavrez_19/100 sutun_X," +
                " coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga) sutun_Y,coalesce(ar.summa_pereocen_zaloga,ar.summa_zaloga) sutun_Z,case when s.item_20='q/y' then '4'" +
                " when s.item_20='y' then '1' end sutun_AA from arh_licschkre ar,regnom r,licsch l,countrycode c,srokpogprockre s,tipzal t,naznackredita n" +
                " where ar.licschkre=l.licsch and l.registrac_nomer=r.regnom and ar.date_oper = to_date('" + txt_hes_tar.Text + "','dd/mm/yyyy') and t.code=ar.tipzaloga and ar.naznackredita=n.code" +
                " and ar.date_close is null and ar.date_oper=to_date('" + txt_hes_tar.Text + "','dd/mm/yyyy') and ar.graphpogkre=2" +
                " and l.countrycode=c.code and ar.subschkre=s.subschkre and ar.licschkre=s.licschkre";

            #endregion
            #region Likvid_azalma
            string likvid= "select case when length(ar.subschkre)=1 then '(0'||ar.subschkre||')'||ar.licschkre else '('||ar.subschkre||')'||ar.licschkre end hesab, "+
            " ar.subschkre sk, max(g.tarix) tarix,ADD_MONTHS(to_date(max('"+txt_hes_tar.Text+"'), 'dd/mm/yyyy'), +36) son_3_il,s.girovun_likvidlik_faiz_derece likvid_faiz "+
            " ,case when " +
            " s.girovun_likvidlik_faiz_derece IS NOT NULL then 1 else 0 end sert " +
            " from arh_licschkre ar,girovun_bazar_deyeri g, srokpogprockre s where ar.date_oper = to_date('"+txt_hes_tar.Text+"', 'dd/mm/yyyy') and ar.date_close is null " +
            " and ar.licschkre = g.licschkre and ar.subschkre = g.subschkre and s.licschkre = ar.licschkre and s.subschkre = ar.subschkre " +
            " and g.tarix <= to_date('"+txt_hes_tar.Text+"', 'dd/mm/yyyy') " +
            " and s.girovun_likvidlik_faiz_derece is not null group by ar.licschkre,ar.subschkre,s.girovun_likvidlik_faiz_derece,ar.date_planclose";
            #endregion
            System.Data.DataTable _dt_ADD1 = new System.Data.DataTable();
            System.Data.DataTable _dt_gecgun_6 = new System.Data.DataTable();
            System.Data.DataTable _dt_gecgun_12 = new System.Data.DataTable();
            System.Data.DataTable _dt_ADD1_2 = new System.Data.DataTable();
            System.Data.DataTable _dt_likvid = new System.Data.DataTable();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(ADD1, connection))
                {

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_ADD1);
                }
                using (OracleCommand command = new OracleCommand(gecgun6, connection))
                {

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_gecgun_6);
                }

                using (OracleCommand command = new OracleCommand(gecgun12, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_gecgun_12);
                }
                using (OracleCommand command = new OracleCommand(ADD1_2, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_ADD1_2);
                }
                using (OracleCommand command = new OracleCommand(likvid, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_likvid);
                }
                connection.Close();
            }

            //DataTable filteredDataTableproqnoz = _dt_ADD1.Clone(); // İlk tablonun şemasını kopyala
            //foreach (var row in _dt_ADD1)
            //{
            //    filteredDataTableproqnoz.Rows.Add(row.ItemArray);
            //}

            //// DataGridView'e yeni DataTable'ı atayarak güncelle
            //dataGridView1.DataSource = filteredDataTableproqnoz;
            dataGridView1.DataSource = _dt_ADD1;
            //dataGridView2.DataSource = _dt_likvid;
            //await Task.Delay(100);

            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txt_hes_tar.Text; // textBoxText dəyəriniz
            //DateTime dateValue = DateTime.ParseExact(textBoxText, "dd-MM-yyyy", null);
            string monthYear = textBoxText.Replace("-", "");
            //monthYear = dateValue.ToString("MMyyyy"); // Ay və il formatı "092024"
            cl.baseFileName = "ADD1.v01.1124m" + monthYear.Substring(2, 6); // Temel dosya adı
            cl.fileName = cl.baseFileName + ".xlsx";
            cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller","ADD1.xlsx");
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
            //"15020",
            FileInfo templateFile = new FileInfo(cl.templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["KT_A7"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["BK_A4"];
                

                int startRow = 9;  // C9 hüceyrəsindən başlamaq üçün
                int startColumn = 3;  // C sütunundan başlayacaq (3-cü sütun)
                worksheet1.Cells[3, 3].Value = txt_hes_tar.Text;

                int startRow2 = 8;  // C9 hüceyrəsindən başlamaq üçün
                int startColumn2 = 2;  // C sütunundan başlayacaq (3-cü sütun)

                // DataTable'dan məlumatları Excel-ə yazmaq
                for (int i = 0; i < _dt_ADD1.Rows.Count; i++)
                {
                    for (int j = 0; j < _dt_ADD1.Columns.Count; j++)
                    {
                        var cellValue = _dt_ADD1.Rows[i][j];

                        // Tarix formatındakı məlumatları saat olmadan yazmaq üçün yoxlama
                        if (cellValue is DateTime)
                        {
                            worksheet1.Cells[startRow + i, startColumn + j].Value = ((DateTime)cellValue).ToString("dd-MM-yyyy");
                        }
                        else if (double.TryParse(cellValue.ToString(), out double numericValue))
                        {
                            // Rəqəmlər üçün formatlama
                            worksheet1.Cells[startRow + i, startColumn + j].Value = numericValue;
                            worksheet1.Cells[startRow + i, startColumn + j].Style.Numberformat.Format = "0";  // Say formatı
                        }
                        else
                        {
                            worksheet1.Cells[startRow + i, startColumn + j].Value = cellValue.ToString();
                        }
                    }


                    // Əgər _dt_ADD1 datatablenin 18-ci sütunu boşdursa
                    if (_dt_ADD1.Rows[i][18] == DBNull.Value || string.IsNullOrEmpty(_dt_ADD1.Rows[i][18].ToString()))
                    {
                        DateTime startDate = Convert.ToDateTime(_dt_ADD1.Rows[i][13]);
                        DateTime endDate = Convert.ToDateTime(_dt_ADD1.Rows[i][14]);

                        // Tarixlər arasındakı ay fərqini hesablayırıq
                        double muddet = ((endDate.Year - startDate.Year) * 12) + endDate.Month - startDate.Month;
                        double faiz = Convert.ToDouble(_dt_ADD1.Rows[i][17]) ;
                        double xhaq = Convert.ToDouble(_dt_ADD1.Rows[i][19]) * 1 / 100;
                        double mebleg = Convert.ToDouble(_dt_ADD1.Rows[i][19]) * -1;
                        double kr = mebleg + xhaq;

                        // Aylıq ödənişi hesablayırıq: aylıq faiz dərəcəsi və müddətə görə
                        //double ay = -(mebleg * faiz) / (1 - Math.Pow(1 + faiz, -muddet));

                        double ay = Convert.ToDouble(calcPayment(-mebleg, muddet, faiz));
                        //string str = pmt.ToString("f2");

                        // muddet sayına uyğun values massivini yaradırıq
                        double[] values = new double[(int)muddet + 1];
                        values[0] = kr;  // İlk dəyər kreditin dəyəri olaraq mənfi olur
                        for (int t = 1; t <= muddet; t++)
                        {
                            values[t] = ay;  // Sonrakı aylıq dəyərlər müsbət olur
                        }

                        // IRR hesablanmasını yoxlayırıq
                        bool hasPositive = values.Any(v => v > 0);
                        bool hasNegative = values.Any(v => v < 0);

                        if (hasPositive && hasNegative)
                        {
                            try
                            {
                                double FIFD = Math.Round(Financial.Irr(values) * 12 * 100, 3);
                                string netice = FIFD.ToString("F2");

                                // Nəticəni lazım olan yerə yazın, məsələn, Excel hüceyrəsinə
                                //worksheet1.Cells[startRow + i, 21].Value = netice; // Misal olaraq
                                worksheet1.Cells[startRow + i, 21].Value = FIFD; // Dəyəri birbaşa atırıq
                                worksheet1.Cells[startRow + i, 21].Style.Numberformat.Format = "0"; // Heç bir onluq rəqəm göstərmir
                                //worksheet1.Cells[startRow + i, 21].Formula = string.Empty; // Formulanı silir ki, dəyər tam olaraq göstərilsin
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("IRR hesablanması zamanı xəta: " + ex.Message);
                            }
                        }
                        else
                        {
                            //Console.WriteLine("IRR hesablanması üçün values massivində həm müsbət, həm də mənfi dəyərlər olmalıdır.");
                        }
                    }

                    // BV sütununa _dt_gecgun_12 ilə uyğun məlumatı yazmaq
                    var matchRow_gecgun_12 = _dt_gecgun_12.AsEnumerable()
                        .FirstOrDefault(row => row.Field<string>(2) == _dt_ADD1.Rows[i][4].ToString());

                    if (matchRow_gecgun_12 != null)
                    {
                        worksheet1.Cells[startRow + i, 73].Value = matchRow_gecgun_12[3].ToString();  // BV sütunu (73-cü sütun)
                    }
                    else
                    {
                        worksheet1.Cells[startRow + i, 73].Value = ""; // Uygun gəlməyənlər üçün boş
                    }

                    // BW sütununa _dt_gecgun_6 ilə uyğun məlumatı yazmaq
                    var matchRow_gecgun_6 = _dt_gecgun_6.AsEnumerable()
                        .FirstOrDefault(row => row.Field<string>(2) == _dt_ADD1.Rows[i][4].ToString());

                    if (matchRow_gecgun_6 != null)
                    {
                        worksheet1.Cells[startRow + i, 74].Value = matchRow_gecgun_6[3].ToString();  // BW sütunu (74-cü sütun)
                    }
                    else
                    {
                        worksheet1.Cells[startRow + i, 74].Value = ""; // Uygun gəlməyənlər üçün boş
                    }

                    // Burada yeni əməliyyatları əlavə edirik
                    var likvidRows = _dt_likvid.AsEnumerable()
                        .Where(rowLikvid =>
                        {
                            bool isValidDate = false;
                            int sert = 0; // sert üçün default dəyər

                            try
                            {
                                // Sert dəyərini int formatına çevirməyi yoxlayırıq
                                sert = int.Parse(rowLikvid[5].ToString());
                                isValidDate = true;
                            }
                            catch (FormatException)
                            {
                                // Format uyğunsuz olsa, isValidDate false qalır və şərtə daxil edilmir.
                            }

                            // Sert 1 olmalıdır, digər şərtlər də yoxlanılır
                            return rowLikvid[0].ToString() == _dt_ADD1.Rows[i][5].ToString() && isValidDate && sert == 1;
                        });

                    if (likvidRows.Any())
                    {
                        // Şərt uyğundur, toplama və azalma əməliyyatını yerinə yetiririk
                        decimal qaliq = Convert.ToDecimal(_dt_ADD1.Rows[i][20]);
                        decimal sumADD1 = Convert.ToDecimal(_dt_ADD1.Rows[i][20]) + Convert.ToDecimal(_dt_ADD1.Rows[i][23]);
                        decimal faiz = Convert.ToDecimal(likvidRows.First()[4]);
                        decimal result = qaliq - (sumADD1 * (faiz / 100));

                        // Nəticəni Excel-in 38-ci sütununa yazırıq
                        worksheet1.Cells[startRow + i, 38].Value = result;
                        // Yeni DataTable yaradılır (DataGridView-ə yazmaq üçün)
                        DataTable table = new DataTable();
                        table.Columns.Add("Sum ADD1", typeof(decimal));
                        table.Columns.Add("Faiz", typeof(decimal));
                        table.Columns.Add("Result", typeof(decimal));

                        // Şərt yerinə yetirilən zaman nəticələri DataTable-a əlavə edirik
                        if (likvidRows.Any())
                        {
                            // Toplama və azalma əməliyyatını yerinə yetiririk
                             sumADD1 = Convert.ToDecimal(_dt_ADD1.Rows[i][20]) + Convert.ToDecimal(_dt_ADD1.Rows[i][23]);
                             faiz = Convert.ToDecimal(likvidRows.First()[4]);
                             result = sumADD1 - (sumADD1 * (faiz / 100));

                            // DataTable-a yeni sətir əlavə edirik
                            table.Rows.Add(sumADD1, faiz, result);
                        }
                        // Sonra table-i DataGridView-ə təyin edirik
                        dataGridView2.DataSource = dtexcelden;
                    }
                    else
                    {
                        // Şərt uyğun gəlmir, 34-cü sütunu Excel-in 38-ci sütununa yazırıq
                        worksheet1.Cells[startRow + i, 38].Value = _dt_ADD1.Rows[i][35];
                    }
                    // ** Əlavə edilən hissə **: 
                    // _dt_ADD1-in 87-ci sütunu ilə (index 91) dtexcelden-in 0-cı sütunu uyğun gəlirsə, 1-ci sütunu Excel-ə yaz
                    var likvidRow = dtexcelden.AsEnumerable()
                        .FirstOrDefault(row => row[0].ToString() == _dt_ADD1.Rows[i][89].ToString());

                    if (likvidRow != null)
                    {
                        // _dt_ADD1-in 91-ci sütununa dtexcelden-in uyğun gələn 1-ci sütun dəyərini yazırıq
                        _dt_ADD1.Rows[i][91] = likvidRow[1].ToString();
                        worksheet1.Cells[startRow + i, 92].Value = likvidRow[1].ToString();
                    }
                }

                // İkinci DataTable olan _dt_ADD1_2 üçün məlumatları Excel-ə yazmaq
                if (_dt_ADD1_2.Rows.Count > 0)  // Cədvəl boşdursa əməliyyat etməyəcək
                {
                    for (int i = 0; i < _dt_ADD1_2.Rows.Count; i++)
                    {
                        for (int j = 0; j < _dt_ADD1_2.Columns.Count; j++)
                        {
                            var cellValue = _dt_ADD1_2.Rows[i][j];

                            // Tarix formatındakı məlumatları saat olmadan yazmaq üçün yoxlama
                            if (cellValue is DateTime)
                            {
                                worksheet2.Cells[startRow2 + i, startColumn2 + j].Value = ((DateTime)cellValue).ToString("dd-MM-yyyy");
                            }
                            else if (double.TryParse(cellValue.ToString(), out double numericValue))
                            {
                                worksheet2.Cells[startRow2 + i, startColumn2 + j].Value = numericValue;
                                worksheet2.Cells[startRow2 + i, startColumn2 + j].Style.Numberformat.Format = "0";  // Say formatı
                            }
                            else
                            {
                                worksheet2.Cells[startRow2 + i, startColumn2 + j].Value = cellValue.ToString();
                            }
                        }
                    }
                }

                // Sətir nömrələrini B sütununda yazmaq
                for (int i = 0; i < _dt_ADD1.Rows.Count; i++)
                {
                    worksheet1.Cells[startRow + i, 2].Value = (i + 1).ToString();  // B sütununda sətir nömrələri
                }

                label1.Text = "Hazırdır.";
                Application.DoEvents();
                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(cl.filePath);
                label1.Visible = false;
                Application.DoEvents();
            }
        }
        private void txt_hes_tar_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_hes_tar.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_hes_tar.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txt_hes_tar_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                btn_sorgu.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }
        private async Task excelAsync()
        {
           
        }

        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            excel();
        }
    }
}
