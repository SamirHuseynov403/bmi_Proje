using DevExpress.XtraEditors;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Vml;
using OfficeOpenXml;
using OfficeOpenXml.ExternalReferences;
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
    //github ucun
    public partial class frm_XBBIS : Form
    {
        public frm_XBBIS()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        string dosyayolu;
        string textBoxText; // TextBox'tan alınan metni sakla
        string yeniMetin;
        string baseFileName; // Temel dosya adı
        string fileName;
        string templateFilePath;
        string filePath;
        private void excel()
        {
            string asa = txt_dov_evvel.Text;
            string asa1 = txt_dov_son.Text;
            #region dovriyye_azalma
            string dovriyye_azalma = "select t.tam_hes,t.balans,t.qeyd_no,t.val,t.G_sut,t.K_sut,NULL as ac_tar,NULL as bag_tar,NULL as muddet,NULL as faiz,NULL as D_suda,NULL as D_faiz,NULL as D_vk,NULL as D_vk_faiz,\r\nt.gir_qal,nvl( t.d_dovr,0)d_dovr,nvl( t.k_dovr,0)k_dovr,t.son_qal ,NULL as gir_qal_suda,NULL as gir_qal_faiz, NULL as gir_qal_vk, NULL as gir_qal_vkfaiz\r\nfrom (select qaliq.tam_hes,qaliq.balans,\r\ncase when substr(qaliq.balans,1,5) in ('10080','10020') then  'nağd xarici valyuta (aktiv üzrə)'\r\nwhen substr(qaliq.balans,1,2) in ('15','35') then  'tələbli depozit (aktiv üzrə)' else 'kredit (aktiv üzrə)' end G_sut,\r\ncase when substr(qaliq.balans,1,5) in ('10080','10020') then  'Mərkəzi bank'\r\nwhen substr(qaliq.balans,1,2) in ('15','35') then 'Bank' else 'Ev təsərrüfatı' end K_sut,\r\nqaliq.qeyd_no,qaliq.val,qaliq.gir_qal,\r\n(select nvl( sum( case when substr(ar.debet,6,2) not in ('00')  then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar where ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy')\r\nand substr(ar.debet, 1, 5) = qaliq.balans and substr(ar.debet,6,2)=qaliq.val and ar.debet=qaliq.tam_hes\r\nand substr(ar.kredit, 1, 2) not in ('86', '66') group by ar.debet) d_dovr,\r\n(select nvl( sum( case when substr(ar.kredit,6,2) not in ('00') then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar where ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+ "','dd/mm/yyyy')\r\nand substr(ar.kredit, 1, 5) = qaliq.balans and substr(ar.kredit,6,2)=qaliq.val and ar.kredit=qaliq.tam_hes\r\nand substr(ar.debet, 1, 2) not in ('86', '66')group by ar.kredit) k_dovr,\r\nqaliq.son_qal from \r\n(select ac.name_latin,p.licsch tam_hes,substr(p.licsch,1,5) balans,substr(p.licsch,10,6) qeyd_no,substr(p.licsch,6,2) val, p.gir_qaliq gir_qal,k.son_qaliq son_qal from licsch lc,\r\n(select t.licsch, case when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_ish_nacval/1000)\r\nelse abs(t.saldo_vhd_inval/1000)end gir_qaliq\r\nfrom odb.arh_saldo_ls t where t.date_oper=to_date( '" + txt_dov_evvel.Text + "','dd/mm/yyyy'  )) p, \r\n(select t.licsch,case when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_ish_nacval/1000)\r\nelse abs(t.saldo_vhd_inval/1000) end son_qaliq\r\nfrom odb.arh_saldo_ls t where t.date_oper=to_date( '" + txt_dov_evvel.Text + "','dd/mm/yyyy'  )) k,odb.accounts ac\r\nwhere  p.licsch=k.licsch and p.licsch=ac.licsch  and p.licsch=lc.licsch and  (lc.date_close_licsch is null or lc.date_close_licsch > to_date( '" + txt_dov_evvel.Text + "','dd/mm/yyyy'  ))\r\norder by substr(p.licsch,1,5),substr(p.licsch,6,2)) qaliq,\r\n(select substr(l.licsch,1,5) hesab,substr(l.licsch,6,2) val,substr(l.licsch,10,6) qeyd_no,\r\nsum(l.oboroti_debet_nacval/1000) Dovriyye_debet_uzre,sum(l.oboroti_kredit_nacval/1000) Dovriyye_kredit_uzre\r\nfrom odb.arh_saldo_ls l where l.date_oper between to_date('" + txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy')\r\ngroup by substr(l.licsch,1,5),substr(l.licsch,10,6),substr(l.licsch,6,2)) dovr\r\nwhere dovr.hesab=qaliq.balans and dovr.qeyd_no=qaliq.qeyd_no and dovr.val=qaliq.val(+)\r\nand qaliq.balans in (10020, 10080, 10089, 15025, 15225, 15227 )\r\nand (qaliq.gir_qal>0 or dovr.Dovriyye_debet_uzre>0 or dovr.Dovriyye_kredit_uzre>0 or qaliq.son_qal>0) order by qaliq.balans,qaliq.qeyd_no,qaliq.val asc)t\r\nUNION ALL\r\nselect p.tam_hes,p.balans,p.qeyd_no,p.val,p.g_sut,p.k_sut,p.ac_tar,p.bag_tar,p.muddet,p.faiz,p.D_suda,p.D_faiz,p.D_vk,p.D_vk_faiz,\r\np.gir_qal,p.d_dovr,p.k_dovr,p.son_qal,p.gir_qal_suda,p.gir_qal_faiz,p.gir_qal_vk,p.gir_qal_vkfaiz from \r\n(select NULL as tam_hes,m.bal balans,m.qeydno qeyd_no,m.val,\r\ncase when substr(m.bal,1,5) in ('10080','10020') then  'nağd xarici valyuta (aktiv üzrə)'\r\nwhen substr(m.bal,1,2) in ('15','35') then  'tələbli depozit (aktiv üzrə)' else 'Ev təsərrüfatı' end G_sut,\r\ncase when substr(m.bal,1,5) in ('10080','10020') then  'Mərkəzi bank'\r\nwhen substr(m.bal,1,2) in ('15','35') then 'Bank' else 'Ev təsərrüfatı' end K_sut,m.ac_tar,m.bag_tar,m.muddet,m.faiz/100 faiz,m.D_suda,m.D_faiz,m.D_vk,m.D_vk_faiz,NULL as gir_qal,NULL as d_dovr,NULL as k_dovr,NULL as son_qal,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21115','21125','21215','21245','21255') and substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end)\r\nfrom odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno and substr(t.licsch,1,5)in ('21115','21125','21215','21245','21255') and t.date_oper=to_date('01-07-2024', 'dd/mm/yyyy')  ) gir_qal_suda,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21117','21127','21217','21247','21257') and substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end)\r\nfrom odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno and substr(t.licsch,1,5)in('21117','21127','21217','21247','21257') and t.date_oper=to_date('01-07-2024', 'dd/mm/yyyy')  ) gir_qal_faiz,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21118','21128','21218','21248','21258') and substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end)\r\nfrom odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno and substr(t.licsch,1,5)in('21118','21128','21218','21248','21258') and t.date_oper=to_date('01-07-2024', 'dd/mm/yyyy')  ) gir_qal_vk,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21119','21129','21219','21249','21259') and substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end)\r\nfrom odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno and substr(t.licsch,1,5)in('21119','21129','21219','21249','21259') and t.date_oper=to_date('01-07-2024', 'dd/mm/yyyy') ) gir_qal_vkfaiz\r\nfrom (select substr(ls.licschkre,1,5) bal, ls.subschkre sub, substr(ls.licschkre,10,6) qeydno, substr(ls.licschkre,6,2) val, ls.date_open ac_tar,\r\ncase when ls.date_prolong is null then ls.date_planclose else ls.date_prolong end bag_tar,\r\ncase when ls.date_prolong is null then MONTHS_BETWEEN(ls.date_planclose, ls.date_open) else round( MONTHS_BETWEEN(ls.date_prolong, ls.date_open),2) end muddet, ls.procstavkre faiz,\r\nSUM(CASE WHEN ar.debet = ls.licschkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) = '00'\r\nTHEN ar.summa_v_nacval/1000\r\nWHEN (ar.debet = ls.licschkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval/1000 ELSE 0 END) D_suda,\r\nSUM(CASE WHEN ar.debet = ls.licschpkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) = '00'\r\nTHEN ar.summa_v_nacval/1000\r\nWHEN (ar.debet = ls.licschpkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval/1000 ELSE 0 END) D_faiz,\r\nSUM(CASE WHEN ar.debet = ls.licsch_19 AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) = '00'\r\nTHEN ar.summa_v_nacval/1000\r\nWHEN (ar.debet = ls.licsch_19 AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval ELSE 0 END) D_vk,\r\nSUM(CASE WHEN ar.debet = ls.licschppkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) = '00'\r\nTHEN ar.summa_v_nacval/1000\r\nWHEN (ar.debet = ls.licschppkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.kredit, 1, 2) not in ('86', '66') AND SUBSTR(ar.debet, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval ELSE 0 END) D_vk_faiz,\r\nSUM(CASE WHEN ar.kredit = ls.licschkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) = '00'\r\nTHEN ar.summa_v_nacval/1000\r\nWHEN (ar.kredit = ls.licschkre AND ar.ssd = ls.subschkre AND ar.date_oper between to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy') and to_date('"+txt_dov_son.Text+ "','dd/mm/yyyy') and substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval/1000 ELSE 0 END) K_suda\r\nfrom arh_dd ar, arh_licschkre ls\r\nwhere substr(ar.debet,10,6)=substr(ls.licschpkre,10,6) \r\nand ar.ssd=ls.subschkre \r\nand ls.date_oper=odb.ish_gun_cari1 (to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy')) \r\nand (ls.date_close is null or ls.date_close>=to_date( '" + txt_dov_evvel.Text + "','dd/mm/yyyy'  ))\r\nand (substr(ls.licschkre,1,5) in ('21115','21125','21215','21245','21255') or substr(ls.licschpkre,1,5) in ('21117','21127','21217','21247','21257')\r\nor substr(ls.licsch_19,1,5) in ('21118','21128','21218','21248','21258') or substr(ls.licschppkre,1,5) in ('21119','21129','21219','21249','21259'))\r\ngroup by substr(ls.licschkre,1,5), ls.subschkre, substr(ls.licschkre,10,6), substr(ls.licschkre,6,2), ls.date_open, ls.procstavkre,ls.summa_19,ls.summa,\r\ncase when ls.date_prolong is null then ls.date_planclose else ls.date_prolong end,\r\ncase when ls.date_prolong is null then MONTHS_BETWEEN(ls.date_planclose, ls.date_open) else round(MONTHS_BETWEEN(ls.date_prolong, ls.date_open),2) end \r\norder by substr(ls.licschkre,1,5), substr(ls.licschkre,10,6), substr(ls.licschkre,6,2) asc)m)p\r\nwhere (p.gir_qal_suda+p.gir_qal_faiz+p.gir_qal_vk+p.gir_qal_vkfaiz)>0 ";
            #endregion
            #region dovriyy_artma
            string dovriyye_artma = "select t.tam_hes,t.balans,t.qeyd_no,t.val,t.G_sut,t.K_sut,NULL as ac_tar,\r\nNULL as bag_tar,NULL as muddet,NULL as faiz,NULL as K_suda,NULL as K_faiz,NULL as K_vk,NULL as K_vk_faiz,\r\nt.gir_qal,nvl( t.d_dovr,0)d_dovr,nvl( t.k_dovr,0)k_dovr,t.son_qal ,NULL as gir_qal_suda,NULL as gir_qal_faiz,\r\nNULL as gir_qal_vk, NULL as gir_qal_vkfaiz from (select qaliq.tam_hes,qaliq.balans,\r\ncase when substr(qaliq.balans,1,5) in ('10080','10020') then  'nağd xarici valyuta (aktiv üzrə)'\r\nwhen substr(qaliq.balans,1,2) in ('15','35') then  'tələbli depozit (aktiv üzrə)' else 'Ev təsərrüfatı' end G_sut,\r\ncase when substr(qaliq.balans,1,5) in ('10080','10020') then  'Mərkəzi bank'\r\nwhen substr(qaliq.balans,1,2) in ('15','35') then 'Bank' else 'Ev təsərrüfatı' end K_sut,\r\nqaliq.qeyd_no,qaliq.val,qaliq.gir_qal,\r\n(select nvl( sum( case when ar.summa_v_inval > 0 then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar where ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' )\r\nand substr(ar.debet, 1, 5) = qaliq.balans and substr(ar.debet,6,2)=qaliq.val and ar.debet=qaliq.tam_hes\r\nand substr(ar.kredit, 1, 2) not in ('86', '66') group by ar.debet) d_dovr,\r\n(select nvl( sum( case when ar.summa_v_inval > 0 then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar where ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' )\r\nand substr(ar.kredit, 1, 5) = qaliq.balans and substr(ar.kredit,6,2)=qaliq.val and ar.kredit=qaliq.tam_hes\r\nand substr(ar.debet, 1, 2) not in ('86', '66')group by ar.kredit) k_dovr,\r\nqaliq.son_qal from (select ac.name_latin,p.licsch tam_hes,substr(p.licsch,1,5) balans,substr(p.licsch,10,6) qeyd_no,\r\nsubstr(p.licsch,6,2) val, p.gir_qaliq gir_qal,k.son_qaliq son_qal from licsch lc,(select t.licsch,\r\ncase when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_ish_nacval/1000) else abs(t.saldo_vhd_inval/1000)end gir_qaliq\r\nfrom odb.arh_saldo_ls t where t.date_oper=odb.ish_gun_cari1( to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) )) p, \r\n(select t.licsch,case when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_ish_nacval/1000)\r\nelse abs(t.saldo_vhd_inval/1000) end son_qaliq from odb.arh_saldo_ls t where t.date_oper=odb.ish_gun_cari1( to_date('"+txt_dov_son.Text+ "','dd/mm/yyyy'))) k,\r\nodb.accounts ac where  p.licsch=k.licsch and p.licsch=ac.licsch  and p.licsch=lc.licsch \r\nand  (lc.date_close_licsch is null or lc.date_close_licsch > to_date('28-06-2024','dd/mm/yyyy') ) order by substr(p.licsch,1,5),\r\nsubstr(p.licsch,6,2)) qaliq,(select substr(l.licsch,1,5) hesab,substr(l.licsch,6,2) val,substr(l.licsch,10,6) qeyd_no,\r\nsum(l.oboroti_debet_nacval/1000) Dovriyye_debet_uzre,sum(l.oboroti_kredit_nacval/1000) Dovriyye_kredit_uzre\r\nfrom odb.arh_saldo_ls l where l.date_oper between to_date( '" + txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+ "','dd/mm/yyyy' )\r\ngroup by substr(l.licsch,1,5),substr(l.licsch,10,6),substr(l.licsch,6,2)) dovr where dovr.hesab=qaliq.balans\r\nand dovr.qeyd_no=qaliq.qeyd_no and dovr.val=qaliq.val(+) and qaliq.balans in (10020, 10080, 10089, 15025, 15225, 15227 )\r\nand (qaliq.gir_qal>0 or dovr.Dovriyye_debet_uzre>0 or dovr.Dovriyye_kredit_uzre>0 or qaliq.son_qal>0) order by qaliq.balans,\r\nqaliq.qeyd_no,qaliq.val asc)t UNION ALL select p.tam_hes,p.balans,p.qeyd_no,p.val,p.g_sut,p.k_sut,p.ac_tar,p.bag_tar,p.muddet,p.faiz*100,\r\np.k_suda,p.k_faiz,p.k_vk,p.k_vk_faiz, p.gir_qal,p.d_dovr,p.k_dovr,p.son_qal,p.gir_qal_suda,p.gir_qal_faiz,p.gir_qal_vk,\r\np.gir_qal_vkfaiz from (select NULL as tam_hes,m.bal balans,m.qeydno qeyd_no,m.val,\r\ncase when substr(m.bal,1,5) in ('10080','10020') then  'nağd xarici valyuta (aktiv üzrə)' when substr(m.bal,1,2) in ('15','35')\r\nthen  'tələbli depozit (aktiv üzrə)' else 'kredit (aktiv üzrə)' end G_sut,\r\ncase when substr(m.bal,1,5) in ('10080','10020') then  'Mərkəzi bank' when substr(m.bal,1,2) in ('15','35')\r\nthen 'Bank' else 'Ev təsərrüfatı' end K_sut,m.ac_tar,m.bag_tar,m.muddet,m.faiz/100 faiz,m.K_suda,m.K_faiz,m.K_vk,m.K_vk_faiz,\r\nNULL as gir_qal,NULL as d_dovr,NULL as k_dovr,NULL as son_qal,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21115','21125','21215','21245','21255') \r\nand substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end) \r\nfrom odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno \r\nand substr(t.licsch,1,5)in ('21115','21125','21215','21245','21255') and t.date_oper= to_date( '" + txt_dov_evvel.Text+"','dd/mm/yyyy'  )  ) gir_qal_suda,\r\n(select sum(case when (substr(t.licsch,1,5)in ('21117','21127','21217','21247','21257') \r\nand substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end) from odb.arh_saldo_ls t \r\nwhere substr(t.licsch,10,6)=m.qeydno and substr(t.licsch,1,5)in('21117','21127','21217','21247','21257') \r\nand t.date_oper= to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy' )) gir_qal_faiz,(select sum(case when (substr(t.licsch,1,5)in ('21118','21128','21218','21248','21258')\r\nand substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end) from odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno\r\nand substr(t.licsch,1,5)in('21118','21128','21218','21248','21258') and t.date_oper= to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  )  ) gir_qal_vk, \r\n(select sum(case when (substr(t.licsch,1,5)in ('21119','21129','21219','21249','21259') \r\nand substr(t.licsch,6,2)=m.val) then abs(t.saldo_vhd_nacval/1000) end) from odb.arh_saldo_ls t where substr(t.licsch,10,6)=m.qeydno\r\nand substr(t.licsch,1,5)in('21119','21129','21219','21249','21259') and t.date_oper= to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  )  ) gir_qal_vkfaiz \r\nfrom (select substr(ls.licschkre,1,5) bal,ls.subschkre sub,substr(ls.licschkre,10,6) qeydno,substr(ls.licschkre,6,2) val,\r\nls.date_open ac_tar,case when ls.date_prolong is null then ls.date_planclose else ls.date_prolong end bag_tar,\r\ncase when ls.date_prolong is null then MONTHS_BETWEEN(ls.date_planclose, ls.date_open) else round( MONTHS_BETWEEN(ls.date_prolong,\r\nls.date_open), 2) end muddet,ls.procstavkre faiz,SUM(CASE WHEN ar.kredit = ls.licschkre AND ar.ssk = ls.subschkre \r\nAND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) \r\nand substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) = '00' \r\nTHEN ar.summa_v_nacval/1000 WHEN (ar.kredit = ls.licschkre AND ar.ssk = ls.subschkre \r\nAND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66')\r\nAND SUBSTR(ar.kredit, 6, 2) <> '00') THEN ar.summa_v_inval/1000 ELSE 0 END) K_suda,SUM(CASE WHEN ar.kredit = ls.licschpkre \r\nAND ar.ssk = ls.subschkre AND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) \r\nand substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) = '00' THEN ar.summa_v_nacval/1000 \r\nWHEN (ar.kredit = ls.licschpkre AND ar.ssk = ls.subschkre AND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) \r\nand to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval/1000 ELSE 0 END) K_faiz, SUM(CASE WHEN ar.kredit = ls.licsch_19 AND ar.ssk = ls.subschkre \r\nAND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66')\r\nAND SUBSTR(ar.kredit, 6, 2) = '00' THEN ar.summa_v_nacval/1000 WHEN (ar.kredit = ls.licsch_19 AND ar.ssk = ls.subschkre \r\nAND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66')\r\nAND SUBSTR(ar.kredit, 6, 2) <> '00') THEN ar.summa_v_inval/1000 ELSE 0 END) K_vk, SUM(CASE WHEN ar.kredit = ls.licschppkre \r\nAND ar.ssk = ls.subschkre AND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) \r\nand substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) = '00' THEN ar.summa_v_nacval/1000 \r\nWHEN (ar.kredit = ls.licschppkre AND ar.ssk = ls.subschkre AND ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) \r\nand to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66') AND SUBSTR(ar.kredit, 6, 2) <> '00') \r\nTHEN ar.summa_v_inval/1000 ELSE 0 END) K_vk_faiz from arh_licschkre ls left join arh_dd ar on substr(ar.kredit, 10, 6) \r\n= substr(ls.licschpkre, 10, 6) and ar.ssk = ls.subschkre and ar.date_oper between to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) \r\nand to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' ) and substr(ar.debet, 1, 2) not in ('86', '66')where ls.date_oper = odb.ish_gun_cari1(to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy')) \r\nand (ls.date_close is null or ls.date_close >=  to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) ) and (substr(ls.licschkre, 1, 5) in ('21115', '21125', '21215', '21245', '21255') \r\nor substr(ls.licschpkre, 1, 5) in ('21117', '21127', '21217', '21247', '21257') or substr(ls.licsch_19, 1, 5) \r\nin ('21118', '21128', '21218', '21248', '21258') or substr(ls.licschppkre, 1, 5) in ('21119', '21129', '21219', '21249', '21259')\r\n)group by ls.procstavkre, substr(ls.licschkre, 1, 5), ls.subschkre, substr(ls.licschkre, 10, 6), substr(ls.licschkre, 6, 2),\r\nls.date_open,case when ls.date_prolong is null then ls.date_planclose else ls.date_prolong end, \r\ncase when ls.date_prolong is null then MONTHS_BETWEEN(ls.date_planclose, ls.date_open) else round(MONTHS_BETWEEN(ls.date_prolong, ls.date_open), 2)\r\nend order by substr(ls.licschkre,1,5), substr(ls.licschkre,10,6), substr(ls.licschkre,6,2) asc)m ) p\r\nwhere (p.gir_qal_suda+p.gir_qal_faiz+p.gir_qal_vk+p.gir_qal_vkfaiz)>0";
            #endregion
            #region cedvel
            string cedvel = "SELECT m.balans,m.val,m.qeydno,m.bag_tar,m.muddet,m.faiz,m.top_qal,\r\nCASE WHEN m.esas3 > m.top_qal THEN m.top_qal ELSE m.esas3 END AS esas3ay,m.faiz3,\r\nCASE WHEN m.esas6 + m.esas3 > m.top_qal THEN GREATEST(m.top_qal - m.esas3, 0) \r\nELSE m.esas6 END AS esas6ay,m.faiz6,\r\nCASE WHEN m.esas9 + m.esas6 + m.esas3 > m.top_qal THEN GREATEST(m.top_qal - m.esas6 - m.esas3, 0)\r\nELSE m.esas9 END AS esas9ay,m.faiz9,CASE \r\nWHEN m.esas12 + m.esas9 + m.esas6 + m.esas3 > m.top_qal THEN GREATEST(m.top_qal - m.esas9 - m.esas6 - m.esas3, 0)\r\nELSE m.esas12 END AS esas12ay,m.faiz12,\r\nCASE WHEN m.esas18 + m.esas12 + m.esas9 + m.esas6 + m.esas3 > m.top_qal THEN GREATEST(m.top_qal - m.esas12 - m.esas9 - m.esas6 - m.esas3, 0)\r\nELSE m.esas18 END AS esas18ay,m.faiz18,\r\nCASE WHEN m.esas24 + m.esas18 + m.esas12 + m.esas9 + m.esas6 + m.esas3 > m.top_qal THEN GREATEST(m.top_qal - m.esas18 - m.esas12 - m.esas9 - m.esas6 - m.esas3, 0)\r\nELSE m.esas24 END AS esas24ay,m.faiz24,\r\nCASE WHEN m.top_qal - (m.esas3 + m.esas6 + m.esas9 + m.esas12 + m.esas18 + m.esas24) < 0 \r\nTHEN 0 ELSE m.top_qal - (m.esas3 + m.esas6 + m.esas9 + m.esas12 + m.esas18 + m.esas24) \r\nEND AS esas25ay,m.faiz25 FROM\r\n(select substr(l.licschkre,1,5) balans,substr(l.licschkre,6,2) val,lc.registrac_nomer qeydno,\r\nl.date_open ac_tar,case when l.date_prolong is null then l.date_planclose else l.date_prolong end bag_tar,\r\ncase when l.date_prolong is null then MONTHS_BETWEEN(l.date_planclose, l.date_open) else round( MONTHS_BETWEEN(l.date_prolong, l.date_open),2) end muddet,\r\nl.procstavkre faiz,\r\nsum(case when g.date_pog between to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 3) then g.summa_pog_kre/1000 else 0 end) as esas3,\r\nsum(case when g.date_pog between to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 3) then g.summa_pog_pro/1000 else 0 end) as faiz3,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 3) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 6) then g.summa_pog_kre/1000 else 0 end) as esas6,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 3) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 6) then g.summa_pog_pro/1000 else 0 end) as faiz6,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 6) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 9) then g.summa_pog_kre/1000 else 0 end) as esas9,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 6) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 9) then g.summa_pog_pro/1000 else 0 end) as faiz9,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 9) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 12) then g.summa_pog_kre/1000 else 0 end) as esas12,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 9) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 12) then g.summa_pog_pro/1000 else 0 end) as faiz12,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 12) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 18) then g.summa_pog_kre/1000 else 0 end) as esas18,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 12) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 18) then g.summa_pog_pro/1000 else 0 end) as faiz18,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 18) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 24) then g.summa_pog_kre/1000 else 0 end) as esas24,\r\nsum(case when g.date_pog between add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 18) + 1 and add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 24) then g.summa_pog_pro/1000 else 0 end) as faiz24,\r\nsum(case when g.date_pog > add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 24)then g.summa_pog_kre/1000 else 0 end) as esas25,\r\nsum(case when g.date_pog > add_months(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'), 24)then g.summa_pog_pro/1000 else 0 end) as faiz25,l.summa_19 vk_qal,(l.summa+l.summa_19)/1000 top_qal\r\nfrom arh_licschkre l,graphpogkre g ,licsch lc\r\nwhere l.date_oper=to_date('"+txt_dov_son.Text+"','dd/mm/yyyy') and substr(l.licschkre,1,5) in (10020, 10080, 10089, 15025, 15225, 15227, 21115, 21117, 21118, 21119, 21125, 21127, 21215, 21217, 21218, 21219, 21225, 21227, 21228, 21229, 21245, 21247, 21248, 21249, 21255, 21257, 21258, 21259)\r\nand (l.date_close is null or l.date_close>=to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy')) and g.licschkre=l.licschkre and g.subschkre=l.subschkre and lc.licsch=l.licschkre\r\ngroup by l.date_open,case when l.date_prolong is null then l.date_planclose else l.date_prolong end,l.summa,l.summa_19,\r\ncase when l.date_prolong is null then MONTHS_BETWEEN(l.date_planclose, l.date_open) else round( MONTHS_BETWEEN(l.date_prolong, l.date_open),2) end,\r\nl.procstavkre,substr(l.licschkre,1,5) ,substr(l.licschkre,6,2) ,lc.registrac_nomer order by lc.registrac_nomer)m";
            #endregion
            #region cari
            string cari = "select t.tam_hes,t.balans,lc.registrac_nomer,t.val,\r\ncase when t.balans in ('39010','39020') then t.qeyd_no || t.val || substr(t.tam_hes, 9, 2)\r\nwhen substr(t.balans,0,3) in ('399','419') then t.balans || t.val else lc.registrac_nomer end C_sutun,\r\ncase when t.balans not IN ('35025', '35026', '39010', '39020', '39939', '39940', '39949', '41931', '41941', '41943')\r\nthen t.balans || substr(t.tam_hes, 16, 2) else t.balans end D_sutun,\r\ncase when lc.registrac_nomer in ('000016') then 'Birbaşa' else 'Digər' end E_sutun,'tələbli depozit (öhdəlik üzrə)' G_sutun,\r\ncase when lc.registrac_nomer in ('000016') then 'Birbaşa investor' else 'Heç bir' end E_sutun,\r\ncase when t.balans in ('35025','35026') then 'Bank'\r\nwhen substr(t.balans,0,2) in ('39','40') then 'Qeyri-maliyyə müəssisəsi'\r\nwhen substr(t.balans,0,2) in ('41') then 'Ev təsərrüfatı' end K_sutun,\r\ncase when lc.countrycode='GEO' then 'Gürcüstan' else c.name end M_sutun,\r\ncase when lc.countrycode='AZE' then 'AZN' else lc.countrycode end O_sutun,lc.date_open_licsch ac_tar,\r\ncase when lc.registrac_nomer in ('000016') then '100%' else '0%' end S_sutun,\r\ncase when lc.registrac_nomer in ('000016') then '100%' else '0%' end T_sutun,\r\ncase when lc.date_open_licsch between to_date(to_date('" + txt_dov_evvel.Text+"','dd/mm/yyyy')) and to_date(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy')) then lc.date_open_licsch  end U_sutun,\r\n'faizsiz' Y_sutun,t.gir_qal AD_sutun,\r\nnvl( t.k_dovr,0) AI_sutun,              \r\nt.gir_qal,nvl( t.d_dovr,0)d_dovr,nvl( t.k_dovr,0)k_dovr,t.son_qal \r\nfrom (select qaliq.tam_hes,qaliq.balans,qaliq.qeyd_no,qaliq.val,qaliq.gir_qal,\r\n(select \r\nnvl( sum( case when substr(ar.debet,6,2) not in ('00') then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar\r\nwhere ar.date_oper between to_date(to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy')) and to_date(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'))\r\nand substr(ar.debet, 1, 5) = qaliq.balans and substr(ar.debet,6,2)=qaliq.val and ar.debet=qaliq.tam_hes\r\nand substr(ar.kredit, 1, 2) not in ('86', '66') group by ar.debet) d_dovr,\r\n(select nvl( sum( case when substr(ar.kredit,6,2) not in ('00') then ar.summa_v_inval/1000 else ar.summa_v_nacval/1000 end),0) d_dovr\r\nfrom arh_dd ar\r\nwhere ar.date_oper between to_date(to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy')) and to_date(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'))\r\nand substr(ar.kredit, 1, 5) = qaliq.balans and substr(ar.kredit,6,2)=qaliq.val and ar.kredit=qaliq.tam_hes\r\nand substr(ar.debet, 1, 2) not in ('86', '66')group by ar.kredit) k_dovr,\r\nqaliq.son_qal from \r\n(select ac.name_latin,p.licsch tam_hes,substr(p.licsch,1,5) balans,substr(p.licsch,10,6) qeyd_no,substr(p.licsch,6,2) val, p.gir_qaliq gir_qal,k.son_qaliq son_qal from licsch lc,\r\n(select t.licsch, case when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_vhd_nacval/1000)\r\nelse abs(t.saldo_vhd_inval/1000)end gir_qaliq\r\nfrom odb.arh_saldo_ls t where t.date_oper=odb.ish_gun_cari1(to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy'))) p, \r\n(select t.licsch,\r\ncase when substr(t.licsch, 6, 2) = '00' then abs(t.saldo_vhd_nacval/1000)\r\nelse abs(t.saldo_vhd_inval/1000) end son_qaliq\r\nfrom odb.arh_saldo_ls t where t.date_oper=odb.ish_gun_cari1(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy'))) k,odb.accounts ac\r\nwhere  p.licsch=k.licsch and p.licsch=ac.licsch  and p.licsch=lc.licsch and  (lc.date_close_licsch is null or lc.date_close_licsch >= to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy'))\r\norder by substr(p.licsch,1,5),substr(p.licsch,6,2)) qaliq,\r\n(select substr(l.licsch,1,5) hesab,substr(l.licsch,6,2) val,substr(l.licsch,10,6) qeyd_no,\r\nsum(l.oboroti_debet_nacval) Dovriyye_debet_uzre,sum(l.oboroti_kredit_nacval) Dovriyye_kredit_uzre\r\nfrom odb.arh_saldo_ls l where l.date_oper between to_date(to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy')) and to_date(to_date('"+txt_dov_son.Text+"','dd/mm/yyyy')) \r\ngroup by substr(l.licsch,1,5),substr(l.licsch,10,6),substr(l.licsch,6,2)) dovr\r\nwhere dovr.hesab=qaliq.balans and dovr.qeyd_no=qaliq.qeyd_no and dovr.val=qaliq.val\r\nand (qaliq.balans in (35025, 35026, 40045, 40055, 40065, 40075, 40145, 40155, 40165, 40175, 41015, 41025, 41026, 41045, 41055, 41115, 41117, 35015, 35941, 35943,\r\n35949, 39930, 39931, 39939, 39940, 39941, 39949, 39950, 39951, 39959, 41931, 41933, 41941, 41943 ) or substr(qaliq.balans,1,4) in (3901, 3902))\r\nand (qaliq.gir_qal>0 or dovr.Dovriyye_debet_uzre>0 or dovr.Dovriyye_kredit_uzre>0 or qaliq.son_qal>0) order by qaliq.balans)t,licsch lc,countrycode c where lc.licsch=t.tam_hes and lc.countrycode=c.code";
            #endregion
            #region olkeler
            string olkeler = "select distinct l.registrac_nomer,\r\ncase when l.countrycode='GEO' then 'Gürcüstan' else c.name end\r\nfrom licsch l,countrycode c\r\nwhere (l.date_close_licsch is null or l.date_close_licsch>to_date('"+txt_dov_evvel.Text+"','dd/mm/yyyy'))\r\nand l.countrycode=c.code order by l.registrac_nomer asc";
            #endregion
            //to_date( '"+txt_dov_evvel.Text+"','dd/mm/yyyy'  ) and to_date( '"+txt_dov_son.Text+"','dd/mm/yyyy' )
            System.Data.DataTable _dt_dovr_artma = new System.Data.DataTable();
            System.Data.DataTable _dt_dovr2 = new System.Data.DataTable();
            System.Data.DataTable _dt_cedvel = new System.Data.DataTable();
            System.Data.DataTable _dt_cari = new System.Data.DataTable();
            System.Data.DataTable _dt_olkeler = new System.Data.DataTable();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(dovriyye_artma, connection))
                {
                    
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_dovr_artma);
                }
                using (OracleCommand command = new OracleCommand(dovriyye_azalma, connection))
                {

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_dovr2);
                }

                using (OracleCommand command = new OracleCommand(cedvel, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_cedvel);
                }
                using (OracleCommand command = new OracleCommand(cari, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_cari);
                }
                using (OracleCommand command = new OracleCommand(olkeler, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_olkeler);
                }
                connection.Close();
            }
            dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            textBoxText = txt_dov_evvel.Text; // TextBox'tan alınan metni sakla
            yeniMetin = textBoxText.Replace("-", "");
            baseFileName = "XBBIS_" + yeniMetin;//.Substring(2, 6); // Temel dosya adı
            fileName = baseFileName + ".xlsm";
            templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "XBBIS_.xlsm");
            filePath = System.IO.Path.Combine(dosyayolu, fileName);

            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }
            //"15020",
            //FileInfo templateFile = new FileInfo(templateFilePath);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var templateFile = new FileInfo(System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "XBBIS_.xlsm")); // Faylın yolunu göstərin
            //@"C:\BMI_\XBBIS_.xlsm
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                #endregion
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["A1 (Aktivlər)"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["A2 (Öhdəliklər)"];

                int startRow = 15; // Başlanğıc satır
                int startColumn = 3; // "C" sütunu Excel-də 3-cü sütundur
                int startColumncedvel = 1; // A sütunu üçün
                int startRowcedvel = 1;    // Birinci sətirdən başlayaraq yazmaq üçün
                int yoxla = 0;
                int carisetirsay= _dt_cari.Rows.Count;
                #region A1
                for (int i = 0; i < _dt_dovr_artma.Rows.Count; i++)
                {
                    // Excel hüceyrəsinə məlumatın yazılması

                    string balans = _dt_dovr_artma.Rows[i][1].ToString();
                    string val = _dt_dovr_artma.Rows[i][3].ToString();
                    string qeydno = _dt_dovr_artma.Rows[i][2].ToString();
                    int sirasay = 0;
                    


                    //C
                    if ((balans.Substring(0, 5) == "10020" || balans.Substring(0, 5) == "10080") && val == "01")
                    {
                        worksheet1.Cells[startRow + i, startColumn].Value = _dt_dovr_artma.Rows[i][1].ToString() + "840";
                    }
                    if ((balans.Substring(0, 5) == "10020" || balans.Substring(0, 5) == "10080") && val == "02")
                    {
                        worksheet1.Cells[startRow + i, startColumn].Value = _dt_dovr_artma.Rows[i][1].ToString() + "978";
                    }
                    if (balans.Substring(0, 5) != "10020" && balans.Substring(0, 5) != "10080")
                    {
                        worksheet1.Cells[startRow + i, startColumn].Value = _dt_dovr_artma.Rows[i][2].ToString();
                    }
                    //D
                    if ((qeydno == "15225" || qeydno == "16160"))
                    {
                        worksheet1.Cells[startRow + i, startColumn + 1].Value = _dt_dovr_artma.Rows[i][2].ToString() + "00";
                    }
                    if ((qeydno != "15225" && qeydno != "16160"))
                    {
                        worksheet1.Cells[startRow + i, startColumn + 1].Value = _dt_dovr_artma.Rows[i][2].ToString();
                    }
                    //E
                    worksheet1.Cells[startRow + i, startColumn + 2].Value = "Digər";
                    //G
                    worksheet1.Cells[startRow + i, startColumn + 4].Value = _dt_dovr_artma.Rows[i][4];
                    //I
                    worksheet1.Cells[startRow + i, startColumn + 6].Value = "Heç bir";
                    //K
                    worksheet1.Cells[startRow + i, startColumn + 8].Value = _dt_dovr_artma.Rows[i][5];
                    //M
                    var dovrArtmaValue = _dt_dovr_artma.Rows[i][2].ToString();

                    // _dt_cari tablosunda bu değeri arıyoruz
                    var matchingRow = _dt_olkeler.AsEnumerable().FirstOrDefault(row => row[0].ToString() == dovrArtmaValue);

                    // beraberlik varsa
                    if (matchingRow != null)
                    {
                        worksheet1.Cells[startRow + i, startColumn + 10].Value = matchingRow[1].ToString();
                    }
                    //worksheet1.Cells[startRow + i, startColumn + 6].Value = "Heç bir"; OLKELER
                    //O
                    if (val == "00")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "AZN";
                    }
                    if (val == "01")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "USD";
                    }
                    if (val == "02")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "EUR";
                    }
                    if (val == "03")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "EUR";
                    }
                    if (val == "04")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "IRR";
                    }
                    if (val == "05")
                    {
                        worksheet1.Cells[startRow + i, startColumn + 12].Value = "AED";
                    }
                    //S
                    worksheet1.Cells[startRow + i, startColumn + 16].Value = "0%";
                    //T
                    worksheet1.Cells[startRow + i, startColumn + 17].Value = "0%";
                    //U
                    worksheet1.Cells[startRow + i, startColumn + 18].Value = _dt_dovr_artma.Rows[i][6];
                    //V
                    worksheet1.Cells[startRow + i, startColumn + 19].Value = _dt_dovr_artma.Rows[i][7];
                    //W
                    worksheet1.Cells[startRow + i, startColumn + 20].Value = _dt_dovr_artma.Rows[i][8];
                    //X
                    if ((balans == "10020" || balans == "10080"))

                    {
                        worksheet1.Cells[startRow + i, startColumn + 21].Value = "";
                        //AD
                        worksheet1.Cells[startRow + i, startColumn + 27].Value = _dt_dovr2.Rows[i][14];
                        //AE
                        worksheet1.Cells[startRow + i, startColumn + 28].Value = 0.0;
                        //AF
                        worksheet1.Cells[startRow + i, startColumn + 29].Value = 0.0;
                        //AG
                        worksheet1.Cells[startRow + i, startColumn + 30].Value = 0.0;
                        //AH
                        worksheet1.Cells[startRow + i, startColumn + 31].Value = 0.0;
                        //AI
                        worksheet1.Cells[startRow + i, startColumn + 32].Value = _dt_dovr2.Rows[i][15];
                        //AJ
                        worksheet1.Cells[startRow + i, startColumn + 33].Value = 0.0;
                        //AK
                        worksheet1.Cells[startRow + i, startColumn + 34].Value = 0.0;
                        //AL
                        worksheet1.Cells[startRow + i, startColumn + 35].Value = 0.0;
                        //AM
                        worksheet1.Cells[startRow + i, startColumn + 36].Value = 0.0;
                        //AN
                        worksheet1.Cells[startRow + i, startColumn + 37].Value = _dt_dovr_artma.Rows[i][16];
                        //AO
                        worksheet1.Cells[startRow + i, startColumn + 38].Value = 0.0;
                        //AP
                        worksheet1.Cells[startRow + i, startColumn + 39].Value = 0.0;
                        //AQ
                        worksheet1.Cells[startRow + i, startColumn + 40].Value = 0.0;
                        //AR
                        worksheet1.Cells[startRow + i, startColumn + 41].Value = 0.0;
                        //AS
                        worksheet1.Cells[startRow + i, startColumn + 42].Value = 0.0;
                        //AT
                        worksheet1.Cells[startRow + i, startColumn + 43].Value = 0.0;
                        //AU
                        worksheet1.Cells[startRow + i, startColumn + 44].Value = 0.0;
                        //AV
                        worksheet1.Cells[startRow + i, startColumn + 45].Value = 0.0;
                        //AW
                        worksheet1.Cells[startRow + i, startColumn + 46].Value = 0.0;
                        //AX
                        worksheet1.Cells[startRow + i, startColumn + 47].Value = 0.0;
                        //AY
                        worksheet1.Cells[startRow + i, startColumn + 48].Value = "";
                        //BG
                        worksheet1.Cells[startRow + i, startColumn + 56].Value = 0.0;
                        //BH
                        worksheet1.Cells[startRow + i, startColumn + 57].Value = 0.0;
                        //BI
                        worksheet1.Cells[startRow + i, startColumn + 58].Value = 0.0;
                        //BJ
                        worksheet1.Cells[startRow + i, 55].Calculate();
                        worksheet1.Cells[startRow + i, startColumn + 59].Value = worksheet1.Cells[startRow + i, startColumn + 52].Value;
                        //BK
                        worksheet1.Cells[startRow + i, startColumn + 60].Value = 0.0;
                        //BL
                        worksheet1.Cells[startRow + i, startColumn + 61].Value = 0.0;
                        //BM
                        worksheet1.Cells[startRow + i, startColumn + 62].Value = 0.0;
                        //BN
                        worksheet1.Cells[startRow + i, startColumn + 63].Value = 0.0;
                        //BO
                        worksheet1.Cells[startRow + i, startColumn + 64].Value = 0.0;
                        //BP
                        worksheet1.Cells[startRow + i, startColumn + 65].Value = 0.0;
                        //BQ
                        worksheet1.Cells[startRow + i, startColumn + 66].Value = 0.0;
                        //BR
                        worksheet1.Cells[startRow + i, startColumn + 67].Value = 0.0;
                        //BS
                        worksheet1.Cells[startRow + i, startColumn + 68].Value = 0.0;
                        //BT
                        worksheet1.Cells[startRow + i, startColumn + 69].Value = 0.0;
                        //BU
                        worksheet1.Cells[startRow + i, startColumn + 70].Value = 0.0;
                        //BV
                        worksheet1.Cells[startRow + i, startColumn + 71].Value = 0.0;
                        //BW
                        worksheet1.Cells[startRow + i, startColumn + 72].Value = 0.0;
                        //BX
                        worksheet1.Cells[startRow + i, startColumn + 73].Value = 0.0;
                        //BY
                        worksheet1.Cells[startRow + i, startColumn + 74].Value = 0.0;

                    }


                    else if (balans == "15025")

                    {
                        worksheet1.Cells[startRow + i, startColumn + 21].Value = "0%";
                        //AA
                        worksheet1.Cells[startRow + i, startColumn + 24].Value = 0.0;
                        //AB
                        worksheet1.Cells[startRow + i, startColumn + 25].Value = 0.0;
                        //AD
                        worksheet1.Cells[startRow + i, startColumn + 27].Value = _dt_dovr2.Rows[i][14];
                        //AE
                        worksheet1.Cells[startRow + i, startColumn + 28].Value = 0.0;
                        //AF
                        worksheet1.Cells[startRow + i, startColumn + 29].Value = 0.0;
                        //AG
                        worksheet1.Cells[startRow + i, startColumn + 30].Value = 0.0;
                        //AH
                        worksheet1.Cells[startRow + i, startColumn + 31].Value = 0.0;
                        //AI
                        worksheet1.Cells[startRow + i, startColumn + 32].Value = _dt_dovr2.Rows[i][15];
                        //AJ
                        worksheet1.Cells[startRow + i, startColumn + 33].Value = 0.0;
                        //AK
                        worksheet1.Cells[startRow + i, startColumn + 34].Value = 0.0;
                        //AL
                        worksheet1.Cells[startRow + i, startColumn + 35].Value = 0.0;
                        //AM
                        worksheet1.Cells[startRow + i, startColumn + 36].Value = 0.0;
                        //AN
                        worksheet1.Cells[startRow + i, startColumn + 37].Value = _dt_dovr_artma.Rows[i][16];
                        //AO
                        worksheet1.Cells[startRow + i, startColumn + 38].Value = 0.0;
                        //AP
                        worksheet1.Cells[startRow + i, startColumn + 39].Value = 0.0;
                        //AQ
                        worksheet1.Cells[startRow + i, startColumn + 40].Value = 0.0;
                        //AR
                        worksheet1.Cells[startRow + i, startColumn + 41].Value = 0.0;
                        //AS
                        worksheet1.Cells[startRow + i, startColumn + 42].Value = 0.0;
                        //AT
                        worksheet1.Cells[startRow + i, startColumn + 43].Value = 0.0;
                        //AU
                        worksheet1.Cells[startRow + i, startColumn + 44].Value = 0.0;
                        //AV
                        worksheet1.Cells[startRow + i, startColumn + 45].Value = 0.0;
                        //AW
                        worksheet1.Cells[startRow + i, startColumn + 46].Value = 0.0;
                        //AX
                        worksheet1.Cells[startRow + i, startColumn + 47].Value = 0.0;
                        //AY
                        worksheet1.Cells[startRow + i, startColumn + 48].Value = "";
                        //BG
                        worksheet1.Cells[startRow + i, startColumn + 56].Value = 0.0;
                        //BH
                        worksheet1.Cells[startRow + i, startColumn + 57].Value = 0.0;
                        //BI
                        worksheet1.Cells[startRow + i, startColumn + 58].Value = 0.0;
                        //BJ
                        worksheet1.Cells[startRow + i, 55].Calculate();
                        worksheet1.Cells[startRow + i, startColumn + 59].Value = worksheet1.Cells[startRow + i, startColumn + 52].Value;
                        //BK
                        worksheet1.Cells[startRow + i, startColumn + 60].Value = 0.0;
                        //BL
                        worksheet1.Cells[startRow + i, startColumn + 61].Value = 0.0;
                        //BM
                        worksheet1.Cells[startRow + i, startColumn + 62].Value = 0.0;
                        //BN
                        worksheet1.Cells[startRow + i, startColumn + 63].Value = 0.0;
                        //BO
                        worksheet1.Cells[startRow + i, startColumn + 64].Value = 0.0;
                        //BP
                        worksheet1.Cells[startRow + i, startColumn + 65].Value = 0.0;
                        //BQ
                        worksheet1.Cells[startRow + i, startColumn + 66].Value = 0.0;
                        //BR
                        worksheet1.Cells[startRow + i, startColumn + 67].Value = 0.0;
                        //BS
                        worksheet1.Cells[startRow + i, startColumn + 68].Value = 0.0;
                        //BT
                        worksheet1.Cells[startRow + i, startColumn + 69].Value = 0.0;
                        //BU
                        worksheet1.Cells[startRow + i, startColumn + 70].Value = 0.0;
                        //BV
                        worksheet1.Cells[startRow + i, startColumn + 71].Value = 0.0;
                        //BW
                        worksheet1.Cells[startRow + i, startColumn + 72].Value = 0.0;
                        //BX
                        worksheet1.Cells[startRow + i, startColumn + 73].Value = 0.0;
                        //BY
                        worksheet1.Cells[startRow + i, startColumn + 74].Value = 0.0;
                    }
                    else
                    {
                        worksheet1.Cells[startRow + i, startColumn + 21].Value = _dt_dovr_artma.Rows[i][9];
                        //Z
                        worksheet1.Cells[startRow + i, startColumn + 22].Value = "dəyişkən";
                        //AA
                        worksheet1.Cells[startRow + i, startColumn + 24].Value = 0.0;
                        //AB
                        worksheet1.Cells[startRow + i, startColumn + 25].Value = 0.0;
                        //AD
                        decimal value18 = _dt_dovr2.Rows[i][18] != DBNull.Value ? Convert.ToDecimal(_dt_dovr2.Rows[i][18]) : 0;
                        decimal value20 = _dt_dovr2.Rows[i][20] != DBNull.Value ? Convert.ToDecimal(_dt_dovr2.Rows[i][20]) : 0;

                        worksheet1.Cells[startRow + i, startColumn + 27].Value = value18 + value20;
                        //AE
                        worksheet1.Cells[startRow + i, startColumn + 28].Value = _dt_dovr2.Rows[i][20];
                        //AF
                        worksheet1.Cells[startRow + i, startColumn + 29].Value =
                        (Convert.ToDecimal(_dt_dovr2.Rows[i][19] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][19])
                        + Convert.ToDecimal(_dt_dovr2.Rows[i][21] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][21]));

                        //AG
                        worksheet1.Cells[startRow + i, startColumn + 30].Value = _dt_dovr2.Rows[i][21];
                        //AH
                        worksheet1.Cells[startRow + i, startColumn + 31].Value = 0.0;
                        //AI
                        worksheet1.Cells[startRow + i, startColumn + 32].Value =
                        (Convert.ToDecimal(_dt_dovr2.Rows[i][10] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][10])
                        + Convert.ToDecimal(_dt_dovr2.Rows[i][12] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][12]));
                        //AJ
                        worksheet1.Cells[startRow + i, startColumn + 33].Value = _dt_dovr2.Rows[i][12];
                        //AK
                        worksheet1.Cells[startRow + i, startColumn + 34].Value =
                        (Convert.ToDecimal(_dt_dovr2.Rows[i][11] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][11])
                        + Convert.ToDecimal(_dt_dovr2.Rows[i][13] == DBNull.Value ? 0 : _dt_dovr2.Rows[i][13]));
                        //AL
                        worksheet1.Cells[startRow + i, startColumn + 35].Value = _dt_dovr2.Rows[i][13];
                        //AM
                        worksheet1.Cells[startRow + i, startColumn + 36].Value = 0.0;
                        //AN
                        worksheet1.Cells[startRow + i, startColumn + 37].Value =
                        (Convert.ToDecimal(_dt_dovr_artma.Rows[i][10] == DBNull.Value ? 0 : _dt_dovr_artma.Rows[i][10])
                        + Convert.ToDecimal(_dt_dovr_artma.Rows[i][12] == DBNull.Value ? 0 : _dt_dovr_artma.Rows[i][12]));
                        //AO
                        worksheet1.Cells[startRow + i, startColumn + 38].Value = _dt_dovr_artma.Rows[i][12];
                        //AP
                        worksheet1.Cells[startRow + i, startColumn + 39].Value =
                        (Convert.ToDecimal(_dt_dovr_artma.Rows[i][11] == DBNull.Value ? 0 : _dt_dovr_artma.Rows[i][11])
                        + Convert.ToDecimal(_dt_dovr_artma.Rows[i][13] == DBNull.Value ? 0 : _dt_dovr_artma.Rows[i][13]));
                        //AQ
                        worksheet1.Cells[startRow + i, startColumn + 40].Value = _dt_dovr_artma.Rows[i][13];
                        //AR
                        worksheet1.Cells[startRow + i, startColumn + 41].Value = 0.0;
                        //AS
                        worksheet1.Cells[startRow + i, startColumn + 42].Value = 0.0;
                        //AT
                        worksheet1.Cells[startRow + i, startColumn + 43].Value = 0.0;
                        //AU
                        worksheet1.Cells[startRow + i, startColumn + 44].Value = 0.0;
                        //AV
                        worksheet1.Cells[startRow + i, startColumn + 45].Value = 0.0;
                        //AW
                        worksheet1.Cells[startRow + i, startColumn + 46].Value = 0.0;
                        //AX
                        worksheet1.Cells[startRow + i, startColumn + 47].Value = 0.0;
                        //AY
                        worksheet1.Cells[startRow + i, startColumn + 48].Value = "";
                        //BG
                        worksheet1.Cells[startRow + i, startColumn + 56].Value = 0.0;
                        //BH
                        worksheet1.Cells[startRow + i, startColumn + 57].Value = 0.0;
                        //BI
                        worksheet1.Cells[startRow + i, startColumn + 58].Value = 0.0;
                        //BJ
                        worksheet1.Cells[startRow + i, 56].Calculate();
                        worksheet1.Cells[startRow + i, startColumn + 59].Value = worksheet1.Cells[startRow + i, startColumn + 53].Value;
                        //BK
                        worksheet1.Cells[startRow + i, startColumn + 60].Value = worksheet1.Cells[startRow + i, startColumn + 55].Value;
                        //BZ
                        worksheet1.Cells[startRow + i, startColumn + 69].Value = "";
                        //CA
                        worksheet1.Cells[startRow + i, startColumn + 70].Value = "";
                        //CB
                        worksheet1.Cells[startRow + i, startColumn + 71].Value = "";
                        //CD
                        worksheet1.Cells[startRow + i, startColumn + 73].Value = "";
                        // Qeyd no üzrə dəyərləri yoxlayaraq uyğun sütunların dəyərlərini toplamaq
                        decimal toplam = 0;
                        decimal toplamfaiz = 0;
                        int currentColumn = 64;
                        int artir = 0;

                        foreach (DataRow row in _dt_cedvel.Rows)
                        {
                            if (row[2].ToString() == qeydno) // 2-ci sütunun dəyərini müqayisə edir
                            {
                                // "for" dövrü şərti kimi icra olunur
                                for (int col = 7; col <= 19; col += 2) // 7-dən 19-a qədər 2 artırırıq
                                {
                                    if (decimal.TryParse(row[col].ToString(), out decimal value)) // Uyğun sütunun dəyərini yoxlayır
                                    {
                                        toplam = value;
                                        worksheet1.Cells[startRow + i, currentColumn + sirasay].Value = toplam;
                                        if (col + 1 < _dt_cedvel.Columns.Count) // Sərhəd yoxlaması
                                        {
                                            if (decimal.TryParse(row[col + 1].ToString(), out decimal faizValue))
                                            {
                                                toplamfaiz = faizValue; // `col + 1` sütunun dəyəri
                                                worksheet1.Cells[startRow + i, currentColumn + sirasay + 1].Value = toplamfaiz;
                                            }
                                        }
                                        sirasay += 2; // Sütunu artırırıq
                                        artir += 1;
                                    }
                                }

                            }
                            else
                            {
                                for (int m = 0; m < 14; m++)
                                {
                                    worksheet1.Cells[startRow + i, 64 + m].Value = 1986;
                                }
                            }
                        }

                        // Dövrün sonunda sirasay-ı sıfırlayırıq
                        sirasay = 0;

                        sirasay = 0; // Növbəti dövr üçün sıfırlayırıq

                    }

                   


                }
                #endregion

                #region A2
                for (int i = 0; i < _dt_cari.Rows.Count; i++)
                {

                    string tamhes_caridovr = _dt_cari.Rows[i][0].ToString();
                    string qeydno_caridovr = _dt_cari.Rows[i][2].ToString();
                    string balans_caridovr = _dt_cari.Rows[i][1].ToString();
                    string val_caridovr = _dt_cari.Rows[i][1].ToString();
                    int qeyd_int_carihes = Convert.ToInt16(qeydno_caridovr);
                    //C
                    if (qeyd_int_carihes == 4801)
                    {
                        worksheet2.Cells[startRow + i, startColumn].Value = qeyd_int_carihes + tamhes_caridovr.Substring(5, 2) + tamhes_caridovr.Substring(8, 2);
                    }
                    else if (balans_caridovr.Substring(0, 3) == "399" || balans_caridovr.Substring(0, 3) == "419")
                    {
                        worksheet2.Cells[startRow + i, startColumn].Value = balans_caridovr + tamhes_caridovr.Substring(6, 2);
                    }
                    else
                    {
                        worksheet2.Cells[startRow + i, startColumn].Value = qeyd_int_carihes;
                    }
                    //D
                    if (balans_caridovr.Substring(0, 3) == "400" || balans_caridovr.Substring(0, 3) == "410")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 1].Value = balans_caridovr + tamhes_caridovr.Substring(16, 2);
                    }
                    else
                    {
                        worksheet2.Cells[startRow + i, startColumn + 1].Value = balans_caridovr;
                    }
                    //E,----- i,s,t
                    if (qeyd_int_carihes == 16)
                    {
                        worksheet2.Cells[startRow + i, startColumn + 2].Value = "Birbaşa";
                        worksheet2.Cells[startRow + i, startColumn + 6].Value = "Birbaşa investor";
                        worksheet2.Cells[startRow + i, startColumn + 16].Value = "100.00%";
                        worksheet2.Cells[startRow + i, startColumn + 17].Value = "100.00%";
                    }
                    else
                    {
                        worksheet2.Cells[startRow + i, startColumn + 2].Value = "Digər";
                        worksheet2.Cells[startRow + i, startColumn + 6].Value = "Heç bir";
                        worksheet2.Cells[startRow + i, startColumn + 16].Value = "0.00%";
                        worksheet2.Cells[startRow + i, startColumn + 17].Value = "0.00%";
                    }
                    //G
                    worksheet2.Cells[startRow + i, startColumn + 4].Value = "tələbli depozit (öhdəlik üzrə)";
                    //I e-de qeyd olunub
                    //K
                    if (balans_caridovr == "35025" || balans_caridovr == "35026")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 8].Value = "Bank";
                    }
                    else if (balans_caridovr.Substring(0, 2) == "39" || balans_caridovr.Substring(0, 2) == "40")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 8].Value = "Qeyri-maliyyə müəssisəsi";
                    }
                    else
                    {
                        worksheet2.Cells[startRow + i, startColumn + 8].Value = "Ev təsərrüfatı";
                    }
                    //M
                    if (balans_caridovr.Substring(0, 3) == "399" || balans_caridovr.Substring(0, 3) == "419")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 10].Value = "İran İslam Respublikası";
                    }
                    else
                    {
                        worksheet2.Cells[startRow + i, startColumn + 10].Value = _dt_cari.Rows[i][10].ToString();
                    }
                    //}

                    //O
                    if (val_caridovr == "00")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "AZN";
                    }
                    if (val_caridovr == "01")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "USD";
                    }
                    if (val_caridovr == "02")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "EUR";
                    }
                    if (val_caridovr == "03")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "RUB";
                    }
                    if (val_caridovr == "04")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "IRR";
                    }
                    if (val_caridovr == "05")
                    {
                        worksheet2.Cells[startRow + i, startColumn + 12].Value = "AED";
                    }
                    //Q
                    //bos qalir +14

                    //S e-de qeyd olununb +16

                    //T e-de qeyd olununb +17

                    //U
                    worksheet2.Cells[startRow + i, startColumn + 18].Value = _dt_cari.Rows[i][12].ToString().Substring(0, 10);
                    //V
                    //bos qalir +19

                    //W
                    //bos qalir +20

                    //X
                    worksheet2.Cells[startRow + i, startColumn + 21].Value = "0%";
                    //Y
                    worksheet2.Cells[startRow + i, startColumn + 22].Value = "faizsiz";
                    //AA
                    worksheet2.Cells[startRow + i, startColumn + 24].Value = 0.0;
                    //AB
                    worksheet2.Cells[startRow + i, startColumn + 25].Value = 0.0;
                    //AD
                    worksheet2.Cells[startRow + i, startColumn + 27].Value = _dt_cari.Rows[i][17].ToString();
                    //AE
                    worksheet2.Cells[startRow + i, startColumn + 28].Value = 0.0;
                    //AF
                    worksheet2.Cells[startRow + i, startColumn + 29].Value = 0.0;
                    //AG
                    worksheet2.Cells[startRow + i, startColumn + 30].Value = 0.0;
                    //AH
                    worksheet2.Cells[startRow + i, startColumn + 31].Value = 0.0;
                    //AI
                    worksheet2.Cells[startRow + i, startColumn + 32].Value = _dt_cari.Rows[i][18].ToString();
                    //AJ
                    worksheet2.Cells[startRow + i, startColumn + 33].Value = 0.0;
                    //AK
                    worksheet2.Cells[startRow + i, startColumn + 34].Value = 0.0;
                    //AL
                    worksheet2.Cells[startRow + i, startColumn + 35].Value = 0.0;
                    //AN
                    worksheet2.Cells[startRow + i, startColumn + 37].Value = _dt_cari.Rows[i][20].ToString();
                    //AO
                    worksheet2.Cells[startRow + i, startColumn + 38].Value = 0.0;
                    //AP
                    worksheet2.Cells[startRow + i, startColumn + 39].Value = 0.0;
                    //AQ
                    worksheet2.Cells[startRow + i, startColumn + 40].Value = 0.0;
                    //AR
                    worksheet2.Cells[startRow + i, startColumn + 41].Value = 0.0;
                    //AS
                    worksheet2.Cells[startRow + i, startColumn + 42].Value = 0.0;
                    //AT
                    worksheet2.Cells[startRow + i, startColumn + 43].Value = 0.0;
                    //AU
                    worksheet2.Cells[startRow + i, startColumn + 44].Value = 0.0;
                    //AV
                    worksheet2.Cells[startRow + i, startColumn + 45].Value = 0.0;
                    //AW
                    worksheet2.Cells[startRow + i, startColumn + 46].Value = 0.0;
                    //AX
                    worksheet2.Cells[startRow + i, startColumn + 47].Value = 0.0;
                    //AY
                    worksheet2.Cells[startRow + i, startColumn + 48].Value = "";
                    //BG
                    worksheet2.Cells[startRow + i, startColumn + 56].Value = 0;
                    //BH
                    worksheet2.Cells[startRow + i, startColumn + 57].Value = 0;
                    //BI
                    worksheet2.Cells[startRow + i, startColumn + 58].Value = 0;
                    //BJ
                    worksheet2.Cells[startRow + i, 55].Calculate();
                    worksheet2.Cells[startRow + i, startColumn + 59].Value = worksheet2.Cells[startRow + i, startColumn + 52].Value;
                    var t1 = worksheet2.Cells[startRow + i, 55];
                    //BK
                    worksheet2.Cells[startRow + i, startColumn + 60].Value = 0;
                    //BL
                    worksheet2.Cells[startRow + i, startColumn + 61].Value = 0;
                    //BM
                    worksheet2.Cells[startRow + i, startColumn + 62].Value = 0;
                    //BN
                    worksheet2.Cells[startRow + i, startColumn + 63].Value = 0;
                    //BO
                    worksheet2.Cells[startRow + i, startColumn + 64].Value = 0;
                    //BP
                    worksheet2.Cells[startRow + i, startColumn + 65].Value = 0;
                    //BQ
                    worksheet2.Cells[startRow + i, startColumn + 66].Value = 0;
                    //BR
                    worksheet2.Cells[startRow + i, startColumn + 67].Value = 0;
                    //BS
                    worksheet2.Cells[startRow + i, startColumn + 68].Value = 0;
                    //BT
                    worksheet2.Cells[startRow + i, startColumn + 69].Value = 0;
                    //BU
                    worksheet2.Cells[startRow + i, startColumn + 70].Value = 0;
                    //BV
                    worksheet2.Cells[startRow + i, startColumn + 71].Value = 0;
                    //BW
                    worksheet2.Cells[startRow + i, startColumn + 72].Value = 0;
                    //BX
                    worksheet2.Cells[startRow + i, startColumn + 73].Value = 0;
                    //BY
                    worksheet2.Cells[startRow + i, startColumn + 74].Value = 0;
                    //BZ
                    worksheet2.Cells[startRow + i, startColumn + 74].Value = 0;
                    //CA
                    worksheet2.Cells[startRow + i, startColumn + 75].Value = 0;
                    //CB
                    worksheet2.Cells[startRow + i, startColumn + 76].Value = 0;
                    //CD
                    worksheet2.Cells[startRow + i, startColumn + 78].Value = 0;

                 
                }
#endregion



                filePath = System.IO.Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private void txt_dov_evvel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_dov_son.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_dov_evvel_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_dov_evvel.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_dov_evvel.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txt_dov_son_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                btn_sorgu.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_dov_son_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_dov_son.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_dov_son.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            excel();
        }
    }
}
