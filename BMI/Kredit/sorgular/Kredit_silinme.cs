using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class Kredit_silinme : Form
    {
        public Kredit_silinme()
        {
            InitializeComponent();
        }

        
        private void listele()
        {
            try
            //to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy')
            {//and to_char(t.date_pog,'yyyy/mm')<=to_char(sysdate,'yyyy/mm')
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy'))s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
                //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select t3.debet, t3.kredit, SUBSTR(odb.func_utf8_to_latin(h.name_licsch),1,40) ad, t3.od_val, t3.od_man,        ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, ROUND(sum(t.summa_v_inval),2) od_val, ROUND(sum(t.summa_v_nacval),2) od_man        from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch        and ((upper(t.primechanie) like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%')        and odb.func_utf8_to_latin(upper(t.primechanie)) not like 'NOVBATI%')group by t.debet,t.kredit ) t3,       (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y        where substr(t1.debet,1,5)=y.balsch and substr(t1.kredit,1,2) in (20,21,23,63,64,65,67)group by t1.debet) t4, odb.licsch h  where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch  union  select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad,         ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man,        ROUND(abs(h1.saldo_ish_inval),2) qal_val, ROUND(abs(h1.saldo_ish_nacval),2) qal_man    from odb.docdna t5, odb.licsch h1, odb.balschkli j   where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch     and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%'      or upper(t5.primechanie) like 'MILLION%') and  odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in  (select t11.debet from odb.docdna t11, odb.balschkli y    where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in (20,21,23,63,64,65,67))      and t5.kredit=h1.licsch group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_char(t.date_pog,'yyyy/mm')<=to_char(sysdate,'yyyy/mm') group by t.licschkre,t.subschkre) p, (select distinct t.licschkre,t.subschkre sk,t.summa_pog_kre+t.summa_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01) qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre  and  k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+) and k.lgotperiod is null and k.bs_vbs is not null and to_date(sysdate,'dd/mm/yyyy')<= to_date(k.date_planclose,'dd/mm/yyyy')", Orcon);
                OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select t3.debet, t3.kredit, SUBSTR(odb.func_utf8_to_latin(h.name_licsch),1,40) ad, t3.od_val, t3.od_man,        ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, ROUND(sum(t.summa_v_inval),2) od_val, ROUND(sum(t.summa_v_nacval),2) od_man        from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch        and ((upper(t.primechanie) like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%')        and odb.func_utf8_to_latin(upper(t.primechanie)) not like 'NOVBATI%')group by t.debet,t.kredit ) t3,       (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y        where substr(t1.debet,1,5)=y.balsch and substr(t1.kredit,1,2) in (20,21,23,63,64,65,67)group by t1.debet) t4, odb.licsch h  where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch  union  select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad,         ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man,        ROUND(abs(h1.saldo_ish_inval),2) qal_val, ROUND(abs(h1.saldo_ish_nacval),2) qal_man    from odb.docdna t5, odb.licsch h1, odb.balschkli j   where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch     and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%'      or upper(t5.primechanie) like 'MILLION%') and  odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in  (select t11.debet from odb.docdna t11, odb.balschkli y    where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in (20,21,23,63,64,65,67))      and t5.kredit=h1.licsch group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_char(t.date_pog,'yyyy/mm')<=to_char(sysdate,'yyyy/mm') group by t.licschkre,t.subschkre) p, (select distinct t.licschkre,t.subschkre sk,t.summa_pog_kre+t.summa_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01) qeyd from odb.srokpogprockre t where not t.item_01  is null) n, (select count( *) kol,substr(t.licschkre,10,6) regnom from odb.licschkre t  where t.date_close is null group by substr(t.licschkre,10,6)) z where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre  and  k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+) and k.lgotperiod is null and k.bs_vbs is not null and to_date(sysdate,'dd/mm/yyyy')<= to_date(k.date_planclose,'dd/mm/yyyy')and z.regnom= substr(k.licschkre,10,6) and z.kol=1", Orcon);

                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dgw1.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                throw;
                
            }
            
        
        }

        DataTable tablo = new DataTable();
        DataTable tablo1 = new DataTable();
        int say = 0;
        int saytopla = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void ayricabaxis()
        {

            //double odenis = 0.0;

            //decimal qaliq = 0.00m;
            //double novbeti_qaliq = 0.00;
            //double qaliqferq = 0.0;
            //decimal  qaliqferq1 = 0.00m;
            //double ayliq = 0.0;
            //double cavab1 = 0.0;
            //double cavab2 = 0.0;
            //double cavab3 = 0.0;
            //double cavab4 = 0.0;

            //double qaliqvkfaiz = 1;
            //double qaliqvk = 1;
            //double qaliqfaiz = 1;
            //double qaliqesas = 1;
            //double sudaborcu = 0.0, faizborcu = 0.0, vkborc = 0.0, vkfaizborc = 0.0;
            //string sub = "", cari = "", suda = "", faiz = "", vk = "", vkfaiz = "", teyinatsuda = "kreditin odenilmesi", teyinatfaiz = "faiz borcun odenilmesi", teyinatvk = "v/k borcun odenilmesi", teyinatvkfaiz = "v/k faiz borcun odenilmesi";


            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{

            //    odenis = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString());
            //    qaliq = Convert.ToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
            //    novbeti_qaliq = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString());
            //    ayliq = Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value.ToString());

            //    sudaborcu = Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString());
            //    faizborcu = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString());
            //    vkborc = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString());
            //    vkfaizborc = Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value.ToString());

            //    cari = dataGridView1.Rows[i].Cells[0].Value.ToString();
            //    suda = dataGridView1.Rows[i].Cells[1].Value.ToString();
            //    faiz = dataGridView1.Rows[i].Cells[2].Value.ToString();
            //    vk = dataGridView1.Rows[i].Cells[3].Value.ToString();
            //    vkfaiz = dataGridView1.Rows[i].Cells[4].Value.ToString();
            //    sub = dataGridView1.Rows[i].Cells[12].Value.ToString();

            //    string debet = "";
            //    string subkod = "";
            //    string kredit = "";
            //    string mebleg = "";
            //    string teyinat = "";
            //    string baglandi = "baglandi";
            //    string gunferq = "gun ferqi";
            //    if (odenis >= sudaborcu + faizborcu + vkborc + vkfaizborc)
            //    {
            //        dataGridView3.Rows.Add(cari, baglandi, "butun qaliqlar sifirlandi");
            //    }

            //    //if (novbeti_qaliq<sudaborcu&&vkborc==0&&vkfaizborc==0&&odenis>=ayliq)
            //    //{
            //    //    qaliqferq = sudaborcu - novbeti_qaliq + faizborcu;
            //    //    if (odenis<qaliqferq)
            //    //    {
            //    //        int ferq = qaliqferq - odenis;
            //    //        dataGridView3.Rows.Add(cari, gunferq,"Gecikmeye dusecek mebleg"+" "+ferq);
            //    //    }
            //    //}
            //    if (vkborc==0)
            //    {
            //        if (odenis > faizborcu + vkborc + vkfaizborc && sudaborcu > novbeti_qaliq)
            //        {
            //            string sebeb = "";

            //            qaliqferq1 = sudaborcu - novbeti_qaliq + faizborcu + vkborc + vkfaizborc;
            //            if (odenis < qaliqferq1)
            //            {
            //                if (odenis < ayliq)
            //                {
            //                    sebeb = "az odenis";
            //                }
            //                else if (odenis >= ayliq)
            //                {
            //                    sebeb = "esas borcdan az silinme";
            //                }
            //                decimal ferq = 0.00m;
            //                ferq = qaliqferq1 - odenis;
            //                dataGridView3.Rows.Add(cari, sebeb, "Gecikmeye dusecek mebleg" + " " + ferq);
            //            }
            //        }
            //    }

            //}
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //dataGridView2.Rows.Add(bos1, bos1, bos1, debet, bos1, subkod, kredit, mebleg, bos1, bos1, teyinat);
            dataGridView2.ColumnCount = 10;
            dataGridView2.Columns[0].Name = "No";
            dataGridView2.Columns[0].Width = 40;
            dataGridView2.Columns[1].Name = "";
            dataGridView2.Columns[1].Width = 1;
            dataGridView2.Columns[2].Name = "";
            dataGridView2.Columns[2].Width = 1;

            dataGridView2.Columns[3].Name = "debet";
            dataGridView2.Columns[3].Width = 150;


            dataGridView2.Columns[4].Name = "subkod";
            dataGridView2.Columns[4].Width = 40;
            dataGridView2.Columns[5].Name = "kredit";
            dataGridView2.Columns[5].Width = 150;
            dataGridView2.Columns[6].Name = "mebleg";
            dataGridView2.Columns[6].Width = 70;
            dataGridView2.Columns[7].Name = "";
            dataGridView2.Columns[7].Width = 1;
            dataGridView2.Columns[8].Name = "";
            dataGridView2.Columns[8].Width = 1;
            dataGridView2.Columns[9].Name = "teyinat";
            dataGridView2.Columns[9].Width = 400;
            dataGridView3.ColumnCount = 3;
            dataGridView3.Columns[0].Name = "Hesab";
            dataGridView3.Columns[0].Width = 130;
            dataGridView3.Columns[1].Name = "Melumat";
            dataGridView3.Columns[1].Width = 160;
            dataGridView3.Columns[2].Name = "Aciqlama";
            dataGridView3.Columns[2].Width = 250;
            // listele();




        }
        DataTable tablodoldur = new DataTable();
        private void listelesil()
        {
            ///dataGridView1.RowCount = 13;
            string hesab = "41010000001253100000";
            string suda = "21210000001253100000";
            string sudafaiz = "21212000001253100000";
            string sudavk = "21213000001253100000";
            string sudavkfaiz = "21214000001253100000";

            string odenis = "300";
            string sudaqaliq = "4650,25";
            string qaliqfaiz = "60,45";
            string qaliqvk = "142,63";
            string qaliqvkfaiz = "12,49";
            string novbetiqaliq = "3892,41";
            string ayliq = "250,63";
            string subkod = "0";

            dgw1.Rows.Add(hesab, suda, sudafaiz, sudavk, sudavkfaiz, odenis, sudaqaliq, qaliqfaiz, qaliqvk, qaliqvkfaiz, novbetiqaliq, ayliq, subkod);
            //dataGridView1.DataSource = tablodoldur;

            string hesab1 = "41010000001240100000";
            string suda1 = "21210000001240100000";
            string sudafaiz1 = "21212000001240100000";
            string sudavk1 = "21213000001240100000";
            string sudavkfaiz1 = "21214000001240100000";

            string odenis1 = "240";
            string sudaqaliq1 = "3250,25";
            string qaliqfaiz1 = "0,45";
            string qaliqvk1 = "412,63";
            string qaliqvkfaiz1 = "62,49";
            string novbetiqaliq1 = "2192,41";
            string ayliq1 = "243,63";
            string subkod1 = "1";
            dgw1.Rows.Add(hesab1, suda1, sudafaiz1, sudavk1, sudavkfaiz1, odenis1, sudaqaliq1, qaliqfaiz1, qaliqvk1, qaliqvkfaiz1, novbetiqaliq1, ayliq1, subkod1);
            // dataGridView1.DataSource = tablodoldur;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ayricabaxis();
        }

        private void Kredit_silinme_Load(object sender, EventArgs e)
        {
            
            listele();
            dataGridView2.ColumnCount = 10;
            dataGridView2.Columns[0].Name = "No";
            dataGridView2.Columns[0].Width = 40;
            dataGridView2.Columns[1].Name = "";
            dataGridView2.Columns[1].Width = 1;
            dataGridView2.Columns[2].Name = "";
            dataGridView2.Columns[2].Width = 1;

            dataGridView2.Columns[3].Name = "debet";
            dataGridView2.Columns[3].Width = 150;


            dataGridView2.Columns[4].Name = "subkod";
            dataGridView2.Columns[4].Width = 40;
            dataGridView2.Columns[5].Name = "kredit";
            dataGridView2.Columns[5].Width = 150;
            dataGridView2.Columns[6].Name = "mebleg";
            dataGridView2.Columns[6].Width = 70;
            dataGridView2.Columns[7].Name = "";
            dataGridView2.Columns[7].Width = 1;
            dataGridView2.Columns[8].Name = "";
            dataGridView2.Columns[8].Width = 1;
            dataGridView2.Columns[9].Name = "teyinat";
            dataGridView2.Columns[9].Width = 400;
            dataGridView3.ColumnCount = 3;
            dataGridView3.Columns[0].Name = "Hesab";
            dataGridView3.Columns[0].Width = 130;
            dataGridView3.Columns[1].Name = "Melumat";
            dataGridView3.Columns[1].Width = 160;
            dataGridView3.Columns[2].Name = "Aciqlama";
            dataGridView3.Columns[2].Width = 250;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                say = Convert.ToInt16(textBox5.Text);
                saytopla = Convert.ToInt16(textBox5.Text);
                decimal odenis = 0.00m;
                decimal kredit_odenis = 0.00m;

                decimal qaliq = 0.00m;
                decimal novbeti_qaliq = 0.00m;
                //decimal qaliqferq = 0.00m;
                decimal qaliqferq1 = 0.00m;
                decimal ayliq = 0.00m;
                decimal cavab1 = 0.00m;
                decimal cavab2 = 0.00m;
                decimal cavab3 = 0.00m;
                decimal cavab4 = 0.00m;
                decimal qaliqcari = 0.00m;


                double qaliqvkfaiz = 1;
                double qaliqvk = 1;
                double qaliqfaiz = 1;
                double qaliqesas = 1;


                decimal sudaborcu = 0.00m, faizborcu = 0.00m, vkborc = 0.00m, vkfaizborc = 0.00m;
                string sub = "", cari = "", suda = "", faiz = "", vk = "", vkfaiz = "", teyinatsuda = "kreditin ödənilməsi", teyinatfaiz = "kreditin % üzrə borсun ödənilməsi", teyinatvk = "kreditin v/k hesabında olan məbləğin ödənilməsi", teyinatvkfaiz = "kreditin v/k % üzrə borсun ödənilməsi";


                for (int i = 0; i < dgw1.Rows.Count - 1; i++)
                {
                    kredit_odenis = Convert.ToDecimal(dgw1.Rows[i].Cells[4].Value.ToString());
                    odenis = Convert.ToDecimal(dgw1.Rows[i].Cells[6].Value.ToString());
                    qaliq = Convert.ToDecimal(dgw1.Rows[i].Cells[13].Value.ToString());
                    novbeti_qaliq = Convert.ToDecimal(dgw1.Rows[i].Cells[19].Value.ToString());
                    ayliq = Convert.ToDecimal(dgw1.Rows[i].Cells[18].Value.ToString());
                    qaliqcari = Convert.ToDecimal(dgw1.Rows[i].Cells[6].Value.ToString());

                    sudaborcu = Convert.ToDecimal(dgw1.Rows[i].Cells[13].Value.ToString());
                    faizborcu = Convert.ToDecimal(dgw1.Rows[i].Cells[15].Value.ToString());
                    vkborc = Convert.ToDecimal(dgw1.Rows[i].Cells[14].Value.ToString());
                    vkfaizborc = Convert.ToDecimal(dgw1.Rows[i].Cells[16].Value.ToString());

                    cari = dgw1.Rows[i].Cells[1].Value.ToString();
                    suda = dgw1.Rows[i].Cells[8].Value.ToString();
                    faiz = dgw1.Rows[i].Cells[10].Value.ToString();
                    vk = dgw1.Rows[i].Cells[9].Value.ToString();
                    vkfaiz = dgw1.Rows[i].Cells[11].Value.ToString();
                    sub = dgw1.Rows[i].Cells[7].Value.ToString();

                    string debet = "";
                    string subkod = "";
                    string kredit = "";
                    string mebleg = "";
                    string teyinat = "";
                    string baglandi = "baglandi";
                    string gunferq = "gun ferqi";
                    string bos1 = "";

                    if (qaliqcari > kredit_odenis)
                    {
                        string sebeb = "";
                        sebeb = "hesab qalğı ödənişdən coxdur";
                        decimal ferqcari = 0.00m;
                        ferqcari = qaliqcari - odenis;
                        dataGridView3.Rows.Add(cari, sebeb, "hesabdakı artıq mebleg" + " " + ferqcari);
                    }

                    if (odenis > kredit_odenis)
                    {
                        string sebeb = "";
                        sebeb = "hesab qalğı ödənişdən coxdur";
                        decimal ferqcari = 0.00m;
                        ferqcari = qaliqcari - odenis;
                        dataGridView3.Rows.Add(cari, sebeb, "hesabdakı artıq mebleg" + " " + ferqcari);
                    }

                    if (odenis >= sudaborcu + faizborcu + vkborc + vkfaizborc)
                    {
                        dataGridView3.Rows.Add(cari, baglandi, "butun qaliqlar sifirlandi");
                    }

                   

                    if (odenis > faizborcu + vkborc + vkfaizborc && sudaborcu > novbeti_qaliq)
                    {
                        string sebeb = "";

                        qaliqferq1 = sudaborcu - novbeti_qaliq + faizborcu + vkborc + vkfaizborc;

                        if (odenis < qaliqferq1)
                        {
                            if (odenis < ayliq)
                            {
                                sebeb = "az odenis";
                            }
                            if (qaliqcari > odenis)
                            {
                                sebeb = "hesab qalğı ödənişdən coxdur";
                                decimal ferqcari = 0.00m;
                                ferqcari = qaliqcari - odenis;
                                dataGridView3.Rows.Add(cari, sebeb, "Gecikmeye dusecek mebleg" + " " + ferqcari);
                            }
                            else if (odenis >= ayliq)
                            {
                                sebeb = "esas borcdan az silinme";
                            }
                            decimal ferq = 0.00m;
                            ferq = qaliqferq1 - odenis;
                            dataGridView3.Rows.Add(cari, sebeb, "Gecikmeye dusecek mebleg" + " " + ferq);
                        }
                    }


                    if (qaliqcari <= odenis)
                    {
                    if (vkfaizborc > 0) 
                    {


                        if (odenis >= vkfaizborc)
                        {
                            say = say + 1;

                            qaliqvkfaiz = 0;
                            cavab1 = vkfaizborc;
                            odenis = odenis - vkfaizborc;
                            debet = cari;
                            subkod = sub;
                            kredit = vkfaiz;
                            mebleg = cavab1.ToString();
                            teyinat = teyinatvkfaiz;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }
                        else if (odenis < vkfaizborc)
                        {
                            say = say + 1;
                            cavab1 = odenis;
                            odenis = odenis - odenis;
                            debet = cari;
                            subkod = sub;
                            kredit = vkfaiz;
                            mebleg = cavab1.ToString();
                            teyinat = teyinatvkfaiz;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }

                    }
                    if (vkborc > 0 && odenis > 0)
                    {
                        if (odenis >= vkborc)
                        {
                            say = say + 1;
                            qaliqvk = 0;
                            qaliqvkfaiz = 0;
                            cavab2 = vkborc;
                            odenis = odenis - vkborc;
                            debet = cari;
                            subkod = sub;
                            kredit = vk;
                            mebleg = cavab2.ToString();
                            teyinat = teyinatvk;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }
                        else if (odenis < vkborc)
                        {
                            say = say + 1;
                            cavab2 = odenis;
                            odenis = odenis - odenis;
                            debet = cari;
                            subkod = sub;
                            kredit = vk;
                            mebleg = cavab2.ToString();
                            teyinat = teyinatvk;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }

                    }
                    if (faizborcu > 0 && odenis > 0)
                    {
                        if (odenis >= faizborcu)
                        {
                            say = say + 1;
                            qaliqfaiz = 0;
                            qaliqvk = 0;
                            qaliqvkfaiz = 0;
                            cavab3 = faizborcu;
                            odenis = odenis - faizborcu;
                            debet = cari;
                            subkod = sub;
                            kredit = faiz;
                            mebleg = decimal.Round(cavab3, 2).ToString();
                            teyinat = teyinatfaiz;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }
                        else if (odenis < faizborcu)
                        {
                            say = say + 1;
                            cavab3 = odenis;
                            odenis = odenis - odenis;
                            debet = cari;
                            subkod = sub;
                            kredit = faiz;
                            mebleg = cavab3.ToString();
                            teyinat = teyinatfaiz;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }
                    }
                    if (sudaborcu > 0 && odenis > 0)
                    {
                        if (odenis >= sudaborcu)
                        {
                            say = say + 1;
                            qaliqfaiz = 0;
                            qaliqvk = 0;
                            qaliqvkfaiz = 0;
                            qaliqesas = 0;
                            cavab4 = sudaborcu;
                            //odenis = odenis - sudaborcu;
                            debet = cari;
                            subkod = sub;
                            kredit = suda;
                            mebleg = cavab4.ToString();
                            teyinat = teyinatsuda;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }
                        else if (odenis < sudaborcu)
                        {
                            say = say + 1;
                            cavab4 = odenis;
                            odenis = odenis - odenis;
                            debet = cari;
                            subkod = sub;
                            kredit = suda;
                            mebleg = cavab4.ToString();
                            teyinat = teyinatsuda;
                            dataGridView2.Rows.Add(say, bos1, bos1, debet, subkod, kredit, mebleg, bos1, bos1, teyinat);
                        }


                    }

                    }









































                    //    cavab1 = vkfaizborc;
                    //    odenis= odenis - vkfaizborc;
                    //    cavab2 = vkborc;
                    //    odenis = odenis - vkborc;
                    //    cavab3 = faizborcu;
                    //    odenis = odenis - faizborcu;
                    //    if (sudaborcu<odenis)
                    //    {
                    //        cavab4 = sudaborcu;
                    //    }

                    //    else if (sudaborcu >= odenis)
                    //    {
                    //        cavab2 = odenis;
                    //    }

                    //    //string debet = cari;
                    //    //string subkod = sub;
                    //    //string kredit = vkfaiz;
                    //    //string mebleg = cavab1.ToString();
                    //    //string teyinat = teyinatvkfaiz;
                    //    //dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);

                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = vk;
                    //    mebleg = cavab2.ToString();
                    //    teyinat = teyinatvk;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);

                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = faiz;
                    //    mebleg = cavab3.ToString();
                    //    teyinat = teyinatfaiz;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);

                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = suda;
                    //    mebleg = cavab4.ToString();
                    //    teyinat = teyinatsuda;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);
                    //}

                    //else if (vkborc>0)//
                    //{
                    //    //cavab1 = vkfaizborc;
                    //    //odenis = odenis - vkfaizborc;
                    //    cavab2 = vkborc;
                    //    odenis = odenis - vkborc;
                    //    cavab3 = faizborcu;
                    //    odenis = odenis - faizborcu;
                    //    if (sudaborcu<odenis)
                    //    {
                    //        cavab4 = sudaborcu;
                    //    }
                    //    else if (sudaborcu >= odenis)
                    //    {
                    //        cavab2 = odenis;
                    //    }

                    //    string debet = cari;
                    //    string subkod = sub;
                    //    string kredit = vk;
                    //    string mebleg = cavab2.ToString();
                    //    string teyinat = teyinatvk;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);



                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = faiz;
                    //    mebleg = cavab3.ToString();
                    //    teyinat = teyinatfaiz;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);

                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = suda;
                    //    mebleg = cavab4.ToString();
                    //    teyinat = teyinatsuda;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);
                    //}

                    //else if (vkfaizborc>0)
                    //{
                    //    cavab1 = vkfaizborc;
                    //    odenis = odenis - vkfaizborc;
                    //    cavab2 = vkborc;
                    //    odenis = odenis - vkborc;
                    //    cavab3 = faizborcu;
                    //    odenis = odenis - faizborcu;
                    //    if (sudaborcu<odenis)
                    //    {
                    //        cavab4 = sudaborcu;
                    //    }
                    //    else if (sudaborcu >= odenis)
                    //    {
                    //        cavab2 = odenis;
                    //    }

                    //    string debet = cari;
                    //    string subkod = sub;
                    //    string kredit = vkfaiz;
                    //    string mebleg = cavab1.ToString();
                    //    string teyinat = teyinatvkfaiz;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);



                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = faiz;
                    //    mebleg = cavab3.ToString();
                    //    teyinat = teyinatfaiz;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);

                    //    debet = cari;
                    //    subkod = sub;
                    //    kredit = suda;
                    //    mebleg = cavab4.ToString();
                    //    teyinat = teyinatsuda;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);
                    //}

                    //else if (sudaborcu>0)
                    //{
                    //    cavab1 = faizborcu;
                    //    odenis = odenis - faizborcu;
                    //    if (sudaborcu<odenis)
                    //    {
                    //        cavab2 = sudaborcu;
                    //    }
                    //    else if (sudaborcu>=odenis)
                    //    {
                    //        cavab2 = odenis;
                    //    }


                    //    string debet = cari;
                    //    string subkod = sub;
                    //    string kredit = faiz;
                    //    string mebleg = cavab1.ToString();
                    //    string teyinat = teyinatfaiz;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit,mebleg,teyinat);

                    //     debet = cari;
                    //     subkod = sub;
                    //     kredit = suda;
                    //     mebleg = cavab2.ToString();
                    //     teyinat = teyinatsuda;
                    //    dataGridView2.Rows.Add(debet, subkod, kredit, mebleg, teyinat);





                }
                //button1.Enabled = false;
            
            

            }
            catch (Exception)
            {
                
                
            }
            button6.Enabled = false;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = desktopFolder + "\\Kredit_sil_ler.xls";

            if (dgw1.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Kredit_sil_ler.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Kredit_sil_ler.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "Kredit_sil_ler.xls";
                    //  worksheet.Cells.Font.Size = 13;

                    //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    //{
                    //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    //    worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                    //    worksheet.Cells[1, i].Font.Bold = true;
                    //}
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                            worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                        }
                    }
                    app.Visible = true;
                    workbook.SaveAs(desktopFolder, Type.Missing);
                    app.Quit();
                }
            }

            fserqli_silinme();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            dgw1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void fserqli_silinme()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = desktopFolder + "\\Sehv_silinme.xls";

            if (dataGridView3.RowCount > 0)
            {
                if (!System.IO.File.Exists(Application.StartupPath + "\\Sehv_silinme.xls"))
                {
                    MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\Sehv_silinme.xls");
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                    worksheet.Name = "Sehv_silinme.xls";
                    //  worksheet.Cells.Font.Size = 13;

                    //for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    //{
                    //    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    //    worksheet.Cells[1, i].Font.name = "A3 Arial Azlat";
                    //    worksheet.Cells[1, i].Font.Bold = true;
                    //}
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView3.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value;
                            worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                        }
                    }
                    app.Visible = true;
                    workbook.SaveAs(desktopFolder, Type.Missing);
                    app.Quit();
                }
            }
        }
    }
}
