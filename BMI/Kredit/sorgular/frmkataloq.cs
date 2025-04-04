using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class frmkataloq : Form
    {
        public frmkataloq()
        {
            InitializeComponent();
        }

        private void kataloq()
        {
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand Orcom = new OracleCommand("select distinct l.name_licsch, lk.region,lk.index_otrasli,lk.naznackredita,lk.subschkre,lk.licschkre,lk.licschpkre,lk.licsch_19,lk.licschppkre,lk.procstavkre,lk.procstav_19,lk.procstavrez,lk.procstavrez_19,lk.summakre,lk.summa,lk.summa_19,lk.date_open,lk.date_close,lk.date_planclose,lk.kolic_prolong,lk.date_prolong,lk.kolic_restructure,lk.date_restructure,lk.day_uderproc,lk.kurator,lk.srok,lk.lgotperiod,lk.tipkredita,lk.tipzaloga,lk.licsch_zaloga,lk.summa_zaloga,lk.kolic_pereocen_zaloga,lk.summa_pereocen_zaloga,lk.data_pereocen_zaloga,fs.gelir_menbeyi, fs.ish_yerinin_adi,fs.vezifesi,r.adress Unvan,r.registrac Qeydiyyat,fs.fin,fs.doguldugu_yer,r.semeynoye_polojenie,fs.pocht_indeksi,sr.item_01,sr.item_02,sr.item_03,sr.item_04,sr.item_05,sr.item_06,sr.item_07,sr.item_08,sr.item_09,sr.item_10 from licschkre lk, regnom r, fiziki_shexs fs, licsch l, srokpogprockre sr where lk.licschkre = l.licsch and substr(lk.licschkre, 10, 6) = l.registrac_nomer and substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 and lk.subschkre = sr.subschkre and lk.licschkre = sr.licschkre and substr(lk.licschkre, 10, 6) = fs.regnom and lk.date_close is null", Orcon);
            //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();

            dataGridView1.Columns[0].HeaderText = "Adı";
            dataGridView1.Columns[0].Width = 350;

            dataGridView1.Columns[1].HeaderText = "Region";
            dataGridView1.Columns[1].Width = 55;

            dataGridView1.Columns[2].HeaderText = "Ind";
            dataGridView1.Columns[2].Width = 55;

            dataGridView1.Columns[3].HeaderText = "K.T";
            dataGridView1.Columns[3].Width = 200;

            dataGridView1.Columns[4].HeaderText = "K.S";
            dataGridView1.Columns[4].Width = 120;

            dataGridView1.Columns[5].HeaderText = "Kredit hesabı";
            dataGridView1.Columns[5].Width = 120;

            dataGridView1.Columns[6].HeaderText = "Faiz hesabı";
            dataGridView1.Columns[6].Width = 120;

            dataGridView1.Columns[7].HeaderText = "VK hesabı";
            dataGridView1.Columns[7].Width = 120;

            dataGridView1.Columns[8].HeaderText = "VK % hesabı";
            dataGridView1.Columns[8].Width = 100;

            dataGridView1.Columns[9].HeaderText = "Faiz";
            dataGridView1.Columns[9].Width = 100;

            dataGridView1.Columns[10].HeaderText = "VK faiz";
            dataGridView1.Columns[10].Width = 100;

            dataGridView1.Columns[11].HeaderText = "Ehtiyyat";
            dataGridView1.Columns[11].Width = 110;

            dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
            dataGridView1.Columns[12].Width = 120;

            dataGridView1.Columns[13].HeaderText = "Kredit məbləği";
            dataGridView1.Columns[13].Width = 120;

            dataGridView1.Columns[14].HeaderText = "Qalıq";
            dataGridView1.Columns[14].Width = 120;

            dataGridView1.Columns[15].HeaderText = "VK qalıq";
            dataGridView1.Columns[15].Width = 200;

            dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
            dataGridView1.Columns[16].Width = 300;

            dataGridView1.Columns[17].HeaderText = "Bağlanma";
            dataGridView1.Columns[17].Width = 300;

            dataGridView1.Columns[18].HeaderText = "Bitmə tarixi";
            dataGridView1.Columns[18].Width = 100;

            dataGridView1.Columns[19].HeaderText = "U.S";
            dataGridView1.Columns[19].Width = 300;

            dataGridView1.Columns[20].HeaderText = "U.T";
            dataGridView1.Columns[20].Width = 200;

            dataGridView1.Columns[21].HeaderText = "RS";
            dataGridView1.Columns[21].Width = 200;

            dataGridView1.Columns[22].HeaderText = "R.T";
            dataGridView1.Columns[22].Width = 200;

            dataGridView1.Columns[23].HeaderText = "Ö.G";
            dataGridView1.Columns[23].Width = 200;

            dataGridView1.Columns[24].HeaderText = "K.K";
            dataGridView1.Columns[24].Width = 200;

            dataGridView1.Columns[25].HeaderText = "K.M";
            dataGridView1.Columns[25].Width = 200;

            dataGridView1.Columns[26].HeaderText = "G.M";
            dataGridView1.Columns[26].Width = 200;

            dataGridView1.Columns[27].HeaderText = "Kreditin növü";
            dataGridView1.Columns[27].Width = 200;

            dataGridView1.Columns[28].HeaderText = "Girovun növü";
            dataGridView1.Columns[28].Width = 200;

            dataGridView1.Columns[22].HeaderText = "Girovun hesabı";
            dataGridView1.Columns[22].Width = 200;

            dataGridView1.Columns[29].HeaderText = "R.T";
            dataGridView1.Columns[29].Width = 200;

            dataGridView1.Columns[30].HeaderText = "Sayı";
            dataGridView1.Columns[30].Width = 200;

            dataGridView1.Columns[31].HeaderText = "G.Y.G";
            dataGridView1.Columns[32].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "Mənbə növü";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "Tranzit hesabı";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "VK tranzit";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "G.Y.G";
            //dataGridView1.Columns[22].Width = 200;

            //dataGridView1.Columns[22].HeaderText = "Mənbə növü";
            //dataGridView1.Columns[22].Width = 200;
        }

        private void frmkataloq_Load(object sender, EventArgs e)
        {
            kataloq();
        }
    }
}
