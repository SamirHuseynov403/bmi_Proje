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
    public partial class Verilmis_kr : Form
    {
        public Verilmis_kr()
        {
            InitializeComponent();
        }
        DataTable Ordt = new DataTable();
        private void testyoxla()
        {
            Ordt.Clear();
            //try
            //{
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select * from odb.graphpogkre t where t.licschkre=21210000001835100000 and t.subschkre=0", Orcon);

                string tarixIl = DateTime.Now.Date.Year.ToString();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa, t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('29-10-2021','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //k.fifd,where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,
                //OracleCommand Orcom = new OracleCommand("select lk.subschkre,l.name_licsch,lk.licschkre,lk.summakre,cri.initialamountofcredit,lk.procstavrez,sr.setmininterestreserves,sr.bgn,sr.aylig_borc_yuku,sr.aylig_gelir,sr.item_09,cig.guarantee_name,cig.guarantee_id,cig.dateofbirth,cig.pincode,LENGTH(cig.pincode) from licschkre lk,creditinfo cri,srokpogprockre sr,licsch l,creditinfoguarantee cig where lk.licschkre = cri.licschkre and lk.subschkre = cri.subschkre and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = l.licsch and lk.licschkre = cig.licschkre and cig.subschkre = lk.subschkre and lk.date_open >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') ", Orcon);
                //OracleCommand Orcom = new OracleCommand(" select m.subschkre,m.name_licsch,m.licschkre,m.summakre,m.initialamountofcredit,m.procstavrez,m.setmininterestreserves,m.bgn,round(nvl((m.ay_hisse+m.aylig_borc_yuku)/NULLIF(m.aylig_gelir,0)*100,0),2),m.bgn-round(nvl((m.ay_hisse+m.aylig_borc_yuku)/m.aylig_gelir*100,0),2),nvl(m.item_09,0),m.guarantee_name,m.guarantee_id,m.dateofbirth,m.pincode,LENGTH(m.pincode),nvl(m.ay_hisse,0),nvl(m.aylig_borc_yuku,0),nvl(m.aylig_gelir,0) from (select lk.subschkre,l.name_licsch,lk.licschkre,lk.summakre,cri.initialamountofcredit,lk.procstavrez,sr.setmininterestreserves,sr.bgn,sr.aylig_borc_yuku, sr.aylig_gelir,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = lk.subschkre and g1.licschkre = lk.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = lk.subschkre and gr3.licschkre = lk.licschkre)))  ay_hisse,sr.item_09,cig.guarantee_name,cig.guarantee_id,cig.dateofbirth,cig.pincode,LENGTH(cig.pincode) from licschkre lk,creditinfo cri,srokpogprockre sr,licsch l,creditinfoguarantee cig where lk.licschkre = cri.licschkre and lk.subschkre = cri.subschkre and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = l.licsch and lk.licschkre = cig.licschkre and cig.subschkre = lk.subschkre and lk.date_open >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy'))m ", Orcon);
                 OracleCommand Orcom = new OracleCommand("select m.subschkre, m.name_licsch, m.licschkre, m.summakre, m.initialamountofcredit, m.procstavrez, m.setmininterestreserves, m.bgn,round(nvl((m.ay_hisse + m.aylig_borc_yuku) / NULLIF(m.aylig_gelir, 0) * 100, 0), 2), m.bgn - round(nvl((m.ay_hisse + m.aylig_borc_yuku) / NULLIF(m.aylig_gelir, 0) * 100, 0), 2),m.item_09, m.guarantee_name, m.guarantee_id, m.dateofbirth, m.pincode, LENGTH(m.pincode),nvl(m.ay_hisse, 0), nvl(m.aylig_borc_yuku, 0), nvl(m.aylig_gelir, 0)from(select lk.subschkre, l.name_licsch, lk.licschkre, lk.summakre, cri.initialamountofcredit, lk.procstavrez, sr.setmininterestreserves, sr.bgn, sr.aylig_borc_yuku, sr.aylig_gelir,(select g1.summa_pog_kre + g1.summa_pog_pro from odb.graphpogkre g1  where g1.subschkre = lk.subschkre  and g1.licschkre = lk.licschkre  and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = lk.subschkre  and gr3.licschkre = lk.licschkre))) ay_hisse, sr.item_09, cig.guarantee_name, cig.guarantee_id, cig.dateofbirth,  cig.pincode,  LENGTH(cig.pincode) from licschkre lk,  creditinfo cri, srokpogprockre  sr, licsch l, creditinfoguarantee cig where lk.licschkre = cri.licschkre(+) and lk.subschkre = cri.subschkre(+)  and lk.licschkre = sr.licschkre(+)    and lk.subschkre = sr.subschkre(+) and lk.licschkre = l.licsch(+) and lk.licschkre = cig.licschkre(+)  and lk.subschkre = cig.subschkre(+) and lk.date_open >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')) m", Orcon);
            //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,k.fifd,t.procstavkre fz, t.procstav_19,t.procstavrez, t.srok,t.kurator, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open > to_date('01-01-2006','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
            //(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();

                dataGridView1.Columns[0].HeaderText = "Sub";
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[1].HeaderText = "Borcalan";
                dataGridView1.Columns[1].Width = 300;
                dataGridView1.Columns[2].HeaderText = "hesab";
                dataGridView1.Columns[2].Width = 150;
                dataGridView1.Columns[3].HeaderText = "Kat_məbləğ";
                dataGridView1.Columns[3].Width = 100;
                dataGridView1.Columns[4].HeaderText = "İnfo məb";
                dataGridView1.Columns[4].Width = 100;
                dataGridView1.Columns[5].HeaderText = "Kat eht";
                dataGridView1.Columns[5].Width = 80;
                dataGridView1.Columns[6].HeaderText = "İnfo eht";
                dataGridView1.Columns[6].Width = 100;
                dataGridView1.Columns[7].HeaderText = "BGN";
                dataGridView1.Columns[7].Width = 80;
                dataGridView1.Columns[8].HeaderText = "Düz BGN";
                dataGridView1.Columns[8].Width = 80;
                dataGridView1.Columns[9].HeaderText = "BGN fərq";
                dataGridView1.Columns[9].Width = 80;
                dataGridView1.Columns[10].HeaderText = "MRT";
                dataGridView1.Columns[10].Width = 100;
                dataGridView1.Columns[11].HeaderText = "Zam ad";
                dataGridView1.Columns[11].Width = 300;
                dataGridView1.Columns[12].HeaderText = "Zam AZE";
                dataGridView1.Columns[12].Width = 100;
                dataGridView1.Columns[13].HeaderText = "Zam dogum";
                dataGridView1.Columns[13].Width = 100;
                dataGridView1.Columns[14].HeaderText = "Zam FİN";
                dataGridView1.Columns[14].Width = 100;
                dataGridView1.Columns[15].HeaderText = "FİN";
                dataGridView1.Columns[15].Width = 50;
                dataGridView1.Columns[16].HeaderText = "Aylıq";
                dataGridView1.Columns[16].Width = 100;
                dataGridView1.Columns[17].HeaderText = "Borc";
                dataGridView1.Columns[17].Width = 100;
                dataGridView1.Columns[18].HeaderText = "Gəlir";
                dataGridView1.Columns[18].Width = 100;
            //}
            //catch (Exception)
            //{
                
            //}
            
            


            ////and t.procstavrez<=25 
        }

        private void Verilmis_kr_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
        }

        private void button1_Click(object sender, EventArgs e)
        {

            testyoxla();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
