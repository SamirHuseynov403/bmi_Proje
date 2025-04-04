using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
//using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;


namespace BMI
{
    public partial class Anakredit : Form
    {
        public Anakredit()
        {
            InitializeComponent();
        }
        public string icraciaditam { get; set; }
        public string icraciaditam2 { get; set; }
        public string icraci_kod { get; set; }
        public string Daxiltarix = string.Empty;
        private void button1_Click(object sender, EventArgs e)
        {
            fr2lbltip.Text = frm2cmbmuqtipi.Text;
            Qizil qzl = new Qizil();
            Menzil mnz = new Menzil();
            frmmenzil frmm = new frmmenzil();
            Zamin zmn = new Zamin();
            Avtomobil avto = new Avtomobil();
            Zaminler zmnlartarix = new Zaminler();
            Zaminler zmnalrmuqtipi=new Zaminler();
            zmn.label24.Text = "111111";
            zmn.test = frm2cmbmuqtipi.Text;
            zmnalrmuqtipi.muqtipi = frm2cmbmuqtipi.Text;
            qzl.test = frm2cmbmuqtipi.Text;
            qzl.muqtipi = frm2cmbmuqtipi.Text;
            qzl.girovnovucombo = fr2lbltip.Text;
            zmn.girovnovucombo = fr2lbltip.Text;
            qzl.girovnovucomboqizil = frm2cmbmuqtipi.Text;
            if (frm2cmbmuqtipi.Text=="Avtomobil")
            {
                Zaminler zmnlar = new Zaminler();
                Avtomobil zaminnn = new Avtomobil();
                zmnlar.subsot = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                zaminnn.icraciadizaminlikde = label12.Text; ;
                avtoat();
                zaminleriat();
            }
            else if (frm2cmbmuqtipi.Text == "Zaminlik")
            {
                Zaminler  zmnlar = new Zaminler();
                Zamin zaminnn = new Zamin();
                zmnlar.subsot = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                zaminnn.icraciadizaminlikde = label12.Text;
                zmnlar.qzlmuqtipi = frm2cmbmuqtipi.Text;
                string altt = "hh";
                zaminlikat();
                zaminleriat();
            }
            else if (frm2cmbmuqtipi.Text == "Zaminlik" || comboBox2.Text=="adi")
            {
                zmn.ShowDialog();
            }
            else if (frm2cmbmuqtipi.Text == "Daşınmaz Əmlak")
            {
                Zaminler zmnlar = new Zaminler();
                frmmenzil zaminnn = new frmmenzil();
                zmnlar.subsot = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                zaminnn.icraciadizaminlikde = label12.Text;
                zmnlar.qzlmuqtipi = frm2cmbmuqtipi.Text;
                zaminnn.test = frm2cmbmuqtipi.Text;
                string altt = "hh";
                menzilat();
            }
            else if (frm2cmbmuqtipi.Text == "Qızıl girovu")
            {
                Zaminler zmnlar = new Zaminler();
                Qizil zaminnn = new Qizil();
                zmnlar.subsot = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                zaminnn.icraciadizaminlikde = label12.Text;
                zmnlar.qzlmuqtipi = frm2cmbmuqtipi.Text;
                string altt = "hh";
                zaminnn.test = frm2cmbmuqtipi.Text;
                qizil_melat();
                zaminleriat();
            }
        }
        DateTime dt = DateTime.Now;
        string tarix_soz;
        public void tarixitapsoz()
        {
            int ay=dt.Month;
            if (ay == 1)
            {
                tarix_soz = dt.Day.ToString() + " Yanvar " + dt.Year.ToString();
            }
            else if (ay == 2)
            {
                tarix_soz = dt.Day.ToString() + " Fevral " + dt.Year.ToString();
            }
            else if (ay == 3)
            {
                tarix_soz = dt.Day.ToString() + " Mart " + dt.Year.ToString();
            }
            else if (ay == 4)
            {
                tarix_soz = dt.Day.ToString() + " Aprel " + dt.Year.ToString();
            }
            else if (ay == 5)
            {
                tarix_soz = dt.Day.ToString() + " May " + dt.Year.ToString();
            }
            else if (ay == 6)
            {
                tarix_soz = dt.Day.ToString() + " Iyun " + dt.Year.ToString();
            }
            else if (ay == 7)
            {
                tarix_soz = dt.Day.ToString() + " Iyul " + dt.Year.ToString();
            }
            else if (ay == 8)
            {
                tarix_soz = dt.Day.ToString() + " Avqust " + dt.Year.ToString();
            }
            else if (ay == 9)
            {
                tarix_soz = dt.Day.ToString() + " Sentyabr " + dt.Year.ToString();
            }
            else if (ay == 10)
            {
                tarix_soz = dt.Day.ToString() + " Oktyabr " + dt.Year.ToString();
            }
            else if (ay == 11)
            {
                tarix_soz = dt.Day.ToString() + " Noyabr " + dt.Year.ToString();
            }
            else if (ay == 12)
            {
                tarix_soz = dt.Day.ToString() + " Dekabr " + dt.Year.ToString();
            }
        }
        private void Anakredit_Load(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            tarixitapsoz();
            dateTimePicker1.Text= DateTime.Now.Date.ToString("dd-MM-yyyy");
            textBox1.Text=DateTime.Now.Date.ToString("dd-MM-yyyy");
            button4.Visible = false;
            button1.Enabled = true;
            kataloqgetir_zamin_ayliq();
            label12.Text = frm1.lblTamad.Text;
            testyoxla();
            if (comboBox2.Text== "Adi")
            {
                button4.Visible = true;
                button1.Enabled = false;
            }
        }
        public string adi;
        private void zaminlikat()
        {
            //try
            //{
                Zamin zmn1 = new Zamin();
                Zaminler zmnlar = new Zaminler();
                zmn1.adi = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                zmn1.subkod_qeyd = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                zmn1.hesab = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                zmn1.teyinat = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                zmn1.mebleg = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                zmn1.ayliq = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                zmn1.fifd = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                zmn1.faiz = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                zmn1.vkfaiz = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                zmn1.ehtfaiz = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                zmn1.muddet = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                zmn1.seriyano = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                zmn1.ver_orqan = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                zmn1.ver_tar = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                zmn1.mobil = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                zmn1.unvan = dataGridView1.CurrentRow.Cells[18].Value.ToString();
                zmn1.krtarixi = dataGridView1.CurrentRow.Cells[4].Value.ToString();

            zmn1.olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();
                zmn1.sudahes = dataGridView1.CurrentRow.Cells[20].Value.ToString();
                zmn1.faizhes = dataGridView1.CurrentRow.Cells[21].Value.ToString();
                zmn1.vkhes = dataGridView1.CurrentRow.Cells[22].Value.ToString();
                zmn1.vkfaizhes = dataGridView1.CurrentRow.Cells[23].Value.ToString();
                zmn1.odgunu = dataGridView1.CurrentRow.Cells[24].Value.ToString();
                zmn1.valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
                zmn1.subhes = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                zmn1.test =frm2cmbmuqtipi.Text;
                zmn1.umumitar = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                zmn1.carihes = dataGridView1.CurrentRow.Cells[26].Value.ToString();
                zmn1.tamhesab = dataGridView1.CurrentRow.Cells[27].Value.ToString();
                zmn1.tamtarix = tarix_soz;
                zmn1.comboadi = comboBox2.Text;
                zmn1.ShowDialog();
            }
        private void qizil_melat()
        {
            //try
            //{

            Qizil zmn1 = new Qizil();
            Zaminler zmnlar = new Zaminler();

            zmn1.adi = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            zmn1.subkod_qeyd = dataGridView1.CurrentRow.Cells[2].Value.ToString();

            zmn1.hesab = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            zmn1.teyinat = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            zmn1.mebleg = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            zmn1.ayliq = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            zmn1.fifd = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            zmn1.faiz = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            zmn1.vkfaiz = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            zmn1.ehtfaiz = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            zmn1.muddet = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            zmn1.seriyano = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            zmn1.ver_orqan = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            zmn1.ver_tar = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            zmn1.mobil = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            zmn1.unvan = dataGridView1.CurrentRow.Cells[18].Value.ToString();

            zmn1.test = frm2cmbmuqtipi.Text;
            zmn1.umumitar = dateTimePicker1.Value.ToString("dd-MM-yyyy");

            zmn1.olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            zmn1.sudahes = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            zmn1.faizhes = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            zmn1.vkhes = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            zmn1.vkfaizhes = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            zmn1.odgunu = dataGridView1.CurrentRow.Cells[24].Value.ToString();
            zmn1.valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            zmn1.subhes = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            zmn1.qzlgirovu = frm2cmbmuqtipi.Text;

            zmn1.carihes = dataGridView1.CurrentRow.Cells[26].Value.ToString();
            zmn1.tamhesab = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            zmn1.tamtarix = tarix_soz;
            zmn1.comboadi = comboBox2.Text;

            zmn1.ShowDialog();
        }
        private void avtoat()
        {
            //try
            //{
            Avtomobil zmn1 = new Avtomobil();
            Zaminler zmnlar = new Zaminler();

            zmn1.adi = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            zmn1.subkod_qeyd = dataGridView1.CurrentRow.Cells[2].Value.ToString();

            zmn1.hesab = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            zmn1.teyinat = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            zmn1.mebleg = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            zmn1.ayliq = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            zmn1.fifd = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            zmn1.faiz = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            zmn1.vkfaiz = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            zmn1.ehtfaiz = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            zmn1.muddet = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            zmn1.seriyano = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            zmn1.ver_orqan = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            zmn1.ver_tar = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            zmn1.mobil = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            zmn1.unvan = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            zmn1.girov_deyeri = dataGridView1.CurrentRow.Cells[28].Value.ToString();

            zmn1.olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            zmn1.sudahes = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            zmn1.faizhes = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            zmn1.vkhes = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            zmn1.vkfaizhes = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            zmn1.odgunu = dataGridView1.CurrentRow.Cells[24].Value.ToString();
            zmn1.valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            zmn1.subhes = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            zmn1.test = frm2cmbmuqtipi.Text;
            zmn1.umumitar = dateTimePicker1.Value.ToString("dd-MM-yyyy");

            zmn1.carihes = dataGridView1.CurrentRow.Cells[26].Value.ToString();
            zmn1.tamhesab = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            zmn1.tamtarix = tarix_soz;
            zmn1.comboadi = comboBox2.Text;
            zmn1.icraci_kod = icraci_kod;

            zmn1.ShowDialog();
        }
        private void menzilat()
        {
            //try
            //{
            frmmenzil zmn1 = new frmmenzil();
            Zaminler zmnlar = new Zaminler();

            zmn1.adi = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            zmn1.subkod_qeyd = dataGridView1.CurrentRow.Cells[2].Value.ToString();

            zmn1.hesab = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            zmn1.teyinat = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            zmn1.mebleg = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            zmn1.ayliq = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            zmn1.fifd = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            zmn1.faiz = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            zmn1.vkfaiz = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            zmn1.ehtfaiz = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            zmn1.muddet = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            zmn1.seriyano = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            zmn1.ver_orqan = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            zmn1.ver_tar = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            zmn1.mobil = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            zmn1.unvan = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            zmn1.girov_deyeri = dataGridView1.CurrentRow.Cells[28].Value.ToString();

            zmn1.olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            zmn1.sudahes = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            zmn1.faizhes = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            zmn1.vkhes = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            zmn1.vkfaizhes = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            zmn1.odgunu = dataGridView1.CurrentRow.Cells[24].Value.ToString();
            zmn1.valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            zmn1.subhes = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            zmn1.test = frm2cmbmuqtipi.Text;
            zmn1.umumitar = dateTimePicker1.Value.ToString("dd-MM-yyyy");

            zmn1.carihes = dataGridView1.CurrentRow.Cells[26].Value.ToString();
            zmn1.tamhesab = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            zmn1.tamtarix = tarix_soz;
            zmn1.comboadi = comboBox2.Text;
            zmn1.icraci_kod = icraci_kod;
            zmn1.ShowDialog();
        }
        private void qizilat()
        {
            //try
            //{
            //Avtomobil zmn1 = new Avtomobil();
            Zaminler zmnlar = new Zaminler();
            Qizil zmn1 = new Qizil();

            zmn1.adi = dataGridView1.CurrentRow.Cells[0].Value.ToString();

            zmn1.subkod_qeyd = dataGridView1.CurrentRow.Cells[2].Value.ToString();

            zmn1.hesab = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            zmn1.teyinat = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            zmn1.mebleg = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            zmn1.ayliq = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            zmn1.fifd = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            zmn1.faiz = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            zmn1.vkfaiz = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            zmn1.ehtfaiz = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            zmn1.muddet = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            zmn1.seriyano = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            zmn1.ver_orqan = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            zmn1.ver_tar = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            zmn1.mobil = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            zmn1.unvan = dataGridView1.CurrentRow.Cells[18].Value.ToString();

            zmn1.olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            zmn1.sudahes = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            zmn1.faizhes = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            zmn1.vkhes = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            zmn1.vkfaizhes = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            zmn1.odgunu = dataGridView1.CurrentRow.Cells[24].Value.ToString();
            zmn1.valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            zmn1.subhes = dataGridView1.CurrentRow.Cells[26].Value.ToString();

            zmn1.carihes = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            zmn1.tamhesab = dataGridView1.CurrentRow.Cells[28].Value.ToString();
            zmn1.tamtarix = tarix_soz;
            zmn1.comboadi = comboBox2.Text;
            zmn1.ShowDialog();
        }
        public void zaminleriat()
        {
            Zaminler zmnlar = new Zaminler();
            zmnlar.zamadi1 = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            zmnlar.zamseriya1 = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            zmnlar.zamtel1 = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            zmnlar.zamunvan1 = dataGridView1.CurrentRow.Cells[20].Value.ToString();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Qizil qzl = new Qizil();
            Menzil mnz = new Menzil();
            Zamin zmn = new Zamin();
            Avtomobil avto = new Avtomobil();

            if (frm2cmbmuqtipi.Text == "Avtomobil")
            {
                avto.ShowDialog();
            }
            else if (frm2cmbmuqtipi.Text == "Zaminlik")
            {
                zmn.ShowDialog();
            }
            else if (frm2cmbmuqtipi.Text == "Daşınmaz Əmlak")
            {
                mnz.ShowDialog();
            }
            else if (frm2cmbmuqtipi.Text == "Qızıl girovu")
            {
                qzl.ShowDialog();
            }
        }
        private void testyoxla()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            //(ar.kredit=l.licschpkre or ar.kredit=l.licschppkre) and

            OracleCommand Orcom = new OracleCommand("select * from odb.arh_licschkre lk where lk.date_oper='04-11-2023' and lk.date_close is null ", Orcon);
            //    "arh_dd ar where date_open> '05-01-2023' and (l.index_otrasli='01902'or index_otrasli='01903') " +
            //    "and substr(l.licschkre,11,6)=substr(ar.kredit,11,6) and substr(ar.debet,0,1)='1' and ar.date_oper=l.date_opeN ", Orcon);

            //OracleCommand Orcom = new OracleCommand(" select r.name_regnom,sum(ar.summa_v_nacval), ar.debet,ar.kredit from licschkre l," +
            //    "arh_dd ar,regnom r where ar.date_oper between to_date('01-10-2023', 'dd/mm/yyyy') and to_date('31-10-2023','dd / mm / yyyy') " +
            //    "and ar.debet=l.licschpkre and substr(ar.kredit,0,5)='64012' and l.tipzaloga='10' and r.regnom=substr(l.licschkre,10,6) and r.svazanniy='1' and l.subschkre=ar.ssd GROUP BY r.name_regnom, ar.debet,ar.kredit", Orcon);

            //OracleCommand verilmis_inex_kr = new OracleCommand("select l.index_otrasli,count(l.summakre),sum(l.summakre) from licschkre l " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') GROUP BY l.index_otrasli ", Orcon);

            //OracleCommand verilmis_girov_kr = new OracleCommand("select g.name,count(l.licschkre),sum(l.summakre) from licschkre l,tipzal g " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') and l.tipzaloga = g.code GROUP BY g.name ", Orcon);

            //OracleCommand verilmis_nov_kr = new OracleCommand("select g.name,count(l.licschkre),sum(l.summakre) from licschkre l,tipkre g " +
            //    " where l.date_open between to_date('01-01-2023', 'dd/mm/yyyy') and to_date('31-01-2023','dd / mm / yyyy') and l.tipkredita = g.code  GROUP BY g.name ", Orcon);

            //OracleCommand Orcom = new OracleCommand(" select ar.date_oper,ar.debet,ar.kredit,ar.summa_v_nacval from  arh_dd ar where ar.date_oper between to_date('01-10-2023', 'dd/mm/yyyy') and to_date('31-10-2023','dd / mm / yyyy') " +
            //    " and ar.debet= and ar.kredit='64012000000000600000' ", Orcon);

            //            OracleCommand odenilmis_nov_kr = new OracleCommand(" SELECT g.name,COUNT(DISTINCT ar.debet)," +
            //             "sum( CASE WHEN ar.kredit = l.licschkre THEN ar.summa_v_nacval ELSE 0 END) AS esas, " +
            //            " sum( CASE WHEN ar.kredit = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //             " sum( CASE WHEN ar.kredit = l.licsch_19 THEN ar.summa_v_nacval ELSE 0 END) AS vk, " +
            //            " sum( CASE WHEN ar.kredit = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //"           FROM licschkre l " +
            //"           JOIN tipkre g ON l.tipkredita = g.code " +
            //"            JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
            //"           WHERE(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
            //            "     AND l.subschkre = ar.ssk and substr(ar.debet,0,1) in (3,4)  " +
            //"           GROUP BY g.name ", Orcon);
            //OracleCommand odenilmis_girov_kr = new OracleCommand("select l.licsch,f.ish_yerinin_adi from licsch l,fiziki_shexs f where substr(l.licsch,0,1)=4 " +
            //    "and substr(l.licsch,16,2)=91 and l.date_close_licsch is null and substr(l.licsch,10,6)=f.regnom ", Orcon);

            //OracleCommand odenilmis_girov_kr = new OracleCommand("select l.registrac_nomer,substr(l.licsch,0,5),l.licsch,l.name_licsch,r.yurik," +
            //    "r.fizik,r.predprinimatel,r.finsektor from licsch l,regnom r" +
            //    " where r.regnom=l.registrac_nomer and l.date_close_licsch is null and (substr(l.licsch,0,3) in(359,389,399,409) " +
            //    "or substr(l.licsch,0,5) in (41932,41942,41943) or r.regnom='000004' or r.yurik=1 or r.predprinimatel=1 or r.finsektor=1 or ( substr(l.licsch,16,2)=91 and r.fizik=1)) " +
            //    " and substr(l.licsch,0,1)in (3,4) and substr(l.licsch,0,2)<>'45' and substr(l.licsch,0,2)<>'44' and r.svazanniy=0 ", Orcon);

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

            OracleCommand sirketin_tesiscileri = new OracleCommand("SELECT * from balsch b where b.balsch='35729' ", Orcon);
            //OracleCommand sirketin_tesiscileri = new OracleCommand("select * from licsch ch " +
            //    "where ch.date_close_licsch is null ", Orcon);


            OracleCommand sirketin_iscileri = new OracleCommand("select distinct ar.date_oper,ar.licsch,-ar.saldo_ish_nacval qaliq, " +
                " ch.name_licsch,z.ish_ad from odb.arh_saldo_ls ar,regnom r,licsch ch,  " +
                " (select f.fin fin,f.ish_yerinin_adi ish_ad  " +
                 "from regnom r,fiziki_shexs f  " +
                 "where f.ish_yerinin_adi=r.name_regnom) z where z.fin=r.pincode and  " +
                "( ch.date_close_licsch is null or  ar.date_oper<= ch.date_close_licsch ) " +
                " and r.regnom=ch.registrac_nomer and substr(ar.licsch,0,1)in (4) and ch.licsch=ar.licsch  " +
                " and substr(ar.licsch,16,2)=91 and ar.date_oper between to_date('23/05/2023')  and to_date('26/05/2023')  " +
                " order by ar.date_oper ", Orcon);

            OracleCommand ikinci_sorgu = new OracleCommand("select date_oper,k.licsch hesab,-sum(k.saldo_vhd_inval) Valyuta_ile,-sum(k.saldo_ish_nacval) AZN_ile from odb.arh_saldo_ls k, odb.licsch p " +
             " where  p.licsch = k.licsch and k.date_oper between to_date('01/10/2023') and to_date('02/10/2023') " +
             " and substr(k.licsch,0,2)in (35) group by k.date_oper,k.licsch " +
             " order by k.date_oper,k.licsch ", Orcon);

            OracleCommand dorduncu_sorgu = new OracleCommand("select ar.date_oper tarix,substr(l.licsch,0,5) hesab,-sum(ar.saldo_vhd_inval) Valyuta_ile," +
                "-sum(ar.saldo_ish_nacval) AZN_ile " +
                " from licsch l, odb.arh_saldo_ls ar " +
                " where ( l.date_close_licsch is null or ar.date_oper<= l.date_close_licsch ) and " +
                " substr(l.licsch,0,5)in (35025,35026,49025,49027) and l.registrac_nomer='000016' " +
                "and ar.licsch=l.licsch group by ar.date_oper, substr(l.licsch,0,5) order by ar.date_oper,substr(l.licsch,0,5) ", Orcon);

            OracleCommand sirketin_tesiscileri1 = new OracleCommand("select ar.debet db,ar.kredit kr,ar.summa_v_nacval meb, " +
"case when(ar.kredit = '21912000000001000001' or ar.kredit = '20910000000001000001') then '1' else substr(l.name_licsch, 1, instr(l.name_licsch, '%') - 1) end tey, " +
 "        case when(substr(ar.debet, 1, 4) = '8916' or substr(ar.debet, 1, 4) = '8911') then 'sah' else 'fizik' end tey " +
 "from arh_dd ar,licsch l " +
 " where ar.date_oper between to_date('29/12/2023', 'DD/MM/YYYY') and to_date('30/12/2023','DD/MM/YYYY') " +
 "and substr(ar.debet,1,2)= '89' and substr(ar.kredit,1,1)= '2' and ar.kredit = l.licsch", Orcon);

            //OracleCommand hesablanmis_faiz_girov_kr = new OracleCommand(" SELECT g.name," +
            //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipzal g ON l.tipzaloga = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('30-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
            //" GROUP BY g.name ", Orcon);

            //OracleCommand hesablanmis_faiz_nov_kr = new OracleCommand(" SELECT g.name," +
            //      " sum( CASE WHEN ar.debet = l.licschpkre THEN ar.summa_v_nacval ELSE 0 END) AS faiz, " +
            //      " sum( CASE WHEN ar.debet = l.licschppkre THEN ar.summa_v_nacval ELSE 0 END) AS vk_faiz " +
            //" FROM licschkre l " +
            //" JOIN tipkre g ON l.tipkredita = g.code " +
            //" JOIN arh_dd ar ON ar.date_oper BETWEEN TO_DATE('01-10-2023', 'dd/mm/yyyy') AND TO_DATE('02-10-2023', 'dd/mm/yyyy') " +
            //" WHERE(ar.debet = l.licschpkre OR ar.debet = l.licschppkre) " +
            //  "     AND l.subschkre = ar.ssd and substr(ar.kredit,0,1) in (6) " +
            //" GROUP BY g.name ", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(sirketin_tesiscileri1);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView2.DataSource = Ordt;
            Orcon.Close();
        }
        
        private void kataloqgetir_zaminsiz()
        {

            try
            {
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
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

                dataGridView1.Columns[1].HeaderText = "KS";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Sub-qeyd";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Hesab No";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "ver.tarixi";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Təyinat";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Məbləğ";
                dataGridView1.Columns[6].Width = 120;

                dataGridView1.Columns[7].HeaderText = "Məbləğ AZN";
                dataGridView1.Columns[7].Width = 120;

                dataGridView1.Columns[8].HeaderText = "Aylıq";
                dataGridView1.Columns[8].Width = 100;

                dataGridView1.Columns[9].HeaderText = "Fifd";
                dataGridView1.Columns[9].Width = 100;

                dataGridView1.Columns[10].HeaderText = "Faiz";
                dataGridView1.Columns[10].Width = 100;

                dataGridView1.Columns[11].HeaderText = "V/K %";
                dataGridView1.Columns[11].Width = 110;

                dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
                dataGridView1.Columns[12].Width = 120;

                dataGridView1.Columns[13].HeaderText = "Müddət";
                dataGridView1.Columns[13].Width = 120;

                dataGridView1.Columns[14].HeaderText = "Seriya No";
                dataGridView1.Columns[14].Width = 120;

                dataGridView1.Columns[15].HeaderText = "Verən orqan";
                dataGridView1.Columns[15].Width = 200;

                dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
                dataGridView1.Columns[16].Width = 300;

                dataGridView1.Columns[17].HeaderText = "Mobil";
                dataGridView1.Columns[17].Width = 300;

                dataGridView1.Columns[18].HeaderText = "Ünvan";
                dataGridView1.Columns[18].Width = 100;

                

                //dataGridView1.Columns[30].HeaderText = "tam hesab";
                //dataGridView1.Columns[30].Width = 200;
            }
            catch (Exception)
            {


            }

            finally
            { }
            

        }
        private void kataloqgetir_zamin_ayliq()
        {//t.subschkre t.licschkre
            try
            {
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select m.*,k.subschkre sk,k.licschkre,k.licsch_19,k.licschpkre,k.licschppkre,k.summakre kred_mab,k.summa kred_qal,k.summa_19 kred_vk, NVL((vk.nacpro_ish-vk.pogpro_ish),0) meb_24, NVL((vk.nacprospro_ish-vk.pogprospro_ish),0) meb19_24,p.mabl qraf_mab,s.ayliq,k.summakre-p.mabl farq,odb.func_utf8_to_latin(n.qeyd) qeyd from odb.licschkre k,odb.nacpogprokre vk,(select.func_utf8_to_latin(h.name_licsch),1,40)ad, t3.od_val, t3.od_man, ROUND(abs(h.saldo_ish_inval),2) qal_val, ROUND(abs(h.saldo_ish_nacval),2) qal_man from (select  t.debet, t.kredit, D(sum(t.summa_v_inval),2) , ROUND(sum(t.summa_v_nacval),2) od_man  from odb.docdna t, odb.balschkli j where substr(t.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t.kredit,1,5)=j.balsch  and ie)like '%KRED%' or upper(t.primechanie) like 'PORTMANAT%' or upper(t.primechanie) like 'EMANAT%' or upper(t.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t.primechanie)) not like'NOVBATI%')group by t.debet,t.kredit ) t3, (select t1.debet,sum(t1.summa_v_inval) val, sum(t1.summa_v_nacval) man from odb.docdna t1, odb.balschkli y  where substr(t1.debet,1,5)=y.balsch andsubstr(t1.kredit,1,2) in (20,21,23,63,64,65,67) group by t1.debet) t4, odb.licsch h where t3.kredit=t4.debet and t3.od_val<>t4.val and t3.od_man<>t4.man and t3.kredit=h.licsch union select t5.debet,t5.kredit, SUBSTR(odb.func_utf8_to_latin(h1.name_licsch),1,40) ad, ROUND(sum(t5.summa_v_inval),2) od_val, ROUND(sum(t5.summa_v_nacval),2) od_man, ROUND(abs(h1.saldo_ish_inval),2) qal_val, 1.saldo_ish_nacval),2)qal_man  from odb.docdna t5, odb.licsch h1, odb.balschkli j  where substr(t5.debet,1,5) in (10010,10020,11010,11020,25019,35090,35100,45050,45079,45089) and substr(t5.kredit,1,5)=j.balsch and ((upper(t5.primechanie) like '%KRED%' or upper(t5.primechanie) like 'PORTMANAT%' or upper(t5.primechanie) like 'EMANAT%' or upper(t5.primechanie) like 'MILLION%') and odb.func_utf8_to_latin(upper(t5.primechanie)) not  like '%NOVBATI%') and t5.kredit not in(select t11.debet from odb.docdna t11, odb.balschkli y where substr(t11.debet,1,5)=y.balsch and substr(t11.kredit,1,2) in(20,21,23,63,64,65,67))  and t5.kredit=h1.licsch  group by t5.debet,t5.kredit,h1.name_licsch,h1.saldo_ish_inval,h1.saldo_ish_nacval) m,(select t.licschkre,t.subschkre sk,sum(t.summa_pog_kre) mabl from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy')group by t.licschkre,t.subschkre) p,(select distinct t.licschkre,t.subschkre mma_pog_pro ayliq from odb.graphpogkre t where length(t.licschkre)=20 and to_date(t.date_pog,'dd/mm/yyyy')<=to_date(sysdate+30,'dd/mm/yyyy') ) s,(select  t.licschkre hes, t.subschkre sk, odb.func_utf8_to_latin(t.item_01)qeyd from odb.srokpogprockre t where not t.item_01  is null) n where k.licschkre=p.licschkre and k.subschkre=p.sk and k.licschpkre=vk.licschpkre and k.subschkre=vk.subschkre and k.licschkre=s.licschkre and k.subschkre=s.sk and substr(m.kredit,10,6)=substr(p.licschkre,10,6) and k.date_close is null and k.licschkre=n.hes(+) and k.subschkre=n.sk(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,y.pincode  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k ,odb.creditinfo y where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and t.licschkre = y.licschkre  and t.subschkre = y.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Adı";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "KS";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Sub-qeyd";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Hesab No";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "ver.tarixi";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Təyinat";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Məbləğ";
                dataGridView1.Columns[6].Width = 120;

                dataGridView1.Columns[7].HeaderText = "Məbləğ AZN";
                dataGridView1.Columns[7].Width = 120;

                dataGridView1.Columns[8].HeaderText = "Aylıq";
                dataGridView1.Columns[8].Width = 100;

                dataGridView1.Columns[9].HeaderText = "Fifd";
                dataGridView1.Columns[9].Width = 100;

                dataGridView1.Columns[10].HeaderText = "Faiz";
                dataGridView1.Columns[10].Width = 100;

                dataGridView1.Columns[11].HeaderText = "V/K %";
                dataGridView1.Columns[11].Width = 110;

                dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
                dataGridView1.Columns[12].Width = 120;

                dataGridView1.Columns[13].HeaderText = "Müddət";
                dataGridView1.Columns[13].Width = 120;

                dataGridView1.Columns[14].HeaderText = "Seriya No";
                dataGridView1.Columns[14].Width = 120;

                dataGridView1.Columns[15].HeaderText = "Verən orqan";
                dataGridView1.Columns[15].Width = 200;

                dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
                dataGridView1.Columns[16].Width = 300;

                dataGridView1.Columns[17].HeaderText = "Mobil";
                dataGridView1.Columns[17].Width = 300;

                dataGridView1.Columns[18].HeaderText = "Ünvan";
                dataGridView1.Columns[18].Width = 100;

                dataGridView1.Columns[19].HeaderText = "Zamin adı";
                dataGridView1.Columns[19].Width = 300;

                dataGridView1.Columns[20].HeaderText = "Zamin seriya No";
                dataGridView1.Columns[20].Width = 200;

                dataGridView1.Columns[21].HeaderText = "Zamin Ünvan";
                dataGridView1.Columns[21].Width = 200;

                dataGridView1.Columns[22].HeaderText = "Zamin telefon";
                dataGridView1.Columns[22].Width = 200;

                //dataGridView1.Columns[30].HeaderText = "tam hesab";
                //dataGridView1.Columns[30].Width = 200;
            }
            catch (Exception)
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre,t.summa,'',k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Adı";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "KS";
                dataGridView1.Columns[1].Width = 55;

                dataGridView1.Columns[2].HeaderText = "Sub-qeyd";
                dataGridView1.Columns[2].Width = 55;

                dataGridView1.Columns[3].HeaderText = "Hesab No";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "ver.tarixi";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Təyinat";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Məbləğ";
                dataGridView1.Columns[6].Width = 120;

                dataGridView1.Columns[7].HeaderText = "Məbləğ AZN";
                dataGridView1.Columns[7].Width = 120;

                dataGridView1.Columns[8].HeaderText = "Aylıq";
                dataGridView1.Columns[8].Width = 100;

                dataGridView1.Columns[9].HeaderText = "Fifd";
                dataGridView1.Columns[9].Width = 100;

                dataGridView1.Columns[10].HeaderText = "Faiz";
                dataGridView1.Columns[10].Width = 100;

                dataGridView1.Columns[11].HeaderText = "V/K %";
                dataGridView1.Columns[11].Width = 110;

                dataGridView1.Columns[12].HeaderText = "Ehtiyyat %";
                dataGridView1.Columns[12].Width = 120;

                dataGridView1.Columns[13].HeaderText = "Müddət";
                dataGridView1.Columns[13].Width = 120;

                dataGridView1.Columns[14].HeaderText = "Seriya No";
                dataGridView1.Columns[14].Width = 120;

                dataGridView1.Columns[15].HeaderText = "Verən orqan";
                dataGridView1.Columns[15].Width = 200;

                dataGridView1.Columns[16].HeaderText = "Verilmə tarixi";
                dataGridView1.Columns[16].Width = 300;

                dataGridView1.Columns[17].HeaderText = "Mobil";
                dataGridView1.Columns[17].Width = 300;

                dataGridView1.Columns[18].HeaderText = "Ünvan";
                dataGridView1.Columns[18].Width = 100;

                dataGridView1.Columns[19].HeaderText = "Zamin adı";
                dataGridView1.Columns[19].Width = 300;

                dataGridView1.Columns[20].HeaderText = "Zamin seriya No";
                dataGridView1.Columns[20].Width = 200;

                dataGridView1.Columns[21].HeaderText = "Zamin Ünvan";
                dataGridView1.Columns[21].Width = 200;

                dataGridView1.Columns[22].HeaderText = "Zamin telefon";
                dataGridView1.Columns[22].Width = 200;

            }

            finally
            { }
            

        }
        private void melumatlaridoldurzamin()

        { 
         
        }
        private void YOXLA()////  Burda qalmisdiq verilme tarixi ilie eyni olanlar gelsin
        {
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse ,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,k.licsch_3 m.az||m.nr||m.bank||m.licsch cari from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('18/08/2021', 'dd/mm/yyyy') and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();

            //dataGridView1.Columns[0].HeaderText = "Ады";
            //dataGridView1.Columns[0].Width = 302;
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            kataloqgetir_zamin_ayliq();
            //testyoxla();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            kataloqgetir_zaminsiz();            
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();

            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand komut = new OracleCommand("insert into muqavile_nomreleri(kr_serencam,kr_zaminlik,kr_menzil,kr_avtomobil,kr_zaminler,depozit,kr_kart,kart_zamin,il) values('1','1','1','1','1','0','0','0','" + tarixIl + "')", Orcon);
            komut.ExecuteNonQuery();
            Orcon.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            qizilmelumatlari qzl = new qizilmelumatlari();
            qzl.ShowDialog();
        }
    }
}
