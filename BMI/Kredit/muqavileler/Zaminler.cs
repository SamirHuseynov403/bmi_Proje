using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class Zaminler : Form
    {
        public Zaminler()
        {
            InitializeComponent();
        }
        public string zamadi1 { get; set; }
        public string zamseriya1 { get; set; }
        public string zamunvan1 { get; set; }
        public string zamtel1 { get; set; }

        public string zamadi2 { get; set; }
        public string zamseriya2 { get; set; }
        public string zamunvan2 { get; set; }
        public string zamtel2 { get; set; }

        public string zamadi3 { get; set; }
        public string zamseriya3 { get; set; }
        public string zamunvan3 { get; set; }
        public string zamtel3 { get; set; }

        public string testNO { get; set; }
        public string qeydnozamin { get; set; }

        public string zmnlartarix { get; set; }
        public string comboadi { get; set; }

        public string muqtipi { get; set; }
        public string muqavileadi { get; set; }

        public string kataloqtarixi { get; set; }
        public string qzlmuqtipi { get; set; }

        int kecencavab = 0;
        public string subsot { get; set; }

        private void zamingetir()
        {
            try
            {
                string tarix = kataloqtarixi;
                dateTimePicker2.Text = zmnlartarix;//DateTime.Now.Date.ToString("dd-MM-yyyy");
                label20.Text = qeydnozamin;
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select g.guarantee_name, g.guarantee_id,  g.adress,g.telefon, r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and substr(t.licschkre, 10, 6) || t.subschkre='" + label20.Text + "'and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);

                //OracleCommand Orcom = new OracleCommand("select g.guarantee_name, g.guarantee_id,  g.adress,g.telefon, r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and substr(t.licschkre, 10, 6) || t.subschkre='" + label20.Text + "' and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
                //OracleCommand Orcom = new OracleCommand("select * from CREDIT_CONTRACT", Orcon);
                kecencavab = Orcom.ExecuteNonQuery();
                //if (kecencavab > 0)
                //{
                    OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                    DataTable Ordt = new DataTable();
                    Orda.Fill(Ordt);
                    dataGridView1.DataSource = Ordt;
                    Orcon.Close();
                    dataGridView1.Columns[0].HeaderText = "Zamin adı";
                    dataGridView1.Columns[0].Width = 300;

                    dataGridView1.Columns[1].HeaderText = "Zamin seriya No";
                    dataGridView1.Columns[1].Width = 200;

                    dataGridView1.Columns[2].HeaderText = "Zamin Ünvan";
                    dataGridView1.Columns[2].Width = 200;

                    dataGridView1.Columns[3].HeaderText = "Zamin telefon";
                    dataGridView1.Columns[3].Width = 200; 
                //}
                //else
                //{
                //    MessageBox.Show("alinmadi");
                //}
                
            }
            catch (Exception)
            {


            }

            finally
            { }
            
        }

        private void Zaminler_Load(object sender, EventArgs e)
        {
            
            //Zamin anak = new Zamin();
            //txbzam1adi.Text = anak.adi;
            //txbzam1Pasport.Text = zamseriya1;
            //txbzam1Telefon.Text = zamtel1;
            //txbzam1Unvan.Text = zamunvan1;
            dateTimePicker2.Text = zmnlartarix; //DateTime.Now.Date.ToString("dd-MM-yyyy");
            if (radioButton1.Checked == true)
            {

                groupBox1.Visible = true;
                groupBox2.Visible = false;
                groupBox3.Visible = false;


            }
            else if (radioButton2.Checked == true)
            {
                groupBox1.Visible = true;
                groupBox2.Visible = true;
                groupBox3.Visible = false;
            }
            else if (radioButton3.Checked == true)
            {
                groupBox1.Visible = true;
                groupBox2.Visible = true;
                groupBox3.Visible = true;
            }
        }


        //and substr(t.licschkre, 10, 6) || t.subschkre='0165701'
        public void zamnoat()
        {
            Zamin anak = new Zamin();
            
        }

        

        private void zamindoldurzamin()
        {

            Zamin zmn = (Zamin)Application.OpenForms["zamin"];
            zmn.zam1adi = this.txbzam1adi.Text;
            zmn.zam1pass = this.txbzam1Pasport.Text;
            zmn.zam1tel = this.txbzam1Telefon.Text;
            zmn.zam1unvan = this.txbzam1Unvan.Text;
            zmn.zam1Pastar = this.txbzam1pastar.Text;
            zmn.zam1Pasorqan = this.cboxzam1Orqan.Text;
            zmn.zam1olkesi = this.cboxzam1Olke.Text;

            zmn.zaminmuqn1 = this.txb_zam1No.Text;
            zmn.zaminmuqn2 = this.txb_zam2No.Text;
            zmn.zaminmuqn3 = this.txb_zam3No.Text;

            zmn.zam2adi = this.txbzam2Zamin.Text;
            zmn.zam2pass = this.txbzam2Pasport.Text;
            zmn.zam2tel = this.txbzam2Telefon.Text;
            zmn.zam2unvan = this.txbzam2Unvan.Text;
            zmn.zam2Pastar = this.txbzam2pastar.Text;
            zmn.zam2Pasorqan = this.cboxzam2Orqan.Text;
            zmn.zam2olkesi = this.cboxzam2Olke.Text;

            zmn.zam3adi = this.txbzam3Zamin.Text;
            zmn.zam3pass = this.txbzam3Pasport.Text;
            zmn.zam3tel = this.txbzam3Telefon.Text;
            zmn.zam3unvan = this.txbzam3Unvan.Text;
            zmn.zam3Pastar = this.txbzam3pastar.Text;
            zmn.zam3Pasorqan = this.cboxzam3Orqan.Text;
            zmn.zam3olkesi = this.cboxzam3Olke.Text;
        }
        private void zamindoldurqizil()
        {

            Qizil zmn = (Qizil)Application.OpenForms["qizil"];
            zmn.zam1adi = this.txbzam1adi.Text;
            zmn.zam1pass = this.txbzam1Pasport.Text;
            zmn.zam1tel = this.txbzam1Telefon.Text;
            zmn.zam1unvan = this.txbzam1Unvan.Text;
            zmn.zam1Pastar = this.txbzam1pastar.Text;
            zmn.zam1Pasorqan = this.cboxzam1Orqan.Text;
            zmn.zam1olkesi = this.cboxzam1Olke.Text;

            zmn.zaminmuqn1 = this.txb_zam1No.Text;
            zmn.zaminmuqn2 = this.txb_zam2No.Text;
            zmn.zaminmuqn3 = this.txb_zam3No.Text;

            zmn.zam2adi = this.txbzam2Zamin.Text;
            zmn.zam2pass = this.txbzam2Pasport.Text;
            zmn.zam2tel = this.txbzam2Telefon.Text;
            zmn.zam2unvan = this.txbzam2Unvan.Text;
            zmn.zam2Pastar = this.txbzam2pastar.Text;
            zmn.zam2Pasorqan = this.cboxzam2Orqan.Text;
            zmn.zam2olkesi = this.cboxzam2Olke.Text;

            zmn.zam3adi = this.txbzam3Zamin.Text;
            zmn.zam3pass = this.txbzam3Pasport.Text;
            zmn.zam3tel = this.txbzam3Telefon.Text;
            zmn.zam3unvan = this.txbzam3Unvan.Text;
            zmn.zam3Pastar = this.txbzam3pastar.Text;
            zmn.zam3Pasorqan = this.cboxzam3Orqan.Text;
            zmn.zam3olkesi = this.cboxzam3Olke.Text;
        }
        private void zamindolduravto()
        {

            Avtomobil zmn = (Avtomobil)Application.OpenForms["avtomobil"];
            zmn.zam1adi = this.txbzam1adi.Text;
            zmn.zam1pass = this.txbzam1Pasport.Text;
            zmn.zam1tel = this.txbzam1Telefon.Text;
            zmn.zam1unvan = this.txbzam1Unvan.Text;
            zmn.zam1Pastar = this.txbzam1pastar.Text;
            zmn.zam1Pasorqan = this.cboxzam1Orqan.Text;
            zmn.zam1olkesi = this.cboxzam1Olke.Text;

            zmn.zaminmuqn1 = this.txb_zam1No.Text;
            zmn.zaminmuqn2 = this.txb_zam2No.Text;
            zmn.zaminmuqn3 = this.txb_zam3No.Text;

            zmn.zam2adi = this.txbzam2Zamin.Text;
            zmn.zam2pass = this.txbzam2Pasport.Text;
            zmn.zam2tel = this.txbzam2Telefon.Text;
            zmn.zam2unvan = this.txbzam2Unvan.Text;
            zmn.zam2Pastar = this.txbzam2pastar.Text;
            zmn.zam2Pasorqan = this.cboxzam2Orqan.Text;
            zmn.zam2olkesi = this.cboxzam2Olke.Text;

            zmn.zam3adi = this.txbzam3Zamin.Text;
            zmn.zam3pass = this.txbzam3Pasport.Text;
            zmn.zam3tel = this.txbzam3Telefon.Text;
            zmn.zam3unvan = this.txbzam3Unvan.Text;
            zmn.zam3Pastar = this.txbzam3pastar.Text;
            zmn.zam3Pasorqan = this.cboxzam3Orqan.Text;
            zmn.zam3olkesi = this.cboxzam3Olke.Text;
        }
        private void zamindoldurmenzil()
        {

            frmmenzil zmn = (frmmenzil)Application.OpenForms["frmmenzil"];
            zmn.zam1adi = this.txbzam1adi.Text;
            zmn.zam1pass = this.txbzam1Pasport.Text;
            zmn.zam1tel = this.txbzam1Telefon.Text;
            zmn.zam1unvan = this.txbzam1Unvan.Text;
            zmn.zam1Pastar = this.txbzam1pastar.Text;
            zmn.zam1Pasorqan = this.cboxzam1Orqan.Text;
            zmn.zam1olkesi = this.cboxzam1Olke.Text;

            zmn.zaminmuqn1 = this.txb_zam1No.Text;
            zmn.zaminmuqn2 = this.txb_zam2No.Text;
            zmn.zaminmuqn3 = this.txb_zam3No.Text;

            zmn.zam2adi = this.txbzam2Zamin.Text;
            zmn.zam2pass = this.txbzam2Pasport.Text;
            zmn.zam2tel = this.txbzam2Telefon.Text;
            zmn.zam2unvan = this.txbzam2Unvan.Text;
            zmn.zam2Pastar = this.txbzam2pastar.Text;
            zmn.zam2Pasorqan = this.cboxzam2Orqan.Text;
            zmn.zam2olkesi = this.cboxzam2Olke.Text;

            zmn.zam3adi = this.txbzam3Zamin.Text;
            zmn.zam3pass = this.txbzam3Pasport.Text;
            zmn.zam3tel = this.txbzam3Telefon.Text;
            zmn.zam3unvan = this.txbzam3Unvan.Text;
            zmn.zam3Pastar = this.txbzam3pastar.Text;
            zmn.zam3Pasorqan = this.cboxzam3Orqan.Text;
            zmn.zam3olkesi = this.cboxzam3Olke.Text;
        }
        private void btnzaminlerYazdir_Click(object sender, EventArgs e)
        {
            if (label22.Text=="Zaminlik")
            {
               zamindoldurzamin();
            }
            else if (label22.Text == "Qızıl girovu")
            {
                zamindoldurqizil();
            }
            else if (label22.Text == "menzil")
            {
                zamindoldurmenzil();
            }
        }

        private void verkreditler()
        {//t.subschkre t.licschkre
            Anakredit anakk = new Anakredit();
            dateTimePicker1.Text = zmnlartarix;// DateTime.Now.Date.ToString("dd-MM-yyyy");
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+) and substr(t.licschkre, 10, 6) || t.subschkre kod='" + subsot + "'", Orcon);
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
        private void kataloqgetir_zamin_ayliq()
        {//t.subschkre t.licschkre
            try
            {
                dateTimePicker1.Text = zmnlartarix;//DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
               // OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);

                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse , k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac, g.guarantee_name, g.guarantee_id,  g.adress,  g.telefon,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k, odb.licsch m  where t.date_open = to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')  and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre  and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);

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


            }

            finally
            { }


        }

        private void muq_no_tap()
        {
            Zamin anak = new Zamin();

            Zamin zmn = (Zamin)Application.OpenForms["Zamin"];
            if (groupBox3.Visible==true)
            {
                zmn.zam_say =Convert.ToInt32( txb_zam3No.Text);
            }
            else if (groupBox2.Visible==true)
            {
                zmn.label23.Text = txb_zam2No.Text;
            }
            else if (groupBox1.Visible == true)
            {
                zmn.label23.Text = txb_zam1No.Text;
            }

        }

        private void zamindoldur()
        {
            try
            {
                Zamin zmn = new Zamin();




                label20.Text = qeydnozamin;
                //dateTimePicker2.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                zamingetir();
                label21.Text = dataGridView1.RowCount.ToString();

                if (comboadi == "Bir nəfərin zəmanəti")
                {
                    txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                    txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                    txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
                }


                else if (comboadi == "İki nəfərin zəmanəti")
                {
                    //MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }
                }


                //else if (comboadi == "Üç nəfərin zəmanəti" && label21.Text == "3")
                //{
                //    MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                //}
                else if (comboadi == "Üç nəfərin zəmanəti")
                {
                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();

                        txbzam3Zamin.Text = dataGridView1.Rows[2].Cells[0].Value.ToString();
                        txbzam3Unvan.Text = dataGridView1.Rows[2].Cells[2].Value.ToString();
                        txbzam3Pasport.Text = dataGridView1.Rows[2].Cells[1].Value.ToString();
                        txbzam3Telefon.Text = dataGridView1.Rows[2].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }

                }






                //Anakredit anakk = new Anakredit();
                //label19.Text = anakk.dataGridView1.Rows[1].Cells[0].Value.ToString();
            }
            catch (Exception)
            {


            }
            finally { }
        }
        private void Zaminler_Load_1(object sender, EventArgs e)
        {
            dateTimePicker2.Text = zmnlartarix;
            //Zamin zmntest = new Zamin();
            label22.Text = muqtipi;
            try
            {
                Zamin zmn = new Zamin();




                label20.Text = qeydnozamin;
                dateTimePicker2.Text = zmnlartarix;//DateTime.Now.Date.ToString("dd-MM-yyyy");
                zamingetir();
                label21.Text = dataGridView1.RowCount.ToString();

                if (comboadi == "Bir nəfərin zəmanəti")
                {
                    txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                    txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                    txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
                    
                }


                else if (comboadi == "İki nəfərin zəmanəti")
                {
                    //MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }
                }


                //else if (comboadi == "Üç nəfərin zəmanəti" && label21.Text == "3")
                //{
                //    MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                //}
                else if (comboadi == "Üç nəfərin zəmanəti")
                {
                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();

                        txbzam3Zamin.Text = dataGridView1.Rows[2].Cells[0].Value.ToString();
                        txbzam3Unvan.Text = dataGridView1.Rows[2].Cells[2].Value.ToString();
                        txbzam3Pasport.Text = dataGridView1.Rows[2].Cells[1].Value.ToString();
                        txbzam3Telefon.Text = dataGridView1.Rows[2].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }

                }






                //Anakredit anakk = new Anakredit();
                //label19.Text = anakk.dataGridView1.Rows[1].Cells[0].Value.ToString();
            }
            catch (Exception)
            {

                
            }
            finally { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            zamingetir();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (groupBox2.Visible==true)
            {
                txbzam2Zamin.Text=dataGridView1.CurrentRow.Cells[0].Value.ToString();
                txbzam2Unvan.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                txbzam2Pasport.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                txbzam2Telefon.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            //zamingetir();
            zamindoldur();

            try
            {
                Zamin zmn = new Zamin();




                label20.Text = qeydnozamin;
                dateTimePicker2.Text = zmnlartarix;//"29-12-2021"; //DateTime.Now.Date.ToString("dd-MM-yyyy");
                zamingetir();
                label21.Text = dataGridView1.RowCount.ToString();

                if (comboadi == "Bir nəfərin zəmanəti")
                {
                    txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                    txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                    txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
                }


                else if (comboadi == "İki nəfərin zəmanəti")
                {
                    //MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }
                }


                //else if (comboadi == "Üç nəfərin zəmanəti" && label21.Text == "3")
                //{
                //    MessageBox.Show("Progess də bir nəfərin zaminliyi doldurulub,Siz iki nəfərin zaminliyini seçmisiniz.");

                //}
                else if (comboadi == "Üç nəfərin zəmanəti")
                {
                    try
                    {
                        txbzam1adi.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        txbzam1Unvan.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        txbzam1Pasport.Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        txbzam1Telefon.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();

                        txbzam2Zamin.Text = dataGridView1.Rows[1].Cells[0].Value.ToString();
                        txbzam2Unvan.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                        txbzam2Pasport.Text = dataGridView1.Rows[1].Cells[1].Value.ToString();
                        txbzam2Telefon.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();

                        txbzam3Zamin.Text = dataGridView1.Rows[2].Cells[0].Value.ToString();
                        txbzam3Unvan.Text = dataGridView1.Rows[2].Cells[2].Value.ToString();
                        txbzam3Pasport.Text = dataGridView1.Rows[2].Cells[1].Value.ToString();
                        txbzam3Telefon.Text = dataGridView1.Rows[2].Cells[3].Value.ToString();
                    }

                    catch (Exception)
                    {


                    }

                    finally
                    { }

                }






                //Anakredit anakk = new Anakredit();
                //label19.Text = anakk.dataGridView1.Rows[1].Cells[0].Value.ToString();
            }
            catch (Exception)
            {

                
            }
            finally { }
        }
        }
    
}
