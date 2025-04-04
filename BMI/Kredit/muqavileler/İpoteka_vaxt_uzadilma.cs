using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Globalization;
using Xceed.Words.NET;
using DevExpress.XtraRichEdit.Export.WordML;

namespace BMI
{
    public partial class İpoteka_vaxt_uzadilma : Form
    {
        public İpoteka_vaxt_uzadilma()
        {
            InitializeComponent();
        }
        Aletler al=new Aletler();

        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;

        string olke = "";
        string ilkin_muqtarix = "";
        string VAL_AD;
        string muqavilNo = "";
        string arayTar = "";
        string teminatsay = "";
        string arayisad = "";
        string arayisNo = "";
        string muqavileili = "";
        public string mek_no { get; set; }

        private void kataloqgetir_zamin_ayliq()
        {//t.subschkre t.licschkre
            try
            {
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse
                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa,t.summa_19  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
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
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse
                OracleCommand Orcom = new OracleCommand("select r.name_regnom, t.subschkre sk,  substr(t.licschkre, 10, 6) || t.subschkre kod, t.licschkre, t.date_open,t.naznackredita, t.summakre, t.summa,t.summakre,k.fifd,t.procstavkre fz, t.procstav_19,   t.procstavrez, t.srok, r.passport, r.senedi_veren_orqaninin_adi ver_orq, r.senedin_verilme_tarixi ver_tar, r.telefon, r.registrac,r.grajdanstvo,t.licschkre,t.licschpkre,t.licsch_19,t.licschppkre,t.day_uderproc,substr(t.licschkre, 6, 2),k.Licsch_3, m.az||m.nr||m.bank||m.licsch cari,t.summa_zaloga,t.summa,t.summa_19  from odb.licschkre t, odb.regnom r,odb.licsch m,odb.srokpogprockre k where t.date_open = to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch=k.licsch_3(+)", Orcon);
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

            finally
            { }


        }
        private void olkeadi()
        {
            if (olke == "AZ")
            {
                comboBox2.Text = "Azərbaycan Respublikası";
            }
            else if (olke == "IRN")
            {
                comboBox2.Text = " İran İslam Respublikası";
            }
        }

        private void doldur()
        {
            double summa = 0.00;

            double summa_19 = 0.00;
            double cemqaliq = 0.00;
            string valyuta = "";
            valyuta = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            if (valyuta == "00")
            {
                cboxkzValyuta.Text = "AZN";
                textBox2.Text = "AZN";
                textBox3.Text = "AZN";
                VAL_AD = "AZN";

            }
            else if (valyuta == "01")
            {
                cboxkzValyuta.Text = "USD";
                textBox2.Text = "USD";
                textBox3.Text = "USD";
                VAL_AD = "USD";
            }
            else if (valyuta == "02")
            {
                cboxkzValyuta.Text = "AVRO";
                textBox2.Text = "AVRO";
                textBox3.Text = "AVRO";
                VAL_AD = "AVRO";

            }
            txbzAyliq.BackColor = Color.White;
            txbzBorcalan.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txbzmuqtar.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(0, 10);
            arayTar = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(0, 10);
            ilkin_muqtarix = txbzmuqtar.Text;
            summa = Convert.ToDouble(dataGridView1.CurrentRow.Cells[29].Value.ToString());
            summa_19 = Convert.ToDouble(dataGridView1.CurrentRow.Cells[30].Value.ToString());
            cemqaliq = summa + summa_19;
            qaliqborc.Text = cemqaliq.ToString();

            txbzCarihesab.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txbzPasport.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            cboxkzOrqan.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            datezPasvertarix.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();

            txbzmebleg.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            //txbzAyliq.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            //if (Convert.ToDouble(txbzAyliq.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString()) == Convert.ToDouble(txbzmebleg.Text))
            //{
            //    txbzAyliq.Text = "";
            //    txbzAyliq.BackColor = Color.Red;
            //}



            txbzFaiz.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();


            int muddeti = Convert.ToInt32(dataGridView1.CurrentRow.Cells[13].Value.ToString()) / 30;
            txbzMuddet.Text = muddeti.ToString();
            txbzTelefon.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            txbzUnvan.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            olke = dataGridView1.CurrentRow.Cells[19].Value.ToString();

            olkeadi();
            //tarixitapsoz();
            //tarixitapsozilkintarix();
            string adhazir;
            string unvanhazir;
            adhazir = txbzBorcalan.Text.ToLower();
            unvanhazir = txbzUnvan.Text.ToLower();
            txbzBorcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);
            txbzUnvan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(unvanhazir);

        }

        private void meknoal()
        {
            int test = 45;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();

            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('DƏDRX', TO_DATE('" + dateavtoMuqtarix.Text + "','dd-MM-yyyy'), 'mənzil girova sal' , '" + icraci_kod + "',  " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();
            //string gond_yer = "DYP";
            //string mezmun = "avto gir salınması";
            //mktb.Xmgonderilenyer = gond_yer.TrimStart();
            //mktb.Xaricmektbtarix = dateavtoMuqtarix.Text;
            //mktb.Xmgisamezmun = mezmun.TrimStart();
            //mktb.icracikodu = icraci_kod;
            ////frmMktbxrc.mktb.Xmmektubmetn = rctbMektubmetn.Text.TrimStart();
            //mktb.Xaricmektinsert();   

        }

        private void muracietsayartimsiz()
        {
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                string proid;
                string hevsecal;
                int mek_no_arti = 0;
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                //string query = "select qey_nom ,il from odb.xaric_mektub where il='2022' order by qey_nom desc";
                string query = "select max(-to_number(substr(t.qey_nom,5,5)))mn from odb.xaric_mektub t where t.il='" + tarixIl + "'";
                OracleCommand cmd = new OracleCommand(query, Orcon);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    int id = int.Parse(dr[0].ToString());
                    proid = id.ToString();
                    mek_no_arti = id + 1;

                    //hevsecal = proid;
                }
                else if (Convert.IsDBNull(dr))
                {
                    proid = ("1");
                }
                else
                {
                    proid = ("1");
                }
                Orcon.Close();
                //txbkavtoModel.Text =tarixIl+"-"+ mek_no_arti.ToString();
                mek_no = tarixIl + "-" + mek_no_arti.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }
        DateTime dt = DateTime.Now;
        string tarix_soz;
        string tarix_sozilkintarix;

        //public void tarixitapsoz()
        //{
        //    int ay = dt.Month;
        //    //textBox1.Text = dt.Day.ToString();
        //    //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
        //    //int ay = 6;

        //    if (ay == 1)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Yanvar " + dt.Year.ToString();
        //    }
        //    else if (ay == 2)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Fevral " + dt.Year.ToString();
        //    }
        //    else if (ay == 3)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Mart " + dt.Year.ToString();
        //    }
        //    else if (ay == 4)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Aprel " + dt.Year.ToString();
        //    }
        //    else if (ay == 5)
        //    {
        //        tarix_soz = dt.Day.ToString() + " May " + dt.Year.ToString();
        //    }
        //    else if (ay == 6)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Iyun " + dt.Year.ToString();
        //    }
        //    else if (ay == 7)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Iyul " + dt.Year.ToString();
        //    }
        //    else if (ay == 8)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Avqust " + dt.Year.ToString();
        //    }
        //    else if (ay == 9)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Sentyabr " + dt.Year.ToString();
        //    }
        //    else if (ay == 10)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Oktyabr " + dt.Year.ToString();
        //    }
        //    else if (ay == 11)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Noyabr " + dt.Year.ToString();
        //    }
        //    else if (ay == 12)
        //    {
        //        tarix_soz = dt.Day.ToString() + " Dekabr " + dt.Year.ToString();
        //    }
        //}
        //public void tarixitapsozilkintarix()
        //{
        //    int ay = Convert.ToInt16(txbzmuqtar.Text.Substring(3, 2));
        //    //textBox1.Text = dt.Day.ToString();
        //    //label14.Text = dt.Day.ToString() + " " + dt.Month.ToString() + "" + dt.Year.ToString();
        //    //int ay = 6;

        //    if (ay == 1)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Yanvar " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 2)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Fevral " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 3)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Mart " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 4)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Aprel " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 5)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " May " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 6)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " İyun " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 7)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " İyul " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 8)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Avqust " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 9)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Sentyabr " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 10)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Oktyabr " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 11)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Noyabr " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //    else if (ay == 12)
        //    {
        //        tarix_sozilkintarix = txbzmuqtar.Text.Substring(0, 2) + " Dekabr " + txbzmuqtar.Text.Substring(6, 4);
        //    }
        //}

        private void word_at()
        {
            Zaminler zmnlar = new Zaminler();
            string yaziile = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzmebleg.Text));
            string yaziileqaliq = yaziyaCevir(Convert.ToDecimal(qaliqborc.Text));
            string ayyazile = yaziyaCevir(Convert.ToDecimal(txbzAyliq.Text));
            string muddetyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzMuddet.Text));
            string faizyazi = yaziyaCevirqepiksiz(Convert.ToDecimal(txbzFaiz.Text));
            string ilsonu = Aletler.TarixinSonReqemineSufiksElaveEt(txt_uzad_tar.Text);


            var adi = txbzBorcalan.Text;
            var passport = txbzPasport.Text;
            var pasvertarix = Aletler.TarixiSozeCevir(datezPasvertarix.Text);
            var Orqan = cboxkzOrqan.Text;
            var Unvan = txbzUnvan.Text;
            var CariHesab = txbzCarihesab.Text;
            var Valyuta = cboxkzValyuta.Text;

            var MuqNo = txbmuqno.Text;
            var xaricmek = mek_no;
            var ElaveNo = Elave.Text;
            var Muqtarix = Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy"));
            var muqilkintar = Aletler.TarixiSozeCevir(txbzmuqtar.Text) ;
            var temsay = teminatsay;
            var arad = txbzBorcalan.Text;
            var İpoteka = comboBox5.Text;
            var İpoteka_unvan = textBox14.Text;
            var Notarius = textBox1.Text;
            var uzad_ay = textBox4.Text;
            var uzad_tarix = Aletler.TarixiSozeCevir(txt_uzad_tar.Text);
            var muqil = txbzmuqtar.Text.Substring(6, 4);

            var Olke = comboBox2.Text;
            var mebleg = txbzmebleg.Text + " " + cboxkzValyuta.Text;
            var qaliqmebleg = qaliqborc.Text + " " + cboxkzValyuta.Text;
            var muddet = txbzMuddet.Text;
            var faiz = txbzFaiz.Text;
            var telf = txbzTelefon.Text;
            var ayliq = txbzAyliq.Text;
            var zam1 = txbzam1adi.Text;
            var zam1pas = txbzam1Pasport.Text;
            var zam1tel = txbzam1Telefon.Text;
            var zam1unvan = txbzam1Unvan.Text;
            var zam1pastarixi = Aletler.TarixiSozeCevir(txbzam1pastar.Text);
            var zam1paso = cboxzam1Orqan.Text;
            var zam1olke = cboxzam1Olke.Text;
            var zammuqavile = txb_zam1No.Text;
            var mebyaziile = yaziile;
            var qalmebyaziile = yaziileqaliq;

            var zamilkadi = textBox8.Text;
            var zamilNo = textBox4.Text;
            var zamilktarix = Aletler.TarixiSozeCevir(textBox1.Text);

            var mevayyazi = ayyazile;
            //var tamcari = tamhesab;
            //var icraci = icraci_adi;
            //var icraciadi = icraciadizaminlikde;



            // TODO: Word Export
            var wordapp = new Microsoft.Office.Interop.Word.Application();
            wordapp.Visible = false;
            string yol = "";
            if (comboBox1.Text == "İki nəfərin zəmanəti")
            {
                yol = "\\Əlavə Vaxt uzatma EYNİ Kredit İpoteka Bir zamin.docx";
            }
            else
            {
                yol = "\\Əlavə Vaxt uzatma EYNİ Kredit İpoteka Bir zamin.docx";
            }

            var wordDocument = wordapp.Documents.Open(Application.StartupPath + yol);
            ReplaceWordStub("{adi}", adi, wordDocument);
            ReplaceWordStub("{passport}", passport, wordDocument);
            ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
            ReplaceWordStub("{Orqan}", Orqan, wordDocument);
            ReplaceWordStub("{Unvan}", Unvan, wordDocument);
            ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
            ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
            ReplaceWordStub("{Elave}", ElaveNo, wordDocument);
            ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
            ReplaceWordStub("{ilkinMuqtarix}", muqilkintar, wordDocument);
            ReplaceWordStub("{Olke}", Olke, wordDocument);
            ReplaceWordStub("{mebleg}", mebleg + "(" + mebyaziile + ")", wordDocument);
            ReplaceWordStub("{qalmebleg}", qaliqmebleg + "(" + qalmebyaziile + ")", wordDocument);
            ReplaceWordStub("{muddet}", muddet + " ay " + "(" + muddetyazi + ")", wordDocument);
            ReplaceWordStub("{faiz}", faiz + "% " + "(" + faizyazi + ")", wordDocument);
            ReplaceWordStub("{tel}", telf, wordDocument);
            ReplaceWordStub("{ayliq}", ayliq + VAL_AD + "(" + mevayyazi + ")", wordDocument);

            ReplaceWordStub("{teminat}", temsay, wordDocument);
            ReplaceWordStub("{aradi}", arad, wordDocument);
            ReplaceWordStub("{İpoteka}", İpoteka, wordDocument);
            ReplaceWordStub("{girovunvan}", İpoteka_unvan, wordDocument);
            ReplaceWordStub("{not_no}", Notarius, wordDocument);
            ReplaceWordStub("{uz_ayi}", uzad_ay, wordDocument);
            ReplaceWordStub("{uz_tarixi}", uzad_tarix, wordDocument);
            ReplaceWordStub("{il}", muqil, wordDocument);
            ReplaceWordStub("{xmekno}", xaricmek, wordDocument);

            ReplaceWordStub("{zammuq1}", zammuqavile, wordDocument);
            ReplaceWordStub("{zam1ad}", zam1, wordDocument);
            ReplaceWordStub("{zam1pas}", zam1pas, wordDocument);
            ReplaceWordStub("{zam1telef}", zam1tel, wordDocument);
            ReplaceWordStub("{zam1unvani}", zam1unvan, wordDocument);
            ReplaceWordStub("{zam1ptarix}", zam1pastarixi, wordDocument);
            ReplaceWordStub("{zam1pasorqan}", zam1paso, wordDocument);
            ReplaceWordStub("{zam1olke}", zam1olke, wordDocument);

            ReplaceWordStub("{zamilkad}", zamilkadi, wordDocument);
            ReplaceWordStub("{zammuq1ilk}", zamilNo, wordDocument);
            ReplaceWordStub("{ilkinmuqtarix1}", zamilktarix, wordDocument);


            //ReplaceWordStub("{zam2ad}", zam2, wordDocument);
            //ReplaceWordStub("{zam2pas}", zam2pas, wordDocument);
            //ReplaceWordStub("{zam2telef}", zam2tel, wordDocument);
            //ReplaceWordStub("{zam2unvani}", zam2unvan, wordDocument);
            //ReplaceWordStub("{zam2ptarix}", zam2pastarixi, wordDocument);
            //ReplaceWordStub("{zam2porqan}", zam2pasorqani, wordDocument);
            //ReplaceWordStub("{zam2olke}", zam2olke, wordDocument);

            //ReplaceWordStub("{zam3ad}", zam3, wordDocument);
            //ReplaceWordStub("{zam3pas}", zam3pas, wordDocument);
            //ReplaceWordStub("{zam3telef}", zam3tel, wordDocument);
            //ReplaceWordStub("{zam3unvani}", zam3unvan, wordDocument);
            //ReplaceWordStub("{zam3ptarix}", zam3pastarixi, wordDocument);
            //ReplaceWordStub("{zam3porqan}", zam3pasorqani, wordDocument);
            //ReplaceWordStub("{zam3olke}", zam3olke, wordDocument);

            //ReplaceWordStub("{zammuq1}", zammuqno1, wordDocument);
            //ReplaceWordStub("{zammuq2}", zammuqno2, wordDocument);
            //ReplaceWordStub("{zammuq3}", zammuqno3, wordDocument);



                wordapp.Visible = true;
                string muqadi = MuqNo + " " + adi + " " + tarix_soz;
                // wordDocument.SaveAs(@"‪‪\\192.168.0.5\kred_sob\Zaminlik" + muqadi);
                //wordDocument.SaveAs(@"‪C:\BMI_\BMI\bin\Debug\zaminlik_muq+"adi);
                wordDocument.SaveAs(Application.StartupPath + muqadi);


            //‪\\192.168.0.5\kred_sob\Kredit zaminlik
            //this.Application.Documents.Open(Application.StartupPath + "adi.docx", ReadOnly: true);
            //‪\\192.168.0.5\kred_sob\Zaminlik
            //wordDocument.Close(Type.Missing, Type.Missing, Type.Missing);
            ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler\Erize1.docx
            ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler

        }

        //[Obsolete]
        public void CreateWordDocument()
        {
            var replacements = new Dictionary<string, string>
{
            { "{adi}", txbzBorcalan.Text },
            { "{passport}", txbzPasport.Text },
            { "{pasvertarix}", Aletler.TarixiSozeCevir(datezPasvertarix.Text) },
            { "{Unvan}", txbzUnvan.Text },
            { "{Valyuta}", cboxkzValyuta.Text },
            { "{muddet}", txbzMuddet.Text },
            { "{faiz}", txbzFaiz.Text + "%" },
            { "{mebleg}", txbzmebleg.Text + " " + cboxkzValyuta.Text },
            { "{tel}", txbzTelefon.Text },
            { "{ayliq}", txbzAyliq.Text }
};

            // Şablon və çıxış yollarını təyin edin
            string templatePath = Application.StartupPath + "\\template.docx";
            string outputPath = Application.StartupPath + "\\" + txbmuqno.Text + " " + txbzBorcalan.Text + ".docx";

            // Word sənədini yaradın
            al.CreateWordDocument(templatePath, outputPath, replacements);
        }
        private void ReplaceWordStub(string stubToReplace, string text, Microsoft.Office.Interop.Word.Document WordDocument)
        {
            var range = WordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }
        private string yaziyaCevir(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için            
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";

            string[] birler = { "", " bir ", " iki ", " üç ", " dörd ", " beş ", " altı ", " yeddi ", " səkkiz ", " doqquz " };
            string[] onlar = { "", " on ", " iyirmi ", " otuz ", " qırx ", " əlli ", " altmış ", " yetmiş ", " səksən ", " doxsan " };
            string[] binler = { "katrilyon", "trilyon", " milyard ", " milyon ", " min ", "" }; //KATRİLYON'un önüne ekleme yapılarak artırabilir.

            int grupSayisi = 6; //sayıdaki 3'lü grup sayısı. katrilyon içi 6. (1.234,00 daki grup sayısı 2'dir.)
            //KATRİLYON'un başına ekleyeceğiniz her değer için grup sayısını artırınız.

            lira = lira.PadLeft(grupSayisi * 3, '0'); //sayının soluna '0' eklenerek sayı 'grup sayısı x 3' basakmaklı yapılıyor.            

            string grupDegeri;

            for (int i = 0; i < grupSayisi * 3; i += 3) //sayı 3'erli gruplar halinde ele alınıyor.
            {
                grupDegeri = "";

                if (lira.Substring(i, 1) != "0")
                    grupDegeri += birler[Convert.ToInt32(lira.Substring(i, 1))] + " yüz "; //yüzler                

                if (grupDegeri == " biryüz ") //biryüz düzeltiliyor.
                    grupDegeri = " yüz ";

                grupDegeri += onlar[Convert.ToInt32(lira.Substring(i + 1, 1))]; //onlar

                grupDegeri += birler[Convert.ToInt32(lira.Substring(i + 2, 1))]; //birler                

                if (grupDegeri != "") //binler
                    grupDegeri += binler[i / 3];

                if (grupDegeri == " birmin ") //birbin düzeltiliyor.
                    grupDegeri = " min ";

                yazi += grupDegeri;
            }

            if (yazi != "")
                yazi += " manat ";

            int yaziUzunlugu = yazi.Length;

            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];

            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];

            if (yazi.Length > yaziUzunlugu)
                yazi += " qəpik.";
            else
                yazi += "sıfır qəpik.";

            return yazi;

        }
        private string yaziyaCevirqepiksiz(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için            
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";

            string[] birler = { "", " bir ", " iki ", " üç ", " dörd ", " beş ", " altı ", " yeddi ", " səkkiz ", " doqquz " };
            string[] onlar = { "", " on ", " iyirmi ", " otuz ", " qırx ", " əlli ", " altmış ", " yetmiş ", " səksən ", " doxsan " };
            string[] binler = { "katrilyon", "trilyon", " milyard ", " milyon ", " min ", "" }; //KATRİLYON'un önüne ekleme yapılarak artırabilir.

            int grupSayisi = 6; //sayıdaki 3'lü grup sayısı. katrilyon içi 6. (1.234,00 daki grup sayısı 2'dir.)
            //KATRİLYON'un başına ekleyeceğiniz her değer için grup sayısını artırınız.

            lira = lira.PadLeft(grupSayisi * 3, '0'); //sayının soluna '0' eklenerek sayı 'grup sayısı x 3' basakmaklı yapılıyor.            

            string grupDegeri;

            for (int i = 0; i < grupSayisi * 3; i += 3) //sayı 3'erli gruplar halinde ele alınıyor.
            {
                grupDegeri = "";

                if (lira.Substring(i, 1) != "0")
                    grupDegeri += birler[Convert.ToInt32(lira.Substring(i, 1))] + " yüz "; //yüzler                

                if (grupDegeri == " biryüz ") //biryüz düzeltiliyor.
                    grupDegeri = " yüz ";

                grupDegeri += onlar[Convert.ToInt32(lira.Substring(i + 1, 1))]; //onlar

                grupDegeri += birler[Convert.ToInt32(lira.Substring(i + 2, 1))]; //birler                

                if (grupDegeri != "") //binler
                    grupDegeri += binler[i / 3];

                if (grupDegeri == " birmin ") //birbin düzeltiliyor.
                    grupDegeri = " min ";

                yazi += grupDegeri;
            }

            if (yazi != "")
                yazi += "";

            int yaziUzunlugu = yazi.Length;

            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];

            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];

            if (yazi.Length > yaziUzunlugu)
                yazi += "";
            else
                yazi += "";

            return yazi;

        }
        private string IlkHarfleriBuyut(string metin)
        {
            System.Globalization.CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Globalization.TextInfo textInfo = cultureInfo.TextInfo;
            return textInfo.ToTitleCase(metin);
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Bir nəfərin zəmanəti")
            {
                groupBox2.Visible = true;
                textBox1.Text = txbzmuqtar.Text;
                teminatsay = "təminatlarından biri";
            }
            else
            { groupBox2.Visible = false; }
            teminatsay = "təminatı";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            kataloqgetir_zamin_ayliq();
        }

        private void button1_Click(object sender, EventArgs e)
        {
                //muracietsayartimsiz();
                //meknoal();
                muqavilNo = txbmuqno.Text;
                word_at();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            doldur();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                groupBox2.Visible = true;
            }
            else if (checkBox1.Checked == false)
            {
                groupBox2.Visible = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                groupBox2.Visible = true;
                groupBox3.Visible = true;
                checkBox2.Checked = true;
            }
            else if (checkBox2.Checked == false)
            {
                groupBox3.Visible = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                groupBox4.Visible = true;
            }
            else if (checkBox3.Checked == false)
            {
                groupBox4.Visible = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // TextBox-dan gələn tarixi daxil edin
            string result = Aletler.TarixiSozeCevir(txt_uzad_tar.Text);
            textBox22.Text = result;
        }
    }
}
