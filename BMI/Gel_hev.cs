using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace BMI
{
    public partial class Gel_hev : Form
    {
        public Gel_hev()
        {
            InitializeComponent();
        }
        SqlConnection baglan = new SqlConnection("Data Source=.\\sqlexpress;Initial Catalog=BMI;Integrated Security=True");
        OracleConnection OrConnect = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass;");
        public string icraci_kod { get; set; }
        public Gelen_hevale gedhev;
        public string Xaricmektbtarix = string.Empty;
        int kecencavab = 0;
        //private void hevnom_al()
        //{

        //    OracleCommand cmd = new OracleCommand();
        //    OracleDataReader SR = null;
        //    cmd.Connection = OrConnect;
        //    cmd.CommandText = "Select HEV_NOM from odb.GEDEN_HEVALE_SAMIR order by HEV_NOM desc";
        //    OrConnect.Open();
        //    SR = cmd.ExecuteReader();
        //    if (SR.Read())
        //    {
        //        string heval = SR.GetValue(0).ToString();
        //        string hevsecal = heval.Substring(5);
        //        int hevno = Convert.ToInt32(hevsecal.ToString());
        //        hevno = hevno + 1;
        //        string cariil = DateTime.Now.Date.Year.ToString().Substring(2);
        //        txbgedHNo.Text = cariil + "-T-" + hevno.ToString();
        //        Pul_Kocurmesi pkh = new Pul_Kocurmesi();
        //        pkh.txbhevale.Text = cariil + "-T-" + hevno.ToString();
        //    }
        //    OrConnect.Close();
        //}


        // gelen hevale ve geden hevale duzelisler islemir ONLARA BAX//////////////////////////

        private void gelenhevaleyazdir()
        {

            if (lblSgldml.Text == "insert")
            {
                //if ( txbgedSAA.Text.Length > 0)
                //{
                    DialogResult dr = MessageBox.Show("Həvalə qeydiyyata alınsın ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        //hvadi = cmbGonderyer.Text.TrimStart();
                        Xaricmektbtarix = dtpTarix.Text;
                        string tarixIl = DateTime.Now.Date.Year.ToString();
                        //TO_DATE('" + dateTimePicker1 + "', 'dd-MM-yyyy')
                        //,'" + txbgedmense.Text + "','" + textBox1.Text + "','" + txbgedtesnifat.Text + "','" + cmbgedGonderen.Text + "','" + txbgedalanb.Text + "'
                        OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                        Ocon.Open();
                        OracleCommand Ocom = new OracleCommand("INSERT INTO ODB.gelen_hevale(hev_nom,hes_nom,saa,tip_res,mebleg,gel_olke,val_tip,tarix,men_olke,hev_tip,gon_tip,al_bank,icra) values ('" + txbgedHNo.Text + "','" + txbgedHesNo.Text + "','" + txbgedSAA.Text + "','" + cmbgedveten.Text + "','" + txbgedmebleg.Text + "','" + txbgedtesnifat.Text + "','" + cmbgedvalyuta.Text + "',TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'),'" + txbgedmense.Text + "','" + cmbgedMhes.Text + "','" + cmbgedGonderen.Text + "','" + txbgedalanb.Text + "','" + icraci_kod + "')", Ocon);
                        //HEV_NOM,HES_NOM,SAA,Mebleg,Val_tip,tarix,hev_tip,men_olke,olke,gon_tip,al_bank,icra
                        //hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank
                        Ocom.ExecuteNonQuery();
                        Ocon.Close();
                        MessageBox.Show("Həvalə qeydiyyata alındı...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        //gedhev.btnYenile.PerformClick();
                        this.Close();
                    }
                //}
                //else
                //{
                //    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //}

            }
            else if (lblSgldml.Text == "update")
            {
                if ( txbgedSAA.Text.Length > 0)
                {
                    DialogResult dr = MessageBox.Show("Məktuba düzəliş edilsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                        Ocon.Open();
                        OracleCommand Ocom = new OracleCommand("Update odb.gelen_hevale x set x.tarix = TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), x.HES_NOM='" + txbgedHesNo.Text + "', x.SAA='" + txbgedSAA.Text + "', x.Mebleg='" + txbgedmebleg.Text + "',x.Val_tip='" + cmbgedvalyuta.Text + "',x.hev_tip='" + txbgedtesnifat.Text + "',x.men_olke='" + txbgedmense.Text + "',x.gon_tip='" + cmbgedGonderen.Text + "',x.al_bank='" + txbgedalanb.Text + "' where x.qey_nom='" + txbgedHNo.Text + "'", Ocon);

                        kecencavab = Ocom.ExecuteNonQuery();

                        if (kecencavab > 0)
                        {
                            MessageBox.Show("Həvaləyə Düzəliş olundu...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Həvaləyə Düzəliş olunmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        Ocon.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            ///////////////////////////
            try
            {
                
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            


        }

        private void hevnom_al()
        {
            string heval = "";
            string tarixIl = DateTime.Now.Date.Year.ToString();
            string boyukil = "";
            string sonuncuil = "";
            OracleCommand cmd = new OracleCommand();
            OracleCommand cmd1 = new OracleCommand();
            OracleDataReader SR = null;
            OracleDataReader SR1 = null;
            cmd.Connection = OrConnect;
            cmd1.Connection = OrConnect;
            cmd.CommandText = "Select max(to_number(substr(hev_nom,6,4))) from odb.GELEN_HEVALe where to_char( tarix,'yyyy')='" + tarixIl + "' ";
            cmd1.CommandText = "Select max(to_char( tarix,'yyyy')) from odb.GELEN_HEVALe ";
            OrConnect.Open();
            SR = cmd.ExecuteReader();
            SR1 = cmd1.ExecuteReader();
            //max(to_number(substr(hev_nom,7,4)))
            if (SR.Read())
            {
                boyukil = SR.GetValue(0).ToString();

                if (boyukil == "")
                {

                }
                else
                {
                    int boyukilsay = Convert.ToInt16(boyukil.ToString());
                }
            }
            if (SR1.Read())
            {
                sonuncuil = SR1.GetValue(0).ToString();
            }

            if (Convert.ToInt16(sonuncuil) < Convert.ToInt16(tarixIl))
            {
                string cariil="";
                txbgedHNo.Text =  cariil = DateTime.Now.Date.Year.ToString().Substring(2) + "-G-1";

            }
            else
            {
                heval = boyukil.ToString();
                int hevno = Convert.ToInt32(heval.ToString());
                hevno = hevno + 1;
                string cariil = DateTime.Now.Date.Year.ToString().Substring(2);
                txbgedHNo.Text = cariil + "-G-" + hevno.ToString();
                Pul_Kocurmesi pkh = new Pul_Kocurmesi();
                txbgedHNo.Text = cariil + "-G-" + hevno.ToString();

            }
            //string hevsecal = heval.Substring(5);


            OrConnect.Close();



        }

        private void hesadgetir()
    {
        try
        {
            if (txbgedHesNo.Text == "")
            {
                txbgedSAA.Text = "";
            }
            else
            {


                OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                con.Open();
                OracleCommand komut = new OracleCommand();
                komut.Connection = con;
                komut.CommandText = "select Licsch,name_licsch from odb.licsch where Licsch='" + txbgedHesNo.Text + "'";
                OracleDataReader dr = komut.ExecuteReader();


                while (dr.Read())
                {

                    txbgedSAA.Text = dr["name_licsch"].ToString();

                }
                con.Close();
            }
        }
        catch (Exception)
        {
            
            
        }
        


        
    }

        private void satirsayisinitap()
        {
            int satirsayi = -1;
            string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection connection = new OracleConnection(connectrionString);
            OracleCommand orcmd = new OracleCommand("select count(*) from odb.geden_hevale_samir", connection);
            connection.Open();
            satirsayi = Convert.ToInt32(orcmd.ExecuteScalar());
            connection.Close();
            string cariil = DateTime.Now.Date.Year.ToString();
            int birartir = satirsayi + 1;
            string nom = cariil + "-G-" + birartir.ToString();
            txbgedHNo.Text = nom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            gelenhevaleyazdir();
            Gelen_hevale glhv = new Gelen_hevale();
            glhv.listeleLedenhevale();
        }

        private void Gel_hev_Load(object sender, EventArgs e)
        {
            if (lblSgldml.Text == "insert")
            {
                hevnom_al();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            hesadgetir();
        }

        private void txbgedHesNo_TextChanged(object sender, EventArgs e)
        {
            hesadgetir();
        }
    }
}
