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
    public partial class hevalegeden : Form
    {
        public hevalegeden()
        {
            InitializeComponent();
        }
        public string Xaricmektbtarix = string.Empty;
        public geden_hevale gedhev;
        public string icraci_kod { get; set; }
        public Mektub mktb;
        

        public string hvadi = string.Empty;
        //public string Xmgonderilenyer = string.Empty;
        //public string Xaricmektbtarix = string.Empty;
        //public string Xmgisamezmun = string.Empty;
        //public string Xmmektubmetn = string.Empty;

        public string gh_hevnom = string.Empty;
        public string gh_hesnom = string.Empty;
        public string gh_adi = string.Empty;
        public string gh_mebleg = string.Empty;
        public string gh_valtip = string.Empty;
        public string gh_tarix = string.Empty;
        public string gh_hevtip = string.Empty;
        public string gh_menolke = string.Empty;
        public string gh_olke = string.Empty;
        public string gh_gontip = string.Empty;
        public string gh_albank = string.Empty;
        

        OracleConnection OrConnect = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass;");

        private void gedhevyazdir()
        {
            int kecencavab = 0;
            if (lblSgldml.Text == "insert")
            {
                if (txbgedSAA.Text.Length > 0 )
                {
                    DialogResult dr = MessageBox.Show("Məktubun qeydiyyata alınsın ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        gh_hevnom = txbgedHNo.Text.TrimStart();
                        gh_mebleg = txbgedmebleg.Text.TrimStart();
                        gh_adi = txbgedSAA.Text.TrimStart();
                       // testet();
                        hvadi = cmbGonderyer.Text.TrimStart();
                        Xaricmektbtarix = dtpTarix.Text;
                        string tarixIl = DateTime.Now.Date.Year.ToString();
                        //TO_DATE('" + dateTimePicker1 + "', 'dd-MM-yyyy')
                        //,'" + txbgedmense.Text + "','" + textBox1.Text + "','" + txbgedtesnifat.Text + "','" + cmbgedGonderen.Text + "','" + txbgedalanb.Text + "'
                        OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                        Ocon.Open();
                        OracleCommand Ocom = new OracleCommand("INSERT INTO ODB.geden_hevale(HEV_NOM,HES_NOM,SAA,Mebleg,TIP_RES,Val_tip,tarix,hev_tip,men_olke,olke,gon_tip,al_bank,icra) values ('" + txbgedHNo.Text + "','" + txbgedHesNo.Text + "','" + txbgedSAA.Text + "','" + txbgedmebleg.Text + "','" + cmbgedveten.Text + "','" + cmbgedvalyuta.Text + "',TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'),'" + txbgedtesnifat.Text + "','" + txbgedmense.Text + "','" + textBox1.Text + "','" + cmbgedGonderen.Text + "','" + txbgedalanb.Text + "','" + icraci_kod + "')", Ocon);
                        //hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank
                        Ocom.ExecuteNonQuery();
                        Ocon.Close();
                        MessageBox.Show("Gedən Həvalə qeydiyyata alındı...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        //gedhev.btnYenile.PerformClick();
                        this.Close();
                        //mktb.gh_hevnom = txbgedHNo.Text;
                        //mktb.gh_hesnom = txbgedHesNo.Text.TrimStart();
                        //mktb.gh_adi = txbgedSAA.Text.TrimStart();
                        //mktb.gh_mebleg = txbgedmebleg.Text.TrimStart();
                        //mktb.gh_valtip = cmbgedvalyuta.Text.TrimStart();
                        //mktb.gh_tarix = dtpTarix.Text;
                        //mktb.gh_hevtip = txbgedtesnifat.Text.TrimStart();
                        //mktb.gh_menolke = txbgedmense.Text.TrimStart();
                        //mktb.gh_olke = textBox1.Text.TrimStart();
                        //mktb.gh_gontip = cmbgedGonderen.Text.TrimStart();
                        //mktb.gh_albank = txbgedalanb.Text.TrimStart();





                        //gedhev.mktb.m_geden_hevale_insert();

                        ////btnTemizle.PerformClick();
                        
                    }
                }
                else
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            else if (lblSgldml.Text == "update")
            {
                if (txbgedSAA.Text.Length > 0 )
                {
                    DialogResult dr = MessageBox.Show("Məktuba düzəliş edilsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        //try
                        {
                            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                            Ocon.Open();
                            OracleCommand Ocom = new OracleCommand("Update odb.geden_hevale x set x.hev_nom='" + txbgedHNo.Text + "',x.hes_nom='" + txbgedHesNo.Text + "',x.saa='" + txbgedSAA.Text + "',x.mebleg='" + txbgedmebleg.Text + "', x.tip_res='" + cmbgedveten.Text + "',x.hev_tip='" + txbgedtesnifat.Text + "',x.men_olke='" + txbgedmense.Text + "',x.olke='" + textBox1.Text + "',x.gon_tip='" + cmbgedGonderen.Text + "',x.al_bank='" + txbgedalanb.Text + "' where x.hev_nom='" + txbgedHNo.Text + "' ", Ocon);
                            //hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank
                            kecencavab = Ocom.ExecuteNonQuery();

                            if (kecencavab > 0)
                            {
                                MessageBox.Show("Məktuba Düzəliş olundu...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Məktuba Düzəliş olunmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            Ocon.Close();
                        }
                        //catch (Exception ex)
                        {
                            //Ocon.Close();
                            //MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        
                    }
                }
                else
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            
            this.Close();
        
        }

        private void testet()
        {
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("INSERT INTO ODB.geden_hevale(HEV_NOM,SAA,Mebleg,tarix) values ('" + gh_hevnom + "','" + gh_adi + "','" + gh_mebleg + "',TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'))", Ocon);
            //hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank
            Ocom.ExecuteNonQuery();
            Ocon.Close();
            MessageBox.Show("yazildi");
            this.Close();
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

        private void button2_Click(object sender, EventArgs e)
        {

            gedhevyazdir();
            //gedhev.btnYenile.PerformClick();


            
            
            

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
                cmd.CommandText = "Select max(to_number(substr(hev_nom,6,4))) from odb.GEDEN_HEVALe where to_char( tarix,'yyyy')='" + tarixIl + "' ";
                cmd1.CommandText = "Select max(to_char( tarix,'yyyy')) from odb.GEDEN_HEVALe ";
                OrConnect.Open();
                SR = cmd.ExecuteReader();
                SR1 = cmd1.ExecuteReader();
                //max(to_number(substr(hev_nom,7,4)))
                if (SR.Read())
                {
                     boyukil = SR.GetValue(0).ToString();
                     if (boyukil=="")
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

                    if (Convert.ToInt16(sonuncuil)<Convert.ToInt16(tarixIl))
                    {
                        string cariil="";
                        txbgedHNo.Text =  cariil = DateTime.Now.Date.Year.ToString().Substring(2) + "-T-1";

                    }
                    else 
                    {
                            heval = boyukil.ToString();
                            int hevno = Convert.ToInt32(heval.ToString());
                            hevno = hevno + 1;
                            string cariil = DateTime.Now.Date.Year.ToString().Substring(2);
                            txbgedHNo.Text = cariil + "-T-" + hevno.ToString();
                            Pul_Kocurmesi pkh = new Pul_Kocurmesi();
                            txbgedHNo.Text = cariil + "-T-" + hevno.ToString();

                    }
                    //string hevsecal = heval.Substring(5);

                
                OrConnect.Close();
            
            
            
        }

        private void hevalegeden_Load(object sender, EventArgs e)
        {
            if (lblSgldml.Text=="insert")
            {
                hevnom_al();
            }
            
        }

        private void txbgedHesNo_TextChanged(object sender, EventArgs e)
        {
            hesadgetir();
        }
    }
}
