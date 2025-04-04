using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Oracle.DataAccess.Client;
using System.IO;
using System.Globalization;
using System.Threading;
using Oracle.ManagedDataAccess;


namespace BMI
{
    public partial class Login : Form
    {
        public Form1 frmana1;
        public Onlayn frmana2;
        public Login()
        {
            InitializeComponent();
        }
        XmlDocument xdoc = new XmlDocument();

        XmlNodeList nodes;

        string adyoxla = "";
        string datasorc = "";
        string userid = "";
        string paswd = "";

        byte status = 0;
        public void BMIxmlbaglanti()
        {
            if (!System.IO.File.Exists(Application.StartupPath + "\\BMIxmlconnection.xml"))
            {
                MessageBox.Show("XML HEYİFKİ tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    xdoc.Load(Application.StartupPath + "\\BMIxmlconnection.xml");
                    nodes = xdoc.SelectNodes("/BMISERVER/OracleServer");

                    foreach (XmlElement element in nodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            datasorc = element.SelectSingleNode("DataSource").InnerText;
                            userid = element.SelectSingleNode("UserID").InnerText;
                            paswd = element.SelectSingleNode("Password").InnerText;
                            adyoxla = element.SelectSingleNode("UserName").InnerText;

                            OracleConnection XmlOcon = new OracleConnection("DATA SOURCE=" + datasorc + ";USER ID=" + userid + ";Password=" + paswd + "");
                            XmlOcon.Open();
                            OracleCommand XMLOrcom = new OracleCommand("select n.name from odb.nameoi n where n.code not in (0, 2, 3, 5, 6, 9, 10, 11, 13, 14, 17, 19, 23, 24, 26, 27, 29, 31, 41) order by n.name asc", XmlOcon);
                            OracleDataReader XMLOrdr = XMLOrcom.ExecuteReader();
                            while (XMLOrdr.Read())
                            {
                                comboBox1.Items.Add(XMLOrdr["name"]).ToString();
                            }
                            XMLOrdr.Close();
                            XmlOcon.Close();

                            if (adyoxla == "1")
                            {
                                comboBox1.SelectedIndex = 0;
                            }
                            else
                            {
                                comboBox1.SelectedItem = adyoxla;
                                checkBox1.Checked = true;
                            }
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public string xmlfilepatch = string.Empty;

        public void BMIxmlupdate()
        {

            if (!System.IO.File.Exists(Application.StartupPath + "\\BMIxmlconnection.xml"))
            {
                MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    xdoc.Load(Application.StartupPath + "\\BMIxmlconnection.xml");
                    nodes = xdoc.SelectNodes("/BMISERVER/OracleServer");
                    foreach (XmlElement element in nodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            element.SelectSingleNode("UserName").InnerText = cmbAdi.Text;
                            xdoc.Save("BMIxmlconnection.xml");
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            status = 0;
            BMIxmlbaglanti();
            ChangeKeyboardLangENG();
                //frmana1.orabag.cmbName = cmbAdi;
                //frmana1.orabag.ckbSaxla = chkAdisaxla;
                //frmana1.orabag.Userlogin();
                //frmana1.orabag.ChangeKeyboardENG();
            
        }
        public void ChangeKeyboardLangENG()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage); InputLanguage.CurrentInputLanguage = l;
        }
        public void metfrmGtamad(ComboBox tmad, Label cavab)
        {
            try
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=" + datasorc + ";USER ID=" + comboBox1.Text + ";Password=" + textBox1.Text + "");
                Orcon.Open();
                OracleCommand Orcom = new OracleCommand("select n.f_i_o  from odb.nameoi n where n.name='" + tmad.Text + "'", Orcon);
                OracleDataReader Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    cavab.Text = Ordr["f_i_o"].ToString();
                }
                Ordr.Close();
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    OracleConnection GirshOcon = new OracleConnection("DATA SOURCE=" + datasorc + ";USER ID=" + comboBox1.Text + ";Password=" + textBox1.Text + "");
            //    GirshOcon.Open();
            //    if (GirshOcon.State == ConnectionState.Open)
            //    {
            //        status = 1;
            //        if (checkBox1.Checked == true)
            //        {
            //            BMIxmlupdate();
            //        }
                    
            //        metfrmGtamad(cmbAdi, frmana1.label1);
                    
            //        GirshOcon.Close();
                    
            //        this.Close();
                    
            //    }
           // }
            //catch (Exception)
            //{
            //    status = 0;
            //    MessageBox.Show("Şifrə doğru deyil", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            ////Form1 frm1 = new Form1();
            //frmana1.orabag.UserID = cmbAdi.Text;
            //frmana1.orabag.Password = txtPassword.Text;
            //frmana1.orabag.Tesdigal();

            //if (frmana1.orabag.tesdig == true)
            //{   
                
            //    frmana1.orabag.Currentuser = cmbAdi.Text;
            //    frmana1.orabag.Xmlyenile();
            //    frmana1.orabag.Adtap();
            //    frmana1.tlsplblAdi.Text = frmana1.orabag.Currentuser;
            //    frmana1.lblicracikodu.Text = frmana1.orabag.icracikodu;
            //    frmana1.frmgirish1.Close();
            //}

            ////Istifadeciler.istifadecigirisi(txtistifadeci, txtPassword);
            ////if (Istifadeciler.netice==true)
            ////{
            ////    Form1 frm1 = new Form1();
            ////    frm1.ShowDialog();
            ////    this.Hide();
            ////}
            ////else if (Istifadeciler.netice == false)
            ////{
            ////    MessageBox.Show("sifre yalnisdir", "Diqqet", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            ////}

        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (status == 0)
            {
                Application.Exit();
            }
        }

        public Form1 frmana { get; set; }
    }
}
