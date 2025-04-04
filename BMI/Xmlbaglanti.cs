using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Data;

namespace BMI
{
    public class Xmlbaglanti
    {
    public XmlDocument xmldoc = new XmlDocument();      
        public XmlNodeList xmlnodes;

        public string xmlfilepatch = string.Empty;

        public string ServerIP = string.Empty;
        public string DataSource = string.Empty;
        public string UserID = string.Empty;
        public string Password = string.Empty;
        public string Currentuser = string.Empty;
        public string Muvekkil = string.Empty;
        public string Vezifesi = string.Empty;
        public string HMpapkayol = string.Empty;
        public string KRDMugpapkayol = string.Empty;

       public bool tesdig;
       public string icracikodu = string.Empty;

       public OracleConnection Orcon;

       public string[] anatarix = new string[3];
       public Label esastarix = new Label();
       public Label savefiletarix = new Label();

       public Xmlbaglanti()
        {
           tesdig = false;
        }

       public void Realtarix()
       {
        esastarix.Text = DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString();
        anatarix = esastarix.Text.Split('-');
           if (anatarix[0].Length == 1)
           {
               anatarix[0] = anatarix[0].Replace(anatarix[0], "0" + anatarix[0]);
           }
           if (anatarix[1].Length == 1)
           {
               anatarix[1] = anatarix[1].Replace(anatarix[1], "0" + anatarix[1]);
           }
           esastarix.Text = anatarix[0] + "-" + anatarix[1] + "-" + anatarix[2];
           savefiletarix.Text = anatarix[0] + anatarix[1] + anatarix[2];
       }

       public void Xmloxu()
        {
            try
            {//‪C:\BMI_\BMI\Servername.xml
                xmlfilepatch = @"C:\BMI_\BMI\Servername.xml";

                if (File.Exists(xmlfilepatch) == false)
                {
                    MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    xmldoc.Load(xmlfilepatch);
                    xmlnodes = xmldoc.SelectNodes("/Server/OracleServer");

                    foreach (XmlElement element in xmlnodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            ServerIP = element.SelectSingleNode("ServerIP").InnerText;
                            DataSource = element.SelectSingleNode("DataSource").InnerText;
                            UserID = element.SelectSingleNode("UserID").InnerText;
                            Password = element.SelectSingleNode("Password").InnerText;
                            Currentuser = element.SelectSingleNode("Currentuser").InnerText;
                            Muvekkil = element.SelectSingleNode("Muvekkil").InnerText;
                            Vezifesi = element.SelectSingleNode("Vezifesi").InnerText;
                            HMpapkayol = element.SelectSingleNode("HMpapkayol").InnerText;
                            KRDMugpapkayol = element.SelectSingleNode("KRDMpapkayol").InnerText;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       public virtual void Xmlyenile()
        {
            try
            {
                xmlfilepatch = @"C:\BMI_\BMI\Servername.xml";

                if (File.Exists(xmlfilepatch) == false)
                {
                    MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else
                {
                    xmldoc.Load(xmlfilepatch);
                    xmlnodes = xmldoc.SelectNodes("/Server/OracleServer");

                    foreach (XmlElement element in xmlnodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            element.SelectSingleNode("ServerIP").InnerText = ServerIP;
                            element.SelectSingleNode("DataSource").InnerText = DataSource;
                            element.SelectSingleNode("UserID").InnerText = UserID;
                            element.SelectSingleNode("Password").InnerText = Password;
                            element.SelectSingleNode("Currentuser").InnerText = Currentuser;
                            xmldoc.Save("Servername.xml");
                            MessageBox.Show("Məlumat yeniləndi", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       public void XmlMyenile()
        {
            try
            {
                xmlfilepatch = Application.StartupPath + @"C:\BMI_\BMI\Servername.xml";

                if (File.Exists(xmlfilepatch) == false)
                {
                    MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    tesdig = false;
                }
                else
                {
                    xmldoc.Load(xmlfilepatch);
                    xmlnodes = xmldoc.SelectNodes("/Server/OracleServer");

                    foreach (XmlElement element in xmlnodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            element.SelectSingleNode("Muvekkil").InnerText = Muvekkil;
                            element.SelectSingleNode("Vezifesi").InnerText = Vezifesi;
                            xmldoc.Save("Servername.xml");
                            MessageBox.Show("Məlumat yeniləndi", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            tesdig = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tesdig = false;
            }
        }

       public virtual void Tesdigal()
        {
            try
            {
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                if (Orcon.State == ConnectionState.Open)
                {
                    tesdig = true;
                }
                Orcon.Close();
            }
            catch (Exception)
            {
                Orcon.Close();
                tesdig = false;
                MessageBox.Show("SYSDBA password və ya DataSource - doğru deyil", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }

}
