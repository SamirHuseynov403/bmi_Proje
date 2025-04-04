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
using System.Threading;
using System.Xml;
using System.IO;
using BMI.Muhasibat;

namespace BMI
{
    public partial class Logiin : Form
    {
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public Logiin()
        {
            InitializeComponent();
            //this.Icon = new System.Drawing.Icon("fon1.ico");
        }
        public Form1 frmana;
        cl_yanasmalar cl = new cl_yanasmalar();
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;
        public string Currentuser = string.Empty;
        public string icracikodu = string.Empty;

        public string ic_kod = string.Empty;

        XmlDocument xdoc = new XmlDocument();

        XmlNodeList nodes;
        public Oraclebaglanti orabag;
        string adyoxla = "";
        string datasorc = "";
        string userid = "";
        string paswd = "";
        public void BMIxmlbaglanti()
        {
            string xmlFilePath = "";
            try
            {
                string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                xmlFilePath = Path.Combine(appDirectory, "Fayllar", "BMIxmlconnection.xml");

                // Əgər fayl mövcud deyilsə, alternativ yolu yoxla
                if (!File.Exists(xmlFilePath))
                {
                    throw new FileNotFoundException("Fayl tapılmadı, alternativ yol seçilir.");
                }
            }
            catch (Exception)
            {
                xmlFilePath = Path.Combine(qovluqyolu, "Fayllar", "BMIxmlconnection.xml");
            }


            //string xmlFilePath = Path.Combine(qovluqyolu, "Fayllar", "BMIxmlconnection.xml");
            // Setup edende bunu bagla yuxaridaki 2 deneni ac

            if (!System.IO.File.Exists(xmlFilePath))
            {
                MessageBox.Show("XML file tapılmadı.......Yoxlama", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    // XML dosyasını yükleyin
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(xmlFilePath);

                    // XPath ile belirli düğümleri seçin
                    XmlNodeList nodes = xdoc.SelectNodes("/BMISERVER/OracleServer");

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

        //public void BMIxmlupdate()
        //{

        //    if (!System.IO.File.Exists(System.IO.Path.Combine(qovluqyolu, "Fayllar", "BMIxmlconnection.xml")))
        //    {
        //        MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        //    }
        //    else
        //    {
        //        try
        //        {
        //            xdoc.Load(Application.StartupPath + "\\BMIxmlconnection.xml");
        //            nodes = xdoc.SelectNodes("/BMISERVER/OracleServer");
        //            foreach (XmlElement element in nodes)
        //            {
        //                if (element.Attributes["ID"].Value == "1")
        //                {
        //                    element.SelectSingleNode("UserName").InnerText = comboBox1.Text;
        //                    xdoc.Save("BMIxmlconnection.xml");
        //                    break;
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //}
        public void BMIxmlupdate()
        {
            string xmlFilePath = Path.Combine(qovluqyolu, "Fayllar", "BMIxmlconnection.xml");

            if (!System.IO.File.Exists(xmlFilePath))
            {
                MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    xdoc.Load(xmlFilePath); // Burada düzgün fayl yolunu istifadə edirik

                    nodes = xdoc.SelectNodes("/BMISERVER/OracleServer");
                    foreach (XmlElement element in nodes)
                    {
                        if (element.Attributes["ID"].Value == "1")
                        {
                            element.SelectSingleNode("UserName").InnerText = comboBox1.Text;
                            xdoc.Save(xmlFilePath); // Faylı eyni yolda saxlayırıq
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
            Form1 frm1 = (Form1)Application.OpenForms["Form1"];
            XaricMektub xrcmek = new XaricMektub();
            try
            {
                OracleConnection GirshOcon = new OracleConnection("DATA SOURCE=" + datasorc + ";USER ID=" + comboBox1.Text + ";Password=" + textBox1.Text + "");
                GirshOcon.Open();
                if (GirshOcon.State == ConnectionState.Open)
                {
                    status = 1;
                    if (checkBox1.Checked == true)
                    {
                        BMIxmlupdate();
                        //MessageBox.Show("yoxlama");
                    }
                    metfrmGtamad(comboBox1, frmana.lblTamad);

                    GirshOcon.Close();

                    this.Close();

                    

                    OracleConnection GirshOcon1 = new OracleConnection("DATA SOURCE=" + datasorc + ";USER ID=" + comboBox1.Text + ";Password=" + textBox1.Text + "");
                    GirshOcon1.Open();
                    Orcom = new OracleCommand("select n.f_i_o, n.code from odb.nameoi n where n.name='" + comboBox1.Text + "'", GirshOcon1);
                    Ordr = Orcom.ExecuteReader();
                    while (Ordr.Read())
                    {
                        Currentuser = Ordr["f_i_o"].ToString();
                        icracikodu = Ordr["code"].ToString();
                        
                        
                    }
                    Ordr.Close();
                    GirshOcon1.Close();
                    SetMenuPermissions(Convert.ToInt32(icracikodu), frmana.menuStrip1);
                    cl.icraci =Convert.ToInt16(icracikodu);
                    frmana.icraci_kod= icracikodu;

                }
            }
            catch (Exception)
            {
                status = 0;
                MessageBox.Show("Şifrə doğru deyil", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            frm1.tlsplblAdi.Text = this.icracikodu;
            frm1.icraci_kod = this.icracikodu;
        }
        public int status { get; set; }
        private void Logiin_Load(object sender, EventArgs e)
        {
            status = 0;
            BMIxmlbaglanti();
            ChangeKeyboardLangENG();
            string filePath = "LastUser.txt";
            if (File.Exists(filePath)) // Fayl varsa
            {
                string lastUser = File.ReadAllText(filePath); // Fayldan istifadəçi adını oxu
                comboBox1.Text = lastUser; // ComboBox-da göstər
                //checkBox1.Checked = true; // Checkbox-u işarələ
            }
        }
        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }
        private void Logiin_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (status == 0)
            {
                Application.Exit();
            }
        }
        public void SetMenuPermissions(int userCode, MenuStrip menuStrip)
        {
            try
            {
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();

                    // İstifadəçi məlumatını gətirən sorğu
                    string userQuery = "SELECT derece FROM bmi_istifadeciler WHERE i_kod = :userCode";

                    // İstifadəçi dərəcəsini əldə et
                    int userDegree = 0;
                    using (OracleCommand userCommand = new OracleCommand(userQuery, connection))
                    {
                        userCommand.Parameters.Add(new OracleParameter("userCode", userCode));
                        object result = userCommand.ExecuteScalar();
                        if (result != null)
                        {
                            userDegree = Convert.ToInt32(result); // İstifadəçinin dərəcəsi
                        }
                        else
                        {
                            MessageBox.Show("İstifadəçi tapılmadı.");
                            return;
                        }
                    }

                    // Elementlər və dərəcələri gətirən sorğu
                    string elementQuery = @"
                    SELECT element_adi,
                    case when e.id=x.form_id and x.icaze='1' then 1 else e.derece end derece
                    FROM bmi_elementler e,bmi_xususi_icazeler x where e.id=x.form_id(+)";

                    // Elementləri əldə et
                    DataTable elementTable = new DataTable();
                    using (OracleCommand elementCommand = new OracleCommand(elementQuery, connection))
                    using (OracleDataAdapter adapter = new OracleDataAdapter(elementCommand))
                    {
                        adapter.Fill(elementTable); // Elementlər cədvələ yüklənir
                    }

                    // Menü elementlərini yoxla
                    foreach (ToolStripMenuItem menuItem in menuStrip.Items)
                    {
                        SetMenuItemPermissions(menuItem, elementTable, userDegree);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message);
            }
        }
        private void SetMenuItemPermissions(ToolStripMenuItem menuItem, DataTable elementTable, int userDegree)
        {
            // Cari menü elementinin `element_adi` ilə uyğunluğu yoxlanılır
            DataRow[] matchingRows = elementTable.Select($"element_adi = '{menuItem.Text}'");
            if (matchingRows.Length > 0)
            {
                int elementDegree = Convert.ToInt32(matchingRows[0]["derece"]);
                menuItem.Enabled = userDegree >= elementDegree; // Dərəcə uyğun gəlirsə, Enabled true
            }
            else
            {
                menuItem.Enabled = false; // Uyğun element tapılmadısa, false
            }

            // Alt menyuları da yoxla (əgər varsa)
            foreach (ToolStripItem subItem in menuItem.DropDownItems)
            {
                if (subItem is ToolStripMenuItem subMenuItem)
                {
                    SetMenuItemPermissions(subMenuItem, elementTable, userDegree);
                }
            }
        }
        private void chkRemember_CheckedChanged(object sender, EventArgs e)
        {
            
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            string filePath = "LastUser.txt";

            if (checkBox1.Checked)
            {
                File.WriteAllText(filePath, comboBox1.Text); // ComboBox-da seçilmiş istifadəçi adını fayla yaz
            }
            else
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath); // Faylı sil
                }
            }
        }
    }
}
