using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using BMI.Muhasibat;

namespace BMI
{
    public partial class DYP : Form
    {
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        public DYP()
        {
            InitializeComponent();
        }
        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;
        public string mek_no { get; set; }
        private void DYP_Load(object sender, EventArgs e)
        {
            string gun, ay, il;
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            txbmektar.Text = gun + "-" + ay + "-" + il;
        }
        private void wordeat()
        {
            var replacements = new Dictionary<string, string>
           {
                { "{mekNo}", mek_no},
                {"{mektarixi}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy"))},
                {"{borcalan}", txbmusad.Text},
                { "{muqtar}", Aletler.TarixiSozeCevir(txbmuqtar.Text)},
                { "{muqNo}", txbmuqNo.Text},
                { "{avtoNo}", txbavtNo.Text},
                { "{marka}", txbMarka.Text},
                { "{avtoil}", txbİli.Text},
                { "{muh}", txbMuherrik.Text},
                { "{ban}", txbBan.Text},
                { "{reng}", txbReng.Text},
           };
            string dypmektub = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Wordler", "DYP arayış girovda çıxma1.doc");
            string outputPath = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis wordler", txbavtNo.Text + ".doc");// + "\\" + txbmuqno.Text + " " + txbzBorcalan.Text + ".docx";
            // Word sənədini yaradın
            aletler.CreateWordDocument(dypmektub, outputPath, replacements);
        }
        string tarix_soz;
        string tarix_sozilkintarix;
        DateTime dt = DateTime.Now;
        private void button1_Click(object sender, EventArgs e)
        {
            
        }
        string musterimekadi = "";
        private void button2_Click(object sender, EventArgs e)
        {
            musterimekadi = txbmusad.Text;
            muracietsayartimsiz();
            meknoal();
            wordeat();
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
        private void meknoal()
        {
            int test = 45;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();
            string qeyd= "avto gir çıx "+ txbmusad.Text;
            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('DYP', TO_DATE('" + dateavtoMuqtarix.Text + "','dd-MM-yyyy'), '"+qeyd+"', '" + icraci_kod + "',  " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();
        }
    }
}
