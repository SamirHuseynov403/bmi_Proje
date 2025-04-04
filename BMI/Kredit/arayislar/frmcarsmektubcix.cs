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
using Word = Microsoft.Office.Interop.Word;
using System.Data.OleDb;


namespace BMI
{
    public partial class frmcarsmektubcix : Form
    {
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI_\BMI\bin\Debug\AtlasCars.accdb");
        string qovluqyolu = Aletler.Layiheanaqovluq();
        Aletler aletler = new Aletler();
        public frmcarsmektubcix()
        {
            InitializeComponent();
        }
        private readonly string saveyolu = @"\\fs\KRED_SOB\1-КРЕДИТ-2023\Atlas Cars gir cix";
        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;
        public string mek_no { get; set; }
        string tarix_soz;
        string tarix_sozilkintarix;
        DateTime dt = DateTime.Now;
        
        private void wordeat_girovdan_cix()
        {
            var replacements = new Dictionary<string, string>
           {
            { "{mekNo}", mek_no },
            { "{mektarixi}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy"))},
            { "{muqtar}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy"))},
            { "{avtoNo}", txbavtNo.Text},
            { "{avtoil}", txbİli.Text},
            { "{muh}", txbMuherrik.Text },
            { "{ban}", txbBan.Text  },
            { "{reng}", txbReng.Text },
           };
            // Şablon və çıxış yollarını təyin edin
            string templatePath = System.IO.Path.Combine(qovluqyolu, "Fayllar","Kredit","Wordler", "carsgirovcix.doc");
            string outputPath = @"\\fs\KRED_SOB\1-КРЕДИТ-2023\Atlas Cars gir cix\" + txbavtNo.Text + ".doc";// + "\\" + txbmuqno.Text + " " + txbzBorcalan.Text + ".docx";

            // Word sənədini yaradın
            aletler.CreateWordDocument(templatePath, outputPath, replacements);
        }
        private void wordeat_texp_deyisme()
        {
            // Əvəz etmələr üçün məlumatları toplayın
            var replacements = new Dictionary<string, string>
    {
        { "{mekNo}", mek_no },
        { "{mektarixi}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy")) },
        { "{muqtar}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy")) },
        { "{avtoNo}", txbavtNo.Text },
        { "{avtoil}", txbİli.Text },
        { "{muh}", txbMuherrik.Text },
        { "{ban}", txbBan.Text },
        { "{reng}", txbReng.Text },
        { "{texpNo}", textBox3.Text }
    };

            // Şablon və çıxış yollarını təyin edin
            string templatePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Wordler", "Cars Şehadetneme deyismesi.doc");
            string outputPath = System.IO.Path.Combine(@"\\fs\KRED_SOB\1-КРЕДИТ-2023\Atlas Cars texpasport deyisme\" + txbavtNo.Text + ".doc");

            // Word sənədini yaradın
            aletler.CreateWordDocument(templatePath, outputPath, replacements);
        }
        
        private void meknoal()
        {
            string gun, ay, il;
            dateavtoMuqtarix.Text= DateTime.Now.Date.ToString("dd-MM-yyyy");
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            textBox2.Text = gun + "-" + ay + "-" + il;
            int test = 45;
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();
            string mezm = "";
            if (comboBox1.Text== "Girovdan çıxma")
            {
                mezm = " avto gir azad Cars";
            }
            else if (comboBox1.Text == "Texpasport dəyişmə")
            {
                mezm = " avto texpasport dəyişmə Cars";
            }
            string tammezmun = txbavtNo.Text + mezm;
            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('DYP', TO_DATE('" + dateavtoMuqtarix.Text + "','dd-MM-yyyy'), '"+ tammezmun + "'   , '" + icraci_kod + "',  " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            muracietsayartimsiz();
            // meknoal();
            aletler.GedenMektubaElaveEt("DYP",comboBox1.Text, txbavtNo.Text, icraci_kod);
            if (comboBox1.Text == "Girovdan çıxma")
            {
                wordeat_girovdan_cix();
                bazaya_at();
            }
            else if (comboBox1.Text == "Texpasport dəyişmə")
            {
                wordeat_texp_deyisme();
            }
       
            temizle();
        }
        private void temizle()
        {
            txbavtNo.Text = "";
            txbBan.Text = "";
            txbMuherrik.Text = "";
            txbReng.Text = "";
            txbİli.Text = "";
            textBox3.Text = "";
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

        private void frmcarsmektubcix_Load(object sender, EventArgs e)
        {

            string gun, ay, il;
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            textBox1.Text = "«"+gun+"»" + "-" + ay + "-" + il;
        }
        private void bazaya_at()
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("insert into Cars_mektublar (avto_no,mek_no,muherrik,ban,tarix,ir_mek_no) values('" + txbavtNo.Text + "','" + mek_no + "','" + txbMuherrik.Text + "','" + txbBan.Text + "','" + dateavtoMuqtarix.Text + "','" + txtirmekno.Text + "')", baglanti);
            komut.ExecuteNonQuery();
            baglanti.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text== "Girovdan çıxma")
            {
                textBox3.Visible = false;
                label3.Visible = false;
            }
            else if (comboBox1.Text == "Texpasport dəyişmə")
            {
                textBox3.Visible = true;
                label3.Visible = true;
            }
        }
    }
}
