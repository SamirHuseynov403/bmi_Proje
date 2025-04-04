using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using BMI.Muhasibat;
using System.IO;

namespace BMI
{
    public partial class zaminarayis : Form
    {
        public zaminarayis()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        public OracleConnection Orcon;
        public OracleCommand Orcom;
        public string icraci_kod = string.Empty;
        public string mek_no { get; set; }
        DateTime dt = DateTime.Now;
        private void kataloqgetir_zamin()
        {//t.subschkre t.licschkre
            try
            {
                //dateTimePicker1.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                //(select  g1.summa_pog_kre+g1.summa_pog_pro  from odb.graphpogkre g1 where g1.subschkre = t.subschkre and g1.licschkre = t.licschkre and g1.date_pog = odb.func_tar((select max(gr3.date_pog) from odb.graphpogkre gr3  where gr3.subschkre = t.subschkre and gr3.licschkre = t.licschkre)))  ay_hisse
                OracleCommand Orcom = new OracleCommand("select r.name_regnom,t.licschkre, t.subschkre sk, g.guarantee_name," +
                    " t.date_open,t.summakre, t.summa from odb.licschkre t, odb.regnom r, odb.creditinfoguarantee g, odb.srokpogprockre k," +
                    " odb.licsch m where lower( g.pincode) like lower('"+txt_axtar.Text+"') " +
                    " and substr(t.licschkre, 10, 6) = r.regnom and t.licschkre = g.licschkre  and t.subschkre = g.subschkre " +
                    " and t.licschkre = k.licschkre  and t.subschkre = k.subschkre and m.licsch = k.licsch_3(+) order by t.subschkre desc", Orcon);
                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                DataTable Ordt = new DataTable();
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Adı";
                dataGridView1.Columns[0].Width = 350;

                dataGridView1.Columns[1].HeaderText = "Hesab";
                dataGridView1.Columns[1].Width = 200;

                dataGridView1.Columns[2].HeaderText = "Sk";
                dataGridView1.Columns[2].Width = 100;

                dataGridView1.Columns[3].HeaderText = "Zamin";
                dataGridView1.Columns[3].Width = 200;

                dataGridView1.Columns[4].HeaderText = "Tarix";
                dataGridView1.Columns[4].Width = 120;

                dataGridView1.Columns[5].HeaderText = "Kredit";
                dataGridView1.Columns[5].Width = 120;

                dataGridView1.Columns[6].HeaderText = "Qalıq";
                dataGridView1.Columns[6].Width = 120;

            }
            catch (Exception)
            {
                MessageBox.Show("Zamin barədə məlumat qeyd olunmayıb.", "Məlumat");
            }

            finally
            {
                
            }

        }
        public string Xmgisamezmun = string.Empty;
        public string Xmmektubmetn = string.Empty;
        private void meknoal()
        {
            Xmgisamezmun = "Zamin borcun baglanması";
            Xmmektubmetn = txbzamin.Text.Trim();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            Orcon.Open();

            Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci,  x.il) values ('" + Xmmektubmetn + "' , TO_DATE('" + dateavtoMuqtarix.Text + "','dd-MM-yyyy'),'"+ Xmgisamezmun + "' ,'" + icraci_kod + "', " + tarixIl + ")", Orcon);
            Orcom.ExecuteNonQuery();
            Orcon.Close();
        }
        private void wordeat()
        {
            
            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis wordler");
            cl.fileName = "Zamin "+txbzamin.Text+" Borcalan "+txbzBorcalan.Text + ".docx"; // Fayl adına `.docx` əlavə edirəm.
            string yaziile = aletler.YaziyaCevir(Convert.ToDecimal( txbmebleg.Text),txt_val.Text);

            var replacements = new Dictionary<string, string>
{
                 { "{mekNo}", mek_no },
                 {"{mektarixi}", Aletler.TarixiSozeCevir(DateTime.Now.ToString("dd-MM-yyyy"))},
                 { "{muqtar}", Aletler.TarixiSozeCevir(txbtarix.Text)},
                 { "{borcalan}", txbzBorcalan.Text},
                 { "{krtar}", txbzamin.Text},
                 { "{zamin}", txbzamin.Text},
                 { "{mebleg}", txbmebleg.Text + " (" + yaziile + ")"},
};

            // Fayl adını yoxlayır və eyni fayl varsa nömrələmə edir
            string fullPath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
            int fileCounter = 1;

            while (File.Exists(fullPath))
            {
                cl.fileName = $"{txbzamin.Text} - {fileCounter}.docx";
                fullPath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                fileCounter++;
            }

            string zaminmektub = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Wordler", "Zaminarayis1.docx");
            string outputPath = fullPath; // Artıq saylı fayl adı düzgün təyin olundu

            // Word sənədini yaradın
            aletler.CreateWordDocument(zaminmektub, outputPath, replacements);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string adhazir;

            txbzBorcalan.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            adhazir = txbzBorcalan.Text.ToLower();
            txbzBorcalan.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(adhazir);
            txbtarix.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString().Substring(0, 10);
            txbzamin.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            //dattar.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString().Substring(0, 10);
            txbmebleg.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();

            string val_kod = dataGridView1.CurrentRow.Cells[1].Value.ToString().Substring(6, 2);
            if (val_kod == "00")
            {
                txt_val.Text = "AZN";
            }
            else if (val_kod == "01")
            {
                txt_val.Text = "USD";
            }
            else if (val_kod == "02")
            {
                txt_val.Text = "AVRO";
            }
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void zaminarayis_Load(object sender, EventArgs e)
        {
            string gun, ay, il;
            gun = DateTime.Now.Date.Day.ToString();
            ay = DateTime.Now.Date.Month.ToString();
            il = DateTime.Now.Date.Year.ToString();
            dateavtoMuqtarix.Text = gun + "-" + ay + "-" + il;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            kataloqgetir_zamin();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            mek_no = aletler.GedenMektubNo("odb.xaric_mektub", "qey_nom", "il");
            //meknoal();
            aletler.GedenMektubaElaveEt(txbzamin.Text,"", "Zamin borcun baglanması", icraci_kod);
            wordeat();
        }
    }
}
