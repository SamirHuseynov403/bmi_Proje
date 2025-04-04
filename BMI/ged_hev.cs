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

namespace BMI
{
    public partial class ged_hev : Form
    {
        Hevale fghev;
        Pul_Kocurmesi fgh;

        public ged_hev(Hevale fd)
        {
            InitializeComponent();
            this.fghev = fd;
        }
        public ged_hev(Pul_Kocurmesi fs)
        {
            InitializeComponent();
            this.fgh = fs;
        }
        SqlConnection baglan = new SqlConnection("Data Source=.\\sqlexpress;Initial Catalog=BMI;Integrated Security=True");
        OracleConnection OrConnect = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass;");


        private void ged_hev_Load(object sender, EventArgs e)
        {
            hevnom_al();
        }
        private void hevnom_al()
        {

            OracleCommand cmd = new OracleCommand();
            OracleDataReader SR = null;
            cmd.Connection = OrConnect;
            cmd.CommandText = "Select HEV_NOM from odb.GEDEN_HEVALE_SAMIR order by HEV_NOM desc";
            OrConnect.Open();
            SR = cmd.ExecuteReader();
            if (SR.Read())
            {
                string heval = SR.GetValue(0).ToString();
                string hevsecal = heval.Substring(5);
                int hevno = Convert.ToInt32(hevsecal.ToString());
                hevno = hevno + 1;
                string cariil = DateTime.Now.Date.Year.ToString().Substring(2);
                txbgedHNo.Text = cariil + "-T-" + hevno.ToString();
                Pul_Kocurmesi pkh = new Pul_Kocurmesi();
                pkh.txbhevale.Text = cariil + "-T-" + hevno.ToString();
            }
            OrConnect.Close();
        }
        private void gedenhevaleyazdir()
        {
            txbgedtarix.Text = DateTime.Now.Date.ToString("dd-MM-yyyy");
            OrConnect.Open();
            OracleCommand command = new OracleCommand("INSERT INTO ODB.geden_hevale_samir(HEV_NOM,HES_NOM,SAA,Mebleg,Val_tip,tarix) values ('" + txbgedHNo.Text + "','" + txbgedHesNo.Text + "','" + txbgedSAA.Text + "','" + txbgedmebleg.Text + "','" + cmbgedvalyuta.Text + "',to_date(sysdate,'dd-mm-yyyy'))", OrConnect);
            command.ExecuteNonQuery();
            OrConnect.Close();

            this.Close();


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
            string nom = cariil + "-T-" + birartir.ToString();
            txbgedHNo.Text = nom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            hevnom_al();
            gedenhevaleyazdir();
        }
    }
}
