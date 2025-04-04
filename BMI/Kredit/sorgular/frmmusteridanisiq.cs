using DevExpress.XtraEditors;
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

namespace BMI
{
    public partial class frmmusteridanisiq : Form
    {
        public frmmusteridanisiq()
        {
            InitializeComponent();
        }
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        OracleCommand komyazdir;
        OracleCommand komutmurraz;
        public string icraci_kod = string.Empty;
        DataTable muracietler = new DataTable();
        string tarixIl = DateTime.Now.Date.Year.ToString();
        string tarixay = DateTime.Now.Date.Month.ToString();
        string tarixgun = DateTime.Now.Date.Day.ToString();

        private void temizlemuraciet()
        {
            try
            {
                txbdanis.Text = "";
                dateEdit2.Text = "";
                txtsaat.Text = "";
                

            }
            catch (Exception)
            {


            }


        }
        public void listelemuraciet()
        {
            //try
            //{
                muracietler.Clear();
                string tarixIl = txbil.Text;
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                OracleDataAdapter isleme = new OracleDataAdapter("select Sira,Tarix,Adi,Raziliq_mebleg,Telefon,danish_qeyd,Gelme_tarixi,Saat,İcraci,il from odb.Muracietler where il>='" + tarixIl + "' and Melumat is null and Komite_qerar <> 'Etiraz' order by sira Desc ", con);
                isleme.Fill(muracietler);
                gridControl1.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                gridmuracietler.Columns[0].Caption = "Sira";
                gridmuracietler.Columns[0].Width = 80;
                gridmuracietler.Columns[1].Caption = "Tarix";
                gridmuracietler.Columns[1].Width = 110;
                gridmuracietler.Columns[2].Caption = "Adı";
                gridmuracietler.Columns[2].Width = 220;
                gridmuracietler.Columns[3].Caption = "Razılıq məbləği";
                gridmuracietler.Columns[3].Width = 100;
                gridmuracietler.Columns[4].Caption = "Telefon";
                gridmuracietler.Columns[4].Width = 150;
                gridmuracietler.Columns[5].Caption = "Danışıq qedləri";
                gridmuracietler.Columns[5].Width = 350;
                gridmuracietler.Columns[6].Caption = "Gəlmə tarixi";
                gridmuracietler.Columns[6].Width = 100;
                gridmuracietler.Columns[7].Caption = "Saat";
                gridmuracietler.Columns[7].Width = 50;
                gridmuracietler.Columns[8].Caption = "İcraçı";
                gridmuracietler.Columns[8].Width = 70;
                gridmuracietler.Columns[9].Caption = "İl";
                gridmuracietler.Columns[9].Width = 50;

            //}
            //catch (Exception)
            //{


            //}
            //finally { };


        }

        private void frmmusteridanisiq_Load(object sender, EventArgs e)
        {
            txbil.Text = DateTime.Now.Date.Year.ToString();
            //string tarixIl = DateTime.Now.Date.Year.ToString();
            listelemuraciet();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            listelemuraciet();
        }

        private void gridmuracietler_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataRow row = gridmuracietler.GetDataRow(gridmuracietler.FocusedRowHandle);
                txbmurad.Text = row[2].ToString();
                txbdanis.Text = row[5].ToString();
                //txbmurmeb.Text = row[3].ToString();
                txtsaat.Text = row[7].ToString();
                txbsiranoelave.Text = row[0].ToString();
                txbil.Text = row[9].ToString();
                dateEdit2.Text = row[6].ToString();
            }
            catch (Exception)
            {


            }
            finally { }
        }

        private void BtnEkle_Click(object sender, EventArgs e)
        {
            try
            {
                
                string sorgu5 = "Update odb.Muracietler set danish_qeyd='" + txbdanis.Text + "', Gelme_tarixi='" + dateEdit2.Text + "',saat='" + txtsaat.Text + "' where Sira=" + txbsiranoelave.Text + " and il='" + txbil.Text + "'";
                komyazdir = new OracleCommand(sorgu5, con);
                con.Open();
                komyazdir.ExecuteNonQuery();
                con.Close();
                XtraMessageBox.Show("Məlumatlar yeniləndi", "Qeyd");
                muracietler.Clear();
                listelemuraciet();
                temizlemuraciet();
                //muracietsay();
            }
            catch (Exception)
            {
                throw;
            }
            finally { }
        }
    }
}
