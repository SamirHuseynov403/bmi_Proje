using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;



namespace Kredit_isler
{
    public partial class Gelme_vaxti : Form
    {
        public Gelme_vaxti()
        {
            InitializeComponent();
        }
        //OleDbConnection baglanti;
        //OleDbDataAdapter da;
        OleDbCommand komut;
        OleDbCommand komut1;
        //DataSet ds;

        OleDbConnection baglanti10 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        OleDbConnection baglan = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        DataTable tablo10=new DataTable();

        int raz_meb = Formesas.raz_mebleg;
        int teyin_sira=Formesas.teyin_sirasi;
        private void Gelme_vaxti_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'kreditDataSet8.Teyin_olumuslar' table. You can move, or remove it, as needed.
            //this.teyin_olumuslarTableAdapter.Fill(this.kreditDataSet8.Teyin_olumuslar);

            Formesas frm1 = new Formesas();
            txb_raz_ad.Text = Formesas.gelmevaxt_adi.ToString();
            txb_gelme_sira.Text = teyin_sira.ToString();
            txb_raz_meb.Text = raz_meb.ToString();
            listele_gelenler();
            //frm1.tarix_goster();
            date_teyin.Value = DateTime.Now;
            
        }

        private void listele_gelenler()
        {
            tablo10.Clear();
            baglan.Open();
            OleDbDataAdapter isleme = new OleDbDataAdapter("select * from Teyin_olumuslar", baglan);
            isleme.Fill(tablo10);
            dgw_gelmetarix.DataSource = tablo10;

            baglan.Close();
        }

        private void testet()
        {
            try
            {
                string sorgu = "Insert into Teyin_olumuslar(sira,Adi_soyadi,Mebleg,gelme_tarixi,gelme_saati,qeyd) values (@sira,@adi,@mebleg,@gelmetarix,@gelmesaati,@qeyd) where Sira != " + txb_gelme_sira.Text + " ";
                komut = new OleDbCommand(sorgu, baglan);
                komut.Parameters.AddWithValue("@sira", txb_gelme_sira.Text);
                komut.Parameters.AddWithValue("@adi", txb_raz_ad.Text);
                komut.Parameters.AddWithValue("@mebleg", txb_raz_meb.Text);
                komut.Parameters.AddWithValue("@gelmetarix", date_teyin.Value.ToString("dd-MM-yyyy"));
                komut.Parameters.AddWithValue("@gelmesaati", txb_raz_saat.Text);
                komut.Parameters.AddWithValue("@qeyd", txb_raz_qeyd.Text);
                baglan.Open();
                komut.ExecuteNonQuery();
                baglan.Close();
                listele_gelenler();
            }
            catch (Exception)
            {
                
             MessageBox.Show("sehv oldu");
            }
            
        }
        private void button9_Click(object sender, EventArgs e)
        {
            //testet();

            //date_teyin.CustomFormat = "dd-MM-yyyy";
            string sorgu = "Insert into Teyin_olumuslar(sira,Adi_soyadi,Mebleg,gelme_tarixi,gelme_saati,qeyd) values (@sira,@adi,@mebleg,@gelmetarix,@gelmesaati,@qeyd)";
            string sorguRAZ = "update Verilmis_kr set Gelme_tarixi=@Tarix,Saat=@saat where Sira=" + txb_gelme_sira.Text + "";
            komut = new OleDbCommand(sorgu, baglan);
            komut1 = new OleDbCommand(sorguRAZ, baglan);

            komut1.Parameters.AddWithValue("@Tarix", date_teyin.Value.ToString("dd-MM-yyyy"));
            komut1.Parameters.AddWithValue("@saat", txb_raz_saat.Text);


            komut.Parameters.AddWithValue("@sira", txb_gelme_sira.Text);
            komut.Parameters.AddWithValue("@adi", txb_raz_ad.Text);
            komut.Parameters.AddWithValue("@mebleg", txb_raz_meb.Text);
            komut.Parameters.AddWithValue("@gelmetarix", date_teyin.Value.ToString("dd-MM-yyyy"));
            komut.Parameters.AddWithValue("@gelmesaati", txb_raz_saat.Text);
            komut.Parameters.AddWithValue("@qeyd", txb_raz_qeyd.Text);


            baglan.Open();
            komut.ExecuteNonQuery();
            komut1.ExecuteNonQuery();
            baglan.Close();
            listele_gelenler();
            Formesas frm1 = new Formesas();
            //frm1.listeleraziliq();
            tarixle_axtar();
            frm1.tarix_goster();
            MessageBox.Show("Məlumat qeydə alındı.");
            this.Close();
        }
        private void tarixle_axtar()
        {
            tablo10.Clear();
            baglan.Open();
            OleDbDataAdapter adtr = new OleDbDataAdapter("select * from Teyin_olumuslar where Gelme_tarixi like '" + date_teyin.Value.ToString("dd-MM-yyyy") + "'", baglan);
            adtr.Fill(tablo10);
            dgw_gelmetarix.DataSource = tablo10;
            baglan.Close();
            //dateTimePicker1.CustomFormat = "dd-MM-yyyy";
            //tablo10.Clear();
            //baglan.Open();
            //OleDbDataAdapter adtr = new OleDbDataAdapter("select * from Teyin_olumuslar where mebleg like '" + txb_raz_meb.Text + "'", baglan);
            //adtr.Fill(tablo10);
            //dgw_gelmetarix.DataSource = tablo10;
            //baglan.Close();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                //date_teyin.CustomFormat = ("dd-MM-yyyy");
                
                tablo10.Clear();
                baglan.Open();
                OleDbDataAdapter adtr = new OleDbDataAdapter("select * from Teyin_olumuslar where Gelme_tarixi like tr1", baglan);
                adtr.SelectCommand.Parameters.AddWithValue("tr1",date_teyin.Value.ToShortDateString());
                adtr.Fill(tablo10);
                dgw_gelmetarix.DataSource = tablo10;
                baglan.Close();

            }
            catch (Exception)
            {

                MessageBox.Show("sehv oldu");
            }
            
        }

        private void txb_raz_meb_TextChanged(object sender, EventArgs e)
        {
            tarixle_axtar();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void date_teyin_ValueChanged(object sender, EventArgs e)
        {
            //Form1 frm1 = new Form1();
            //frm1.tarix_goster();
            tarixle_axtar();
        }
    }
}
