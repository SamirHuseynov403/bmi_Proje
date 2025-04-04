using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Data.OleDb;
using System.IO;

namespace BMI
{
    public partial class Onlayn : Form
    {
        public Form1 frmana1;
        Form1 fgh;

        public void onlayn(Form1 fs)
        {
            InitializeComponent();
            this.fgh = fs;
        }
        public Onlayn()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        DataTable cedvelonlayn = new DataTable();
        OleDbCommand yaz;
        OleDbCommand yaz1;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                cedvelonlayn.Clear();
                Form1 frm1 = new Form1();
                string sorgu = "Insert into Onlayn(sira,adi,mebleg,fin,tarix,baxildi,icraci) values (@sira,@adi,@mebleg,@fin,@tarix,@baxildi,@icraci)";
                yaz = new OleDbCommand(sorgu, baglanti);
                yaz.Parameters.AddWithValue("@sira", txbonsira.Text);
                yaz.Parameters.AddWithValue("@adi", txbonad.Text);
                yaz.Parameters.AddWithValue("@mebleg", txbonmebleg.Text);
                yaz.Parameters.AddWithValue("@fin", txbonfin.Text);
                yaz.Parameters.AddWithValue("@tarix", dateTimePicker1.Value.ToString("dd-MM-yyyy"));
                yaz.Parameters.AddWithValue("@baxildi", "baxilmamis");
                yaz.Parameters.AddWithValue("@icraci", label7.Text);
                baglanti.Open();
                yaz.ExecuteNonQuery();
                baglanti.Close();
                listeleonlayn();
                sayıonlayn();
                temizleonlayn();
            }
            catch (Exception)
            {
              
            }
            
        }
        private void sayıonlayn()
        {
            string proid;
            string query = "select sira from onlayn order by sira Desc";
            baglanti.Open();
            OleDbCommand cmd = new OleDbCommand(query, baglanti);
            OleDbDataReader dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                int id = int.Parse(dr[0].ToString()) + 1;
                proid = id.ToString();
            }
            else if (Convert.IsDBNull(dr))
            {
                proid = ("1");
            }
            else
            {
                proid = ("1");
            }
            baglanti.Close();
            txbonsira.Text = proid.ToString();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            
        }

        void listeleonlayn()
        {
            cedvelonlayn.Clear();
            baglanti.Open();
            OleDbDataAdapter islek = new OleDbDataAdapter("select * from onlayn", baglanti);
            islek.Fill(cedvelonlayn);
            dataGridView1.DataSource = cedvelonlayn;
            baglanti.Close();
            dataGridView1.Columns[0].HeaderText = "Sıra";
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].HeaderText = "Adı";
            dataGridView1.Columns[1].Width = 320;
            dataGridView1.Columns[2].HeaderText = "Məbləğ";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].HeaderText = "FİN";
            dataGridView1.Columns[3].Width = 110;
            dataGridView1.Columns[4].HeaderText = "Tarix";
            dataGridView1.Columns[4].Width = 110;
            dataGridView1.Columns[5].HeaderText = "Status";
            dataGridView1.Columns[5].Width = 150;
            dataGridView1.Columns[6].HeaderText = "İcraçı";
            dataGridView1.Columns[6].Width = 250;
        }

        void listeleonlaynbaxilmamis()

        {
            cedvelonlayn.Clear();
            baglanti.Open();
            OleDbDataAdapter islek = new OleDbDataAdapter("select * from onlayn where baxildi='baxilmamis'", baglanti);
            islek.Fill(cedvelonlayn);
            dataGridView1.DataSource = cedvelonlayn;
            baglanti.Close();
            dataGridView1.Columns[0].HeaderText = "Sıra";
            dataGridView1.Columns[0].Width = 60;
            dataGridView1.Columns[1].HeaderText = "Adı";
            dataGridView1.Columns[1].Width = 320;
            dataGridView1.Columns[2].HeaderText = "Məbləğ";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].HeaderText = "FİN";
            dataGridView1.Columns[3].Width = 110;
            dataGridView1.Columns[4].HeaderText = "Tarix";
            dataGridView1.Columns[4].Width = 110;
            dataGridView1.Columns[5].HeaderText = "Status";
            dataGridView1.Columns[5].Width = 150;
            dataGridView1.Columns[6].HeaderText = "İcraçı";
            dataGridView1.Columns[6].Width = 250;
        }

        private void Onlayn_Load(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();
            listeleonlayn();
            sayıonlayn();
            onlaynsay();
            label7.Text = frm1.lblTamad.Text;
            txbonfin.MaxLength=7;
            dateTimePicker1.Value = DateTime.Now;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                OleDbCommand sorgu = new OleDbCommand("delete  from onlayn where sira=" + dataGridView1.CurrentRow.Cells[0].Value.ToString() + "", baglanti);
                sorgu.ExecuteNonQuery();
                baglanti.Close();
                MessageBox.Show("Məlumat silindi");
                cedvelonlayn.Clear();
                listeleonlayn();
                sayıonlayn();
                temizleonlayn();
            }
            catch (Exception)
            {
                
                
            }
            
        }

        private void onlaynsay() 
        {
            int satirsayi = -1;

            OleDbCommand orcmd = new OleDbCommand("select count(*) from onlayn where baxildi='baxilmamis'", baglanti);
            baglanti.Open();
            satirsayi = Convert.ToInt32(orcmd.ExecuteScalar());
            baglanti.Close();
            //string cariil = DateTime.Now.Date.Year.ToString();
            int birartir = satirsayi;
            string nom = birartir.ToString();
            textBox1.Text = nom;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sorgu1 = "update onlayn set adi=@adi,mebleg=@mebleg,fin=@fin where Sira=" + textBox1.Text + "";
            yaz1 = new OleDbCommand(sorgu1, baglanti);
            yaz1.Parameters.AddWithValue("@adi", txbonad.Text);
            yaz1.Parameters.AddWithValue("@mebleg", txbonmebleg.Text);
            yaz1.Parameters.AddWithValue("@fin", txbonfin.Text);
            baglanti.Open();
            yaz1.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Məlumat yeniləndi");
            listeleonlayn();
            temizleonlayn();
        }

        private void temizleonlayn()
        {
            txbonad.Text = "";
            txbonfin.Text = "";
            txbonmebleg.Text = "";
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                baglanti.Open();
                OleDbDataAdapter axtar = new OleDbDataAdapter("select * from onlayn where  Adi like '%" + textBox2.Text + "%' or fin like '%" + textBox2.Text + "%'", baglanti);
                DataTable tablo2 = new DataTable();
                axtar.Fill(tablo2);
                dataGridView1.DataSource = tablo2;
                baglanti.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("Xəta baş verdi");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listeleonlaynbaxilmamis();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listeleonlayn();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            txbonsira.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txbonad.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txbonmebleg.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            txbonfin.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            

        }
    }
}
