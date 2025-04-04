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
    public partial class Hevale : Form
    {
        public Hevale()
        {
            InitializeComponent();
        }
        int alinancavab = 0;

        private void button7_Click(object sender, EventArgs e)
        {
            ged_hev frmyged = new ged_hev(this);
            frmyged.Show();
        }

        private void button8_Click(object sender, EventArgs e) //Silmek
        {
            try
            {
                string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                OracleConnection connection = new OracleConnection(connectrionString);


                connection.Open();
                OracleCommand komut = new OracleCommand("delete from odb.gelen_hevale where hev_nom='" + textBox3.Text + "'", connection);
                
                if (alinancavab > 0)
                {
                    MessageBox.Show("Məktub məlumat bazasından silindi...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub silinmədi... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                connection.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show("hazir deyil");
            }
            

            //baglan.Open();
            //SqlCommand komut = new SqlCommand("delete from geden_hevale where HEV_NOM ='"+dataGridView2.SelectedCells+"'", baglan);
            //komut.ExecuteNonQuery();
            //baglan.Close();

            yenilegeden();
            textBox4.Text = "";
        }
        private void yenilegeden()
        {
            try
            {
                OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Ocon.Open();
                OracleCommand Ocom = new OracleCommand("Select * From odb.geden_hevale_samir", Ocon);

                OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
                DataTable Odt = new DataTable();
                Oda.Fill(Odt);
                dataGridView2.DataSource = Odt;
                Ocon.Close();

                lblSetirsayı.Text = "";
            lblSetirsayı.Text = dataGridView2.RowCount.ToString();
            if (Convert.ToInt32(lblSetirsayı.Text) > 1)
            {
                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }



                //string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                //OracleConnection connection = new OracleConnection(connectrionString);
                //connection.Open();
                //OracleDataAdapter ordtr = new OracleDataAdapter("Select * From odb.geden_hevale_samir", connection);
                //DataTable tablo = new DataTable();
                //ordtr.Fill(tablo);
                //dataGridView2.DataSource = tablo;
                //connection.Close();
                //OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                //Orcon.Open();
                //OracleCommand Orcom = new OracleCommand("select t.licschkre,t.subschkre sk,substr(t.licschkre,10,6)||t.subschkre kod,r.name_regnom,r.telefon,r.mobilniy,r.passport,r.senedi_veren_orqaninin_adi ver_orq,r.senedin_verilme_tarixi ver_tar,t.summa,t.procstavkre fz,t.date_open,g.guarantee_name,g.guarantee_id from odb.licschkre t,odb.regnom r,odb.creditinfoguarantee g where t.date_open=to_date('22/06/2021', 'dd/mm/yyyy') and substr(t.licschkre,10,6)=r.regnom and t.licschkre=g.licschkre and t.subschkre=g.subschkre", Orcon);

                //OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                //DataTable Ordt = new DataTable();
                //Orda.Fill(Ordt);
                //dataGridView2.DataSource = Ordt;
                //Orcon.Close();
                //baglan.Open();
                //SqlDataAdapter adtr = new SqlDataAdapter("Select * From geden_hevale", baglan);
                //DataTable tablo = new DataTable();
                //adtr.Fill(tablo);
                //dataGridView2.DataSource = tablo;
                //baglan.Close();
                //return tablo;
            //}
            //catch (Exception)
            //{

            //    MessageBox.Show("hazir deyil");
            //}
            
        }

        private void yenilegelen()
        {
            try
            {
                string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                OracleConnection connection = new OracleConnection(connectrionString);
                connection.Open();
                OracleDataAdapter ordtr = new OracleDataAdapter("Select * From odb.gelen_hevale", connection);//sehv verse gelen_hevale adina duz bax
                DataTable tablo = new DataTable();
                ordtr.Fill(tablo);
                dataGridView1.DataSource = tablo;
                connection.Close();

            }
            catch (Exception)
            {

                MessageBox.Show("hazir deyil");
            }
            
            
        }

        private void Hevale_Load(object sender, EventArgs e)
        {
            yenilegeden();
            yenilegelen();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                OracleConnection connection = new OracleConnection(connectrionString);
                connection.Open();
                OracleCommand isleme = new OracleCommand("select * from odb.geden_hevale_samir where SAA like '%" + textBox1.Text + "%' or hev_nom like '" + textBox1.Text + "%'", connection);
                isleme.ExecuteNonQuery();
                OracleDataAdapter da = new OracleDataAdapter(isleme);
                DataTable tablo2 = new DataTable();

                da.Fill(tablo2);
                dataGridView2.DataSource = tablo2;
                connection.Close();

            }
            catch (Exception ex)
            {

                MessageBox.Show("hazir deyil");
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Gel_hev glh = new Gel_hev();
            glh.ShowDialog();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                OracleConnection connection = new OracleConnection(connectrionString);
                connection.Open();
                OracleCommand isleme = new OracleCommand("select * from odb.gelen_hevale where SAA like '%" + textBox2.Text + "%' or hev_nom like '" + textBox2.Text + "%'", connection);
                isleme.ExecuteNonQuery();
                OracleDataAdapter da = new OracleDataAdapter(isleme);
                DataTable tablo2 = new DataTable();

                da.Fill(tablo2);
                dataGridView1.DataSource = tablo2;
                connection.Close();

            }
            catch (Exception ex)
            {

                MessageBox.Show("hazir deyil");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
                OracleConnection connection = new OracleConnection(connectrionString);


                connection.Open();
                OracleCommand komut = new OracleCommand("delete from odb.gelen_hevale where hev_nom='" + textBox3.Text + "' or tarix='" + textBox5.Text + "' ", connection);
                alinancavab = komut.ExecuteNonQuery();
                
                alinancavab = komut.ExecuteNonQuery();
                if (alinancavab > 0)
                {
                    MessageBox.Show("Məktub məlumat bazasından silindi...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub silinmədi... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                connection.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show("hazir deyil");
            }


            //baglan.Open();
            //SqlCommand komut = new SqlCommand("delete from geden_hevale where HEV_NOM ='"+dataGridView2.SelectedCells+"'", baglan);
            //komut.ExecuteNonQuery();
            //baglan.Close();

            yenilegelen();
            textBox3.Text = "";
            textBox5.Text = "";
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox4.Text = dataGridView2.CurrentRow.Cells["HEV_NOM"].Value.ToString();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView1.CurrentRow.Cells["HEV_NOM"].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells["TARIX"].Value.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
