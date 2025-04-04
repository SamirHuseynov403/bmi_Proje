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
    public partial class Mebleg_uzre : Form
    {
        public Mebleg_uzre()
        {
            InitializeComponent();
        }
        public int zam_say=0;

        private void gr_uzre()
        {
            string tarixIl = DateTime.Now.Date.Year.ToString();

            ////string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            //con.Open();
            //OracleCommand komut = new OracleCommand();
            //komut.Connection = con;
            //komut.CommandText = "select sum(summa) from odb.licschkre where tipzaloga='" + comboBox1.Text + "' and date_open >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')";
            //OracleDataReader dr = komut.ExecuteReader();

            //while (dr.Read())
            //{

            //    zam_say =Convert.ToUInt16 (dr["summa"].ToString());
            //    txb1_qaliq.Text = zam_say.ToString();

            //}


            //con.Close();

            string Kasa = "Select SUM(summa)from odb.licschkre where tipzaloga='" + comboBox1.Text + "' and date_open >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy')";

            con.Open();
            OracleCommand cmd = new OracleCommand(Kasa, con);
            cmd.CommandType = CommandType.Text;
            OracleDataReader reader = cmd.ExecuteReader();

            reader.Read();
            txb1_qaliq.Text = reader["summa"].ToString();

            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //gr_uzre();
            
            muqavile_nom1();
            test();
            toplama();
            toplam2();
            toplam3();
            toplam4();
            toplam5();
            toplam6();
            
        }

        private void test()
        {
            string avto = comboBox1.Text.ToLower();
            string say1=txbmeb1.Text, say2=txbmeb2.Text;
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            //and lk.summakre between "+say1+" and "+say2+"
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleCommand Orcom = new OracleCommand("select lk.summakre,lk.summa, lk.summa_19, g.name from odb.licschkre lk , odb.tipzal g where g.code=lk.tipzaloga and lk.tipzaloga = '"+label12.Text+"' and lk.date_close is null   ", Orcon);
            //OracleCommand Orcom = new OracleCommand("select * from tipzal", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();

            dataGridView1.Columns[0].HeaderText = "Kredit məbləği";
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].HeaderText = "Qalıq məbləği";
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].HeaderText = "V/K məbləğ";
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].HeaderText = "Girovun növü";
            dataGridView1.Columns[3].Width = 200;
        }

        private void novlergetir()
        {
            OracleConnection baglanti = new OracleConnection();
            baglanti.ConnectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            OracleCommand komut = new OracleCommand();
            komut.CommandText = "SELECT *FROM tipzal";
            komut.Connection = baglanti;
            komut.CommandType = CommandType.Text;

            OracleDataReader dr;
            baglanti.Open();
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["name"]);
            }

            baglanti.Close();
            
        }

        
        public void muqavile_nom1()
        {
            string sayitap;
            string avto = comboBox1.Text.ToLower().Replace('ə', 'a');
            System.DateTime moment = new System.DateTime(
                                1999, 1, 13, 3, 57, 32, 11);
            // Year gets 1999.
            int year = moment.Year;
            string tarixIl = DateTime.Now.Date.Year.ToString();

            ////string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
            OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
            con.Open();
            OracleCommand komut = new OracleCommand();
            komut.Connection = con;
            komut.CommandText = "select g.code from odb.tipzal g where odb.func_utf8_to_latin(g.name)=odb.func_utf8_to_latin(upper('" + avto + "'))";
            OracleDataReader dr = komut.ExecuteReader();


            while (dr.Read())
            {
             zam_say = Convert.ToInt32(dr["code"].ToString());
                
            }


            con.Close();
            label12.Text = zam_say.ToString();


        }

        private void Mebleg_uzre_Load(object sender, EventArgs e)
        {
            novlergetir();
        }

        private void toplama()
        {
            int satirsayi = -1;
            Decimal tx1 =Convert.ToDecimal( txbmeb1.Text);
            Decimal tx2 = Convert.ToDecimal(txbmeb2.Text);
            Decimal dtg1 = 0;
            Decimal dtg2 = 0;
            Decimal toplam = 0;
            Decimal toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1= Convert.ToDecimal(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1>=tx1 && dtg1 <=tx2)
                {
                    toplam += Convert.ToDecimal(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDecimal(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi =satirsayi+ 1;
                }
                
            }
            if (satirsayi.ToString()=="-1")
            {
                satirsayi = 0;
            }
             toplam = toplam + toplamvk;
            if (checkBox7.Checked==true)
            {
                txb1_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000,2).ToString();
            }
            else
            {
                txb1_qaliq.Text = toplam.ToString();
            }
            

            txbsay1.Text=satirsayi.ToString();
        }
        private void toplam2()
        {
            int satirsayi = -1;
            double tx1 = Convert.ToDouble(txbmeb3.Text);
            double tx2 = Convert.ToDouble(txbmeb4.Text);
            double dtg1 = 0;
            double dtg2 = 0;
            double toplam = 0;
            double toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1 >= tx1 && dtg1 <= tx2)
                {
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi = satirsayi + 1;
                }

            }
            if (satirsayi.ToString() == "-1")
            {
                satirsayi = 0;
            }
            toplam = toplam + toplamvk;
            if (checkBox7.Checked == true)
            {
                txb2_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000, 2).ToString();
            }
            else
            {
                txb2_qaliq.Text = toplam.ToString();
            }

            txbsay2.Text = satirsayi.ToString();
        }
        private void toplam3()
        {
            int satirsayi = -1;
            double tx1 = Convert.ToDouble(txbmeb5.Text);
            double tx2 = Convert.ToDouble(txbmeb6.Text);
            double dtg1 = 0;
            double dtg2 = 0;
            double toplam = 0;
            double toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1 >= tx1 && dtg1 <= tx2)
                {
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi = satirsayi + 1;
                }

            }
            if (satirsayi.ToString() == "-1")
            {
                satirsayi = 0;
            }
            toplam = toplam + toplamvk;
            if (checkBox7.Checked == true)
            {
                txb3_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000, 2).ToString();
            }
            else
            {
                txb3_qaliq.Text = toplam.ToString();
            }

            txbsay3.Text = satirsayi.ToString();
        }
        private void toplam4()
        {
            int satirsayi = -1;
            double tx1 = Convert.ToDouble(txbmeb7.Text);
            double tx2 = Convert.ToDouble(txbmeb8.Text);
            double dtg1 = 0;
            double dtg2 = 0;
            double toplam = 0;
            double toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1 >= tx1 && dtg1 <= tx2)
                {
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi = satirsayi + 1;
                }

            }
            if (satirsayi.ToString() == "-1")
            {
                satirsayi = 0;
            }
            toplam = toplam + toplamvk;
            if (checkBox7.Checked == true)
            {
                txb4_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000, 2).ToString();
            }
            else
            {
                txb4_qaliq.Text = toplam.ToString();
            }

            txbsay4.Text = satirsayi.ToString();
        }
        private void toplam5()
        {
            int satirsayi = -1;
            double tx1 = Convert.ToDouble(txbmeb9.Text);
            double tx2 = Convert.ToDouble(txbmeb10.Text);
            double dtg1 = 0;
            double dtg2 = 0;
            double toplam = 0;
            double toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1 >= tx1 && dtg1 <= tx2)
                {
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi = satirsayi + 1;
                }

            }
            if (satirsayi.ToString() == "-1")
            {
                satirsayi = 0;
            }
            toplam = toplam + toplamvk;
            if (checkBox7.Checked == true)
            {
                txb5_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000, 2).ToString();
            }
            else
            {
                txb5_qaliq.Text = toplam.ToString();
            }

            txbsay5.Text = satirsayi.ToString();
        }
        private void toplam6()
        {
            int satirsayi = -1;
            double tx1 = Convert.ToDouble(txbmeb11.Text);
            //double tx2 = Convert.ToDouble(txbmeb12.Text);
            double dtg1 = 0;
            double dtg2 = 0;
            double toplam = 0;
            double toplamvk = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                dtg1 = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                if (dtg1 > tx1 )
                {
                    toplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    toplamvk += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    satirsayi = satirsayi + 1;
                }

            }
            if (satirsayi.ToString() == "-1")
            {
                satirsayi = 0;
            }
            toplam = toplam + toplamvk;
            if (checkBox7.Checked == true)
            {
                txb6_qaliq.Text = decimal.Round(Convert.ToDecimal(toplam) / 1000, 2).ToString();
            }
            else
            {
                txb6_qaliq.Text = toplam.ToString();
            }

            txbsay6.Text = satirsayi.ToString();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        
    }
}
