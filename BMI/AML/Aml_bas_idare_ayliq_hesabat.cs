using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.AML
{
    public partial class Aml_bas_idare_ayliq_hesabat : Form
    {
        public Aml_bas_idare_ayliq_hesabat()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        cl_aletler cl = new cl_aletler();

        private void axtar()
        {
            try
            {
                string startDate = txt_giris.Text;
                string endDate = txt_cixis.Text;

                string sorgu = "select count(*) say,sum( d.summa_v_nacval) meb " +
                 " from arh_dd d " +
                 " where " +
                 " d.date_oper between to_date(:start_date, 'dd-mm-yyyy') and to_date(:end_date, 'dd-mm-yyyy') and " +
                 " substr(d.debet, 1, 1) not in (5, 6, 7, 8, 9) and " +
                 " substr(d.kredit, 1, 1) not in (5, 6, 7, 8, 9) and " +
                 " (substr(d.debet, 1, 1) not in (1) or substr(d.kredit, 1, 1) not in (1)) and " +
                 " (substr(d.debet, 1, 1) not in (2) or substr(d.kredit, 1, 1) not in (2)) and " +
                 " (substr(d.debet, 1, 1) not in (4) or substr(d.kredit, 1, 1) not in (2))";

                using (OracleConnection connection = new OracleConnection(cl.baglan))
                {
                    OracleCommand command = new OracleCommand(sorgu, connection);
                    command.Parameters.Add(new OracleParameter("start_date", startDate.Substring(0, 10)));
                    command.Parameters.Add(new OracleParameter("end_date", endDate.Substring(0, 10)));
                    connection.Open();
                    OracleDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Toplam satır sayısını textbox1'e yazdır
                        txt_say.Text = reader["say"].ToString();
                        // Toplam meblağı textbox2'ye yazdır
                        //txt_mebleg.Text = reader["meb"].ToString();
                        decimal totalCount = Convert.ToDecimal(reader["meb"]);
                        txt_mebleg.Text = string.Format("{0:#,0.00}", totalCount);
                    }
                    reader.Close();
                }
            }
            catch (Exception)
            {


            }
            finally { }
            
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            axtar();
        }

        private void txt_giris_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txt_cixis.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_giris_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_giris.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_giris.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void txt_cixis_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_cixis_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txt_cixis.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txt_cixis.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AML.frm_aml_bas_idare_yanasma frm = new frm_aml_bas_idare_yanasma();
            frm.ShowDialog();
        }
    }
}
