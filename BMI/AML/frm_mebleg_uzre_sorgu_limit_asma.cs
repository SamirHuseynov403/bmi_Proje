using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
//using System.Text.RegularExpressions;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.AML
{
    public partial class frm_mebleg_uzre_sorgu_limit_asma : Form
    {
        public frm_mebleg_uzre_sorgu_limit_asma()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        DataTable dtcem = new DataTable();
        DataTable dtetrafli = new DataTable();
        string hesab = "";
        string tarix = "";
        private void axtar()
        {
            try
            {
                string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";

                dtcem.Clear();
                dtcem.Columns.Clear();
                string daily_report = " select d.date_oper tar,l.registrac_nomer q_no,r.name_regnom adi," +
                    " r.passport pas,r.pincode fin,r.adress unvan, " +
                    " case " +
                    " when r.m = 1 then 'kisi' " +
                    " when r.f = 1 then 'qadin' end cins," +
                    " r.grajdanstvo,sum(d.summa_v_nacval)mebleg,count(*) ms " +
                    " from arh_dd d, regnom r,licsch l " +
                    " where d.date_oper between TO_DATE('" + txt_giris.Text + "', 'dd/mm/yyyy') and TO_DATE('" + txt_cixis.Text + "', 'dd/mm/yyyy') " +
                    " and l.registrac_nomer not in '000004' and substr(d.kredit,1,2) in (38,39,40,41) and d.kredit = l.licsch and l.registrac_nomer = r.regnom and substr(d.kredit,1,2) not in ('86','66') and substr(d.debet,1,2) not in ('86','66')" +
                    " group by d.date_oper,l.registrac_nomer ,r.name_regnom ,r.passport ,r.pincode ,r.gender ,r.grajdanstvo,r.adress ," +
                    " case " +
                    " when r.m = 1 then 'kisi' " +
                    " when r.f = 1 then 'qadin' end " +
                    " having sum(d.summa_v_nacval) >= '" + txtmebleg.Text + "' order by d.date_oper asc ";


                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    using (OracleCommand command = new OracleCommand(daily_report, connection))
                    {
                        connection.Open();
                        OracleDataAdapter adapter = new OracleDataAdapter(command);
                        adapter.Fill(dtcem);
                        dataGridView1.DataSource = dtcem;
                    }
                    connection.Close();
                    dataGridView1.Columns[0].HeaderText = "Tarix";
                    dataGridView1.Columns[0].Width = 80;
                    dataGridView1.Columns[1].HeaderText = "Qeyd No";
                    dataGridView1.Columns[1].Width = 80;
                    //dataGridView1.Columns[2].HeaderText = "Hesab";
                    //dataGridView1.Columns[2].Width = 160;
                    dataGridView1.Columns[2].HeaderText = "Adı";
                    dataGridView1.Columns[2].Width = 280;
                    dataGridView1.Columns[3].HeaderText = "Pasport";
                    dataGridView1.Columns[3].Width = 90;
                    dataGridView1.Columns[4].HeaderText = "FİN";
                    dataGridView1.Columns[4].Width = 75;
                    dataGridView1.Columns[5].HeaderText = "Ünvan";
                    dataGridView1.Columns[5].Width = 435;
                    dataGridView1.Columns[6].HeaderText = "Cinsi";
                    dataGridView1.Columns[6].Width = 60;
                    dataGridView1.Columns[7].HeaderText = "V.";
                    dataGridView1.Columns[7].Width = 30;
                    dataGridView1.Columns[8].HeaderText = "Məbləğ";
                    dataGridView1.Columns[8].Width = 95;
                    dataGridView1.Columns[9].HeaderText = "M.S";
                    dataGridView1.Columns[9].Width = 55;
                }

            }
            catch (Exception)
            {
            }
            finally { }
        }

        private void axtaretrafli()
        {
            //try
            //{
                string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";

                //dataGridView1.Columns.Clear();
                dtetrafli.Clear();
                dtetrafli.Columns.Clear();
                string daily_report = " select d.date_oper,l.registrac_nomer q_no,d.kredit hesab,r.name_regnom adi," +
                    " r.passport pas,r.pincode fin,r.adress unvan, " +
                    " case " +
                    " when r.m = 1 then 'kisi' " +
                    " when r.f = 1 then 'qadin' end cins," +
                    " r.grajdanstvo,d.summa_v_nacval mebleg " +
                    " from arh_dd d, regnom r,licsch l " +
                    " where d.date_oper between TO_DATE('" + txt_giris.Text + "', 'dd/mm/yyyy') and TO_DATE('" + txt_cixis.Text + "', 'dd/mm/yyyy')" +
                    " and substr(d.kredit,1,2) in (38,39,40,41) and d.kredit = l.licsch " +
                    " and l.registrac_nomer = r.regnom and l.registrac_nomer='" + hesab+ "' and d.date_oper=TO_DATE('" + tarix + "','dd/mm/yyyy')";

                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    using (OracleCommand command = new OracleCommand(daily_report, connection))
                    {
                        connection.Open();
                        OracleDataAdapter adapter = new OracleDataAdapter(command);
                        adapter.Fill(dtetrafli);
                        dataGridView1.DataSource = dtetrafli;
                    }
                    connection.Close();
                    dataGridView1.Columns[0].HeaderText = "Tarix";
                    dataGridView1.Columns[0].Width = 80;
                    dataGridView1.Columns[1].HeaderText = "Qeyd No";
                    dataGridView1.Columns[1].Width = 80;
                    dataGridView1.Columns[2].HeaderText = "Hesab";
                    dataGridView1.Columns[2].Width = 180;
                    dataGridView1.Columns[3].HeaderText = "Adı";
                    dataGridView1.Columns[3].Width = 220;
                    dataGridView1.Columns[4].HeaderText = "Pasport";
                    dataGridView1.Columns[4].Width = 90;
                    dataGridView1.Columns[5].HeaderText = "FİN";
                    dataGridView1.Columns[5].Width = 75;
                    dataGridView1.Columns[6].HeaderText = "Ünvan";
                    dataGridView1.Columns[6].Width = 380;
                    dataGridView1.Columns[7].HeaderText = "Cinsi";
                    dataGridView1.Columns[7].Width = 60;
                    dataGridView1.Columns[8].HeaderText = "V.";
                    dataGridView1.Columns[8].Width = 30;
                    dataGridView1.Columns[9].HeaderText = "Məbləğ";
                    dataGridView1.Columns[9].Width = 100;
            }

            //}
            //catch (Exception)
            //{
            //}
            //finally { }
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            axtar();
            button2.Visible = false;
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
                    //txt_giris.SelectAll();
                }
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
                    //txt_cixis.SelectAll();
                }
            }
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

        private void txt_cixis_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                txtmebleg.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txtmebleg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txt_giris_Enter(object sender, EventArgs e)
        {
            if (txt_giris.Text == "dd-mm-yyyy")
            {
                txt_giris.Text = "";
            }
            txt_giris.ForeColor = Color.Black;
        }

        private void txt_cixis_Enter(object sender, EventArgs e)
        {
            if (txt_cixis.Text == "dd-mm-yyyy")
            {
                txt_cixis.Text = "";
            }
            txt_cixis.ForeColor = Color.Black;
        }

        private void txtmebleg_Enter(object sender, EventArgs e)
        {
            txtmebleg.Text = "";
            txtmebleg.ForeColor = Color.Black;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            axtar();
            button2.Visible = false;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            hesab = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            string value = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            DateTime dateTime = DateTime.Parse(value);
            tarix = dateTime.ToString("dd-MM-yyyy");
            button2.Visible = true;
             axtaretrafli();
            //}
            //catch (Exception)
            //{
            //}
            //finally { }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void txtmebleg_KeyPress(object sender, KeyPressEventArgs e)
        {
            //// Girilen karakterin sayı olup olmadığını ve pozitif sayı olup olmadığını kontrol et
            //if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            //{
            //    e.Handled = true; // Geçersiz karakteri engelle
            //}

            //// "." karakterine izin ver, ancak sadece bir kere kullanılabilir ve sadece sayıdan sonra gelmeli
            //if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
            //{
            //    e.Handled = true;
            //}

            //// Eğer ilk karakter "-" ise, sadece birinci karakter olabilir
            //if (e.KeyChar == '-' && (sender as TextBox).Text.Length > 0)
            //{
            //    e.Handled = true;
            //}
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void txt_giris_Leave(object sender, EventArgs e)
        {
            
        }
    }
}
