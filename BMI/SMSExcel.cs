using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BMI
{
    public partial class SMSExcel : Form
    {
        public SMSExcel()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\192.168.0.5\kred_sob\EL VURMA\Kredit.accdb");
        DataTable cedvelSMS = new DataTable();

        private void SMSExcel_Load(object sender, EventArgs e)
        {
            ////SMSExcel smexc = new SMSExcel();
            //baglanti.Open();
            ////cedvel.Clear();
            //OleDbDataAdapter isleme = new OleDbDataAdapter("select Adi_soyadi,Telefon from Verilmis_kr ", baglanti);
            //isleme.Fill(cedvelSMS);
            //dataGridView1.DataSource = cedvelSMS;
            //baglanti.Close();
        }
    }
}
