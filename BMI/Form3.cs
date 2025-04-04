using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;

namespace BMI
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        
        private void Form3_Load(object sender, EventArgs e)
        {
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand Orcom = new OracleCommand("select t.licschkre,t.subschkre sk,substr(t.licschkre,10,6)||t.subschkre kod,r.name_regnom,r.telefon,r.mobilniy,r.passport,r.senedi_veren_orqaninin_adi ver_orq,r.senedin_verilme_tarixi ver_tar,t.summa,t.procstavkre fz,t.date_open,g.guarantee_name,g.guarantee_id from odb.licschkre t,odb.regnom r,odb.creditinfoguarantee g where t.date_open=to_date('22/06/2021', 'dd/mm/yyyy') and substr(t.licschkre,10,6)=r.regnom and t.licschkre=g.licschkre and t.subschkre=g.subschkre", Orcon);

            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();
        }
    }
}
