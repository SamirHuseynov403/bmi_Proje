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
using DevExpress.DataProcessing.InMemoryDataProcessor;
using DevExpress.XtraEditors;

namespace BMI
{
    public partial class frm_odenis_ferqlisilinme : Form
    {
        public frm_odenis_ferqlisilinme()
        {
            InitializeComponent();
        }
        DataTable Ordt = new DataTable();
        void axtar()
        {
            //try
            //{
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                OracleCommand Orcom = new OracleCommand("SELECT m.kredit, m.mab1, "+
                " (m.mab1 - NVL(n.mab2, 0) - NVL(e.mab4, 0)) AS farq "+
                " FROM(SELECT t.kredit, SUM(t.summa_v_nacval) mab1 " +
                " FROM odb.docdna t, odb.balschkli k WHERE k.balsch = SUBSTR(t.kredit, 1, 5) " +
                " GROUP BY t.kredit) m, (SELECT t.debet, SUM(NVL(t.summa_v_nacval, 0)) mab2 " +
                " FROM odb.docdna t, odb.balschkli k WHERE k.balsch = SUBSTR(t.debet, 1, 5) " +
                " GROUP BY t.debet) n,(SELECT t.debet, SUM(NVL(t.summa_v_nacval, 0)) mab4 " +
                " FROM odb.view_tek_docdna t JOIN odb.balschkli k ON k.balsch = SUBSTR(t.debet, 1, 5) where t.otvet_ispoln = 0 " +
                " GROUP BY t.debet) e,(SELECT t.debet, SUM(t.summa_v_nacval) mab3 FROM odb.doccom t " +
                " WHERE SUBSTR(t.debet,1,1) = '4' AND SUBSTR(t.kredit,1,2) = '67' " +
                " GROUP BY t.debet) x WHERE m.kredit = n.debet(+) " +
                " AND m.kredit = x.debet(+)AND m.kredit = e.debet(+) " +
                " AND((m.mab1 - NVL(n.mab2, 0) - NVL(e.mab4, 0)) <> 0) " +
                " AND SUBSTR(m.kredit,16,2) <> '91'", Orcon);

                OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
                dataGridView1.Columns[0].HeaderText = "Hesab no No";
                dataGridView1.Columns[1].HeaderText = "Ödəniş";
                dataGridView1.Columns[2].HeaderText = "Fərq";
            //}
            //catch (Exception)
            //{
            //}
            //finally { }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            axtar();
            int sutunSayısı = dataGridView1.RowCount-1;

            // Label kontrolüne sütun sayısını atayın
            label1.Text = "Sətir Sayı: " + sutunSayısı.ToString();
        }

        private void frm_odenis_ferqlisilinme_Load(object sender, EventArgs e)
        {
            
        }
    }
}
