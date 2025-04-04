using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class Ged_hev_duzelis : Form
    {
        public Ged_hev_duzelis()
        {
            InitializeComponent();
        }

        public string Xmgeydiyyatnom = string.Empty;
        public string Xmgonderilenyer = string.Empty;
        public string Xaricmektbtarix = string.Empty;
        public string Xmgisamezmun = string.Empty;
        public string Xmmektubmetn = string.Empty;

        public string il = string.Empty;
        int kecencavab = 0;
        public BindingSource mybing;

        private void button1_Click(object sender, EventArgs e) //duzelis etmek  odb nin adini duzelt
        {
        //    try
        //    {
        //        string connectrionString = "Data Source=BMI;User ID=FOXPRO;Password=pass";
        //        OracleConnection connection = new OracleConnection(connectrionString);
        //        connection.Open();
        //        OracleCommand komut = new OracleCommand("Update odb.xaric_mektub x set x.tarix = TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), x.gon_yer='" + Xmgonderilenyer + "', x.qisa_mez='" + Xmgisamezmun + "', x.mektub_metn='" + Xmmektubmetn + "' where x.qey_nom='" + Xmgeydiyyatnom + "'", connection);
        //        kecencavab = komut.ExecuteNonQuery();

        //        if (kecencavab > 0)
        //        {
        //            MessageBox.Show("Məktuba Düzəliş olundu...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //        else
        //        {
        //            MessageBox.Show("Məktuba Düzəliş olunmadı... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        }
        //        connection.Close();
        //    }
        //    catch (Exception ex)
        //    {
                
        //        MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
            
        }
    }
}

