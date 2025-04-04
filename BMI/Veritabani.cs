using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Windows.Forms;
using Oracle.DataAccess.Client;



namespace BMI
{
    class Veritabani
    {
        public static OracleConnection OrConnect = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass;");
        //Data Source=SAMIR-PC\SQLEXPRESS;Initial Catalog=bmi;Integrated Security=True
        //Data Source=BMI;User ID=FOXPRO;Password=pass
        //oracle adlarini yaz bu sqldi
        public readonly string zamarayis = @"C:\BMI\borcalan arayis.docx";
        public readonly string saveyolu = @"\\fs\KRED_SOB\1-КРЕДИТ-2022\Atlas Cars gir cix";

        OleDbConnection baglan = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI\BMI\bin\Debug\Duzelisler.accdb");
        public int yas_hedd, hey_sig, x_h, avto12, avto18, avto24, avto36, avto48, avto60, men12, men18, men24, men36, men48, men60;
        public string mudir, kr_dep_reis,adam;
        public string Currentuser = string.Empty;

        public void testc()
        {
            adam = "samir";
        }
        public void yaz()
        {
            baglan.Open();
            OleDbCommand komut = new OleDbCommand("Select yas_heddi,heyat_sig,xidmet_haqq,avto_sig_12ay,avto_sig_18ay,avto_sig_24ay,avto_sig_36ay,avto_sig_48ay,avto_sig_60ay,menzil_sig_12ay,menzil_sig_18ay,menzil_sig_24ay,menzil_sig_36ay,menzil_sig_48ay,menzil_sig_60ay,Mudir,Kr_dep_reisi from Duzelisler where ID=1", baglan);
            OleDbDataReader dr = komut.ExecuteReader();
            while (dr.Read())
            {
                yas_hedd=Convert.ToInt32( dr["yas_heddi"]);
                hey_sig = Convert.ToInt32(dr["heyat_sig"]);
                x_h = Convert.ToInt32(dr["xidmet_haqq"]);
                avto12 = Convert.ToInt32(dr["avto_sig_12ay"]);
                avto18 = Convert.ToInt32(dr["avto_sig_18ay"]);
                avto24 = Convert.ToInt32(dr["avto_sig_24ay"]);
                avto36 = Convert.ToInt32(dr["avto_sig_36ay"]);
                avto48 = Convert.ToInt32(dr["avto_sig_48ay"]);
                avto60 = Convert.ToInt32(dr["avto_sig_60ay"]);
                men12 = Convert.ToInt32(dr["menzil_sig_12ay"]);
                men18 = Convert.ToInt32(dr["menzil_sig_18ay"]);
                men24 = Convert.ToInt32(dr["menzil_sig_24ay"]);
                men36 = Convert.ToInt32(dr["menzil_sig_36ay"]);
                men48 = Convert.ToInt32(dr["menzil_sig_48ay"]);
                men60 = Convert.ToInt32(dr["menzil_sig_60ay"]);
                mudir = dr["Mudir"].ToString();
                kr_dep_reis = dr["Kr_dep_reisi"].ToString();
            }
            
            baglan.Close();
        }
    }
}
