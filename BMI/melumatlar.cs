using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI
{
    class melumatlar
    {
        public int cavab;
        public void test()
        {
            cavab = 5;
        }
        OleDbConnection baglan = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI\BMI\Duzelisler.accdb");
        public float yas_hedd, hey_sig, x_h, avto12, avto18, avto24, avto36, avto48, avto60, men12, men18, men24, men36, men48, men60;
        public string mudir, kr_dep_reis, adam;

        public void yaz()
        {
            baglan.Open();
            OleDbCommand komut = new OleDbCommand("Select ID, yas_heddi,heyat_sig,xidmet_haqq,avto_sig_12ay,avto_sig_18ay,avto_sig_24ay,avto_sig_36ay,avto_sig_48ay,avto_sig_60ay,menzil_sig_12ay,menzil_sig_18ay,menzil_sig_24ay,menzil_sig_36ay,menzil_sig_48ay,menzil_sig_60ay,Mudir,Kr_dep_reisi from Duzelisler where ID=1", baglan);
            OleDbDataReader dr = komut.ExecuteReader();
            while (dr.Read())
            {
                yas_hedd = Convert.ToSingle(dr["yas_heddi"]);
                hey_sig = Convert.ToSingle(dr["heyat_sig"]);
                x_h = Convert.ToSingle(dr["xidmet_haqq"]);
                avto12 = Convert.ToSingle(dr["avto_sig_12ay"]);
                avto18 = Convert.ToSingle(dr["avto_sig_18ay"]);
                avto24 = Convert.ToSingle(dr["avto_sig_24ay"]);
                avto36 = Convert.ToSingle(dr["avto_sig_36ay"]);
                avto48 = Convert.ToSingle(dr["avto_sig_48ay"]);
                avto60 = Convert.ToSingle(dr["avto_sig_60ay"]);
                men12 = Convert.ToSingle(dr["menzil_sig_12ay"]);
                men18 = Convert.ToSingle(dr["menzil_sig_18ay"]);
                men24 = Convert.ToSingle(dr["menzil_sig_24ay"]);
                men36 = Convert.ToSingle(dr["menzil_sig_36ay"]);
                men48 = Convert.ToSingle(dr["menzil_sig_48ay"]);
                men60 = Convert.ToSingle(dr["menzil_sig_60ay"]);
                mudir = dr["Mudir"].ToString();
                kr_dep_reis = dr["Kr_dep_reisi"].ToString();
            }

            baglan.Close();
        }
    }
}
