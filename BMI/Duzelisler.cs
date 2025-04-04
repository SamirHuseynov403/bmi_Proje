using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Threading;

namespace BMI
{
    public partial class Duzelisler : Form
    {
        

        public Duzelisler()
        {
            InitializeComponent();
        }
        public Oraclebaglanti orabag;
        

        //OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\bin\\Debug\\Duzelisler.accdb");
        OleDbConnection baglan = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\BMI\BMI\Duzelisler.accdb");
        private void groupBox1_Enter(object sender, EventArgs e)
        {
            
        }

        private void Duzelisler_Load(object sender, EventArgs e)
        {
            melumatlar mlmt = new melumatlar();
            mlmt.yaz();
            yas_hed.Text = mlmt.yas_hedd.ToString();
            hey_sig.Text = mlmt.hey_sig.ToString();
            x_h.Text = mlmt.x_h.ToString();
            avto12.Text = mlmt.avto12.ToString();
            avto18.Text = mlmt.avto18.ToString();
            avto24.Text = mlmt.avto24.ToString();
            avto36.Text = mlmt.avto36.ToString();
            avto48.Text = mlmt.avto48.ToString();
            avto60.Text = mlmt.avto60.ToString();
            men12.Text = mlmt.men12.ToString();
            men18.Text = mlmt.men18.ToString();
            men24.Text = mlmt.men24.ToString();
            men36.Text = mlmt.men36.ToString();
            men48.Text = mlmt.men48.ToString();
            men60.Text = mlmt.men60.ToString();
            mudir.Text = mlmt.mudir;
            kr_dep_reis.Text = mlmt.kr_dep_reis;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            baglan.Open();
            OleDbCommand cmd = baglan.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "update Duzelisler Set yas_heddi='" + yas_hed.Text + "',heyat_sig='" + hey_sig.Text + "',xidmet_haqq='" + x_h.Text + "',avto_sig_12ay='" + avto12.Text + "',avto_sig_18ay='" + avto18.Text + "',avto_sig_24ay='" + avto24.Text + "',avto_sig_36ay='" + avto36.Text + "',avto_sig_48ay='" + avto48.Text + "',avto_sig_60ay='" + avto60.Text + "',menzil_sig_12ay='" + men12.Text + "',menzil_sig_18ay='" + men18.Text + "',menzil_sig_24ay='" + men24.Text + "',menzil_sig_36ay='" + men36.Text + "',menzil_sig_48ay='" + men48.Text + "',menzil_sig_60ay='" + men60.Text + "',Mudir='" + mudir.Text + "',Kr_dep_reisi='" + kr_dep_reis.Text + "' where ID=1";
            cmd.ExecuteNonQuery();
            baglan.Close();
            this.Close();
            //OleDbCommand komut = new OleDbCommand("Update Duzelisler Set ()values() where ID=1", baglan);
            //komut.ExecuteNonQuery();
            //
            //
            //baglan.Close();
            //this.Close();
        }//
        //
    }//yas_heddi,heyat_sig%,xidmet_haqq%,avto_sig_12ay%,avto_sig_18ay%,avto_sig_24ay%,avto_sig_36ay%,avto_sig_48ay%,avto_sig_60ay%,menzil_sig_12ay%,menzil_sig_18ay%,menzil_sig_24ay%,menzil_sig_36ay%,menzil_sig_48ay%,menzil_sig_60ay%,Mudir,Kr_dep_reisi
}
