using BMI.Muhasibat;
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

namespace BMI.Emeliyyat
{
    public partial class frm_cari_hesab_melumati : Form
    {

        public frm_cari_hesab_melumati()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        //private cl_yanasmalar my_cl = new cl_yanasmalar();
        string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";


        private void excel()
        {
            DataTable dt_med_mex = new DataTable();
            DataTable dt_qaliqlar = new DataTable();
            string mexaric = "select " +
                    "case " +
                    "when r.finsektor = 1 then 'Maliyyə sektoru' " +
                    "when r.yurik = 1  then 'Hüquqi şəxs' " +
                    "when r.predprinimatel = 1 then 'Sahibkar' " +
                    "when r.fizik = 1 then 'Fiziki şəxs' end nov, " +
                    "round(sum(d.summa_v_inval), 2) val_meb, " +
                    "round(sum(d.summa_v_nacval), 2) meb, " +
                    "count(d.debet)" +
                    "from arh_dd d, regnom r,balschkli b " +
                    "where substr(d.debet,10,6)= r.regnom " +
                    "and substr(d.kredit,1,5) not in ('66220', '86220') " +
                    "and d.date_oper BETWEEN TO_DATE('01-02-2024', 'dd/mm/yyyy') AND TO_DATE('13-02-2024', 'dd/mm/yyyy') " +
                    "and substr(d.debet,1,5)= b.balsch " +
                    "group by " +
                    "case " +
                    "when r.finsektor = 1 then 'Maliyyə sektoru' " +
                    "when r.yurik = 1  then 'Hüquqi şəxs' " +
                    "when r.predprinimatel = 1 then 'Sahibkar' " +
                    "when r.fizik = 1 then 'Fiziki şəxs' end";
            string medaxil = "select "+
                    "case "+
                    "when r.finsektor = 1 then 'Maliyyə sektoru' " +
                    "when r.yurik = 1  then 'Hüquqi şəxs' " +
                    "when r.predprinimatel = 1 then 'Sahibkar' " +
                    "when r.fizik = 1 then 'Fiziki şəxs' end nov, " +
                    "round(sum(d.summa_v_inval), 2) val_meb, " +
                    "round(sum(d.summa_v_nacval), 2) meb, " +
                    "count(d.debet)" +
                    "from arh_dd d, regnom r,balschkli b "+
                    "where substr(d.kredit,10,6)= r.regnom "+
                    "and substr(d.debet,1,5) not in ('66220', '86220') "+
                    "and d.date_oper BETWEEN TO_DATE('01-02-2024', 'dd/mm/yyyy') AND TO_DATE('13-02-2024', 'dd/mm/yyyy') "+
                    "and substr(d.kredit,1,5)= b.balsch "+
                    "group by "+
                    "case " +
                    "when r.finsektor = 1 then 'Maliyyə sektoru' " +
                    "when r.yurik = 1  then 'Hüquqi şəxs' " +
                    "when r.predprinimatel = 1 then 'Sahibkar' " +
                    "when r.fizik = 1 then 'Fiziki şəxs' end";

            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                if (checkBox1.Checked==true)
                {
                    using (OracleCommand command = new OracleCommand(medaxil, connection))
                    {
                        connection.Open();
                        OracleDataAdapter adapter = new OracleDataAdapter(command);
                        adapter.Fill(dt_med_mex);
                        dataGridView1.DataSource = dt_med_mex;
                        
                    }
                }
                else
                {
                    using (OracleCommand command = new OracleCommand(mexaric, connection))
                    {
                        connection.Open();
                        OracleDataAdapter adapter = new OracleDataAdapter(command);
                        adapter.Fill(dt_med_mex);
                        dataGridView1.DataSource = dt_med_mex;

                    }
                }
                dataGridView1.Columns[0].HeaderText = "Müştərilər";
                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].HeaderText = "Valyuta məbləğində";
                dataGridView1.Columns[1].DefaultCellStyle.Format = "N2";
                dataGridView1.Columns[1].Width = 140;
                dataGridView1.Columns[2].HeaderText = "Manat məbləğində";
                dataGridView1.Columns[2].DefaultCellStyle.Format = "N2";
                dataGridView1.Columns[2].Width = 140;
                dataGridView1.Columns[3].HeaderText = "Say";
                dataGridView1.Columns[3].Width = 67;
            }
            }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {

                lblmedmex.Text = "Mədaxil";
                checkBox2.Checked = false;
                excel();
            }
            if (checkBox1.Checked == false)
            {
                lblmedmex.Text = "Məxaric";
                checkBox2.Checked = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                lblmedmex.Text = "Məxaric";
                checkBox1.Checked = false;
                excel();
            }
            if (checkBox2.Checked == false)
            {
                lblmedmex.Text = "Mədaxil";
                checkBox1.Checked = true;
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel();
        }
    }
}
