using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class AML_elaqeli_sexs : Form
    {
        public AML_elaqeli_sexs()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        private void axtar()
        {
            DataTable Ordt = new DataTable();
            DataTable Ordtdebet = new DataTable();
            Ordt.Clear();
            Ordtdebet.Clear();
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleCommand Orcom = new OracleCommand("select distinct substr(b.date_oper, 0,10) , b.debet, b.kredit,b.summa_v_inval, b.summa_v_nacval," +
                " b.primechanie,d.name_debet,d.inn_debet,d.name_credit, d.inn_credit from arh_dd a, arh_dd b, doc_vnesh_nacval d where " +
                "a.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and a.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy') and substr(a.debet, 10,6)= '" + textBox1.Text + "' " +
                " and  b.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and b.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')and substr(b.debet, 10,6)= '" + textBox2.Text + "' " +
                " and substr(a.kredit, 10,6) = substr(b.kredit, 10,6) and substr(a.kredit, 0,1)!='6'and substr(b.kredit, 0,1)!='6'and substr(a.kredit, 0,5)!='35025'and substr(b.kredit, 0,5)!='35025'and " +
                "substr(a.kredit, 0,5)!='15025'and substr(b.kredit, 0,5)!='15025'and substr(a.kredit, 0,5)!='45019'and substr(b.kredit, 0,5)!='45019'and substr(a.kredit, 0,5)!='25019'and substr(b.kredit, 0,5)!='25019'" +
                "and substr(a.kredit, 10,6)!='000001'and substr(b.kredit, 10,6)!='000001'and substr(a.kredit, 0,5)!='45105'and substr(b.kredit, 0,5)!='45105'and substr(a.kredit, 0,5)!='45023'and substr(b.kredit, 0,5)!='45023'" +
                "and substr(a.kredit, 0,5)!='45011'and substr(b.kredit, 0,5)!='45011' and a.debet=d.debet AND B.date_oper='04.03.2022' and d.date_oper='04.03.2022'", Orcon);
            OracleCommand Orcomdebet = new OracleCommand("select distinct substr(a.date_oper, 0,10) , a.debet, a.kredit,a.summa_v_inval, a.summa_v_nacval, a.primechanie from arh_dd a, arh_dd b where a.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and a.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy') and substr(a.debet, 10,6)= '"+textBox1.Text+"'  and  b.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and b.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')and substr(b.debet, 10,6)= '" + textBox2.Text + "'  and substr(a.kredit, 10,6) = substr(b.kredit, 10,6) and substr(a.kredit, 0,1)!='6'and substr(b.kredit, 0,1)!='6'and substr(a.kredit, 0,5)!='35025'and substr(b.kredit, 0,5)!='35025'and substr(a.kredit, 0,5)!='15025'and substr(b.kredit, 0,5)!='15025'and substr(a.kredit, 0,5)!='45019'and substr(b.kredit, 0,5)!='45019'and substr(a.kredit, 0,5)!='25019'and substr(b.kredit, 0,5)!='25019'and substr(a.kredit, 10,6)!='000001'and substr(b.kredit, 10,6)!='000001'and substr(a.kredit, 0,5)!='45105'and substr(b.kredit, 0,5)!='45105'and substr(a.kredit, 0,5)!='45023'and substr(b.kredit, 0,5)!='45023'and substr(a.kredit, 0,5)!='45011'and substr(b.kredit, 0,5)!='45011'", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            OracleDataAdapter Ordadebet = new OracleDataAdapter(Orcomdebet);
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Ordadebet.Fill(Ordtdebet);
            dataGridView2.DataSource = Ordtdebet;
            Orcon.Close();

            dataGridView1.Columns[0].HeaderText = "Tarix";
            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].HeaderText = "Debet";
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].HeaderText = "Kredit";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].HeaderText = "Valyuta";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "AZN";
            dataGridView1.Columns[4].Width = 100;
            dataGridView1.Columns[5].HeaderText = "Təyinat";
            dataGridView1.Columns[5].Width = 500;

            dataGridView2.Columns[0].HeaderText = "Tarix";
            dataGridView2.Columns[0].Width = 100;
            dataGridView2.Columns[1].HeaderText = "Debet";
            dataGridView2.Columns[1].Width = 150;
            dataGridView2.Columns[2].HeaderText = "Kredit";
            dataGridView2.Columns[2].Width = 150;
            dataGridView2.Columns[3].HeaderText = "Valyuta";
            dataGridView2.Columns[3].Width = 100;
            dataGridView2.Columns[4].HeaderText = "AZN";
            dataGridView2.Columns[4].Width = 100;
            dataGridView2.Columns[5].HeaderText = "Təyinat";
            dataGridView2.Columns[5].Width = 500;
            //button1.Text = "Sorğula";
        }
        private void axtarhesdehese()
        {
            DataTable Ordt = new DataTable();
            DataTable Ordtdebet = new DataTable();
            Ordt.Clear();
            Ordtdebet.Clear();
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleCommand Orcom = new OracleCommand("select substr(a.date_oper, 0,10) , a.debet, a.kredit,a.summa_v_inval, a.summa_v_nacval, a.primechanie from arh_dd a where a.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and a.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy') and substr(a.debet, 10,6)= '" + textBox1.Text + "' and  substr(a.kredit, 10,6) = '" + textBox2.Text + "' ", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            dataGridView2.DataSource = Ordtdebet;
            Orcon.Close();

            dataGridView1.Columns[0].HeaderText = "Tarix";
            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].HeaderText = "Debet";
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].HeaderText = "Kredit";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].HeaderText = "Valyuta";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "AZN";
            dataGridView1.Columns[4].Width = 100;
            dataGridView1.Columns[5].HeaderText = "Təyinat";
            dataGridView1.Columns[5].Width = 500;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked==true)
            {
                axtarhesdehese();
            }
            if (radioButton2.Checked==true)
            {
                axtar();
            }
            if (radioButton3.Checked == true)
            {
                axtarkredit();
            }
            button1.Text = "Sorğula";
        }
        private void axtarkredit()
        {
            DataTable Ordt = new DataTable();
            DataTable Ordtdebet = new DataTable();
            Ordt.Clear();
            Ordtdebet.Clear();
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleCommand Orcom = new OracleCommand("select distinct substr(b.date_oper, 0,10), b.debet, b.kredit,b.summa_v_inval, b.summa_v_nacval, b.primechanie from arh_dd a, arh_dd b where a.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and a.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy') and substr(a.kredit, 10,6)= '" + textBox1.Text + "'  and  b.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and b.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')and substr(b.kredit, 10,6)= '" + textBox2.Text + "'  and substr(a.debet, 10,6) = substr(b.debet, 10,6) and substr(a.kredit, 0,1)!='6'and substr(b.kredit, 0,1)!='6'and substr(a.kredit, 0,5)!='35025'and substr(b.kredit, 0,5)!='35025'and substr(a.kredit, 0,5)!='15025'and substr(b.kredit, 0,5)!='15025'and substr(a.kredit, 0,5)!='45019'and substr(b.kredit, 0,5)!='45019'and substr(a.kredit, 0,5)!='25019'and substr(b.kredit, 0,5)!='25019'and substr(a.kredit, 10,6)!='000001'and substr(b.kredit, 10,6)!='000001' and substr(a.kredit, 0,5)!='45105'and substr(b.kredit, 0,5)!='45105'and substr(a.kredit, 0,5)!='45023'and substr(b.kredit, 0,5)!='45023'and substr(a.kredit, 0,5)!='45011'and substr(b.kredit, 0,5)!='45011'", Orcon);
            OracleCommand Orcomdebet = new OracleCommand("select distinct substr(a.date_oper, 0,10), a.debet, a.kredit,a.summa_v_inval, a.summa_v_nacval, a.primechanie from arh_dd a, arh_dd b where a.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and a.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy') and substr(a.kredit, 10,6)= '" + textBox1.Text + "'  and  b.date_oper >= to_date('" + dateTimePicker1.Text + "','dd-MM-yyyy') and b.date_oper<= to_date('" + dateTimePicker2.Text + "','dd-MM-yyyy')and substr(b.kredit, 10,6)= '" + textBox2.Text + "'  and substr(a.debet, 10,6) = substr(b.debet, 10,6) and substr(a.kredit, 0,1)!='6'and substr(b.kredit, 0,1)!='6'and substr(a.kredit, 0,5)!='35025'and substr(b.kredit, 0,5)!='35025'and substr(a.kredit, 0,5)!='15025'and substr(b.kredit, 0,5)!='15025'and substr(a.kredit, 0,5)!='45019'and substr(b.kredit, 0,5)!='45019'and substr(a.kredit, 0,5)!='25019'and substr(b.kredit, 0,5)!='25019'and substr(a.kredit, 10,6)!='000001'and substr(b.kredit, 10,6)!='000001' and substr(a.kredit, 0,5)!='45105'and substr(b.kredit, 0,5)!='45105'and substr(a.kredit, 0,5)!='45023'and substr(b.kredit, 0,5)!='45023'and substr(a.kredit, 0,5)!='45011'and substr(b.kredit, 0,5)!='45011'", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
            OracleDataAdapter Ordadebet = new OracleDataAdapter(Orcomdebet);

            Orda.Fill(Ordt);
            dataGridView2.DataSource = Ordt;
            Ordadebet.Fill(Ordtdebet);
            dataGridView1.DataSource = Ordtdebet;
            Orcon.Close();

            dataGridView1.Columns[0].HeaderText = "Tarix";
            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].HeaderText = "Debet";
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].HeaderText = "Kredit";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[3].HeaderText = "Valyuta";
            dataGridView1.Columns[3].Width = 100;
            dataGridView1.Columns[4].HeaderText = "AZN";
            dataGridView1.Columns[4].Width = 100;
            dataGridView1.Columns[5].HeaderText = "Təyinat";
            dataGridView1.Columns[5].Width = 500;

            dataGridView2.Columns[0].HeaderText = "Tarix";
            dataGridView2.Columns[0].Width = 100;
            dataGridView2.Columns[1].HeaderText = "Debet";
            dataGridView2.Columns[1].Width = 150;
            dataGridView2.Columns[2].HeaderText = "Kredit";
            dataGridView2.Columns[2].Width = 150;
            dataGridView2.Columns[3].HeaderText = "Valyuta";
            dataGridView2.Columns[3].Width = 100;
            dataGridView2.Columns[4].HeaderText = "AZN";
            dataGridView2.Columns[4].Width = 100;
            dataGridView2.Columns[5].HeaderText = "Təyinat";
            dataGridView2.Columns[5].Width = 500;
        }
        
        private void exceleattelebe()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "AML";
            // storing header part in Excel  
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
        }
        private void exceleattelebe1()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "AML";
            // storing header part in Excel  
            
            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            exceleattelebe();
            exceleattelebe1();
        }
    }
}
