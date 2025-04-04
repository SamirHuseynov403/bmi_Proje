using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BMI
{
    public partial class Telebe : Form
    {
        public Telebe()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        private void temizle()
        {
            txtadi.Text = "";
            txthev.Text = "";
            txtpas.Text = "";
            txtref.Text = "";
            txtmeb.Text = "";

        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string adi = "", pasp = "", hevno = "", refno = "", filial, uniad, alanbank, teyinat1 = "", teyinat2 = "", teyinat3 = "";
                string hes35025, hes45023, hes45011, hes67013, tel_kurs;
                double mebleg = 0, kurs, tammebleg, xhaqq, tamxh;
                string bos1 = "";

                hes35025 = txt35025.Text;
                hes45011 = txt45011.Text;
                hes45023 = txt45023.Text;
                hes67013 = txt67013.Text;

                tel_kurs = txbtelkurs.Text;
                adi = txtadi.Text;
                pasp = txtpas.Text;
                mebleg = Convert.ToDouble(txtmeb.Text);
                filial = cmbfilial.Text;
                hevno = txthev.Text;
                refno = txtref.Text;
                uniad = cmbuni.Text;
                kurs = Convert.ToDouble(txtkurs.Text);
                alanbank = cmbalanbank.Text;
                xhaqq = Convert.ToDouble(txtxh.Text);

                tammebleg = mebleg;
                tamxh = tammebleg * kurs * xhaqq / 100;

                if (tammebleg * kurs * xhaqq / 100 >= 0.50)
                {
                    tamxh = tammebleg * kurs * xhaqq / 100;
                }
                else
                    tamxh = 0.5;

                teyinat1 = " G/H " + hevno + " REF " + refno + " BMI " + filial + " " + alanbank + " Tranz. " + uniad + " t/h " + tel_kurs + " -kurs " + adi + " " + pasp;
                teyinat2 = "1USD=" + kurs + " AZN " + " G/H " + hevno + " " + adi + " -" + cmbuni.Text + " kurs " + tel_kurs + " təh/h pasp " + pasp;
                teyinat3 = " G/H " + hevno + " REF " + refno + " BMI " + filial + " " + alanbank + " Tranz. " + uniad + " t/h " + tel_kurs + " -kurs " + adi + " " + pasp + " x/h ";

                dataGridView1.Rows.Add(bos1, bos1, bos1, hes35025, bos1, hes45023, mebleg, bos1, bos1, teyinat1);
                dataGridView1.Rows.Add(bos1, bos1, bos1, hes45023, bos1, hes45011, tammebleg, bos1, bos1, teyinat2);
                dataGridView1.Rows.Add(bos1, bos1, bos1, hes45011, bos1, hes67013, tamxh, bos1, bos1, teyinat3);
                temizle();
            }
            catch (Exception)
            {
                MessageBox.Show("Məbləğ düzgün yazılmayıb");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if
(dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.SelectedRows[0].Index);
            }

            else
            {

                MessageBox.Show("Silinəcək sətiri seçin");

            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            exceleattelebe();
        }
        private void exceleattelebe()
        {
            try
            {
                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopFolder = desktopFolder + "\\TƏLƏBƏ İMPORT.xls";

                if (dataGridView1.RowCount > 0)
                {
                    if (!System.IO.File.Exists(Application.StartupPath + "\\TƏLƏBƏ İMPORT.xls"))
                    {
                        MessageBox.Show("Excel file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Application.StartupPath + "\\TƏLƏBƏ İMPORT.xls");
                        Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                        worksheet.Name = "TƏLƏBƏ İMPORT.xls";
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                                worksheet.Cells[i + 2, j + 1].Font.Name = "A3 Arial Azlat";
                            }
                        }
                        app.Visible = true;
                        workbook.SaveAs(desktopFolder, Type.Missing);
                        app.Quit();
                    }
                }
            }
            catch (Exception)
            {
                
                
            }
            
        }
        

    }
}
