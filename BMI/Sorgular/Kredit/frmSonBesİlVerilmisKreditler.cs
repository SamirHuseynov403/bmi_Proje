using BMI.Muhasibat;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Globalization;

namespace BMI.Sorgular.Kredit
{
    public partial class frmSonBesİlVerilmisKreditler : Form
    {
        public frmSonBesİlVerilmisKreditler()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();

        void excel()
        {
            DataTable _dt_kredit_son_5_verilmis = new DataTable();

            string kredit_son_5_verilmis = @"select TO_CHAR(l.date_open, 'MM-YYYY') AS ay_il, SUM(l.summakre * ROUND(odb.func_get_kurval(SUBSTR(l.licschkre, 6, 2), l.date_oper), 6))meb 
                from arh_licschkre l where
                l.date_oper = (SELECT MAX(c.date_oper)
                FROM odb.calendar c
                WHERE (c.space_or_star IS NULL or c.space_or_star='d') AND c.date_oper < TRUNC(SYSDATE))
                and EXTRACT(YEAR FROM l.date_open) BETWEEN EXTRACT(YEAR FROM SYSDATE) - 9 AND EXTRACT(YEAR FROM SYSDATE)
                GROUP BY TO_CHAR(l.date_open, 'MM-YYYY')
                ORDER BY MIN(l.date_open)";

            int currentYear = DateTime.Now.Year;

            int year1 = currentYear;
            int year2 = currentYear - 1;
            int year3 = currentYear - 2;
            int year4 = currentYear - 3;
            int year5 = currentYear - 4;
            int year6 = currentYear - 5;
            int year7 = currentYear - 6;
            int year8 = currentYear - 7;
            int year9 = currentYear - 8;
            int year10 = currentYear - 9;

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(kredit_son_5_verilmis, connection))
                {

                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_kredit_son_5_verilmis);
                }

                connection.Close();

            }
            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = year5.ToString()+" - "+year1.ToString(); // TextBox'tan alınan metni sakla
                                                                                         //string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = textBoxText + " qaliqlar"; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kredit", "Exceller", "Verilmis kreditler son bes il.xlsx");
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);

            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);
            FileInfo newFile = new FileInfo(filePath);
            templateFile.CopyTo(newFile.FullName, true);

            using (ExcelPackage package = new ExcelPackage(newFile, true))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets["Granted"];

                // İl başlıqları B5, C5, D5, E5, F5, 
                ws.Cells["B5"].Value = year10;
                ws.Cells["C5"].Value = year9;
                ws.Cells["D5"].Value = year8;
                ws.Cells["E5"].Value = year7;
                ws.Cells["F5"].Value = year6;

                ws.Cells["G5"].Value = year5;
                ws.Cells["H5"].Value = year4;
                ws.Cells["I5"].Value = year3;
                ws.Cells["J5"].Value = year2;
                ws.Cells["K5"].Value = year1;

                // Aylara görə datanı yaz
                foreach (DataRow row in _dt_kredit_son_5_verilmis.Rows)
                {
                    string ayIl = row["ay_il"].ToString(); // "04-2025"
                    DateTime dateOper = DateTime.ParseExact(ayIl, "MM-yyyy", CultureInfo.InvariantCulture);
                    int year = dateOper.Year;
                    int month = dateOper.Month;

                    double toplam = row["meb"] == DBNull.Value ? 0 : Convert.ToDouble(row["meb"]);

                    int yearColOffset = (year.ToString() == year10.ToString()) ? 0 :
                                        (year.ToString() == year9.ToString()) ? 1 :
                                        (year.ToString() == year8.ToString()) ? 2 :
                                        (year.ToString() == year7.ToString()) ? 3 :

                                        (year.ToString() == year6.ToString()) ? 4 :
                                        (year.ToString() == year5.ToString()) ? 5 :
                                        (year.ToString() == year4.ToString()) ? 6 :
                                        (year.ToString() == year3.ToString()) ? 7 :
                                        (year.ToString() == year2.ToString()) ? 8 :

                                        (year.ToString() == year1.ToString()) ? 9 : -1;

                    if (yearColOffset != -1)
                    {
                        ws.Cells[6 + (month - 1), 2 + yearColOffset].Value = toplam;
                    }
                }


                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private void frmSonBesİlVerilmisKreditler_Load(object sender, EventArgs e)
        {
            excel();
        }
    }
}
