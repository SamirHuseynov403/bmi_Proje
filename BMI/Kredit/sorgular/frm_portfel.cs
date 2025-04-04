using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI.Kredit
{
    public partial class frm_portfel : Form
    {
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;
        string projeanasehife = AppDomain.CurrentDomain.BaseDirectory;
        string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
        public frm_portfel()
        {
            InitializeComponent();
        }

        private void excel()
        {
            string kr_portfel_TIP = "select t.name," +
                "sum(l.summakre * ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6))ver," +
                "sum(l.summa * ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6)) qaliq," +
                "sum(l.summa_19 * ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6)) vk_qaliq,count(l.licschkre) say " +
                "from arh_licschkre l ,tipkre t "+
               " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and t.code = l.tipkredita and l.date_close is null" +
               " and l.tipkredita = t.code "+
              "  group by t.name";

            string kr_portfel_ZALOQ = "select t.name,sum(l.summakre* ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6)) ekvver, "+
            "sum(l.summa * ROUND(odb.func_get_kurval(substr(l.licschkre, 6, 2), l.date_oper), 6)) qaliq, " +
            "sum(l.summa_19 * ROUND(odb.func_get_kurval(substr(l.licschkre, 6, 2), l.date_oper), 6)) vk_qaliq,count(l.licschkre) say from arh_licschkre l, tipzal t " +
            "where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and l.date_close is null " +
            "and l.tipzaloga = t.code " +
            "group by t.name";

            string kr_portfel_SEKTOR = "select f.name,tp.name,t.name,i.name_index_otrasli," +
                "sum(l.summakre * ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6)) ekv , "+
            "sum(l.summa * ROUND(odb.func_get_kurval(substr(l.licschkre, 6, 2), l.date_oper), 6)) ekv_qaliq, " +
            "sum(l.summa_19 * ROUND(odb.func_get_kurval(substr(l.licschkre, 6, 2), l.date_oper), 6)) ekv_vk_qaliq,count(l.licschkre),substr(l.licschkre, 6, 2) val " +
            "from arh_licschkre l,faaliyyat_sektoru f, regnom r,tipzal t, tipkre tp,index_otrasli i " +
            "where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and l.date_close is null " +
            "and substr(l.licschkre,10,6)= r.regnom and t.code = l.tipkredita and f.code = r.faaliyyat_sektoru and tp.code = l.tipkredita and l.index_otrasli = i.index_otrasli " +
            "group by t.name,f.name,substr(l.licschkre, 6, 2),tp.name,l.index_otrasli,i.name_index_otrasli " +
            "  order by f.name ";

            string kr_portfel_MIN_QALIQ = "select r.name_regnom,l.subschkre,l.licschkre," +
                "(l.summa+l.summa_19)* ROUND(odb.func_get_kurval(substr(l.licschkre,6,2),l.date_oper),6) ekv from arh_licschkre l,regnom r "+
            " where l.date_oper = TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') and l.date_close is null " +
            " and(l.summa + l.summa_19) <= 10 and(l.summa + l.summa_19) <> 0 and substr(l.licschkre,10,6)= r.regnom";

            
            DataTable _dt_portfel_tip = new DataTable();
            DataTable _dt_portfel_zaloq = new DataTable();
            DataTable _dt_portfel_sektor = new DataTable();
            DataTable _dt_portfel_qaliq = new DataTable();
            DataTable _dt_qaliqlar = new DataTable();
            DataTable _dt_likvid = new DataTable();
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(kr_portfel_TIP, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_portfel_tip);
                }
                using (OracleCommand command = new OracleCommand(kr_portfel_ZALOQ, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_portfel_zaloq);
                }
                using (OracleCommand command = new OracleCommand(kr_portfel_SEKTOR, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_portfel_sektor);
                }
                using (OracleCommand command = new OracleCommand(kr_portfel_MIN_QALIQ, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_portfel_qaliq);
                }
                //using (OracleCommand command = new OracleCommand(gunun_qaliqlari, connection))
                //{
                //    //connection.Open();
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(_dt_qaliqlar);
                //}
                //using (OracleCommand command = new OracleCommand(likvidler, connection))
                //{
                //    //connection.Open();
                //    OracleDataAdapter adapter = new OracleDataAdapter(command);
                //    adapter.Fill(_dt_likvid);
                //}
            }
            string baseFileName;
            string d_yolu=lbld_yolu.Text;
            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string d_yoluna_baglan ="";
            if (d_yolu== "tip")
            {
                d_yoluna_baglan = @"C:\BMI_\Kr_portfel_tip.xlsx";

            }
            else if (d_yolu == "zaloq")
            {
                d_yoluna_baglan = @"C:\BMI_\Kr_portfel_zaloq.xlsx";
            }
            else if (d_yolu == "sektor")
            {
                d_yoluna_baglan = @"C:\BMI_\Kr_portfel_sektor.xlsx";
            }
            else if (d_yolu == "qaliq")
            {
                d_yoluna_baglan = @"C:\BMI_\Kr_portfel_10_az.xlsx";
            }

            baseFileName = "Kredit sorğu" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = d_yoluna_baglan;
            string filePath = Path.Combine(dosyayolu, fileName);

            if (File.Exists(Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);

            if (d_yolu == "tip")
            {
                using (ExcelPackage package = new ExcelPackage(templateFile))
                {
                    ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Məlumat"];
                    int rowCount = _dt_portfel_tip.Rows.Count;

                    int startRow = 6; // Başlangıç satırı
                    int startColumn = 1; // Başlangıç sütunu

                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < _dt_portfel_tip.Columns.Count; j++)
                        {
                            double toplam1 = 0;
                            double toplam2 = 0;
                            double toplam3 = 0;
                            double toplam4 = 0;
                            foreach (DataRow row in _dt_portfel_tip.Rows)
                            {
                                toplam1 += Convert.ToDouble(row[1]);
                                toplam2 += Convert.ToDouble(row[2]);
                                toplam3 += Convert.ToDouble(row[3]);
                                toplam4 += Convert.ToDouble(row[4]);
                            }
                            worksheet1.Cells[startRow + i, startColumn + j].Value = _dt_portfel_tip.Rows[i][j];
                            worksheet1.Cells[startRow + i, startColumn + j].Style.Numberformat.Format = "#,##0.00";
                            if (i + 1 == rowCount)
                            {
                                int endColumn = 5; // 5. sütuna kadar olan hücreleri belirtmek için son sütun indeksi
                                ExcelRange range = worksheet1.Cells[startRow + i + 1, 1, startRow + i + 1, endColumn];
                                worksheet1.Cells[startRow + i + 1, 1].Value = "Toplam";
                                worksheet1.Cells[startRow + i + 1, 2].Value = toplam1;
                                worksheet1.Cells[startRow + i + 1, 3].Value = toplam2;
                                worksheet1.Cells[startRow + i + 1, 4].Value = toplam3;
                                worksheet1.Cells[startRow + i + 1, 5].Value = toplam4;
                                range.Style.Numberformat.Format = "#,##0.00";
                            }
                        }
                    }
                    worksheet1.Cells[1, 2].Value = textBox2.Text + " tarix məlumatı";

                    filePath = Path.Combine(dosyayolu, fileName);
                    package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                    System.Diagnostics.Process.Start(filePath);
                }

            }
            else if (d_yolu == "zaloq")
            {
                using (ExcelPackage package = new ExcelPackage(templateFile))
                {
                    ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Məlumat"];
                    int rowCount = _dt_portfel_zaloq.Rows.Count;

                    int startRow = 6; // Başlangıç satırı
                    int startColumn = 1; // Başlangıç sütunu

                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < _dt_portfel_zaloq.Columns.Count; j++)
                        {
                            double toplam1 = 0;
                            double toplam2 = 0;
                            double toplam3 = 0;
                            double toplam4 = 0;
                            foreach (DataRow row in _dt_portfel_zaloq.Rows)
                            {
                                toplam1 += Convert.ToDouble(row[1]);
                                toplam2 += Convert.ToDouble(row[2]);
                                toplam3 += Convert.ToDouble(row[3]);
                                toplam4 += Convert.ToDouble(row[4]);
                            }

                            worksheet1.Cells[startRow + i, startColumn + j].Value = _dt_portfel_zaloq.Rows[i][j];
                            ExcelRange rangefor = worksheet1.Cells[startRow + i, startColumn + j];
                            rangefor.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Font.Size = 13;
                            rangefor.Style.Font.Italic = true;
                            //rangefor.Style.Font.Bold = true;
                            rangefor.Style.Numberformat.Format = "#,##0.00";
                            if (i+1 == rowCount)
                            {

                                int endColumn = 5; // 5. sütuna kadar olan hücreleri belirtmek için son sütun indeksi
                                ExcelRange range = worksheet1.Cells[startRow + i + 1, 1, startRow + i + 1, endColumn];
                                worksheet1.Cells[startRow + i + 1, 1].Value = "Toplam";
                                worksheet1.Cells[startRow + i + 1, 2].Value = toplam1;
                                worksheet1.Cells[startRow + i + 1, 3].Value = toplam2;
                                worksheet1.Cells[startRow + i + 1, 4].Value = toplam3;
                                worksheet1.Cells[startRow + i + 1, 5].Value = toplam4;
                                // Stil özelliklerini belirtilen hücre aralığına uygulayın
                                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                range.Style.Font.Size = 13;
                                range.Style.Font.Bold = true;
                                range.Style.Numberformat.Format = "#,##0.00";
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Gold);

                            }
                            
                        }
                        
                    }
                    worksheet1.Cells["A2"].Value = textBox2.Text + " tarix məlumatı";

                    filePath = Path.Combine(dosyayolu, fileName);
                    package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                    System.Diagnostics.Process.Start(filePath);
                }
            }
            else if (d_yolu == "sektor")
            {
                using (ExcelPackage package = new ExcelPackage(templateFile))
                {
                    ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Məlumat"];
                    int rowCount = _dt_portfel_sektor.Rows.Count;

                    int startRow = 6; // Başlangıç satırı
                    int startColumn = 1; // Başlangıç sütunu

                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < _dt_portfel_sektor.Columns.Count; j++)
                        {
                            double toplam1 = 0;
                            double toplam2 = 0;
                            double toplam3 = 0;
                            double toplam4 = 0;
                            foreach (DataRow row in _dt_portfel_sektor.Rows)
                            {
                                toplam1 += Convert.ToDouble(row[4]);
                                toplam2 += Convert.ToDouble(row[5]);
                                toplam3 += Convert.ToDouble(row[6]);
                                toplam4 += Convert.ToDouble(row[7]);
                            }

                            worksheet1.Cells[startRow + i, startColumn + j].Value = _dt_portfel_sektor.Rows[i][j];
                            ExcelRange rangefor = worksheet1.Cells[startRow + i, startColumn + j];
                            rangefor.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            rangefor.Style.Font.Size = 13;
                            rangefor.Style.Font.Italic = true;
                            //rangefor.Style.Font.Bold = true;
                            rangefor.Style.Numberformat.Format = "#,##0.00";
                            if (i + 1 == rowCount)
                            {

                                int endColumn = 9; // 5. sütuna kadar olan hücreleri belirtmek için son sütun indeksi
                                ExcelRange range = worksheet1.Cells[startRow + i + 1, 1, startRow + i + 1, endColumn];
                                worksheet1.Cells[startRow + i + 1, 4].Value = "Toplam";
                                worksheet1.Cells[startRow + i + 1, 5].Value = toplam1;
                                worksheet1.Cells[startRow + i + 1, 6].Value = toplam2;
                                worksheet1.Cells[startRow + i + 1, 7].Value = toplam3;
                                worksheet1.Cells[startRow + i + 1, 8].Value = toplam4;
                                // Stil özelliklerini belirtilen hücre aralığına uygulayın
                                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                range.Style.Font.Size = 12;
                                range.Style.Font.Bold = true;
                                range.Style.Numberformat.Format = "#,##0.00";
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Gold);

                            }

                        }

                    }
                    worksheet1.Cells["A2"].Value = textBox2.Text + " tarix məlumatı";

                    filePath = Path.Combine(dosyayolu, fileName);
                    package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                    System.Diagnostics.Process.Start(filePath);
                }
            }
            else if (d_yolu == "qaliq")
            {
                using (ExcelPackage package = new ExcelPackage(templateFile))
                {
                    ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Məlumat"];
                    int rowCount = _dt_portfel_qaliq.Rows.Count;

                    int startRow = 6; // Başlangıç satırı
                    int startColumn = 1; // Başlangıç sütunu

                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < _dt_portfel_qaliq.Columns.Count; j++)
                        {
                            //double toplam1 = 0;
                            //double toplam2 = 0;
                            //double toplam3 = 0;
                            double toplam4 = 0;
                            foreach (DataRow row in _dt_portfel_qaliq.Rows)
                            {
                                //toplam1 += Convert.ToDouble(row[0]);
                                //toplam2 += Convert.ToDouble(row[1]);
                                //toplam3 += Convert.ToDouble(row[2]);
                                toplam4 += Convert.ToDouble(row[3]);
                            }
                            worksheet1.Cells[startRow + i, startColumn + j].Value = _dt_portfel_qaliq.Rows[i][j];
                            worksheet1.Cells[startRow + i, startColumn + j].Style.Numberformat.Format = "#,##0.00";
                            if (i + 1 == rowCount)
                            {
                                int endColumn = 5; // 5. sütuna kadar olan hücreleri belirtmek için son sütun indeksi
                                ExcelRange range = worksheet1.Cells[startRow + i + 1, 1, startRow + i + 1, endColumn];
                                //worksheet1.Cells[startRow + i + 1, 1].Value = "Toplam";
                                //worksheet1.Cells[startRow + i + 1, 2].Value = toplam1;
                                //worksheet1.Cells[startRow + i + 1, 3].Value = toplam2;
                                //worksheet1.Cells[startRow + i + 1, 4].Value = toplam4;
                                //worksheet1.Cells[startRow + i + 1, 5].Value = toplam4;
                                range.Style.Numberformat.Format = "#,##0.00";
                            }
                        }
                    }
                    worksheet1.Cells[1, 2].Value = textBox2.Text + " tarix məlumatı";

                    filePath = Path.Combine(dosyayolu, fileName);
                    package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                    System.Diagnostics.Process.Start(filePath);
                }
            }

            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = textBox2.Text;
            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    textBox2.Text = yeniFormatliTarih;
                }
                else
                {
                }
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
                if (e.KeyCode == Keys.Enter)
                {
                    // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                    button1.Focus();
                    e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
                }
        }
    }
}
