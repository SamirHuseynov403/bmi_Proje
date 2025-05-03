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
using OfficeOpenXml;
using BMI.Muhasibat;

namespace BMI
{
    public partial class frm_AML_huquqi_hesab_imza : Form
    {
        public frm_AML_huquqi_hesab_imza()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();

        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        private void ExportToExcel(DataGridView dataGridView, string excelFilePath)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                for (int i = 1; i < dataGridView.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        if (dataGridView.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }

                workbook.SaveAs(excelFilePath);
                workbook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veriler Excel dosyasına aktarılamadı: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        private void MatchDataAndAddToExcel(string excelFilePath, DataGridView dataGridView1, DataGridView dataGridView2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(excelFilePath);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                int idColumnIndex1 = dataGridView1.Columns["ID"].Index;
                int idColumnIndex2 = dataGridView2.Columns["ID"].Index;

                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    if (row2.Cells[idColumnIndex2].Value != null)
                    {
                        string idToMatch = row2.Cells[idColumnIndex2].Value.ToString();

                        foreach (DataGridViewRow row1 in dataGridView1.Rows)
                        {
                            if (row1.Cells[idColumnIndex1].Value != null && row1.Cells[idColumnIndex1].Value.ToString() == idToMatch)
                            {
                                for (int i = 0; i < row2.Cells.Count; i++)
                                {
                                    worksheet.Cells[row1.Index + 2, dataGridView1.Columns.Count + i + 1] = row2.Cells[i].Value != null ? row2.Cells[i].Value.ToString() : "";
                                }
                            }
                        }
                    }
                }

                workbook.Save();
                workbook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eşleştirme işlemi sırasında hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        //using OfficeOpenXml; // EPPlus kütüphanesinin namespace'i
        private void MatchAndExportToExcel(DataGridView dataGridView1, DataGridView dataGridView3, string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Veriler");

                // DataGridView1 verilerini başlık satırıyla birlikte ekle
                for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i].Value = dataGridView1.Columns[i - 1].HeaderText;
                }

                // DataGridView1 verilerini ekle
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                int colOffset = dataGridView1.Columns.Count + 1;
                int idColumnIndex1 = dataGridView1.Columns["ID"].Index;
                int idColumnIndex2 = dataGridView2.Columns["ID"].Index;

                int excelbaslangic = 2; // İlk verinin yazılacağı satır

                foreach (DataGridViewRow row1 in dataGridView1.Rows)
                {
                    string idInGridView1 = row1.Cells[idColumnIndex1].Value.ToString();

                    foreach (DataGridViewRow row2 in dataGridView2.Rows)
                    {
                        var cell = row2.Cells[idColumnIndex2];
                        if (cell.Value != null)
                        {
                            string idInGridView2 = cell.Value.ToString();

                            if (idInGridView1 == idInGridView2)
                            {
                                for (int k = 0; k < dataGridView2.Columns.Count; k++)
                                {
                                    worksheet.Cells[excelbaslangic, k + colOffset].Value = row2.Cells[k].Value;
                                }
                                excelbaslangic++;
                            }
                        }
                    }
                }

                package.Save();
            }
        }
        void ekle()
        {
            int idColumnIndex1 = dataGridView1.Columns[0].Index;

            // İkinci DataGridView'deki ID sütununun indeksi
            int idColumnIndex2 = dataGridView3.Columns[0].Index;

            // DataGridView3 için yeni bir DataTable oluşturun
            DataTable dataTable3 = new DataTable();

            // DataGridView1'deki ID sütununu DataGridView3'e ekleyin
            dataTable3.Columns.Add("ID", typeof(string));

            // DataGridView1'deki AD ve Yas sütunlarını DataGridView3'e ekleyin
            dataTable3.Columns.Add("AD1", typeof(string));
            dataTable3.Columns.Add("Yas1", typeof(string));

            // DataGridView2'deki AD ve Yas sütunlarını DataGridView3'e ekleyin
            dataTable3.Columns.Add("AD2", typeof(string));
            dataTable3.Columns.Add("Yas2", typeof(string));

            // DataGridView1'deki her bir satırı döngüye alın
            foreach (DataGridViewRow row1 in dataGridView1.Rows)
            {
                string idToMatch1 = row1.Cells[idColumnIndex1].Value?.ToString();

                if (!string.IsNullOrEmpty(idToMatch1))
                {
                    // ID'ye sahip bir satır oluşturun
                    DataRow newRow = dataTable3.NewRow();
                    newRow[0] = idToMatch1;

                    // İlgili verileri yeni satıra ekleyin
                    newRow["AD1"] = row1.Cells[1].Value?.ToString();
                    //newRow["Yas1"] = row1.Cells["Yas"].Value?.ToString();
                    newRow["Yas1"] = row1.Cells[2].Value?.ToString();

                    // ID'ye sahip verilere karşılık gelen satırları bulun
                    foreach (DataGridViewRow row2 in dataGridView3.Rows)
                    {
                        string idToMatch2 = row2.Cells[idColumnIndex2].Value?.ToString();

                        if (idToMatch1 == idToMatch2)
                        {
                            // ID'ye sahip verileri yeni satır olarak ekleyin
                            newRow["Qeydno"] = row2.Cells[0].Value?.ToString();
                            newRow["Adı"] = row2.Cells[1].Value?.ToString();
                            newRow["VÖEN-i"] = row2.Cells[2].Value?.ToString();
                            newRow["Mülkiyyət növü"] = row2.Cells[3].Value?.ToString();
                            newRow["Təşkilati-hüquqi forması"] = row2.Cells[4].Value?.ToString();
                            newRow["Adı"] = row2.Cells[5].Value?.ToString();
                            newRow["VÖEN-i"] = row2.Cells[6].Value?.ToString();
                            newRow["Filialı"] = row2.Cells[7].Value?.ToString();
                            newRow["Nömrəsi"] = row2.Cells[8].Value?.ToString();
                            newRow["Tarixi"] = row2.Cells[9].Value?.ToString();
                            newRow["Növü"] = row2.Cells[10].Value?.ToString();
                            newRow["Valyutası"] = row2.Cells[11].Value?.ToString();
                            newRow["Nömrəsi"] = row2.Cells[12].Value?.ToString();
                            newRow["Açılma tarixi"] = row2.Cells[13].Value?.ToString();
                            //newRow["Yas2"] = row2.Cells[1].Value?.ToString();
                            //newRow["AD2"] = row2.Cells[1].Value?.ToString();
                        }
                    }

                    dataTable3.Rows.Add(newRow);
                }
            }

            // DataGridView3'e DataTable'dan verileri ekleyerek sonuçları gösterebilirsiniz
            dataGridView4.DataSource = dataTable3;
        }
        private void bitmis_imza_secilmisler()
        {
            // DataGridView4 için yeni bir DataTable oluşturun
            DataTable dt4 = new DataTable();

            // DataGridView1'deki sütunları DataTable4'e ekleyin
            foreach (DataGridViewColumn col in dataGridView5.Columns)
            {
                dt4.Columns.Add(col.Name);
            }

            // Daha önce eklenen "hesab_nov" değerlerini saklamak için bir HashSet kullanın
            HashSet<string> eklenenHesaplar = new HashSet<string>();

            // dataGridView5 ve DataGridView3'yi karşılaştırın ve eşleşen satırları DataTable4'e ekleyin
            foreach (DataGridViewRow row1 in dataGridView5.Rows)
            {
                string id1 = row1.Cells[0].Value?.ToString();

                foreach (DataGridViewRow row2 in dataGridView3.Rows)
                {
                    string id2 = row2.Cells[0].Value?.ToString();

                    if (id1 == id2)
                    {
                        // "hesab_nov" değerini kontrol edin ve daha önce eklenmemişse ekleyin
                        string hesabNov = row1.Cells[12].Value?.ToString();
                        if (!eklenenHesaplar.Contains(hesabNov))
                        {
                            DataRow newRow = dt4.NewRow();

                            // dataGridView5'den verileri yeni satıra ekle
                            foreach (DataGridViewCell cell in row1.Cells)
                            {
                                newRow[cell.ColumnIndex] = cell.Value;
                            }

                            // DataGridView3'den verileri yeni satıra ekle
                            foreach (DataGridViewCell cell in row1.Cells)
                            {
                                newRow[cell.ColumnIndex] = cell.Value;
                            }

                            dt4.Rows.Add(newRow);

                            // Eklenen "hesab_nov" değerini HashSet'e ekleyin
                            eklenenHesaplar.Add(hesabNov);
                        }
                    }
                }
            }

            // DataGridView4'ü DataTable4 ile doldurun
            dataGridView4.DataSource = dt4;
        }
        private void yeni_imza_secilmisler()
        {
            // DataGridView7 için yeni bir DataTable oluşturun
            DataTable dt4 = new DataTable();

            // DataGridView5'deki sütunları DataTable4'e ekleyin
            foreach (DataGridViewColumn col in dataGridView5.Columns)
            {
                dt4.Columns.Add(col.Name);
            }

            // Daha önce eklenen "hesab_nov" değerlerini saklamak için bir HashSet kullanın
            HashSet<string> eklenenHesaplar = new HashSet<string>();

            // dataGridView5 ve DataGridView3'yi karşılaştırın ve eşleşen satırları DataTable4'e ekleyin
            foreach (DataGridViewRow row1 in dataGridView5.Rows)
            {
                string id1 = row1.Cells[0].Value?.ToString();

                foreach (DataGridViewRow row2 in dataGridView6.Rows)
                {
                    string id2 = row2.Cells[0].Value?.ToString();

                    if (id1 == id2)
                    {
                        // "hesab_nov" değerini kontrol edin ve daha önce eklenmemişse ekleyin
                        string hesabNov = row1.Cells[12].Value?.ToString();
                        if (!eklenenHesaplar.Contains(hesabNov))
                        {
                            DataRow newRow = dt4.NewRow();

                            // dataGridView5'den verileri yeni satıra ekle
                            foreach (DataGridViewCell cell in row1.Cells)
                            {
                                newRow[cell.ColumnIndex] = cell.Value;
                            }

                            // DataGridView6'den verileri yeni satıra ekle
                            foreach (DataGridViewCell cell in row1.Cells)
                            {
                                newRow[cell.ColumnIndex] = cell.Value;
                            }

                            dt4.Rows.Add(newRow);

                            // Eklenen "hesab_nov" değerini HashSet'e ekleyin
                            eklenenHesaplar.Add(hesabNov);
                        }
                    }
                }
            }

            // DataGridView4'ü DataTable4 ile doldurun
            dataGridView7.DataSource = dt4;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            lbl1.Text = "0";
            lbl2.Text = "0";
            lbl3.Text = "0";
            acilmis_huquqi_hesablar();
            imza_huqulu_numayendeler();
            imza_huquqi_vaxti_biten();
            imza_huquqi_vaxti_yeni();
            aktiv_olan_acilmis_huquqi_hesablar();
            bitmis_imza_secilmisler();
            yeni_imza_secilmisler();
            if (dataGridView1.RowCount > 1)
            {
                button2.Enabled = true;
                sayelave1();
            }
            if (dataGridView3.RowCount > 1)
            {
                button3.Enabled = true;
                sayelave2();
            }
            if (dataGridView6.RowCount > 1)
            {
                button4.Enabled = true;
                sayelave3();
            }
            else if (dataGridView1.RowCount < 1 || dataGridView3.RowCount < 1)
            {
                MessageBox.Show("Daxil etdiyiniz tarixlər üzrə hər hansı bir məlumat mövcud deyil.", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        void test()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string sourceFilePath = Path.Combine(desktopPath, "excel1.xlsx"); // Masaüstündeki kaynak dosya
            string targetFilePath = Path.Combine(desktopPath, "excel2.xlsx");   // Masaüstündeki hedef dosya

            // Excel uygulamasını başlat
            Microsoft.Office.Interop. Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Kaynak Excel dosyasını aç
            Microsoft.Office.Interop.Excel.Workbook sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
            Microsoft.Office.Interop.Excel.Worksheet sourceWorksheet = sourceWorkbook.Worksheets[1]; // İlgili sayfa numarası

            // Hedef Excel dosyasını aç
            Microsoft.Office.Interop.Excel.Workbook targetWorkbook = excelApp.Workbooks.Open(targetFilePath);
            Microsoft.Office.Interop.Excel.Worksheet targetWorksheet = targetWorkbook.Worksheets[1]; // İlgili sayfa numarası

            // Verileri kopyala
            sourceWorksheet.UsedRange.Copy();

            // Formüllerin yapışmasını önlemek için yapıştırma seçeneklerini ayarla
            targetWorksheet.UsedRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);

            // Değişiklikleri kaydet
            targetWorkbook.Save();

            // Excel uygulamasını kapat
            excelApp.Quit();

            Console.WriteLine("Veriler başarıyla kopyalandı.");
        }
        void sayelave1()
        {
            int benzersizIDSayisi = dataGridView1.Rows
            .Cast<DataGridViewRow>()
            .Select(row => row.Cells[1].Value)
            .Where(id => id != null)
            .Select(id => id.ToString())
            .Distinct()
            .Count();
            lbl1.Text = benzersizIDSayisi.ToString();
        }
        void sayelave2()
        {
            int benzersizIDSayisi = dataGridView3.Rows
            .Cast<DataGridViewRow>()
            .Select(row => row.Cells[1].Value)
            .Where(id => id != null)
            .Select(id => id.ToString())
            .Distinct()
            .Count();
            lbl2.Text = benzersizIDSayisi.ToString();
        }
        void sayelave3()
        {
            int benzersizIDSayisi = dataGridView6.Rows
            .Cast<DataGridViewRow>()
            .Select(row => row.Cells[1].Value)
            .Where(id => id != null)
            .Select(id => id.ToString())
            .Distinct()
            .Count();
            lbl3.Text = benzersizIDSayisi.ToString();
        }
        private void frm_AML_huquqi_hesab_imza_Load(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
           
        }
        private void acilmis_huquqi_hesablar()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select t.registrac_nomer Rn, odb.func_utf8_to_latin(t.name_licsch) ad, r.inn_regnom voen," +
       " ' ' mulk_nov, ' ' tes_nov,'Bank Melli Iran' bank_ad, '1300036291' bank_voen, ' Bakı filialı' bank_fil, ' ' seh_nom, ' ' seh_tar, ' ' hesab_nov, "+
       " case when substr(t.licsch, 6, 2) = 00 then 'AZN' " +
       " else case when substr(t.licsch, 6, 2) = 01 then 'USD' " +
       " else case when substr(t.licsch, 6, 2) = 02 then 'EUR' " +
       " else case when substr(t.licsch, 6, 2) = 03 then 'RUB' " +
       " else case when substr(t.licsch, 6, 2) = 04 then 'IRR' " +
       " else case when substr(t.licsch, 6, 2) = 05 then 'AED'" +
       " end end end end end end hes_valuta, t.az || t.nr || t.bank || t.licsch Hesab_nom, t.date_open_licsch ac_tar" +
       " from odb.licsch t, odb.regnom r" +
       " where t.registrac_nomer = r.regnom and not substr(t.licsch, 1, 5)  in (99999)  and length(t.licsch) = 20  " +
       " and t.date_close_licsch is null and r.yurik = 1 and " +
       " t.date_open_licsch between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd/mm/yyyy')" +
       " order by t.name_licsch ASC", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                 MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Text = "Sorğu";
            }
            //finally { }
        }
        private void aktiv_olan_acilmis_huquqi_hesablar()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select t.registrac_nomer Rn, odb.func_utf8_to_latin(t.name_licsch) ad, r.inn_regnom voen," +
       " ' ' mulk_nov, ' ' tes_nov,'Bank Melli Iran' bank_ad, '1300036291' bank_voen, ' Bakı filialı' bank_fil, ' ' seh_nom, ' ' seh_tar, ' ' hesab_nov, " +
       " case when substr(t.licsch, 6, 2) = 00 then 'AZN' " +
       " else case when substr(t.licsch, 6, 2) = 01 then 'USD' " +
       " else case when substr(t.licsch, 6, 2) = 02 then 'EUR' " +
       " else case when substr(t.licsch, 6, 2) = 03 then 'RUB' " +
       " else case when substr(t.licsch, 6, 2) = 04 then 'IRR' " +
       " else case when substr(t.licsch, 6, 2) = 05 then 'AED'" +
       " end end end end end end hes_valuta, t.az || t.nr || t.bank || t.licsch Hesab_nom, substr(t.date_open_licsch,0,10)ac_tar" +
       " from odb.licsch t, odb.regnom r" +
       " where t.registrac_nomer = r.regnom and not substr(t.licsch, 1, 5)  in (99999)  and length(t.licsch) = 20  " +
       " and t.date_close_licsch is null and r.yurik = 1 order by t.registrac_nomer", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView5.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Text = "Sorğu";
            }
            //finally { }
        }
        private void imza_huqulu_numayendeler()
        {
            //try
            //{
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                // CheckBox'ların durumlarını saklayan bir dizi oluşturun
                bool[] checkBoxStates = new bool[5]; // 5 CheckBox olduğunu varsayalım

                // CheckBox'ların metinlerini saklayan bir dizi oluşturun
                string[] checkBoxTexts = new string[5]; // CheckBox metinlerini burada saklayın

                // CheckBox'ları ve metinlerini diziye atayın (örnek olarak)
                checkBoxStates[0] = checkBox1.Checked;
                checkBoxTexts[0] = checkBox1.Text;

                checkBoxStates[1] = checkBox2.Checked;
                checkBoxTexts[1] = checkBox2.Text;

                checkBoxStates[2] = checkBox3.Checked;
                checkBoxTexts[2] = checkBox3.Text;

                checkBoxStates[3] = checkBox4.Checked;
                checkBoxTexts[3] = checkBox4.Text;

                checkBoxStates[4] = checkBox5.Checked;
                checkBoxTexts[4] = checkBox5.Text;

                // Şimdi sorgunuzu oluşturabilirsiniz
                Orcon.Open();

                string sqlQuery = "select distinct r.name_regnom, r.regnom, ih.soyadi soy,ih.adi ad,ih.ata_adi ata," +
                    " case when ih.vetendashligi = 'AZ' then ih.fin " +
                    " else case when ih.vetendashligi <> 'AZ' then ' ' " +
                    " end end fin,ih.vetendashligi veten, " +
                    " case when ih.vetendashligi = 'AZ' then ' ' " +
                    " else case when ih.vetendashligi <> 'AZ' then ih.seriyasi_ve_nomresi " +
                    " end end passp, " +
                    " ih.id_aml, j.descript elaqeli " +
                    " from regnom r, licsch l, huquqi_shexs hs, imza_huquqi_olan_shexsler ih, aml_related_persons f, aml_setup_related_persons j, " +
                    " (select r.regnom, min(t.date_open_licsch) qeyd_tar from odb.licsch t, odb.regnom r " +
                    " where t.registrac_nomer = r.regnom group by r.regnom, r.name_regnom order by r.regnom) s" +
                    " where r.yurik = 1 and r.regnom = l.registrac_nomer(+) and r.regnom = hs.regnom(+)  and r.regnom = ih.regnom(+) " +
                    " and ih.id_aml = f.id_aml(+) and f.id_related_persons = j.id(+) and r.regnom = s.regnom(+)";

                string checkBoxConditions = "";

                for (int i = 0; i < checkBoxStates.Length; i++)
                {
                    if (checkBoxStates[i])
                    {
                        if (checkBoxConditions != "")
                        {
                            checkBoxConditions += " OR ";
                        }
                        checkBoxConditions += "j.descript = '" + checkBoxTexts[i] + "'";
                    }
                }

                if (!string.IsNullOrEmpty(checkBoxConditions))
                {
                    sqlQuery += " AND (" + checkBoxConditions + ")";
                }

                sqlQuery += " order by r.regnom";

                Orcom = new OracleCommand(sqlQuery, Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Ordt.Clear(); // Önceki verileri temizle
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;

                Orcon.Close();
            //}
            //catch (Exception)
            //{
            //    // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    button1.Text = "Sorğu";
            //}
            //finally { }
        }
        //private void imza_huquqi_vaxti_biten()
        //{
        //    try
        //    {
        //        DataTable Ordt = new DataTable();
        //        Ordt.Clear();
        //        OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
        //        Orcon.Open();
        //        Orcom = new OracleCommand("select t.regnom,t.SOYADI,T.ADI,T.ATA_ADI,T.FIN,T.VETENDASHLIGI,' ' DIGER,substr(T.IMZA_BITME,0,10), ROWNUM AS SIRA_NO from odb.imza_huquqi_olan_shexsler t " +
        //        "where t.imza_bitme BETWEEN TO_DATE('"+textBox2.Text+"', 'dd/mm/yyyy') AND TO_DATE('"+textBox3.Text+"', 'dd/mm/yyyy')", Orcon);
        //        Orda = new OracleDataAdapter(Orcom);

        //        Orda.Fill(Ordt);
        //        dataGridView3.DataSource = Ordt;
        //        Orcon.Close();
        //    }
        //    catch (Exception)
        //    {
        //        // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        button1.Text = "Sorğu";
        //    }
        //    finally { }
        //}
        private void imza_huquqi_vaxti_biten()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                using (OracleConnection Orcon = new OracleConnection(cl.con))
                {
                    Orcon.Open();

                    string query = @"SELECT t.regnom, t.soyadi, t.adi, t.ata_adi, t.fin, t.vetendasHligi, 
                                    ' ' AS diger, 
                                    SUBSTR(t.imza_bitme, 0, 10) AS imza_bitme, 
                                    ROWNUM AS sira_no 
                             FROM odb.imza_huquqi_olan_shexsler t 
                             WHERE t.imza_bitme 
                             BETWEEN TO_DATE(:girisTarix, 'dd/mm/yyyy') AND TO_DATE(:cixisTarix, 'dd/mm/yyyy')";

                    using (OracleCommand Orcom = new OracleCommand(query, Orcon))
                    {
                        Orcom.Parameters.Add(new OracleParameter("girisTarix", textBox2.Text));
                        Orcom.Parameters.Add(new OracleParameter("cixisTarix", textBox3.Text));

                        OracleDataAdapter Orda = new OracleDataAdapter(Orcom);
                        Orda.Fill(Ordt);
                        dataGridView3.DataSource = Ordt;
                    }
                }
            }
            catch (Exception)
            {
                button1.Text = "Sorğu";
            }
        }

        private void imza_huquqi_vaxti_yeni()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select t.regnom,t.SOYADI,T.ADI,T.ATA_ADI,T.FIN,T.VETENDASHLIGI,' ' DIGER, substr(T.imza_bashlama,0,10) from odb.imza_huquqi_olan_shexsler t " +
                " where t.imza_bashlama BETWEEN TO_DATE('" + textBox2.Text + "', 'dd/mm/yyyy') AND TO_DATE('" + textBox3.Text + "', 'dd/mm/yyyy')", Orcon);
                Orda = new OracleDataAdapter(Orcom);

                Orda.Fill(Ordt);
                dataGridView6.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Text = "Sorğu";
            }
            finally { }
        }
        private void imza_numayende_vaxti()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select * from odb.numayende t where t.imza_bashlama is not null or t.imza_bitme is not null", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView4.DataSource = Ordt;
                Orcon.Close();
        }
            catch (Exception)
            {
                 MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Text = "Sorğu";
            }
            finally { }
        }
        private void imza_huqulu_numayendeler1()
        {
            //try
            //{
                bool[] checkBoxStates = new bool[5];

                checkBoxStates[0] = checkBox1.Checked;
                checkBoxStates[1] = checkBox2.Checked;
                checkBoxStates[2] = checkBox3.Checked;
                checkBoxStates[3] = checkBox4.Checked;
                checkBoxStates[4] = checkBox5.Checked;

                int selectedCheckBoxCount = checkBoxStates.Count(state => state);

                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                string sqlQuery1 = "SELECT distinct r.name_regnom,r.regnom, ih.soyadi soy,ih.adi ad,ih.ata_adi ata,ih.vetendashligi veten," +
    "ih.fin,ih.seriyasi_ve_nomresi pasport,ih.id_aml" +
    " from regnom r, licsch l, huquqi_shexs hs, imza_huquqi_olan_shexsler ih, aml_related_persons f, aml_setup_related_persons j, " +
    " (select  r.regnom, min(t.date_open_licsch) qeyd_tar from odb.licsch t, odb.regnom r   " +
    " where t.registrac_nomer = r.regnom  group by r.regnom, r.name_regnom  order by r.regnom) s" +
    " where r.yurik = 1 and r.regnom = l.registrac_nomer(+) and r.regnom = hs.regnom(+)  and r.regnom = ih.regnom(+) " +
    " and ih.id_aml = f.id_aml(+) and f.id_related_persons = j.id(+) and r.regnom = s.regnom(+) ";

                string checkBoxConditions = "";

                for (int i = 0; i < checkBoxStates.Length; i++)
                {
                    if (checkBoxStates[i])
                    {
                        if (checkBoxConditions != "")
                        {
                            checkBoxConditions += " OR ";
                        }
                        checkBoxConditions += "j.descript = 'Checbox" + (i + 1) + "'";
                    }
                }

                // Seçili CheckBox'lar varsa, j.descript koşulunu sorguya ekleyin
                if (!string.IsNullOrEmpty(checkBoxConditions))
                {
                    sqlQuery1 += " and " + checkBoxConditions + "";
                }

                sqlQuery1 += " order by r.regnom";

                Orcom = new OracleCommand(sqlQuery1, Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;

                // Diğer sorguları benzer şekilde güncelleyebilirsiniz
                // Sorguları bu şablona göre düzenleyin ve tekrar eden kodları kopyala/yapıştır yaparak oluşturun

                Orcon.Close();
            
            //catch (Exception)
            //{
            //    // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    button1.Text = "Sorğu";
            //}
            //finally { }
        }
        private void exceleat()
        {
            //try
            //{
            string valkod = "";
            string adelave = "MMX form";
            int setirsay = dataGridView1.RowCount;
            int dtg1qeydno = 0;
            int dtg2qeydno = 2;
            int dtg1qeydno20 = 21;
            int dtg1qeydno9 = 9;
            int dtg1qeydno21 = 21;
            int dtg1qeydnoCARI = 20;
            string hesad = dataGridView2.Rows[0].Cells[0].Value.ToString();
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            string baseFileName = $"{valkod} - {adelave}  "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\AML_Huquqi_hesablar.xlsx";
            if (File.Exists(Path.Combine(desktopFolder, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(desktopFolder, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            if (dataGridView1.RowCount > 0)
            {
                FileInfo templateFile = new FileInfo(templateFilePath);
                using (ExcelPackage package = new ExcelPackage(templateFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    worksheet.Name = "Hesab çıxarışı";

                    int baslamasetiri = 3;
                    int setir = 1;
                    int sutun = 1;
                    //worksheet.Cells[7, 4].Value = textBox2.Text;
                    //worksheet.Cells[7, 5].Value = textBox3.Text;
                    //worksheet.Cells[8, 4].Value = dataGridView2.Rows[0].Cells[1].Value;
                    //worksheet.Cells[8, 5].Value = dataGridView2.Rows[0].Cells[2].Value;

                    int topexcelsutunsayi = 39;
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        //string deger = dataGridView1.Rows[i].Cells[8].Value.ToString().Substring(0,2);
                        object huquqiqeydno = dataGridView1.Rows[i].Cells[dtg1qeydno].Value;
                        setirsay = setirsay - 1;
                        //label5.Text = setirsay.ToString();

                        for (int j = 0; j < topexcelsutunsayi; j++)
                        {
                            worksheet.Cells[baslamasetiri + i, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                            if (dataGridView1.Rows[i].Cells[dtg2qeydno].Value != null)
                            {
                                object imzalarqeydno = dataGridView2.Rows[i].Cells[dtg2qeydno].Value;
                                
                                if (imzalarqeydno != null && imzalarqeydno.ToString() == huquqiqeydno.ToString())
                                {
                                    worksheet.Cells[7, 4].Value = textBox2.Text;
                                }
                                //}
                            }
                        }
                    }

                    string filePath = Path.Combine(desktopFolder, fileName);
                    button1.Text = "Sorğu";
                    package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                                                            //MessageBox.Show("Excel dosyası başarıyla oluşturuldu ve kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    System.Diagnostics.Process.Start(filePath);

                }
            }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show("Xeta oldu: " + ex.Message);
            //}
            //finally { }
        }
        private void excelat()
        {
            DataTable dataTable1 = new DataTable();
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataTable1.Columns.Add(dataGridView1.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dataRow = dataTable1.NewRow();
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable1.Rows.Add(dataRow);
            }

            // DataGridView2'deki verileri bir DataTable'a al
            DataTable dataTable2 = new DataTable();
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                dataTable2.Columns.Add(dataGridView2.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                DataRow dataRow = dataTable2.NewRow();
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable2.Rows.Add(dataRow);
            }

            // Excel dosyasını aç
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(@"C:\BMI_\AML_Huquqi_hesablar.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk (veya tek) sayfa

                // DataGridView1'deki verileri Excel sayfasına ekleyin
                int excelbaslangic = 1;
                foreach (DataRow dataRow in dataTable1.Rows)
                {
                    for (int i = 0; i < dataTable1.Columns.Count; i++)
                    {
                        worksheet.Cells[excelbaslangic, i + 1].Value = dataRow[i].ToString();
                    }
                    excelbaslangic++;
                }

                // DataGridView2'deki belirli koşulu sağlayan satırları ekleyin
                foreach (DataRow dataRow in dataTable2.Rows)
                {
                    if (dataRow["qeyno"].ToString() == "belirli_değer")
                    {
                        for (int i = 0; i < dataTable2.Columns.Count; i++)
                        {
                            worksheet.Cells[excelbaslangic, i + 1].Value = dataRow[i].ToString();
                        }
                        excelbaslangic++;
                    }
                }

                // Excel dosyasını kaydet
                package.Save();
            }

            MessageBox.Show("Veriler Excel dosyasına aktarıldı.");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // excelat();
            excelElave1();
        }
        void excelElave1()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            DataTable dataTable1 = new DataTable();
            string baseFileName = $"{textBox2.Text} - {textBox3.Text}   "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            if (File.Exists(Path.Combine(desktopFolder, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(Path.Combine(desktopFolder, $"{baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsx";
            }
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataTable1.Columns.Add(dataGridView1.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dataRow = dataTable1.NewRow();
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable1.Rows.Add(dataRow);
            }

            // DataGridView2'deki verileri bir DataTable'a al
            DataTable dataTable2 = new DataTable();
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                dataTable2.Columns.Add(dataGridView2.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                DataRow dataRow = dataTable2.NewRow();
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable2.Rows.Add(dataRow);
            }

            // Excel dosyasını aç
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(@"C:\BMI_\AML_Huquqi_hesablar.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk (veya tek) sayfa

                int excelbaslangic = 4;
                bool dataTable2IsEmpty = true; // dataTable2 içindeki tüm satırların boş olduğunu varsayalım

                for (int row2 = 0; row2 < dataTable2.Rows.Count; row2++)
                {
                    for (int i = 0; i < dataTable2.Columns.Count; i++)
                    {
                        string cellValue = dataTable2.Rows[row2][i].ToString();

                        // Eğer herhangi bir hücre değeri boş değilse, dataTable2 boş değildir
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            dataTable2IsEmpty = false;
                            break; // Bir hücre dolu olduğunda döngüyü durdurabiliriz
                        }
                    }

                    // Eğer tüm satırlar boşsa, dataTable2 işlenmeyecek
                    if (!dataTable2IsEmpty)
                    {
                        break;
                    }
                }

                // dataTable2 boş değilse işlem yapabiliriz
                if (!dataTable2IsEmpty)
                {
                    // dataTable2 içeriğini işlemeye devam edin
                    for (int row1 = 0; row1 < dataTable1.Rows.Count; row1++)
                    {
                        int excimzalar = 14;
                        bool rowIsEmpty = true;

                        // DataGridView1'deki verileri Excel sayfasına ekleyin
                        for (int i = 0; i < dataTable1.Columns.Count; i++)
                        {
                            string cellValue = dataTable1.Rows[row1][i].ToString();
                            worksheet.Cells[excelbaslangic, i + 1].Value = cellValue;
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                rowIsEmpty = false;
                            }
                        }

                        // Eğer satır boşsa işlem yapmayın
                        if (!rowIsEmpty)
                        {
                            string qeyno1 = dataTable1.Rows[row1][0].ToString();

                            int qeyno2Count = 0;

                            for (int row2 = 0; row2 < dataTable2.Rows.Count; row2++)
                            {
                                string qeyno2 = dataTable2.Rows[row2][1].ToString();
                                if (qeyno1 == qeyno2)
                                {
                                    for (int i = 2; i < 8; i++)
                                    {
                                        worksheet.Cells[excelbaslangic, i + excimzalar].Value = dataTable2.Rows[row2][i].ToString();
                                    }
                                    excimzalar = excimzalar + 6;
                                    if (!string.IsNullOrWhiteSpace(dataTable2.Rows[row2][2].ToString()))
                                    {
                                        qeyno2Count++;
                                    }

                                }
                            }

                            if (qeyno2Count > 0)
                            {
                                worksheet.Cells[excelbaslangic, 15].Value = qeyno2Count;
                            }
                        }

                        excelbaslangic++;
                        if (!string.IsNullOrWhiteSpace(dataTable1.Rows[row1][1].ToString()))
                        {
                            worksheet.Cells[4 + row1, 1].Value = (row1 + 1).ToString();
                        }
                    }
                }

                // Excel dosyasını kaydet
                string filePath = Path.Combine(desktopFolder, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
                //package.Save();
            }

            
        }
        void excelElave2() //həmçinin imza səlahiyyəti bitmiş
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            DataTable dataTable1 = new DataTable();
            //string baseFileName = $"{textBox2.Text} - {textBox3.Text}   "; // Temel dosya adı
            //string fileName = baseFileName + ".xlsx";

            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = $"{textBox2.Text} - {textBox3.Text}   "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\AML_Əlavə 2.xlsx";
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
            for (int i = 0; i < dataGridView4.ColumnCount; i++)
            {
                dataTable1.Columns.Add(dataGridView4.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView4.Rows)
            {
                DataRow dataRow = dataTable1.NewRow();
                for (int i = 0; i < dataGridView4.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable1.Rows.Add(dataRow);
            }

            // DataGridView2'deki verileri bir DataTable'a al
            DataTable dataTable2 = new DataTable();
            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                dataTable2.Columns.Add(dataGridView3.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                DataRow dataRow = dataTable2.NewRow();
                for (int i = 0; i < dataGridView3.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable2.Rows.Add(dataRow);
            }

            // Excel dosyasını aç
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(@"C:\BMI_\AML_Əlavə 2.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk (veya tek) sayfa

                int excelbaslangic = 4;

                for (int row1 = 0; row1 < dataTable1.Rows.Count; row1++)
                {
                    
                    int excimzalar = 14;
                    int sira = 0;
                    // dataGridView5'deki verileri Excel sayfasına ekleyin
                    for (int i = 0; i < dataTable1.Columns.Count; i++)
                    {
                        worksheet.Cells[excelbaslangic, i + 1].Value = dataTable1.Rows[row1][i].ToString();
                    }
                    string qeyno1 = dataTable1.Rows[row1][0].ToString();

                    for (int row2 = 0; row2 < dataTable2.Rows.Count-1; row2++)
                    {
                        string qeyno2 = dataTable2.Rows[row2][0].ToString();
                        if (qeyno1 == qeyno2)
                        {

                            // Eşleşme durumunda DataGridView2'den gelen verileri Excel sayfasına ekleyin, 16. sütundan başlayarak
                            for (int i = 1; i < 8; i++) // 3. sütundan başlayarak
                            {
                                worksheet.Cells[excelbaslangic, i + excimzalar].Value = dataTable2.Rows[row2][i].ToString();
                            }
                            excimzalar = excimzalar + 7;
                        }
                    }
                    excelbaslangic++;
                    if (!string.IsNullOrWhiteSpace(dataTable1.Rows[row1][1].ToString()))
                    {
                        worksheet.Cells[4 + row1, 1].Value = (row1 + 1).ToString();
                    }
                }

                // Excel dosyasını kaydet
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
                //package.Save();
            }

            //MessageBox.Show("Veriler Excel dosyasına aktarıldı.");
        }
        void excelElave3() //həmçinin imza səlahiyyəti verilmis
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //desktopFolder = Path.Combine(desktopFolder, "AML cixaris");
            DataTable dataTable1 = new DataTable();

            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            string textBoxText = textBox2.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = $"{textBox2.Text} - {textBox3.Text}   "; // Temel dosya adı
            string fileName = baseFileName + ".xlsx";
            string templateFilePath = @"C:\BMI_\AML_Əlavə 3.xlsx";
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
            for (int i = 0; i < dataGridView5.ColumnCount; i++)
            {
                dataTable1.Columns.Add(dataGridView7.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView7.Rows)
            {
                DataRow dataRow = dataTable1.NewRow();
                for (int i = 0; i < dataGridView7.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable1.Rows.Add(dataRow);
            }

            // DataGridView7'deki verileri bir DataTable'a al
            DataTable dataTable2 = new DataTable();
            for (int i = 0; i < dataGridView6.ColumnCount; i++)
            {
                dataTable2.Columns.Add(dataGridView6.Columns[i].HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView6.Rows)
            {
                DataRow dataRow = dataTable2.NewRow();
                for (int i = 0; i < dataGridView6.ColumnCount; i++)
                {
                    dataRow[i] = row.Cells[i].Value;
                }
                dataTable2.Rows.Add(dataRow);
            }

            // Excel dosyasını aç
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(@"C:\BMI_\AML_Əlavə 3.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk (veya tek) sayfa

                int excelbaslangic = 4;

                for (int row1 = 0; row1 < dataTable1.Rows.Count; row1++)
                {
                    int excimzalar = 14;
                    int sira = 0;
                    // dataGridView5'deki verileri Excel sayfasına ekleyin
                    for (int i = 0; i < dataTable1.Columns.Count; i++)
                    {
                        worksheet.Cells[excelbaslangic, i + 1].Value = dataTable1.Rows[row1][i].ToString();
                    }

                    string qeyno1 = dataTable1.Rows[row1][0].ToString();

                    for (int row2 = 0; row2 < dataTable2.Rows.Count - 1; row2++)
                    {
                        string qeyno2 = dataTable2.Rows[row2][0].ToString();
                        if (qeyno1 == qeyno2)
                        {

                            // Eşleşme durumunda DataGridView7'den gelen verileri Excel sayfasına ekleyin, 16. sütundan başlayarak
                            for (int i = 1; i < 8; i++) // 3. sütundan başlayarak
                            {
                                worksheet.Cells[excelbaslangic, i + excimzalar].Value = dataTable2.Rows[row2][i].ToString();
                            }
                            excimzalar = excimzalar + 7;
                        }
                    }
                    excelbaslangic++;
                    if (!string.IsNullOrWhiteSpace(dataTable1.Rows[row1][1].ToString()))
                    {
                        worksheet.Cells[4 + row1, 1].Value = (row1 + 1).ToString();
                    }
                }

                // Excel dosyasını kaydet
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
                //package.Save();
            }

            //MessageBox.Show("Veriler Excel dosyasına aktarıldı.");
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void button2_MouseHover(object sender, EventArgs e)
        {
            
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
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = textBox3.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    textBox3.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox3.Clear(); // Hatalı girişi temizle
                }
            }
        }
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                textBox3.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }
        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            excelElave2();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            excelElave3();
        }
    }
}
