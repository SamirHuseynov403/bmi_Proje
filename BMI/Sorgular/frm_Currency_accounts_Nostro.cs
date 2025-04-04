using BMI.Muhasibat;
using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Windows.Documents;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace BMI.Sorgular
{
    public partial class frm_Currency_accounts_Nostro : Form
    {
        public frm_Currency_accounts_Nostro()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        DataTable dataTable = new DataTable();
        private void word_at()
        {
            dataTable.Clear();
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy");

            string query = @"SELECT 
    ad, 
    SUM(CASE WHEN val = '00' THEN qq ELSE 0 END) AS AZN,
    SUM(CASE WHEN val = '01' THEN qqv ELSE 0 END) AS USD,
    SUM(CASE WHEN val = '02' THEN qqv ELSE 0 END) AS EURO,
    SUM(CASE WHEN val = '03' THEN qqv ELSE 0 END) AS RUB,
    SUM(CASE WHEN val = '04' THEN qqv ELSE 0 END) AS İİR,
    SUM(CASE WHEN val = '05' THEN qqv ELSE 0 END) AS AER  
FROM (
    SELECT odb.func_utf8_to_latin(upper(r.name_regnom)) ad,         
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) IN (11010,11020) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
     UNION 
    SELECT odb.func_utf8_to_latin(upper(r.name_regnom)) ad,         
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) = z.balsch  
       AND substr(l.licsch,1,5) NOT IN (15210,15213,15215,15220,15223,15225,11710,15770) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY') 
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
     UNION 
    SELECT 
           CASE 
               WHEN substr(l.licsch,1,5) IN (15770,11710,11110) THEN 'REVERSE REPO AND OVERNIGHT OPERATIONS' 
               ELSE odb.func_utf8_to_latin(upper(r.name_regnom)) 
           END ad,       
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) = z.balsch  
       AND substr(l.licsch,1,5) IN (15770,11710,11110) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')  
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
) m 
GROUP BY ad"; // Lazım olan SQL sorğusu


            string[] kurslarValues = new string[5]; // USD, EUR, RUB, IRR, AED dəyərləri
                                                    // Tarixi C# MonthCalendar-dan götür

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();

                using (OracleCommand command = new OracleCommand(query, connection))
                {
                    command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                // Kurslar üçün yeni SELECT sorğusunu icra et
                string kurslarQuery = @"
            SELECT 
                '1 USD   ' || TO_CHAR(1 * func_get_kurval('01', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS usd,
                '1 EUR   ' || TO_CHAR(1 * func_get_kurval('02', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS avro,
                '100 RUB   ' || TO_CHAR(100 * func_get_kurval('03', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS rub,
                '10000 IIR   ' || TO_CHAR(10000 * func_get_kurval('04', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS irr,
                '1 AED   ' || TO_CHAR(1 * func_get_kurval('05', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS aed
            FROM dual";



                using (OracleCommand kurslarCommand = new OracleCommand(kurslarQuery, connection))
                {
                    // Tarixi parametr olaraq əlavə et
                    kurslarCommand.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                    using (OracleDataReader reader = kurslarCommand.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            kurslarValues[0] = reader["usd"].ToString();
                            kurslarValues[1] = reader["avro"].ToString();
                            kurslarValues[2] = reader["rub"].ToString();
                            kurslarValues[3] = reader["irr"].ToString();
                            kurslarValues[4] = reader["aed"].ToString();
                        }
                    }
                }
            }

            // Word sənədini aç və məlumatları əlavə et
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Word.Document doc = wordApp.Documents.Open(@"C:\BMI_\Nostro Accounts.docx");
            Microsoft.Office.Interop.Word.Table table = doc.Tables[1]; // Sənəddəki birinci tabloyu seçirik

            // Word sənədində {tarix} mətni tap və əvəz et
            Word.Find findObject = wordApp.Selection.Find;
            findObject.Text = "{tarix}"; // Axtarılan mətn
            findObject.Replacement.Text = formattedDate; // Əvəz ediləcək mətn
            findObject.Forward = true; // İrəli axtarış
            findObject.Wrap = Word.WdFindWrap.wdFindContinue; // Bütün sənəddə davam etdir
            findObject.Format = false; // Format nəzərə alınmadan
            findObject.MatchCase = false; // Böyük-kiçik hərf fərqi olmadan
            findObject.MatchWholeWord = true; // Tam söz kimi axtar


            // Mətni əvəz et
            bool isReplaced = findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);

            if (table.Columns.Count == 7) // Tabloda 7 sütun varsa
            {
                // İlk sütunun eni (3.55 inches = 237 points)
                table.Columns[1].Width = 230;

                // Qalan sütunların enini təyin edin (1.06 inches = 76 points)
                for (int colIndex = 2; colIndex <= 7; colIndex++)
                {
                    table.Columns[colIndex].Width = 81;
                }
                //Tablonun hər bir sətirinin içindəki aralığı təyin et
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell.TopPadding = 0.25f;    // Üst aralığı (0.2 inch)
                        cell.BottomPadding = 0.25f; // Alt aralığı (0.2 inch)
                        cell.LeftPadding = 0.25f;   // Sol aralığı (0.2 inch)
                        cell.RightPadding = 0.25f;  // Sağ aralığı (0.2 inch)
                    }
                }
                table.AllowAutoFit = false; // Avtomatik uyğunlaşmanı deaktiv et
            }

            // Kurslardan gələn məlumatları tablonun 1-ci sətirinin 3-7-ci sütunlarına yaz
            for (int i = 0; i < kurslarValues.Length; i++)
            {
                table.Cell(1, 3 + i).Range.Text = kurslarValues[i]; // 3-cü sütundan başlayaraq yazırıq
            }

            // Query-dən gələn məlumatları tablonun qalan hissəsinə yaz
            int rowIndex = 2; // İkinci sətrdən başlayaraq
            foreach (DataRow row in dataTable.Rows)
            {
                // Yeni sətir əlavə et
                table.Rows.Add();

                for (int colIndex = 1; colIndex <= dataTable.Columns.Count; colIndex++)
                {
                    table.Cell(rowIndex, colIndex).Range.Text = row[colIndex - 1].ToString();

                }
                rowIndex++;
            }

            // Sonuncu sətir əlavə et
            table.Rows.Add();
            Word.Row lastRow = table.Rows[table.Rows.Count];

            // Sonuncu sətirin birinci sütununa "Total" yaz
            lastRow.Cells[1].Range.Text = "Total";
            lastRow.Cells[1].Range.Font.Bold = 1; // Qalın şrift

            // Digər sütunlara toplamlara uyğun dəyər yaz
            for (int colIndex = 2; colIndex <= table.Columns.Count; colIndex++)
            {
                double sum = 0;

                // Hər sütunun toplamasını hesabla
                for (int rowIndex_ = 2; rowIndex_ < table.Rows.Count - 1; rowIndex_++) // 2-dən başlayırıq çünki başlıq var
                {
                    string cellValue = table.Cell(rowIndex_, colIndex).Range.Text.Trim().Replace("\r\a", "");
                    if (double.TryParse(cellValue, out double value))
                    {
                        sum += value; // Toplamı hesabla
                    }
                }

                // Sonuncu sətirdə nəticəni yaz
                lastRow.Cells[colIndex].Range.Text = sum.ToString("F2"); // İki onluq göstərəcək
                lastRow.Cells[colIndex].Range.Font.Bold = 1; // Qalın şrift
            }

            // Sonuncu sətir üçün şrift və fon rəngi dəyiş
            Word.Row lastRow1 = table.Rows[table.Rows.Count];
            foreach (Word.Cell cell in lastRow1.Cells)
            {
                cell.Range.Font.Size = 9; // Şrift ölçüsü 12
                cell.Range.Font.Bold = 1; // Qalın şrift
                //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlueGray; // Fon rəngi: Sarı
            }

            // Rəqəmləri xüsusi formatda yazmaq üçün NumberFormatInfo təyin et
            var numberFormat = new System.Globalization.NumberFormatInfo
            {
                NumberGroupSeparator = " ", // Qruplaşdırma üçün boşluq istifadə edilir
                NumberDecimalSeparator = ".", // Ondalık üçün nöqtə istifadə edilir
                NumberGroupSizes = new[] { 3 } // 3 rəqəmdən bir qruplaşdır
            };

            // Tabloda rəqəmləri xüsusi formatda dəyişdir
            foreach (Word.Row row in table.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    string cellValue = cell.Range.Text.Trim().Replace("\r\a", "");
                    if (double.TryParse(cellValue, out double value))
                    {
                        // Əgər dəyər 0-dırsa, hüceyrəni boş qoy
                        if (value == 0)
                        {
                            cell.Range.Text = ""; // Boş məzmun
                        }
                        else
                        {
                            // Əks halda formatla
                            cell.Range.Text = value.ToString("N2", numberFormat);
                        }
                    }
                }
            }
            // İlk sətir (başlıq) üçün şrift və fon rəngi dəyiş
            Word.Row firstRow = table.Rows[1];
            foreach (Word.Cell cell in firstRow.Cells)
            {
                cell.Range.Font.Size = 9; // Şrift ölçüsü 12
                cell.Range.Font.Bold = 1; // Qalın şrift
                //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue; // Fon rəngi: Mavi
            }

            // Faylı saxla
            string filePath = @"C:\Users\mehemmed\Desktop\Updated_Nostro_Accounts.docx";
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath); // Köhnə faylı sil
            }
            doc.SaveAs2(filePath);

            wordApp.Visible = true;
        }
        private void word_atyeni()
        {
            dataTable.Clear();
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy");

            string query = @"SELECT 
    ad, 
    SUM(CASE WHEN val = '00' THEN qq ELSE 0 END) AS AZN,
    SUM(CASE WHEN val = '01' THEN qqv ELSE 0 END) AS USD,
    SUM(CASE WHEN val = '02' THEN qqv ELSE 0 END) AS EURO,
    SUM(CASE WHEN val = '03' THEN qqv ELSE 0 END) AS RUB,
    SUM(CASE WHEN val = '04' THEN qqv ELSE 0 END) AS İİR,
    SUM(CASE WHEN val = '05' THEN qqv ELSE 0 END) AS AER  
FROM (
    SELECT odb.func_utf8_to_latin(upper(r.name_regnom)) ad,         
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) IN (11010,11020) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
     UNION 
    SELECT odb.func_utf8_to_latin(upper(r.name_regnom)) ad,         
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) = z.balsch  
       AND substr(l.licsch,1,5) NOT IN (15210,15213,15215,15220,15223,15225,11710,15770,11110) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY') 
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
     UNION 
    SELECT 
           CASE 
               WHEN substr(l.licsch,1,5) IN (15770,11710,11110) THEN 'REVERSE REPO AND OVERNIGHT OPERATIONS' 
               ELSE odb.func_utf8_to_latin(upper(r.name_regnom)) 
           END ad,       
           l.saldo_ish_nacval qq, 
           l.saldo_ish_inval qqv,
           substr(l.licsch,6,2) val
      FROM odb.arh_saldo_ls l, odb.licsch l1, odb.regnom r, odb.balsch_rs z 
     WHERE substr(l.licsch,1,5) = z.balsch  
       AND substr(l.licsch,1,5) IN (15770,11710,11110) 
       AND l.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')  
       AND l.licsch = l1.licsch 
       AND r.regnom = l1.registrac_nomer 
       AND l.saldo_ish_nacval > 0 
) m 
GROUP BY ad"; // Lazım olan SQL sorğusu


            string[] kurslarValues = new string[5]; // USD, EUR, RUB, IRR, AED dəyərləri
                                                    // Tarixi C# MonthCalendar-dan götür

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();

                using (OracleCommand command = new OracleCommand(query, connection))
                {
                    command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                    using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                // Kurslar üçün yeni SELECT sorğusunu icra et
                string kurslarQuery = @"
            SELECT 
                '1 USD   ' || TO_CHAR(1 * func_get_kurval('01', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS usd,
                '1 EUR   ' || TO_CHAR(1 * func_get_kurval('02', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS avro,
                '100 RUB   ' || TO_CHAR(100 * func_get_kurval('03', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS rub,
                '10000 IIR   ' || TO_CHAR(10000 * func_get_kurval('04', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS irr,
                '1 AED   ' || TO_CHAR(1 * func_get_kurval('05', TO_DATE(:selectedDate, 'DD-MM-YYYY')), '0.0000') AS aed
            FROM dual";



                using (OracleCommand kurslarCommand = new OracleCommand(kurslarQuery, connection))
                {
                    // Tarixi parametr olaraq əlavə et
                    kurslarCommand.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                    using (OracleDataReader reader = kurslarCommand.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            kurslarValues[0] = reader["usd"].ToString();
                            kurslarValues[1] = reader["avro"].ToString();
                            kurslarValues[2] = reader["rub"].ToString();
                            kurslarValues[3] = reader["irr"].ToString();
                            kurslarValues[4] = reader["aed"].ToString();
                        }
                    }
                }
            }
            // Word sənədini aç və məlumatları əlavə et
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            string qovluqyoludoc = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Sorgular", "Wordler", "Nostro Accounts.docx");
            Word.Document doc = wordApp.Documents.Open(qovluqyoludoc);
            Microsoft.Office.Interop.Word.Table table = doc.Tables[1]; // Sənəddəki birinci tabloyu seçirik

            // Word sənədində {tarix} mətni tap və əvəz et
            Word.Find findObject = wordApp.Selection.Find;
            findObject.Text = "{tarix}"; // Axtarılan mətn
            findObject.Replacement.Text = formattedDate; // Əvəz ediləcək mətn
            findObject.Forward = true; // İrəli axtarış
            findObject.Wrap = Word.WdFindWrap.wdFindContinue; // Bütün sənəddə davam etdir
            findObject.Format = false; // Format nəzərə alınmadan
            findObject.MatchCase = false; // Böyük-kiçik hərf fərqi olmadan
            findObject.MatchWholeWord = true; // Tam söz kimi axtar


            // Mətni əvəz et
            bool isReplaced = findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);

            if (table.Columns.Count == 7) // Tabloda 7 sütun varsa
            {
                // İlk sütunun eni (3.55 inches = 237 points)
                table.Columns[1].Width = 230;

                // Qalan sütunların enini təyin edin (1.06 inches = 76 points)
                for (int colIndex = 2; colIndex <= 7; colIndex++)
                {
                    table.Columns[colIndex].Width = 81;
                }
                //Tablonun hər bir sətirinin içindəki aralığı təyin et
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell.TopPadding = 0.25f;    // Üst aralığı (0.2 inch)
                        cell.BottomPadding = 0.25f; // Alt aralığı (0.2 inch)
                        cell.LeftPadding = 0.25f;   // Sol aralığı (0.2 inch)
                        cell.RightPadding = 0.25f;  // Sağ aralığı (0.2 inch)
                    }
                }
                table.AllowAutoFit = false; // Avtomatik uyğunlaşmanı deaktiv et
            }

            // Kurslardan gələn məlumatları tablonun 1-ci sətirinin 3-7-ci sütunlarına yaz
            for (int i = 0; i < kurslarValues.Length; i++)
            {
                table.Cell(1, 3 + i).Range.Text = kurslarValues[i]; // 3-cü sütundan başlayaraq yazırıq
            }

            // Query-dən gələn məlumatları tablonun qalan hissəsinə yaz
            int rowIndex = 2; // İkinci sətrdən başlayaraq
            foreach (DataRow row in dataTable.Rows)
            {
                // Yeni sətir əlavə et
                table.Rows.Add();

                for (int colIndex = 1; colIndex <= dataTable.Columns.Count; colIndex++)
                {
                    table.Cell(rowIndex, colIndex).Range.Text = row[colIndex - 1].ToString();

                }
                rowIndex++;
            }

            // Sonuncu sətir əlavə et
            table.Rows.Add();
            Word.Row lastRow = table.Rows[table.Rows.Count];

            // Sonuncu sətirin birinci sütununa "Total" yaz
            lastRow.Cells[1].Range.Text = "Total";
            lastRow.Cells[1].Range.Font.Bold = 1; // Qalın şrift

            // Digər sütunlara toplamlara uyğun dəyər yaz
            for (int colIndex = 2; colIndex <= table.Columns.Count; colIndex++)
            {
                double sum = 0;

                // Hər sütunun toplamasını hesabla
                for (int rowIndex_ = 2; rowIndex_ <= table.Rows.Count; rowIndex_++)  // 2-dən başlayırıq çünki başlıq var
                {
                    string cellValue = table.Cell(rowIndex_, colIndex).Range.Text.Trim().Replace("\r\a", "");
                    if (double.TryParse(cellValue, out double value))
                    {
                        sum += value; // Toplamı hesabla
                    }
                }

                // Sonuncu sətirdə nəticəni yaz
                lastRow.Cells[colIndex].Range.Text = sum.ToString("F2"); // İki onluq göstərəcək
                lastRow.Cells[colIndex].Range.Font.Bold = 1; // Qalın şrift
            }

            // Sonuncu sətir üçün şrift və fon rəngi dəyiş
            Word.Row lastRow1 = table.Rows[table.Rows.Count];
            foreach (Word.Cell cell in lastRow1.Cells)
            {
                cell.Range.Font.Size = 9; // Şrift ölçüsü 12
                cell.Range.Font.Bold = 1; // Qalın şrift
                //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlueGray; // Fon rəngi: Sarı
            }

            // Rəqəmləri xüsusi formatda yazmaq üçün NumberFormatInfo təyin et
            var numberFormat = new System.Globalization.NumberFormatInfo
            {
                NumberGroupSeparator = " ", // Qruplaşdırma üçün boşluq istifadə edilir
                NumberDecimalSeparator = ".", // Ondalık üçün nöqtə istifadə edilir
                NumberGroupSizes = new[] { 3 } // 3 rəqəmdən bir qruplaşdır
            };

            // Tabloda rəqəmləri xüsusi formatda dəyişdir
            foreach (Word.Row row in table.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    string cellValue = cell.Range.Text.Trim().Replace("\r\a", "");
                    if (double.TryParse(cellValue, out double value))
                    {
                        // Əgər dəyər 0-dırsa, hüceyrəni boş qoy
                        if (value == 0)
                        {
                            cell.Range.Text = ""; // Boş məzmun
                        }
                        else
                        {
                            // Əks halda formatla
                            cell.Range.Text = value.ToString("N2", numberFormat);
                        }
                    }
                }
            }
            // İlk sətir (başlıq) üçün şrift və fon rəngi dəyiş
            Word.Row firstRow = table.Rows[1];
            foreach (Word.Cell cell in firstRow.Cells)
            {
                cell.Range.Font.Size = 9; // Şrift ölçüsü 12
                cell.Range.Font.Bold = 1; // Qalın şrift
                //cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue; // Fon rəngi: Mavi
            }

            // Faylı saxla
            string folderPath = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis wordler");
            string baseFileName = "Yaradilmis Nostro Accounts";
            string fileExtension = ".docx";
            string filePath = System.IO.Path.Combine(folderPath, baseFileName + fileExtension);

            // Eyni adda fayl varsa, adın sonuna nömrə əlavə et
            int fileCounter = 1;
            while (System.IO.File.Exists(filePath))
            {
                filePath = System.IO.Path.Combine(folderPath, $"{baseFileName} - {fileCounter}{fileExtension}");
                fileCounter++;
            }

            doc.SaveAs2(filePath);

            wordApp.Visible = true;
        }
        private void frm_CorresPondent_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now.Date;


            // Keçmiş günləri seçməyə icazə vermək üçün heç bir məhdudiyyət qoymuruq
            // Amma bu məhdudiyyətlərə riayət etmək üçün əlavə nəzarət kodu yazaq.
        }
        private void CreateCustomCalendar(int year, int month)
        {
            // Əvvəlcə köhnə düymələri təmizləyirik
            panelCalendar.Controls.Clear();

            // Ayın ilk gününü və gün sayını tapırıq
            DateTime firstDay = new DateTime(year, month, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Təqvim üçün başlanğıc koordinatları
            int startX = 10, startY = 10;
            int buttonSize = 40;
            int margin = 5;

            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime currentDay = new DateTime(year, month, day);

                Button btnDay = new Button
                {
                    Text = day.ToString(),
                    Width = buttonSize,
                    Height = buttonSize,
                    Left = startX + ((day - 1) % 7) * (buttonSize + margin),
                    Top = startY + ((day - 1) / 7) * (buttonSize + margin),
                    Tag = currentDay
                };

                // Bugünkü və gələcək günlər üçün rəngləmə
                if (currentDay >= DateTime.Now.Date)
                {
                    btnDay.BackColor = Color.LightGray; // Gələcək günlər boz rəngdə
                    btnDay.Enabled = false; // Gələcək günləri deaktiv edirik
                }
                else
                {
                    btnDay.BackColor = Color.LightGreen; // Keçmiş günlər yaşıl rəngdə
                }

                // Düyməni panellə əlavə edirik
                panelCalendar.Controls.Add(btnDay);
            }
        }
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            btn_sorgu.Enabled = true;
            // Bugünkü və gələcək tarixlər seçilibsə:
            if (e.Start >= today)
            {
                MessageBox.Show("Bugünkü və gələcək tarixləri seçmək mümkün deyil!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Seçimi geri qaytarırıq (istifadəçiyə xəbərdarlıq veririk və seçimi keçmiş tarixə qaytarırıq).
                monthCalendar1.SetSelectionRange(today.AddDays(-1), today.AddDays(-1)); // Dünənki tarixi seçili edir
                btn_sorgu.Enabled = false;
            }
            else
            {
            // MonthCalendar-dan seçilən tarixi al
            DateTime selectedDate = monthCalendar1.SelectionStart;
            string formattedDate = selectedDate.ToString("dd-MM-yyyy"); // Oracle üçün uyğun tarix formatı

            // Oracle sorğusu
            string query = @"
        SELECT c.space_or_star AS gun
        FROM odb.calendar c
        WHERE c.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')";

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                try
                {

                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        // Tarixi parametr olaraq ötür
                        command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                        // Sorğunu icra et
                        object result = command.ExecuteScalar();

                        if (result != null && result.ToString() == "*")
                        {
                            btn_sorgu.Enabled = false;
                            // Əgər nəticə "*" isə, qeyri-iş günü mesajını göstər
                            MessageBox.Show("Seçdiyiniz tarix qeyri-iş günüdür!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }
                    }
                }
                catch (Exception ex)
                {
                    // Xəta baş verdikdə mesaj göstər
                    MessageBox.Show($"Xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                }
            }
        }

        private void btn_sorgu_Click(object sender, EventArgs e)
        {
            word_atyeni();
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }
    }
}
