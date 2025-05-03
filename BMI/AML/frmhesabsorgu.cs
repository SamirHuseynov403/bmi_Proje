using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Controls;
using BMI.Muhasibat;


namespace BMI
{
    public partial class frmhesabsorgu : Form
    {
        public frmhesabsorgu()
        {
            InitializeComponent();
        }
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();

        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;
        private void exceleat2()
        {
            try
            {
                DataTable Hes_adlari = new DataTable();
                Hes_adlari.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select t.licsch,t.name_licsch,r.pincode," +
                " case when r.fizik = 1 then '' else t.inn_licsch end voen " +
                " from odb.licsch t,odb.regnom r " +
                " where substr(t.licsch, 11, 5) = substr(r.regnom, 2.5) and t.date_close_licsch is null " +
                " order by t.licsch", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Hes_adlari);
                Orcon.Close();

                string valkod = "";
                string adelave = "MMX form";
                int setirsay = dataGridView1.RowCount;
                if (textBox1.Text.Substring(5, 2) == "00")
                {
                    valkod = "AZN";
                }
                else if (textBox1.Text.Substring(5, 2) == "01")
                {
                    valkod = "USD";
                }
                else if (textBox1.Text.Substring(5, 2) == "02")
                {
                    valkod = "EURO";
                }
                else if (textBox1.Text.Substring(5, 2) == "03")
                {
                    valkod = "RUBL";
                }
                else if (textBox1.Text.Substring(5, 2) == "04")
                {
                    valkod = "IRR";
                }
                else if (textBox1.Text.Substring(5, 2) == "05")
                {
                    valkod = "BƏƏ";
                }

                int columnNo = 38;
                int columnNo1 = 39;
                int columnNo20 = 21;
                int columnNo9 = 9;
                int columnNo21 = 21;
                int columnNoCARI = 20;
                string hesad = dataGridView2.Rows[0].Cells[0].Value.ToString();

                cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
                cl.fileName = "AML cixaris" + ".xlsx";
                cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "AML", "Exceller", "AML_Hesab.xlsx");
                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                cl.baseFileName = hesad;

                if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
                {
                    int fileCounter = 1;
                    while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.baseFileName} - {fileCounter}.xlsx")))
                    {
                        fileCounter++;
                    }
                    cl.fileName = $"{cl.baseFileName} - {fileCounter}.xlsx";
                }
                if (dataGridView1.RowCount > 0)
                {
                    FileInfo templateFile = new FileInfo(cl.templateFilePath);
                    using (ExcelPackage package = new ExcelPackage(templateFile))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                        worksheet.Name = "Hesab çıxarışı";

                        int startRow = 12;
                        worksheet.Cells[5, 4].Value = textBox2.Text;//baslama tarixi
                        worksheet.Cells[5, 5].Value = textBox3.Text;//bitme tarixi
                        worksheet.Cells[6, 4].Value = dataGridView2.Rows[0].Cells[1].Value;//saldo baslama
                        worksheet.Cells[6, 5].Value = dataGridView2.Rows[0].Cells[2].Value;//saldo bitme
                        worksheet.Cells[7, 5].Value=textBox1.Text;

                        int excelColumnCount = 37;
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            int yoxla21_ = 0;
                            int yoxla9_ = 0;
                            //string deger = dataGridView1.Rows[i].Cells[8].Value.ToString().Substring(0,2);
                            //object cellValue21 = dataGridView1.Rows[i].Cells[columnNo1 - 1].Value;
                            setirsay = setirsay - 1;
                            label5.Text = setirsay.ToString();

                            for (int j = 0; j < excelColumnCount; j++)
                            {
                                object cellValue = dataGridView1.Rows[i].Cells[j].Value;

                                if ((j == 0 || j == 1) && DateTime.TryParse(cellValue?.ToString(), out DateTime parsedDate))
                                {
                                    worksheet.Cells[startRow + i, j + 1].Value = parsedDate.ToString("yyyy.MM.dd");
                                }
                                else
                                {
                                    worksheet.Cells[startRow + i, j + 1].Value = cellValue;
                                }
                                if (dataGridView1.Rows[i].Cells[columnNo1 - 1].Value != null)
                                {

                                    object cellValue21 = dataGridView1.Rows[i].Cells[columnNo1 - 1].Value;
                                    object cellValue9 = dataGridView1.Rows[i].Cells[columnNo - 1].Value;
                                    object cellValue20 = dataGridView1.Rows[i].Cells[columnNo20 - 1].Value;
                                    object cellValue9_ = dataGridView1.Rows[i].Cells[8].Value;
                                    object cellValue21_ = dataGridView1.Rows[i].Cells[columnNo21 - 1].Value;
                                    object cellValueCARI = dataGridView1.Rows[i].Cells[columnNo20 - 1].Value;
                                    object cellValuedebCARI = dataGridView1.Rows[i].Cells[columnNo9 - 2].Value;
                                    string ilkbes9 = cellValue9.ToString().Substring(0, 5);
                                    string ilkbes21 = cellValue21.ToString().Substring(0, 5);

                                    if (cellValue21 != null && cellValue21.ToString() == "25010000000000300000")
                                    {
                                        //gridmuracietler.SetRowCellValue(i, gridmuracietler.Columns[3], "ATM");
                                        dataGridView1.Rows[i].Cells[3].Value = "ATM";
                                    }
                                    if (cellValue9_ != null && cellValue9_.ToString() == "25052000040000300000")
                                    {
                                        dataGridView1.Rows[i].Cells[4].Value = "İŞÇİLƏRƏ AVANS MÜKAFAT";
                                    }
                                    if (cellValue9_ != null && cellValue9_.ToString() == "25019000000000300006")
                                    {
                                        dataGridView1.Rows[i].Cells[4].Value = "E-MANAT İLƏ ƏMƏLİYYATLAR";
                                        dataGridView1.Rows[i].Cells[11].Value = "AZE";
                                    }
                                    if (cellValue9 != null && cellValue9.ToString() == "11010000020000200000" || cellValue21 != null && cellValue21.ToString() == "11010000020000200000")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "AZP";
                                    }
                                    if (cellValue9 != null && cellValue9.ToString() == "11010000030000200000" || cellValue21 != null && cellValue21.ToString() == "11010000030000200000")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "XON";
                                    }
                                    if (cellValue9 != null && cellValue9.ToString() == "11010000050000200000" || cellValue21 != null && cellValue21.ToString() == "11010000050000200000")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "AOS";
                                    }
                                    if (ilkbes9 == "35025" || ilkbes21 == "35025")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "SFT";
                                    }
                                    if (ilkbes9 == "35020" || ilkbes21 == "35020")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "SFT";
                                    }
                                    if (ilkbes9 == "15025" || ilkbes21 == "15025")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "SFT";
                                    }
                                    if (ilkbes9 == "15020" || ilkbes21 == "15020")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "SFT";
                                    }
                                    if (ilkbes9 == "10010" || ilkbes21 == "10010")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "CAS";
                                    }
                                    if (ilkbes9 == "10020" || ilkbes21 == "10020")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "CAS";
                                    }
                                    if (ilkbes9 == "25019")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "PTR";
                                    }
                                    if (cellValue21 != null && cellValue9.ToString() == "25010000000000300000" || cellValue21 != null && cellValue21.ToString() == "25020010000000300002" || cellValue21 != null && cellValue21.ToString() == "25020020000000300002")
                                    {
                                        dataGridView1.Rows[i].Cells[3].Value = "POS";
                                    }
                                    if (cellValue21_ != null && yoxla21_ < 1)
                                    {
                                        string hesabAdi = string.Empty; // Boş bir string yaradırıq
                                        string fin = string.Empty;
                                        string voen = string.Empty;

                                        // DataTable-dən "cellValue21_" ilə eyni hesab adını əldə etmək
                                        foreach (DataRow row in Hes_adlari.Rows)
                                        {
                                            // 1-ci sütuna görə müqayisə edirik
                                            if (cellValue21_.ToString() == row[0].ToString())
                                            {
                                                hesabAdi = row[1].ToString();
                                                fin = row[2].ToString();
                                                voen = row[3].ToString();
                                                yoxla21_ = yoxla21_ + 1;
                                                break; // Hesab tapılınca loop-dan çıxırıq
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(hesabAdi))
                                        {
                                            dataGridView1.Rows[i].Cells[16].Value = hesabAdi; // Hesab adını dataGridView-ə yazırıq
                                            dataGridView1.Rows[i].Cells[18].Value = fin;
                                            dataGridView1.Rows[i].Cells[17].Value = voen;
                                        }
                                    }
                                    if (cellValue9_ != null && yoxla9_ < 1)
                                    {
                                        string hesabAdi = string.Empty; // Boş bir string yaradırıq

                                        // DataTable-dən "cellValue9_" ilə eyni hesab adını əldə etmək
                                        foreach (DataRow row in Hes_adlari.Rows)
                                        {
                                            // 1-ci sütuna görə müqayisə edirik
                                            if (cellValue9_.ToString() == row[0].ToString())
                                            {
                                                hesabAdi = row[1].ToString();
                                                yoxla9_ = yoxla9_ + 1;
                                                break; // Hesab tapılınca loop-dan çıxırıq
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(hesabAdi))
                                        {
                                            dataGridView1.Rows[i].Cells[4].Value = hesabAdi; // Hesab adını dataGridView-ə yazırıq
                                        }
                                    }

                                    if (cellValue21_ != null && cellValue21_.ToString().Length >= 2)
                                    {
                                        string firstTwoCharacters = cellValue21_.ToString().Substring(0, 2);

                                        if (firstTwoCharacters == "40" || firstTwoCharacters == "41" || firstTwoCharacters == "38" || firstTwoCharacters == "39")
                                        {
                                            if (cellValueCARI.ToString() != "P/k")
                                            {
                                                dataGridView1.Rows[i].Cells[19].Value = "Cari";
                                            }
                                        }
                                    }
                                    if (cellValue9_ != null && cellValue9_.ToString().Length >= 2)
                                    {
                                        string firstTwoCharacters = cellValue9_.ToString().Substring(0, 2);

                                        if (firstTwoCharacters == "40" || firstTwoCharacters == "41" || firstTwoCharacters == "38" || firstTwoCharacters == "39")
                                        {
                                            if (cellValueCARI.ToString() != "P/k")
                                            {
                                                dataGridView1.Rows[i].Cells[19].Value = "Cari";
                                            }
                                        }
                                    }
                                    //if (cellValue9_ != null && (cellValue9_.ToString().Substring(0, 2) == "40" || cellValue9_.ToString().Substring(0, 2) == "41" || cellValue9_.ToString().Substring(0, 2) == "38" || cellValue9_.ToString().Substring(0, 2) == "39") && cellValuedebCARI.ToString() != "P/k")
                                    //{
                                    //    //string test = cellValue9_.ToString().Substring(0, 2);
                                    //    dataGridView1.Rows[i].Cells[7].Value = "Cari";
                                    //}
                                }
                            }
                        }
                        cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                        package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet
                        System.Diagnostics.Process.Start(cl.filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Xeta oldu: " + ex.Message);
            }
            finally { }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            button1.Text = "Məlumatlar hazırlanır...";
            Thread backgroundThread = new Thread(() =>
            {

                this.Invoke((MethodInvoker)delegate
                    {
                        // Ana iş parçacığında çalışan kod burada olmalıdır

                        // Diğer kodlar burada devam eder
                        hesabad_qaliq();


                        if (radioButton1.Checked == true)
                        {
                            axtar();
                            hesabad_qaliq();
                            if (dataGridView1.RowCount == 0 || dataGridView2.RowCount <= 1)
                            {
                                button1.Text = "Sorğu";
                                MessageBox.Show("Nəticə yoxdur və ya hesab daxil etdiyiniz giriş tarixindən sonra açılıb.", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                object value = dataGridView2.Rows[0].Cells[0].Value;
                                txtad.Text = value.ToString();
                                exceleat2();
                            }
                        }
                        else if (radioButton2.Checked == true)
                        {
                            axtarhuquqi();
                            hesabad_qaliq();
                            if (dataGridView1.RowCount == 0 || dataGridView2.RowCount <= 1)
                            {
                                button1.Text = "Sorğu";
                                MessageBox.Show("Nəticə yoxdur və ya hesab daxil etdiyiniz giriş tarixindən sonra açılıb.", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                this.Invoke((MethodInvoker)delegate
                                {

                                });
                            }
                            else
                            {
                                object value = dataGridView2.Rows[0].Cells[0].Value;
                                txtad.Text = value.ToString();
                                exceleat2();
                            }
                        }
                    });
            });

            backgroundThread.Start();
        }
        private void gonder()
        {
            OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Orcon.Open();
            OracleCommand Orcom = new OracleCommand("select  x.*  from (select to_char(t.date_oper, 'dd/mm/yyyy') qeb_tarix," +
                " to_char(t.date_oper, 'dd/mm/yyyy') icra_tarix, t.nomer_docum san_nom, '  ' cat_kan," +
                " TRIM(f.soyadi) || ' ' || TRIM(f.adi) || ' ' || TRIM(f.ata_adi) gon_ad, '   ' gon_voen," +
                " f.fin gon_fin, case when substr(t.debet, 16, 1) = '9' then 'P/k'  else  'Cari'  end gon_hesnov," +
                " t.debet gon_hes, 'Bank Melli Iran' gon_bank, 'Baki fil. ' gon_fil, f.verildiyi_olke gon_olke," +
                " case when        t.kod_valuti = '00' then '944' else case when t.kod_valuti = '01'  then '840' else case when t.kod_valuti = '02'" +
                "  then '978' else case when t.kod_valuti = '03'  then '643' else case when t.kod_valuti = '04'  then '364'" +
                " else case when t.kod_valuti = '05'  then '784' end end end end end end gon_valuta, '   ' gon_bos1, '   ' gon_bos2," +
                " '   ' gon_bos3, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.kredit, 1, 1) <> '4')  " +
                "then TRIM(f.soyadi) || ' ' || TRIM(f.adi) || ' ' || TRIM(f.ata_adi) else  case when t.id_vd is null and  substr(t.kredit, 1, 1) = '4' " +
                " then  h.name_regnom  else  vd.ben_ad  end end alan_ad, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6)" +
                " or substr(t.kredit, 1, 1) <> '4')  then '    '  else case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or" +
                " substr(t.kredit, 1, 1) = '4')  then h.inn_regnom else  vd.inn_kred  end end  alan_voen, case when t.id_vd is null" +
                " and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.kredit, 1, 1) <> '4')  then  f.fin  else  case when t.id_vd is null " +
                "and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.kredit, 1, 1) = '4')  then  h.pincode else  vd.fin_kredit  end end " +
                " alan_fin, '   ' alan_hes_nov, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or" +
                " substr(t.kredit, 1, 1) <> '4')   then  t.kredit else  case when t.id_vd is null and substr(t.kredit, 1, 1) = '4'  " +
                " then  t.kredit else   vd.kredit  end  end alan_hes, case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.ben_bank  " +
                "end  alan_bank, case when  t.id_vd is null  then  'Baki fil. ' else  'diger'  end  alan_fil, case when t.id_vd is null" +
                " and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.kredit, 1, 1) <> '4')   then  f.verildiyi_olke " +
                "else  case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.kredit, 1, 1) = '4')  " +
                " then  f.verildiyi_olke else  '   '  end  end alan_olke, case when        t.kod_valuti = '00' then '944' else case" +
                " when t.kod_valuti = '01'  then '840' else case when t.kod_valuti = '02'  then '978' else case when t.kod_valuti = '03'  " +
                "then '643' else case when t.kod_valuti = '04'  then '364' else case when t.kod_valuti = '05'  then '784' end end end end end end " +
                "alan_valuta, '   ' alan_bos1, '   ' alan_bos2, '   ' alan_bos3, 0 med_val, t.summa_v_inval max_val, t.kod_valuti, 0 med_azn," +
                " t.summa_v_nacval max_azn, '  ' emel, '  ' kommun, substr(t.debet, 1, 5) dt, substr(t.kredit, 1, 5) kt, t.id_vd from odb.arh_dd t," +
                " odb.emitent_benefisiar k, odb.fiziki_shexs f, odb.regnom h,(select z.* from (select v.date_oper tarix, v.value_date, r.inn_regnom," +
                " r.pincode fin, odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name, v.nomer_docum," +
                " v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account kredit," +
                " odb.func_utf8_to_latin(v.beneficiary_name) ben_ad, v.amount, odb.func_utf8_to_latin(v.comments) cmnt, case " +
                "when v.currency = 'AZN' then '944' else case when v.currency = 'USD'  then '840' else case when v.currency = 'EUR'  then" +
                " '978' else case when v.currency = 'RUB'  then '643' else case when v.currency = 'IRR'  then '364' else case when v.currency = 'AED' " +
                " then '784' end end end end end end valuta, '  ' inn_kred, '  ' fin_kredit from odb.doc_vnesh_inval v," +
                " odb.regnom r  where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') " +
                "and substr(v.account_no,10,6)=r.regnom union" +
                " all select v.date_oper tarix, v.date_oper val_tar, r.inn_regnom, r.pincode fin," +
                " odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, 'Bank Melli Iran Bakı filialı' filial_name, v.nomer_docum,v.mfo_credit," +
                " odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit, v.summa_v_nacval amount, 'Odemeler ' cmnt, '944' valuta," +
                " v.inn_credit, '  ' fin_kredit from odb.doc_vnesh_nacval v, odb.regnom r  where v.date_oper between  to_date('textBox2.Text'," +
                " 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') and substr(v.debet,10,6)=r.regnom union" +
                " all select v.date_oper tarix," +
                " v.date_oper val_tar, v.inn_debet, '     '  fin_debet, odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet," +
                " 'Bank Melli Iran Bakı filialı' filial_name, v.nomer_docum, v.mfo_kredit, odb.func_utf8_to_latin(v.kredit_name) ben_name," +
                " v.kredit, v.kredit_name, v.sum1 amount, 'Daxilolma' cmnt, '944' valuta, r.inn_regnom, r.pincode fin_kred from odb.doc_vnesh_postupl v," +
                " odb.regnom r where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and " +
                "to_date('textBox3.Text','dd / mm / yyyy') and odb.left(odb.right(lpad(v.kredit,28,'0'),11),6)=r.regnom union" +
                " all select v.date_oper tarix," +
                " v.date_oper val_tar, '  ' inn_deb, '  ' fin_debet, odb.func_utf8_to_latin(v.sender_name) emit_name, v.sender_account debet," +
                " v.sender_bank_name filial_name, v.nomer_docum, v.beneficiary_bank_name, odb.func_utf8_to_latin(v.beneficiary_name) ben_name," +
                " v.beneficiary_account, v.beneficiary_name, v.amount, 'Daxilolma' cmnt, v.currency valuta, r.inn_regnom," +
                " r.pincode fin_kred from odb.doc_vnesh_swift v, odb.regnom r where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') " +
                "and to_date('textBox3.Text','dd / mm / yyyy') and odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),6)=r.regnom   ) z" +
                " where z.debet like '%' || &'41045000001857600000' || '%' or  z.kredit like '%' || &'41045000001857600000' || '%' order by z.tarix," +
                " z.cmnt)  vd  where t.recnum = k.doc_id(+) and t.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and" +
                " to_date('textBox3.Text','dd / mm / yyyy') and t.debet=&'41045000001857600000'  and substr(t.debet,10,6)=f.regnom(+) and " +
                " substr(t.kredit,10,6)=h.regnom(+) and t.date_oper = vd.tarix(+) and t.debet = vd.debet(+) and t.summa_v_nacval = vd.amount(+) " +
                "UNION" +
                " ALL select to_char(t.date_oper, 'dd/mm/yyyy') qeb_tarix, to_char(t.date_oper, 'dd/mm/yyyy') icra_tarix, t.nomer_docum san_nom," +
                " '  ' cat_kan, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.debet, 1, 3) in (100)) " +
                " then h.name_regnom else  case when t.id_vd is null and  substr(t.debet, 1, 1) in (4, 6, 7, 8, 9)  then" +
                "  h.name_regnom  else  vd.ben_ad  end end gon_ad, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or" +
                " substr(t.debet, 1, 3) in (100))   then h.inn_regnom  else case when t.id_vd is null " +
                "and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.debet, 1, 1) in (4, 6, 7, 8, 9))  then " +
                "'   ' else  vd.inn_regnom  end end  gon_voen, case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) " +
                "or substr(t.debet, 1, 3) in (100))   then  h.pincode--f.fin  else  case when t.id_vd is null " +
                "and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.debet, 1, 1) in (4, 6, 7, 8, 9)) " +
                " then  h.pincode  else  vd.fin  end end  alan_fin, '   ' gon_hes_nov, case when t.id_vd is null  " +
                "then  t.debet  else  trim(vd.debet) end gon_hes, case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.filial_name  " +
                "end  gon_bank, case when  t.id_vd is null  then  'Baki fil. ' else  vd.filial_name  end  gon_fil, case when t.id_vd is null " +
                "and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6)  or substr(t.debet, 1, 3) in (100))  then  f.verildiyi_olke else  " +
                "case when t.id_vd is null and(substr(t.debet, 10, 6) = substr(t.kredit, 10, 6) or substr(t.debet, 1, 1) in (4, 6, 7, 8, 9))   " +
                "then  f.verildiyi_olke  else  '   '  end  end gon_olke, case when        t.kod_valuti = '00' then" +
                " '944' else case when t.kod_valuti = '01'  then '840' else case when t.kod_valuti = '02'  then '978' else case when " +
                "t.kod_valuti = '03'  then '643' else case when t.kod_valuti = '04'  then '364' else case when t.kod_valuti = '05'  then" +
                " '784'end end end end end end alan_valuta, '   ' gon_bos1, '   ' gon_bos2, '   ' gon_bos3,TRIM(f.soyadi) || ' ' || TRIM(f.adi) ||" +
                " ' ' || TRIM(f.ata_adi) gon_ad, '   ' alan_voen, f.fin alan_fin, case when substr(t.kredit, 16, 1) = '9' then 'P/k'  else" +
                "  'Cari'  end alan_hesnov, t.kredit alan_hes, 'Bank Melli Iran' alan_bank, 'Baki fil. ' alan_fil, f.verildiyi_olke alan_olke," +
                " case when        t.kod_valuti = '00' then '944' else case when t.kod_valuti = '01'  then '840' else case when t.kod_valuti = '02'  " +
                "then '978' else case when t.kod_valuti = '03'  then '643' else case when t.kod_valuti = '04'  then '364' else case" +
                " when t.kod_valuti = '05'  then '784' end end end end end end alan_valuta, '   ' alan_bos1, '   ' alan_bos2, '   ' alan_bos3," +
                "t.summa_v_inval med_val, 0 max_val, t.kod_valuti, t.summa_v_nacval med_azn, 0 max_azn, '  ' emel, '  ' kommun," +
                " substr(t.debet, 1, 5) dt, substr(t.kredit, 1, 5) kt, t.id_vd from odb.arh_dd t, odb.emitent_benefisiar k, odb.fiziki_shexs f," +
                " odb.regnom h,(select z.* from (select v.date_oper tarix, v.value_date, r.inn_regnom, r.pincode fin," +
                " odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name, v.nomer_docum," +
                "v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account kredit," +
                " odb.func_utf8_to_latin(v.beneficiary_name) ben_ad,v.amount, odb.func_utf8_to_latin(v.comments) cmnt," +
                "case when v.currency = 'AZN' then '944' else case when v.currency = 'USD'  then '840' else case when v.currency = 'EUR'  then" +
                " '978' else case when v.currency = 'RUB'  then '643' else case when v.currency = 'IRR'  then '364' else case " +
                "when v.currency = 'AED'  then '784' end end end end end end valuta, '  ' inn_kred, '  ' fin_kredit from odb.doc_vnesh_inval v," +
                " odb.regnom r  where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') and" +
                " substr(v.account_no,10,6)=r.regnom union" +
                " all select v.date_oper tarix, v.date_oper val_tar, r.inn_regnom, r.pincode fin," +
                " odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, v.mfo_debet filial_name, v.nomer_docum,v.mfo_credit," +
                " odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit, v.summa_v_nacval amount, 'Odemeler ' cmnt," +
                " '944' valuta, v.inn_credit, '  ' fin_kredit from odb.doc_vnesh_nacval v," +
                " odb.regnom r  where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') " +
                "and substr(v.debet,10,6)=r.regnom union" +
                " all select v.date_oper tarix, v.date_oper val_tar, v.inn_debet, '     '  fin_debet," +
                " odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, v.mfo_debet filial_name, v.nomer_docum, v.mfo_kredit," +
                " odb.func_utf8_to_latin(v.kredit_name) ben_name, substr(v.kredit, 9, 20) kredit, v.kredit_name, v.sum1 amount, 'Daxilolma' cmnt," +
                " '944' valuta, r.inn_regnom, r.pincode fin_kred from odb.doc_vnesh_postupl v," +
                " odb.regnom r where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') " +
                "and odb.leftodb.(right(lpad(v.kredit,28,'0'),11),6)=r.regnom union" +
                " all select v.date_oper tarix, v.date_oper val_tar, '  ' inn_deb," +
                " '  ' fin_debet, odb.func_utf8_to_latin(v.sender_name) emit_name, v.sender_account debet, v.sender_bank_name filial_name," +
                " v.nomer_docum,v.beneficiary_bank_name, odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account, " +
                "v.beneficiary_name, v.amount, 'Daxilolma' cmnt, v.currency valuta, r.inn_regnom, r.pincode fin_kred from odb.doc_vnesh_swift v," +
                " odb.regnom r where v.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') and" +
                " odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),6)=r.regnom   ) z where z.debet like '%' || &'41045000001857600000' || '%' or" +
                "  z.kredit like '%' || &'41045000001857600000' || '%' order by z.tarix, z.cmnt)  vd where t.recnum = k.doc_id(+) and" +
                " t.date_oper between  to_date('textBox2.Text', 'dd/mm/yyyy') and to_date('textBox3.Text','dd / mm / yyyy') and" +
                " t.kredit=&'41045000001857600000'  and substr(t.kredit,10,6)=f.regnom(+) and  substr(t.debet, 10, 6) = h.regnom(+) and" +
                " t.date_oper = vd.tarix(+) and t.kredit = vd.kredit(+) and t.summa_v_nacval = vd.amount(+))  x order by x.qeb_tarix asc", Orcon);
            OracleDataAdapter Orda = new OracleDataAdapter(Orcom);

            DataTable Ordt = new DataTable();
            Orda.Fill(Ordt);
            dataGridView1.DataSource = Ordt;
            Orcon.Close();
        }
        private void axtar()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select  x.*  from  " +
" (select t.date_oper qeb_tarix, to_char(t.date_oper,'dd/mm/yyyy') icra_tarix, t.recnum san_nom,'  ' cat_kan," +
"        TRIM(f.name_regnom)  gon_ad,'   ' gon_voen, f.pincode gon_fin," +
"       case when substr(t.debet,16,1)='9' then 'P/k'  else  'Cari'  end gon_hesnov,t.debet gon_hes ,'Bank Melli Iran' gon_bank, 'Baki fil. ' gon_fil,p.countrycode gon_olke," +
"       case when         t.kod_valuti='00' then '944' " +
"       else case when t.kod_valuti='01'  then '840' " +
"       else case when t.kod_valuti='02'  then '978' " +
"       else case when t.kod_valuti='03'  then '643' " +
"       else case when t.kod_valuti='04'  then '364'          " +
"       else case when t.kod_valuti='05'  then '784'   " +
"       end end end end end end gon_valuta,'   ' gon_bos1,'   ' gon_bos2, vd.sender_bic ,   " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')  then TRIM(f.name_regnom) " +
"               else  case when t.id_vd is null and  substr(t.kredit,1,1)='4'  then  h.name_regnom " +
"               else  vd.ben_ad  end end alan_ad," +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')  then '    '  " +
"               else case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)='4')  then h.inn_regnom " +
"               else  vd.inn_kred  end end  alan_voen, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')  then  f.pincode " +
"               else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)='4')  then  h.pincode  " +
"               else  vd.fin_kredit  end end  alan_fin, '   ' alan_hes_nov," +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')   then  t.kredit " +
"               else  case when t.id_vd is null and substr(t.kredit,1,1)='4'   then  t.kredit " +
"               else   vd.kredit  end  end alan_hes ," +
"       case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.ben_bank  end  alan_bank," +
"       case when  t.id_vd is null  then  'Baki fil. ' else  'diger'  end  alan_fil, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')   then  s.countrycode " +
"               else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)='4')   then  s.countrycode " +
"               else  '   '  end  end alan_olke,  " +
"       case when        t.kod_valuti='00' then '944' " +
"       else case when t.kod_valuti='01'  then '840' " +
"       else case when t.kod_valuti='02'  then '978' " +
"       else case when t.kod_valuti='03'  then '643' " +
"       else case when t.kod_valuti='04'  then '364'          " +
"       else case when t.kod_valuti='05'  then '784'   " +
"       end end end end end end alan_valuta,'   ' alan_bos1,'   ' alan_bos2, vd.receiver_bic,  " +
 "    0 med_val,t.summa_v_inval max_val, t.kod_valuti, 0 med_azn,t.summa_v_nacval max_azn,  t.primechanie emel, '  ' kommun, substr(t.debet,1,5) dt,substr(t.kredit,1,5) kt ,t.debet dbtam,t.kredit krtam, t.id_vd  " +
 "    from odb.arh_dd t, odb.emitent_benefisiar k, odb.regnom f, odb.regnom h, odb.licsch p, odb.licsch s," +
" (select z.*  from " +
" (select v.date_oper tarix,v.value_date, r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name,  v.beneficiary_account kredit,odb.func_utf8_to_latin(v.beneficiary_name) ben_ad, " +
" v.amount, odb.func_utf8_to_latin(v.comments) cmnt, " +
  "      case when        v.currency='AZN' then '944'  " +
    "    else case when v.currency='USD'  then '840'  " +
     "   else case when v.currency='EUR'  then '978'  " +
      "  else case when v.currency='RUB'  then '643'  " +
      " else case when v.currency='IRR'  then '364' " +
      " else case when v.currency='AED'  then '784'    " +
     " end end end end end end valuta,'  ' inn_kred, '  ' fin_kredit , v.sender_bic, v.receiver_bic  " +
    " from odb.doc_vnesh_inval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.account_no,10,6)=r.regnom(+)  " +

" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, mf.bank_large_name  filial_name, v.nomer_docum, " +
" v.mfo_credit,odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit,v.summa_v_nacval amount, 'Odemeler ' cmnt, '944' valuta,v.inn_credit, '  ' fin_kredit, ' ' sender_bic, ' ' receiver_bic  " +
 " from odb.doc_vnesh_nacval v, odb.regnom r, odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.debet,10,6)=r.regnom  and  v.mfo_credit=mf.mfo(+) " +
" union all " +
" select v.date_oper tarix, v.date_oper val_tar,v.inn_debet,'     '  fin_debet,odb.func_utf8_to_latin(v.name_debet) emit_name,v.debet, mf.bank_large_name  filial_name,  v.nomer_docum, " +
"  v.mfo_kredit,odb.func_utf8_to_latin(v.kredit_name) ben_name,v.kredit,v.kredit_name,v.sum1 amount,  'Daxilolma' cmnt, '944' valuta,r.inn_regnom, r.pincode fin_kred, v.bic_debet,v.bic_kredit   " +
" from odb.doc_vnesh_postupl v, odb.regnom r,odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.kredit,28,'0'),11),6)=r.regnom and v.mfo_debet=mf.mfo(+) " +

" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,'  ' inn_deb, '  ' fin_debet , odb.func_utf8_to_latin(v.sender_name) emit_name,v.sender_account debet, v.sender_bank_name filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name,odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account,v.beneficiary_name ,v.amount, 'Daxilolma' cmnt, v.currency valuta,r.inn_regnom, r.pincode fin_kred,v.sender_bank_bic,v.beneficiary_bank_bic " +
" from odb.doc_vnesh_swift v, odb.regnom r where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),6)=r.regnom ) z  " +
" where z.debet like '%'||'" + textBox1.Text + "'||'%' or  z.kredit like '%'||'" + textBox1.Text + "'||'%' " +
" order by z.tarix ,z.cmnt )  vd " +
"   where t.recnum=k.doc_id(+) and t.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and t.debet='" + textBox1.Text + "'  and substr(t.debet,10,6)=f.regnom(+) and  substr(t.kredit,10,6)=h.regnom(+) and t.debet=p.licsch(+)  and t.kredit=s.licsch(+)  " +
"      and t.date_oper=vd.tarix(+) and t.debet=vd.debet(+) and t.summa_v_nacval=vd.amount(+) and (t.vid_operacii <> 97 or t.vid_operacii is null)  " +

" UNION ALL    " +

" select t.date_oper qeb_tarix, to_char(t.date_oper,'dd/mm/yyyy') icra_tarix,  t.recnum san_nom,'  ' cat_kan, " +
 "       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150,350))  then h.name_regnom  " +
  "              else  case when t.id_vd is null and  substr(t.debet,1,1) in (4,6,7,8,9)  then  h.name_regnom  " +
   "             else  vd.ben_ad  end end gon_ad, " +
  "      case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150,350))   then h.inn_regnom  " +
  "              else case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))  then s.inn_licsch  " +
  "              else  vd.inn_regnom  end end  gon_voen,  " +
  "      case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150,350))   then  h.pincode   " +
  "              else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))  then  h.pincode   " +
  "              else  vd.fin  end end  alan_fin, '   ' gon_hes_nov,  " +
  "      case when t.id_vd is null  then  t.debet  else  trim(vd.debet) end gon_hes , " +
  "      case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.filial_name  end  gon_bank, " +
  "      case when  t.id_vd is null  then  'Baki fil. ' else  vd.filial_name  end  gon_fil,  " +
  "      case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6)  or substr(t.debet,1,3) in (100,150,350))  then  p.countrycode  " +
  "              else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))   then  s.countrycode  else  '   '  end  end gon_olke,  " +
  "      case when        t.kod_valuti='00' then '944'  " +
  "      else case when t.kod_valuti='01'  then '840'  " +
  "      else case when t.kod_valuti='02'  then '978'  " +
  "      else case when t.kod_valuti='03'  then '643'  " +
  "      else case when t.kod_valuti='04'  then '364'           " +
  "      else case when t.kod_valuti='05'  then '784'    " +
  "      end end end end end end alan_valuta,  '   ' gon_bos1,'   ' gon_bos2, vd.sender_bic ,  " +
  "      TRIM(f.name_regnom) gon_ad,'   ' alan_voen, f.pincode alan_fin, " +
  "      case when substr(t.kredit,16,1)='9' then 'P/k'  else  'Cari'  end alan_hesnov,t.kredit alan_hes ,'Bank Melli Iran' alan_bank, 'Baki fil. ' alan_fil,p.countrycode alan_olke, " +
  "      case when        t.kod_valuti='00' then '944'  " +
  "      else case when t.kod_valuti='01'  then '840'  " +
  "      else case when t.kod_valuti='02'  then '978'  " +
  "      else case when t.kod_valuti='03'  then '643'  " +
  "      else case when t.kod_valuti='04'  then '364'           " +
  "      else case when t.kod_valuti='05'  then '784'    " +
  "      end end end end end end alan_valuta,'   ' alan_bos1,'   ' alan_bos2, vd.receiver_bic ,  " +

  "    t.summa_v_inval med_val,0 max_val, t.kod_valuti,t.summa_v_nacval med_azn,0 max_azn,  t.primechanie emel, '  ' kommun, substr(t.debet,1,5) dt,substr(t.kredit,1,5) kt ,t.debet dbtam,t.kredit krtam, t.id_vd   " +
  "      from odb.arh_dd t, odb.emitent_benefisiar k, odb.regnom f, odb.regnom h,   odb.licsch p, odb.licsch s, " +
" (select z.*  from " +
" (select v.date_oper tarix,v.value_date, r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name,  v.beneficiary_account kredit,odb.func_utf8_to_latin(v.beneficiary_name) ben_ad, " +
" v.amount, odb.func_utf8_to_latin(v.comments) cmnt, " +
"        case when         v.currency='AZN' then '944'  " +
 "       else case when v.currency='USD'  then '840'  " +
  "      else case when v.currency='EUR'  then '978'  " +
   "     else case when v.currency='RUB'  then '643'  " +
   "     else case when v.currency='IRR'  then '364'           " +
   "     else case when v.currency='AED'  then '784'    " +
   "     end end end end end end valuta,'  ' inn_kred, '  ' fin_kredit , v.sender_bic, v.receiver_bic  " +
   "  from odb.doc_vnesh_inval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.account_no,10,6)=r.regnom  " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, v.mfo_debet filial_name, v.nomer_docum, " +
" v.mfo_credit,odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit,v.summa_v_nacval amount, 'Odemeler ' cmnt, '944' valuta,v.inn_credit, '  ' fin_kredit ,' ' sender_bic, ' ' receiver_bic   " +
"  from odb.doc_vnesh_nacval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.debet,10,6)=r.regnom " +
" union all " +
" select v.date_oper tarix, v.date_oper val_tar,v.inn_debet,'     '  fin_debet,odb.func_utf8_to_latin(v.name_debet) emit_name,v.debet,mf.bank_large_name  filial_name, v.nomer_docum, " +
"  v.mfo_kredit,odb.func_utf8_to_latin(v.kredit_name) ben_name,substr(v.kredit,9,20) kredit,v.kredit_name,v.sum1 amount,  'Daxilolma' cmnt, '944' valuta,r.inn_regnom, r.pincode fin_kred, v.bic_debet,v.bic_kredit  " +
" from odb.doc_vnesh_postupl v, odb.regnom r, odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.kredit,28,'0'),11),6)=r.regnom and v.mfo_debet=mf.mfo(+) " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,'  ' inn_deb, '  ' fin_debet , odb.func_utf8_to_latin(v.sender_name) emit_name,v.sender_account debet, v.sender_bank_name filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name,odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account,v.beneficiary_name ,v.amount, 'Daxilolma' cmnt, v.currency valuta,r.inn_regnom, r.pincode fin_kred,v.sender_bank_bic,v.beneficiary_bank_bic " +
" from odb.doc_vnesh_swift v, odb.regnom r where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),6)=r.regnom   ) z  " +
" where z.debet like '%'||'" + textBox1.Text + "'||'%' or  z.kredit like '%'||'" + textBox1.Text + "'||'%' " +
" order by z.tarix ,z.cmnt )  vd " +
"   where t.recnum=k.doc_id(+) and t.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and t.kredit='" + textBox1.Text + "'  and substr(t.kredit,10,6)=f.regnom(+) and t.kredit=p.licsch(+)  and t.debet=s.licsch(+)  " +
"    and  substr(t.debet,10,6)=h.regnom(+) " +
"    and t.date_oper=vd.tarix(+) and t.kredit=vd.kredit(+) and t.summa_v_nacval=vd.amount(+)  and (t.vid_operacii <> 97 or t.vid_operacii is null)   )  x " +
"    order by x.qeb_tarix asc ", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Text = "Sorğu";
            }
            finally { }
        }
        private void hesabad_qaliq()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select ac.name_latin,p.gir_qaliq,k.son_qaliq from " +
" (select t.licsch,abs(t.saldo_vhd_inval) gir_qaliq from odb.arh_saldo_ls t where t.licsch='" + textBox1.Text + "' and t.date_oper=odb.ish_gun_cari1(to_date('" + textBox2.Text + "', 'dd/mm/yyyy'))) p, " +
" (select t.licsch,abs(t.saldo_ish_inval) son_qaliq    from odb.arh_saldo_ls t where t.licsch='" + textBox1.Text + "' and t.date_oper=odb.ish_gun_cari1(to_date('" + textBox3.Text + "', 'dd/mm/yyyy'))) k, odb.accounts ac " +
" where  p.licsch=k.licsch and p.licsch=ac.licsch", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView2.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { }
        }
        private void hesabad_adlari()
        {
            try
            {
                DataTable Ordt = new DataTable();
                Ordt.Clear();
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select  t.licsch,t.name_licsch,r.pincode from odb.licsch t,odb.regnom r " +
                " where substr(t.licsch, 11, 5) = substr(r.regnom, 2.5) and t.date_close_licsch is null " +
                " order by t.licsch", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                Orcon.Close();
            }
            catch (Exception)
            {
                // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { }
        }
        private void axtarhuquqi()
        {
            DataTable Ordt = new DataTable();
            Ordt.Clear();
            try
            {
                OracleConnection Orcon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select  x.*  from  " +
" (select t.date_oper qeb_tarix, to_char(t.date_oper,'dd/mm/yyyy') icra_tarix, t.recnum san_nom,'  ' cat_kan, " +
"       TRIM(f.name_regnom)  gon_ad, f.inn_regnom  gon_voen, f.pincode gon_fin, " +
"       case when substr(t.debet,16,1)='9' then 'P/k'  else  'Cari'  end gon_hesnov,t.debet gon_hes ,'Bank Melli Iran' gon_bank, 'Baki fil. ' gon_fil, p.countrycode gon_olke, " +
"       case when        t.kod_valuti='00' then '944'  " +
"       else case when t.kod_valuti='01'  then '840'  " +
"       else case when t.kod_valuti='02'  then '978'  " +
"       else case when t.kod_valuti='03'  then '643'  " +
"       else case when t.kod_valuti='04'  then '364'           " +
"       else case when t.kod_valuti='05'  then '784'    " +
"       end end end end end end gon_valuta, '   ' gon_bos1, '   ' gon_bos2, vd.sender_bic ,    " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150))  then TRIM(f.name_regnom)  " +
"               else  case when t.id_vd is null and  substr(t.debet,1,1) in (4,6,7,8,9)  then  h.name_regnom  " +
"               else  vd.ben_ad  end end  alan_ad, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150)) then f.inn_regnom " +
"               else case when t.id_vd is null and substr(t.debet,1,1) in (4,6,7,8,9)  then h.inn_regnom  " +
"               else  vd.inn_kred  end end  alan_voen,  " +
"       case when t.id_vd is null and  (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150))  then  f.pincode   " +
"               else  case when t.id_vd is null and substr(t.debet,1,1) in (4,6,7,8,9)  then  h.pincode   " +
"               else  vd.fin_kredit  end end  alan_fin, '   ' alan_hes_nov, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')   then  t.kredit  " +
"               else  case when t.id_vd is null and substr(t.kredit,1,1)='4'   then  t.kredit  " +
"               else   vd.kredit  end  end alan_hes , " +
"     case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.filial_name  end  alan_bank, " +
"       case when  t.id_vd is null  then  'Baki fil. ' else  vd.filial_name  end  alan_fil,  " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)<>'4')   then  s.countrycode  " +
"               else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.kredit,1,1)='4')   then  s.countrycode  " +
"               else  '   '  end  end alan_olke,   " +
"       case when        t.kod_valuti='00' then '944'  " +
"       else case when t.kod_valuti='01'  then '840'  " +
"       else case when t.kod_valuti='02'  then '978'  " +
"       else case when t.kod_valuti='03'  then '643'  " +
"       else case when t.kod_valuti='04'  then '364'           " +
"       else case when t.kod_valuti='05'  then '784'    " +
"       end end end end end end alan_valuta,'   ' alan_bos1,'   ' alan_bos2, vd.receiver_bic,                         " +
"     0 med_val,t.summa_v_inval max_val, t.kod_valuti, 0 med_azn,t.summa_v_nacval max_azn, t.primechanie emel, '  ' kommun, substr(t.debet,1,5) dt,substr(t.kredit,1,5) kt, t.debet dbtam,t.kredit krtam   " +
"       from odb.arh_dd t, odb.emitent_benefisiar k, odb.regnom f, odb.regnom h, odb.licsch p, odb.licsch s, " +
" (select z.*  from " +
" (select v.date_oper tarix,v.value_date, r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name,  v.beneficiary_account kredit,odb.func_utf8_to_latin(v.beneficiary_name) ben_ad, " +
" v.amount, odb.func_utf8_to_latin(v.comments) cmnt, " +
"       case when         v.currency='AZN' then '944'  " +
"       else case when v.currency='USD'  then '840'  " +
"       else case when v.currency='EUR'  then '978'  " +
"       else case when v.currency='RUB'  then '643'  " +
"      else case when v.currency='IRR'  then '364'           " +
"      else case when v.currency='AED'  then '784'    " +
"     end end end end end end valuta,'  ' inn_kred, '  ' fin_kredit , v.sender_bic,v.receiver_bic  " +
"    from odb.doc_vnesh_inval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.account_no,10,6)=r.regnom(+)  " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, mf.bank_large_name  filial_name, v.nomer_docum, " +
" v.mfo_credit,odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit,v.summa_v_nacval amount, 'Odemeler ' cmnt, '944' valuta,v.inn_credit, '  ' fin_kredit, ' ' sender_bic, ' ' receiver_bic  " +
"  from odb.doc_vnesh_nacval v, odb.regnom r, odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.debet,10,6)=r.regnom  and  v.mfo_credit=mf.mfo(+) " +
" union all " +
" select v.date_oper tarix, v.date_oper val_tar,v.inn_debet,'     '  fin_debet,odb.func_utf8_to_latin(v.name_debet) emit_name,v.debet, mf.bank_large_name  filial_name,  v.nomer_docum, " +
" v.mfo_kredit,odb.func_utf8_to_latin(v.kredit_name) ben_name,v.kredit,v.kredit_name,v.sum1 amount,  'Daxilolma' cmnt, '944' valuta,r.inn_regnom, r.pincode fin_kred, v.bic_debet,v.bic_kredit   " +
" from odb.doc_vnesh_postupl v, odb.regnom r,odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.kredit,28,'0'),11),6)=r.regnom and v.mfo_debet=mf.mfo(+) " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,'  ' inn_deb, '  ' fin_debet , odb.func_utf8_to_latin(v.sender_name) emit_name,v.sender_account debet, v.sender_bank_name filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name,odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account,v.beneficiary_name ,v.amount, 'Daxilolma' cmnt, v.currency valuta,r.inn_regnom, r.pincode fin_kred,v.sender_bank_bic,v.beneficiary_bank_bic " +
" from odb.doc_vnesh_swift v, odb.regnom r where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),6)=r.regnom ) z  " +
" where z.debet like '%'||'" + textBox1.Text + "'||'%' or  z.kredit like '%'||'" + textBox1.Text + "'||'%' " +
" order by z.tarix ,z.cmnt )  vd " +
"  where t.recnum=k.doc_id(+) and t.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and t.debet='" + textBox1.Text + "'  and substr(t.debet,10,6)=f.regnom(+) and  substr(t.kredit,10,6)=h.regnom(+) and t.debet=p.licsch(+)  and t.kredit=s.licsch(+)  " +
"     and t.date_oper=vd.tarix(+) and t.debet=vd.debet(+) and t.summa_v_nacval=vd.amount(+) and (t.vid_operacii <> 97 or t.vid_operacii is null) " +
" UNION ALL    " +
" select t.date_oper qeb_tarix, to_char(t.date_oper,'dd/mm/yyyy') icra_tarix, t.recnum san_nom,'  ' cat_kan, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150))  then h.name_regnom  " +
"               else  case when t.id_vd is null and  substr(t.debet,1,1) in (4,6,7,8,9)  then  h.name_regnom  " +
"               else  vd.emit_name  end end gon_ad, " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150))   then h.inn_regnom  " +
"               else case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))  then f.inn_regnom  " +
"               else  vd.inn_regnom  end end  gon_voen,  " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,3) in (100,150))   then  h.pincode    " +
"               else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))  then  h.pincode   " +
"               else  vd.fin  end end  gon_fin, '   ' gon_hes_nov,  " +
"       case when  t.id_vd is null  then  t.debet  else  trim(vd.debet) end gon_hes , " +
"       case when  t.id_vd is null  then  'Bank Melli Iran'  else vd.filial_name  end  gon_bank, " +
"       case when  t.id_vd is null  then  'Baki fil. ' else  vd.filial_name  end  gon_fil,  " +
"       case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6)  or substr(t.debet,1,3) in (100,150))  then  p.countrycode  " +
"               else  case when t.id_vd is null and (substr(t.debet,10,6)=substr(t.kredit,10,6) or substr(t.debet,1,1) in (4,6,7,8,9))   then  s.countrycode  else  '   '  end  end gon_olke,  " +
"       case when        t.kod_valuti='00' then '944'  " +
"       else case when t.kod_valuti='01'  then '840'  " +
"       else case when t.kod_valuti='02'  then '978'  " +
"       else case when t.kod_valuti='03'  then '643'  " +
"       else case when t.kod_valuti='04'  then '364'           " +
"       else case when t.kod_valuti='05'  then '784'    " +
"       end end end end end end gon_valuta,  '   ' gon_bos1,'   ' gon_bos2,vd.sender_bic , " +
"       TRIM(f.name_regnom) alan_ad, f.inn_regnom alan_voen, f.pincode alan_fin, " +
"       case when substr(t.kredit,16,1)='9' then 'P/k'  else  'Cari'  end alan_hesnov,t.kredit alan_hes ,'Bank Melli Iran' alan_bank, 'Baki fil. ' alan_fil,p.countrycode alan_olke, " +
 "      case when        t.kod_valuti='00' then '944'  " +
 "      else case when t.kod_valuti='01'  then '840'  " +
"       else case when t.kod_valuti='02'  then '978'  " +
"       else case when t.kod_valuti='03'  then '643'  " +
"       else case when t.kod_valuti='04'  then '364'           " +
"       else case when t.kod_valuti='05'  then '784'    " +
"       end end end end end end alan_valuta,'   ' alan_bos1,'   ' alan_bos2,vd.receiver_bic , " +
"     t.summa_v_inval med_val,0 max_val, t.kod_valuti,t.summa_v_nacval med_azn,0 max_azn, t.primechanie emel, '  ' kommun, substr(t.debet,1,5) dt,substr(t.kredit,1,5) kt,t.debet dbtam,t.kredit krtam  " +
"       from odb.arh_dd t, odb.emitent_benefisiar k, odb.regnom f, odb.regnom h,  odb.licsch p, odb.licsch s, " +
" (select z.*  from " +
" (select v.date_oper tarix,v.value_date, r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.account_name) emit_name, v.account_no debet, v.filial_name,v.nomer_docum," +
" v.beneficiary_bank_name ben_bank, odb.func_utf8_to_latin(v.beneficiary_name) ben_name,  v.beneficiary_account kredit,odb.func_utf8_to_latin(v.beneficiary_name) ben_ad, " +
" v.amount, odb.func_utf8_to_latin(v.comments) cmnt, " +
 "      case when         v.currency='AZN' then '944'  " +
 "      else case when v.currency='USD'  then '840'  " +
 "      else case when v.currency='EUR'  then '978'  " +
 "      else case when v.currency='RUB'  then '643'  " +
 "     else case when v.currency='IRR'  then '364'           " +
 "     else case when v.currency='AED'  then '784'    " +
 "    end end end end end end valuta,'  ' inn_kred, '  ' fin_kredit, v.sender_bic,v.receiver_bic   " +
 "   from odb.doc_vnesh_inval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.account_no,10,6)=r.regnom  " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,r.inn_regnom,r.pincode fin, odb.func_utf8_to_latin(v.name_debet) emit_name, v.debet, v.mfo_debet filial_name, v.nomer_docum, " +
" v.mfo_credit,odb.func_utf8_to_latin(v.name_credit) ben_name, v.kredit, v.name_credit,v.summa_v_nacval amount, 'Odemeler ' cmnt, '944' valuta,v.inn_credit, '  ' fin_kredit ,' ' sender_bic, ' ' receiver_bic   " +
" from odb.doc_vnesh_nacval v, odb.regnom r  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and substr(v.debet,10,6)=r.regnom " +
" union all " +
" select v.date_oper tarix, v.date_oper val_tar,v.inn_debet,'     '  fin_debet,odb.func_utf8_to_latin(v.name_debet) emit_name,v.debet,mf.bank_large_name  filial_name, v.nomer_docum, " +
" v.mfo_kredit,odb.func_utf8_to_latin(v.kredit_name) ben_name,odb.right(v.kredit,20) kredit,v.kredit_name,v.sum1 amount,  'Daxilolma' cmnt, '944' valuta,r.inn_regnom, r.pincode fin_kred, v.bic_debet,v.bic_kredit  " +
" from odb.doc_vnesh_postupl v, odb.regnom r, odb.mfo mf  where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(v.kredit,11),6)=r.regnom and v.mfo_debet=mf.mfo(+) " +
" union all " +
" select v.date_oper tarix,  v.date_oper val_tar,'  ' inn_deb, '  ' fin_debet , odb.func_utf8_to_latin(v.sender_name) emit_name,v.sender_account debet, v.sender_bank_name filial_name,v.nomer_docum, " +
" v.beneficiary_bank_name,odb.func_utf8_to_latin(v.beneficiary_name) ben_name, v.beneficiary_account,v.beneficiary_name ,v.amount, 'Daxilolma' cmnt, v.currency valuta,r.inn_regnom, r.pincode fin_kred,v.sender_bank_bic,v.beneficiary_bank_bic " +
" from odb.doc_vnesh_swift v, odb.regnom r where v.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and odb.left(odb.right(lpad(v.beneficiary_account,28,'0'),11),5)=r.regnom   ) z  " +
" where z.debet like '%'||'" + textBox1.Text + "'||'%' or  z.kredit like '%'||'" + textBox1.Text + "'||'%' " +
" order by z.tarix ,z.cmnt )  vd " +
"  where t.recnum=k.doc_id(+) and t.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd / mm / yyyy') and t.kredit='" + textBox1.Text + "'   and substr(t.kredit,10,6)=f.regnom(+)  and t.kredit=p.licsch(+)  and t.debet=s.licsch(+)  " +
"   and  substr(t.debet,10,6)=h.regnom(+) " +
"   and t.date_oper=vd.tarix(+) and t.kredit=vd.kredit(+) and t.summa_v_nacval=vd.amount(+)  and (t.vid_operacii <> 97 or t.vid_operacii is null)   )  x  " +
 "  order by x.qeb_tarix asc", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                Orda.Fill(Ordt);
                dataGridView1.DataSource = Ordt;
                Orcon.Close();
            }
            catch (Exception)
            {
                button1.Text = "Sorğu";
                // MessageBox.Show("Xəta baş verdi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { }
        }
        private bool IsValidDate(string input)
        {
            // Örnek olarak tarih doğrulaması yapabilirsiniz. Bu, gerçek bir tarih doğrulaması değildir.
            // Gerçek tarih doğrulaması yapmak için DateTime.TryParse kullanmalısınız.
            string[] parts = input.Split('-');
            if (parts.Length != 3)
            {
                return false;
            }

            int day, month, year;
            if (!int.TryParse(parts[0], out day) || !int.TryParse(parts[1], out month) || !int.TryParse(parts[2], out year))
            {
                return false;
            }

            if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900)
            {
                return false;
            }

            return true;
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
        private void frmhesabsorgu_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;

            // İlk TextBox'a Enter tuşuna basıldığında bir sonraki TextBox'a odaklanma işlemini ekle
            textBox1.KeyDown += textBox1_KeyDown;

            // İkinci TextBox'a Enter tuşuna basıldığında bir sonraki işlemi gerçekleştir
            textBox2.KeyDown += textBox2_KeyDown;
            textBox3.KeyDown += textBox3_KeyDown;
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                textBox2.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
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
        private void OpenInExcel(DataGridView dataGridView)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Veriler");

                // DataGridView sütun başlıklarını Excel'e aktar
                for (int col = 1; col <= dataGridView.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dataGridView.Columns[col - 1].HeaderText;
                }

                // DataGridView verilerini Excel'e aktar
                for (int row = 0; row < dataGridView.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataGridView.Rows[row].Cells[col].Value;
                    }
                }

                // Excel dosyasını açma işlemi
                string excelFilePath = "temp.xlsx"; // Geçici bir dosya yolunu belirleyebilirsiniz
                package.SaveAs(new FileInfo(excelFilePath));

                // Excel dosyasını varsayılan Excel uygulamasıyla aç
                Process.Start(excelFilePath);
            }
        }
    }
}


