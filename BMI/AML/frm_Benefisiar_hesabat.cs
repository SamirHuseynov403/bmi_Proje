using BMI.Muhasibat;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using Excel = Microsoft.Office.Interop.Excel;
using Excel.FinancialFunctions;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;

namespace BMI.AML
{
    public partial class frm_Benefisiar_hesabat : Form
    {
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        cl_yanasmalar cl = new cl_yanasmalar();
        public frm_Benefisiar_hesabat()
        {
            InitializeComponent();
        }
        private void Hesabati_hazirla_ve_excele_at()
        {
            #region sql_aciq huquqi sexs hasablari
            string h_h_hesablar = @"select r.name_regnom h_h_adi,h.incorporation_country_code olke,r.inn_regnom voen,r.regnom qeyd_no
                    from regnom r,huquqi_shexs h,
                    (select distinct l.registrac_nomer qeyd_no from licsch l
                    where l.date_close_licsch is null and substr(l.licsch,1,1) in ('3','4'))hes
                    where r.regnom=hes.qeyd_no and r.yurik=1 and r.regnom=h.regnom
                    ";
            #endregion
            #region Benefisiarlar
            string benefisiarlar = @"select h.qeydiyyat_tarixi q_tar,i.soyadi soyad,i.adi ad,i.ata_adi ata,i.fin fin,i.vetendashligi vettendasliq,
                i.olke olke,i.doguldugu_tarix d_tar,t.izah,
                REGEXP_SUBSTR(i.seriyasi_ve_nomresi, '^[A-Z]+') seriya,REGEXP_SUBSTR(i.seriyasi_ve_nomresi, '[0-9]+') seriya_no,
                i.verilme_tarixi ver_tar_ves,i.tesischinin_payi pay,'təsisçi',i.regnom qeyd_no
                from imza_huquqi_olan_shexsler i,huquqi_shexs h,types_of_documents t
                where i.tesischinin_payi>=10
                and i.regnom=h.regnom and t.isare=senedin_novu";
            #endregion

            System.Data.DataTable _dt_hesablar = new System.Data.DataTable();
            System.Data.DataTable _dt_benefisiarlar = new System.Data.DataTable();

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                connection.Open();
                using (OracleCommand command = new OracleCommand(h_h_hesablar, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_hesablar);
                }
                using (OracleCommand command = new OracleCommand(benefisiarlar, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_benefisiarlar);
                }
                connection.Close();
            }

            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            cl.fileName = "Benefisiar mülkiyətçi" + ".xlsx";
            cl.templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "AML", "Exceller", "benefisiar.xlsx");
            cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);

            if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, cl.fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.baseFileName} - {fileCounter}.xlsx")))
                {
                    fileCounter++;
                }
                cl.fileName = $"{cl.baseFileName} - {fileCounter}.xlsx";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(cl.templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Əlavə "];

                int startRow = 5;
                int startColumn = 6;

                int maxBenefCount = _dt_hesablar.AsEnumerable()
                .Select(row =>
                {
                    string qeyd_no = row["qeyd_no"].ToString();
                    return _dt_benefisiarlar.AsEnumerable()
                        .Count(b => b["qeyd_no"].ToString() == qeyd_no);
                })
                .Max();

                if (maxBenefCount > 2)
                {
                    int fromRow = 3;
                    int toRow = 4;
                    int fromCol = 19;  // S
                    int toCol = 31;    // AE

                    int targetStartCol = 32; // AF = 32-ci sütun

                    for (int i = 0; i < maxBenefCount - 2; i++) // artıq 2-si əsasda var, əlavə qədər davam et
                    {
                        int pasteFromCol = targetStartCol + (13 * i);
                        int pasteToCol = pasteFromCol + (toCol - fromCol); // 13 sütun genişlik

                        worksheet1.Cells[fromRow, fromCol, toRow, toCol] // mənbə
                                   .Copy(worksheet1.Cells[fromRow, pasteFromCol, toRow, pasteToCol]); // hədəf
                    }
                }


                for (int i = 0; i < _dt_hesablar.Rows.Count; i++)
                {
                    string qeyd_no = _dt_hesablar.Rows[i]["qeyd_no"].ToString();
                    var benefRows = _dt_benefisiarlar.AsEnumerable().Where(b => b["qeyd_no"].ToString() == qeyd_no).ToList();
                    if (benefRows.Count > 0)
                    {
                        worksheet1.Cells[startRow, 2].Value = _dt_hesablar.Rows[i][0].ToString();
                        worksheet1.Cells[startRow, 3].Value = _dt_hesablar.Rows[i][1].ToString();
                        worksheet1.Cells[startRow, 4].Value = _dt_hesablar.Rows[i][2].ToString();
                        worksheet1.Cells[startRow, 5].Value = "";
                        for (int j = 0; j < benefRows.Count; j++)
                        {
                            var benef = benefRows[j];
                            worksheet1.Cells[startRow, 1].Value = benef[0].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[1].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[2].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[3].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[4].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[5].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[6].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[7].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[8].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[9].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[10].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[11].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[12].ToString();
                            worksheet1.Cells[startRow, startColumn++].Value = benef[13].ToString();
                        }
                        startColumn = 6;
                    }
                    startRow++;
                }

                cl.filePath = System.IO.Path.Combine(cl.dosyayolu, cl.fileName);
                package.SaveAs(new FileInfo(cl.filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(cl.filePath);
                label1.Visible = false;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            label1.Visible = true;
            label1.Text = "Hazırlanır gözləyin...";
            Application.DoEvents();
            Hesabati_hazirla_ve_excele_at();
            label1.Text = "Hazırdır.";
            Application.DoEvents();
        }
    }
}
