using OfficeOpenXml;
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

namespace BMI.Muhasibat
{
    public partial class frm_elaqeli_dep_cem : Form
    {
        public frm_elaqeli_dep_cem()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;
        private void excel()
        {
            #region sql_kodlar

            
            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            //duz olan
            string sirketin_tesiscileri = "select ar.date_oper," +
                " l.licsch,-ar.saldo_vhd_inval Valyuta_ile,-ar.saldo_ish_nacval AZN_ile,l.name_licsch,i.regnom" +
                "  from licsch l,odb.arh_saldo_ls ar,regnom r,imza_huquqi_olan_shexsler i" +
                " where r.regnom=l.registrac_nomer and ( l.date_close_licsch is null or ar.date_oper<= l.date_close_licsch )  " +
                " and substr(l.licsch,0,1)in (3,4) and i.customer_regnom=r.regnom and" +
                " ar.licsch=l.licsch and ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd/mm/yyyy')";

            string sirketin_tesisci_plus_sirket = "select sal.date_oper,t.qeyd_no,sum(-sal.saldo_ish_nacval)+t.AZN_ile topqal1 from odb.arh_saldo_ls sal, "+
    " (select ar.date_oper tar, sum(-ar.saldo_ish_nacval) AZN_ile, i.regnom qeyd_no " +
    "    from licsch l, odb.arh_saldo_ls ar, regnom r, imza_huquqi_olan_shexsler i " +
    "   where r.regnom = l.registrac_nomer and(l.date_close_licsch is null or ar.date_oper <= l.date_close_licsch) " +
    "   and substr(l.licsch, 0, 1)in (3, 4) and i.customer_regnom = r.regnom and " +
    "   ar.licsch = l.licsch and  ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "', 'dd/mm/yyyy') group by ar.date_oper, i.regnom) t " +
    "   where substr(sal.licsch, 10, 6) = t.qeyd_no and t.tar = sal.date_oper and substr(sal.licsch,0,1) in (3, 4) " +
    "   group by sal.date_oper,t.qeyd_no,t.AZN_ile order by sal.date_oper,topqal1 desc";
            
            string sirketin_qaliq = "select sal.date_oper,t.qeyd_no,sum(-sal.saldo_ish_nacval) qal_1 from odb.arh_saldo_ls sal, " +
" (select ar.date_oper tar, sum(-ar.saldo_ish_nacval) AZN_ile, i.regnom qeyd_no " +
 "    from licsch l, odb.arh_saldo_ls ar, regnom r, imza_huquqi_olan_shexsler i " +
  "   where r.regnom = l.registrac_nomer and(l.date_close_licsch is null or ar.date_oper <= l.date_close_licsch) " +
  "   and substr(l.licsch, 0, 1)in (3, 4) and i.customer_regnom = r.regnom and " +
  "   ar.licsch = l.licsch and  ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "', 'dd/mm/yyyy') group by ar.date_oper, i.regnom) t " +
  "       where substr(sal.licsch, 10, 6) = t.qeyd_no and t.tar = sal.date_oper and substr(sal.licsch,0,1) in (3, 4) " +
  "   group by sal.date_oper,t.qeyd_no order by sal.date_oper,qal_1 desc";
            //evvelki hazir
            //        string xususi_ceki = "select sal.date_oper,t.qeyd_no,sum(-sal.saldo_ish_nacval)+t.AZN_ile,xs.top_qal,round(((sum(-sal.saldo_ish_nacval)+t.AZN_ile)/xs.top_qal*100),2) orta_chechi from odb.arh_saldo_ls sal, " +
            //" (select ar.date_oper tar, sum(-ar.saldo_ish_nacval) AZN_ile, i.regnom qeyd_no " +
            //"    from licsch l, odb.arh_saldo_ls ar, regnom r, imza_huquqi_olan_shexsler i " +
            //"   where r.regnom = l.registrac_nomer and(l.date_close_licsch is null or ar.date_oper <= l.date_close_licsch) " +
            //"   and substr(l.licsch, 0, 1)in (3, 4) and i.customer_regnom = r.regnom and " +
            //"   ar.licsch = l.licsch and  ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "', 'dd/mm/yyyy') group by ar.date_oper, i.regnom) t" +
            //",(select -sum(k.saldo_ish_nacval) top_qal,k.date_oper tar1 from odb.arh_saldo_ls k, odb.licsch p " +
            //         " where  p.licsch = k.licsch  " +
            //         //"and k.date_oper between to_date('01/11/2023', 'dd/mm/yyyy') and to_date('25/11/2023','dd/mm/yyyy') " +
            //         " and substr(k.licsch,0,2)in (35,36,38,39,40,41) and ( p.date_close_licsch is null or  k.date_oper<= p.date_close_licsch )group by k.date_oper )xs  " +
            //"       where substr(sal.licsch, 10, 6) = t.qeyd_no and t.tar = sal.date_oper and substr(sal.licsch,0,1) in (3, 4) and sal.date_oper=xs.tar1 " +
            //"   group by sal.date_oper,t.qeyd_no,t.AZN_ile,xs.top_qal order by sal.date_oper,orta_chechi desc";

            //yenisi

            string xususi_ceki = "select sal.date_oper,t.qeyd_no,sum(-sal.saldo_ish_nacval)+t.AZN_ile+CASE WHEN p.qeyd=t.qeyd_no THEN p.AZN_ile ELSE 0 END" +
                ",xs.top_qal,round(((sum(-sal.saldo_ish_nacval)+t.AZN_ile+CASE WHEN p.qeyd=t.qeyd_no THEN p.AZN_ile ELSE 0 END)/xs.top_qal),4) orta_chechi" +
                " from odb.arh_saldo_ls sal, " +
    " (select ar.date_oper tar, sum(-ar.saldo_ish_nacval) AZN_ile, i.regnom qeyd_no " +
    "    from licsch l, odb.arh_saldo_ls ar, regnom r, imza_huquqi_olan_shexsler i " +
    "   where r.regnom = l.registrac_nomer and(l.date_close_licsch is null or ar.date_oper <= l.date_close_licsch) " +
    "   and substr(l.licsch, 0, 1)in (3, 4) and i.customer_regnom = r.regnom and " +
    "   ar.licsch = l.licsch and  ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "', 'dd/mm/yyyy') group by ar.date_oper, i.regnom) t" +
    ",(select -sum(k.saldo_ish_nacval) top_qal,k.date_oper tar1 from odb.arh_saldo_ls k, odb.licsch p " +
             " where  p.licsch = k.licsch  " +
             //"and k.date_oper between to_date('01/11/2023', 'dd/mm/yyyy') and to_date('25/11/2023','dd/mm/yyyy') " +
             " and substr(k.licsch,0,2)in (35,36,38,39,40,41,49) and ( p.date_close_licsch is null or  k.date_oper<= p.date_close_licsch )group by k.date_oper )xs," +

             " (select distinct ar.date_oper tarix,ar.licsch,-ar.saldo_vhd_inval Valyuta_ile,-ar.saldo_ish_nacval AZN_ile, " +
                " ch.name_licsch,z.ish_ad,s.qno qeyd from odb.arh_saldo_ls ar,regnom r,licsch ch,  " +
                " (select f.fin fin,f.ish_yerinin_adi ish_ad  " +
                 "from regnom r,fiziki_shexs f  " +
                 "where f.ish_yerinin_adi=r.name_regnom) z," +
                 " (select rt.regnom qno,rt.name_regnom from regnom rt ) s where z.fin=r.pincode and  " +
                "( ch.date_close_licsch is null or  ar.date_oper<= ch.date_close_licsch ) " +
                " and r.regnom=ch.registrac_nomer and substr(ar.licsch,0,1)in (4) and ch.licsch=ar.licsch and s.name_regnom=z.ish_ad  " +
                " and substr(ar.licsch,16,2)=91   " +
                " order by ar.date_oper) p " +

             "       where substr(sal.licsch, 10, 6) = t.qeyd_no and t.tar = sal.date_oper and sal.date_oper=p.tarix" +
             " and substr(sal.licsch,0,1) in (3, 4) and sal.date_oper=xs.tar1 " +
    "   group by sal.date_oper,t.qeyd_no,t.AZN_ile,xs.top_qal,p.AZN_ile,p.qeyd,t.qeyd_no" +
    " order by sal.date_oper,orta_chechi desc";




            string sirketin_iscileri = "select distinct ar.date_oper,ar.licsch,-ar.saldo_vhd_inval Valyuta_ile,-ar.saldo_ish_nacval AZN_ile, " +
                " ch.name_licsch,z.ish_ad,s.qno from odb.arh_saldo_ls ar,regnom r,licsch ch,  " +
                " (select f.fin fin,f.ish_yerinin_adi ish_ad  " +
                 "from regnom r,fiziki_shexs f  " +
                 "where f.ish_yerinin_adi=r.name_regnom) z," +
                 " (select rt.regnom qno,rt.name_regnom from regnom rt ) s where z.fin=r.pincode and  " +
                "( ch.date_close_licsch is null or  ar.date_oper<= ch.date_close_licsch ) " +
                " and r.regnom=ch.registrac_nomer and substr(ar.licsch,0,1)in (4) and ch.licsch=ar.licsch and s.name_regnom=z.ish_ad  " +
                " and substr(ar.licsch,16,2)=91 and ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd/mm/yyyy')  " +
                " order by ar.date_oper";

            string ikinci_sorgu_35_36_38_39_40_41 = "select k.date_oper,k.licsch hesab," +
                "-sum(k.saldo_vhd_inval) Valyuta_ile,-sum(k.saldo_ish_nacval) AZN_ile from odb.arh_saldo_ls k," +
                " odb.licsch p " +
             " where  p.licsch = k.licsch and " +
             " k.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd/mm/yyyy') " +
             " and substr(k.licsch,0,2)in (35,36,38,39,40,41,49) and ( p.date_close_licsch is null or  k.date_oper<= p.date_close_licsch )" +
             " group by k.date_oper,k.licsch " +
             " order by k.date_oper,k.licsch";

            string dorduncu_sorgu_35025_26_49025_27 = "select ar.date_oper tarix," +
                "-sum(ar.saldo_ish_nacval) AZN_ile,xs.top_qal" +
                " from licsch l, odb.arh_saldo_ls ar," +
                "(select -sum(k.saldo_ish_nacval) top_qal,k.date_oper tar1 from odb.arh_saldo_ls k, odb.licsch p " +
             " where  p.licsch = k.licsch  " +
             //"and k.date_oper between to_date('01/11/2023', 'dd/mm/yyyy') and to_date('25/11/2023','dd/mm/yyyy') " +
             " and substr(k.licsch,0,2)in (35,36,38,39,40,41,49) and ( p.date_close_licsch is null or  k.date_oper<= p.date_close_licsch )group by k.date_oper )xs " +
                " where ( l.date_close_licsch is null or ar.date_oper<= l.date_close_licsch ) and " +
                " substr(l.licsch,0,5)in (35025,35026,49025,49027) and l.registrac_nomer='000016' " +
                " and ar.licsch=l.licsch and ar.date_oper=xs.tar1" +
                " and ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "','dd/mm/yyyy') " +
                "  group by ar.date_oper,xs.top_qal order by ar.date_oper ";

            string hazir_hesabat = "select ff.etarix es_tar,MAX(ff.orta_chechi) from (select sal.date_oper as etarix,t.qeyd_no,sum(-sal.saldo_ish_nacval)+t.AZN_ile+CASE WHEN p.qeyd=t.qeyd_no THEN p.AZN_ile ELSE 0 END" +
                ",xs.top_qal,round(((round(sum(-sal.saldo_ish_nacval),4)+round(t.AZN_ile,4)+CASE WHEN p.qeyd=t.qeyd_no THEN round(p.AZN_ile,4) ELSE 0 END)/round(xs.top_qal,4)),4) orta_chechi" +
                " from odb.arh_saldo_ls sal, " +
    " (select ar.date_oper tar, sum(-ar.saldo_ish_nacval) AZN_ile, i.regnom qeyd_no " +
    "    from licsch l, odb.arh_saldo_ls ar, regnom r, imza_huquqi_olan_shexsler i " +
    "   where r.regnom = l.registrac_nomer and(l.date_close_licsch is null or ar.date_oper <= l.date_close_licsch) " +
    "   and substr(l.licsch, 0, 1)in (3, 4) and i.customer_regnom = r.regnom and " +
    "   ar.licsch = l.licsch and  ar.date_oper between to_date('" + textBox2.Text + "', 'dd/mm/yyyy') and to_date('" + textBox3.Text + "', 'dd/mm/yyyy') group by ar.date_oper, i.regnom) t" +
    ",(select -sum(k.saldo_ish_nacval) top_qal,k.date_oper tar1 from odb.arh_saldo_ls k, odb.licsch p " +
             " where  p.licsch = k.licsch  " +
             //"and k.date_oper between to_date('01/11/2023', 'dd/mm/yyyy') and to_date('25/11/2023','dd/mm/yyyy') " +
             " and substr(k.licsch,0,2)in (35,36,38,39,40,41,49) and ( p.date_close_licsch is null or  k.date_oper<= p.date_close_licsch )group by k.date_oper )xs," +

             " (select distinct ar.date_oper tarix,ar.licsch,-ar.saldo_vhd_inval Valyuta_ile,-ar.saldo_ish_nacval AZN_ile, " +
                " ch.name_licsch,z.ish_ad,s.qno qeyd from odb.arh_saldo_ls ar,regnom r,licsch ch,  " +
                " (select f.fin fin,f.ish_yerinin_adi ish_ad  " +
                 "from regnom r,fiziki_shexs f  " +
                 "where f.ish_yerinin_adi=r.name_regnom) z," +
                 " (select rt.regnom qno,rt.name_regnom from regnom rt ) s where z.fin=r.pincode and  " +
                "( ch.date_close_licsch is null or  ar.date_oper<= ch.date_close_licsch ) " +
                " and r.regnom=ch.registrac_nomer and substr(ar.licsch,0,1)in (4) and ch.licsch=ar.licsch and s.name_regnom=z.ish_ad  " +
                " and substr(ar.licsch,16,2)=91   " +
                " order by ar.date_oper) p " +
             "       where substr(sal.licsch, 10, 6) = t.qeyd_no and t.tar = sal.date_oper and sal.date_oper=p.tarix" +
             " and substr(sal.licsch,0,1) in (3, 4) and sal.date_oper=xs.tar1 " +
    "   group by sal.date_oper,t.qeyd_no,t.AZN_ile,xs.top_qal,p.AZN_ile,p.qeyd,t.qeyd_no" +
    " ) ff group by ff.etarix order by ff.etarix";

            #endregion


            DataTable _dt_sirket_tesisci = new DataTable();
            DataTable _dt_sirket_isci = new DataTable();
            DataTable _dt_qaliq_35_36 = new DataTable();
            DataTable _dt_qaliq_35025_26 = new DataTable();
            DataTable _dt_sirketin_tesisci_plus_sirket = new DataTable();
            DataTable _dt_sirketin_qaliq = new DataTable();
            DataTable _dt_xususi_ceki = new DataTable();
            DataTable _dt_hazir_hesabat = new DataTable();
            //DataTable hesablanmis_faiz_GIROV = new DataTable();
            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                using (OracleCommand command = new OracleCommand(sirketin_tesiscileri, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_sirket_tesisci);
                }

                using (OracleCommand command = new OracleCommand(sirketin_iscileri, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_sirket_isci);
                }
                using (OracleCommand command = new OracleCommand(ikinci_sorgu_35_36_38_39_40_41, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qaliq_35_36);
                }

                using (OracleCommand command = new OracleCommand(dorduncu_sorgu_35025_26_49025_27, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qaliq_35025_26);
                }
                using (OracleCommand command = new OracleCommand(sirketin_tesisci_plus_sirket, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_sirketin_tesisci_plus_sirket);
                }
                using (OracleCommand command = new OracleCommand(sirketin_qaliq, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_sirketin_qaliq);
                }
                using (OracleCommand command = new OracleCommand(xususi_ceki, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_xususi_ceki);
                }
                using (OracleCommand command = new OracleCommand(hazir_hesabat, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_hazir_hesabat);
                }
            }

            
                string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
                string baseFileName = "Əlaqəli depozitlərin cəmi"; // Temel dosya adı
                string fileName = baseFileName + ".xlsx";
                string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "Elaqeli_depozitlerin_cemi.xlsx");
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
            FileInfo templateFile = new FileInfo(templateFilePath);
            
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["_dt_sirket_tesisci"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["_dt_sirket_isci"];
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets["_dt_qaliq_35_36"];
                ExcelWorksheet worksheet4 = package.Workbook.Worksheets["_dt_qaliq_35025_26"];
                ExcelWorksheet worksheet5 = package.Workbook.Worksheets["sirket_qaliq"];
                ExcelWorksheet worksheet6 = package.Workbook.Worksheets["tarix_uzre_en_boyukler"];
                ExcelWorksheet worksheet7 = package.Workbook.Worksheets["Xususi_ceki_siyahi"];
                ExcelWorksheet worksheet8 = package.Workbook.Worksheets["Hesabat"];
                // Tüm sayfalara DataTable'ları ekleyin
                if (_dt_sirket_tesisci.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_sirket_tesisci.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_sirket_tesisci.Columns.Count; col++)
                        {
                            worksheet1.Cells[row + 2, col + 1].Value = _dt_sirket_tesisci.Rows[row][col];
                        }
                    }
                }

                if (_dt_sirket_isci.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_sirket_isci.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_sirket_isci.Columns.Count; col++)
                        {
                            worksheet2.Cells[row + 2, col + 1].Value = _dt_sirket_isci.Rows[row][col];
                        }
                    }
                }
                if (_dt_qaliq_35_36.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_qaliq_35_36.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_qaliq_35_36.Columns.Count; col++)
                        {
                            worksheet3.Cells[row + 2, col + 1].Value = _dt_qaliq_35_36.Rows[row][col];
                        }
                    }
                }
                if (_dt_qaliq_35025_26.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_qaliq_35025_26.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_qaliq_35025_26.Columns.Count; col++)
                        {
                            worksheet4.Cells[row + 2, col + 1].Value = _dt_qaliq_35025_26.Rows[row][col];
                        }
                    }
                }
                if (_dt_sirketin_qaliq.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_sirketin_qaliq.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_sirketin_qaliq.Columns.Count; col++)
                        {
                            worksheet5.Cells[row + 2, col + 1].Value = _dt_sirketin_qaliq.Rows[row][col];
                        }
                    }
                }
                if (_dt_sirketin_tesisci_plus_sirket.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_sirketin_tesisci_plus_sirket.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_sirketin_tesisci_plus_sirket.Columns.Count; col++)
                        {
                            worksheet6.Cells[row + 2, col + 1].Value = _dt_sirketin_tesisci_plus_sirket.Rows[row][col];
                        }
                    }
                }

                if (_dt_xususi_ceki.Rows.Count > 0)
                {
                    for (int row = 0; row < _dt_xususi_ceki.Rows.Count; row++)
                    {
                        for (int col = 0; col < _dt_xususi_ceki.Columns.Count; col++)
                        {
                            worksheet7.Cells[row + 2, col + 1].Value = _dt_xususi_ceki.Rows[row][col];
                        }
                    }
                }

                
                    
                if (_dt_hazir_hesabat.Rows.Count > 0 || _dt_qaliq_35025_26.Rows.Count > 0)
                {
                    int startColumn = 3; // C column
                    int endColumn = 34; // AG column

                    // Iterate over rows in DataTable
                    for (int row = 0; row < 32; row++)
                    {
                        int sayi = _dt_hazir_hesabat.Rows.Count;
                        // Get the date from DataTable

                        // Iterate over date columns in Excel
                        if (startColumn <= endColumn)
                        {
                            // Get the date from Excel

                            //string yoxla = _dt_hazir_hesabat.Rows[row][0].ToString();
                            // Compare the days only

                            for (int i = 0; i < _dt_hazir_hesabat.Rows.Count; i++)
                            {

                                int exgun = Convert.ToInt32(worksheet8.Cells[1, startColumn].Text);
                                DateTime dtFromDataTable = Convert.ToDateTime(_dt_hazir_hesabat.Rows[i][0]).Date;
                                int sqlgun = dtFromDataTable.Day;

                                if (sqlgun == exgun)
                                {
                                    // If the days match, set the value to Excel
                                    worksheet8.Cells[2, startColumn].Value = _dt_hazir_hesabat.Rows[i][1];
                                    //worksheet8.Cells[2, startColumn].Style.Numberformat.Format = "0.0000%";
                                    worksheet8.Cells[3, startColumn].Value =Math.Round (Math.Round( Convert.ToDouble( _dt_qaliq_35025_26.Rows[i][1]),4) /Math.Round( Convert.ToDouble(_dt_qaliq_35025_26.Rows[i][2]),4),4);
                                    //worksheet8.Cells[3, startColumn].Style.Numberformat.Format = "0.0000%";
                                }
                                else if (sqlgun != exgun)
                                {
                                    // If the DataTable day is greater, set the value from the previous day in Excel
                                    worksheet8.Cells[2, startColumn].Value = worksheet8.Cells[2, startColumn - 1].Value;

                                    worksheet8.Cells[3, startColumn].Value = worksheet8.Cells[3, startColumn - 1].Value;

                                    worksheet8.Cells[2, startColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                    worksheet8.Cells[3, startColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                                    //startColumn += 1;
                                    i -= 1;
                                }
                                startColumn += 1;
                            }

                            // If _dt_hazir_hesabat.Rows[row][0] is null or DBNull.Value, handle accordingly
                            int say = startColumn;
                            for (int i = say; i < endColumn; i++)
                            {
                                string test = worksheet8.Cells[2, i - 1].ToString();
                                worksheet8.Cells[2, i].Value = worksheet8.Cells[2, i - 1].Value;
                                worksheet8.Cells[3, i].Value = worksheet8.Cells[3, i - 1].Value;
                                worksheet8.Cells[2, i].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                worksheet8.Cells[3, i].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                startColumn += 1;
                            }
                            break;
                        }
                    }
                }
                #endregion
                filePath = Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }

        }
            
    private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "Məlumatlar hazırlanır...";
            excel();
            button1.Text = "Ümumi sorğu";
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
    }
}
