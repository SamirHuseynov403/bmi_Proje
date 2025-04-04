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
using System.Data.SQLite;
using DevExpress.CodeParser;
using DocumentFormat.OpenXml.Vml;


namespace BMI.Muhasibat
{
    public partial class frm_report_comments_Yeni : Form
    {
        public frm_report_comments_Yeni()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        //string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
        string lcrcem,lcrazn,lcrval;
        string lcrcemd, lcraznd, lcrvald, lcr4, lcr5;


        string[] A1 = { "100" };
        string[] A2 = { "110" };
        string[] A2a = { "11010000040000200000" };
        string[] A2b = { "11010000010000200000","11020020010000200000" };
        string[] A3a = { "159" };
        string[] A3a_qeydno = { "000010" };
        string[] A3a_1 = { "15020" };
        string[] A3a_1a = { "15910" };
        string[] A3a_1a_yanasma = { "10" };
        //A3a.1b bos
        string[] A3a_2 = { "15025" };
        string[] A3a_2a = { "15911" };
        string[] A3a_2a_yanasma = { "10" };
        //A3a_2b bos
        //A4a bos
        string[] A4_1a = { "14030", "14032"};
        string[] A4_1a1 = { "14030", "14032" };
        string[] A4_1b = { "14010", "14012" };
        string[] A4_1b1 = { "14010", "14012" };
        string[] A4_1c = { };
        //A4_2a-dan A4_3c kimi bos
        string[] A5 = { "15770","11710"};//yalniz manat
        string[] A6 = { "11110", "15210", "15213", "15220", "15220", "15223", "15223", "15225", "15225" };
        string[] A6a = { "11110" };
        //A6a,A6b bos
        string[] A6c = { "15910000000001200100" };
        string[] A7 = {"15620"};
        //A7a bos
        string[] A8_1 = {};
        string[] A8_2 = {};
        //A83 bos
        string[] A8b = {"209","219" };// son iki reqemi 25=> ve ya üç reqemi 100 olanlar
        string[] A8b_yanasma = { "11", "12","16" };

        string[] setir_9_a = {
                "15022", "15027", "15028", "15212",
                "15214", "15772", "11712", "15222", "15222", "15224", "15224", "15227", "15227", "15622", "15624",
                "20392", "20394", "20632", "20634", "20652", "20654", "20662", "20664", "20682", "20764",
                "21072", "21074", "21077", "21079", "21112", "21114", "21117", "21119", "21127", "21127",
                "21129", "21129", "21142", "21144", "21212", "21214", "21217", "21219", "21222", "21224",
                "21227", "21227", "21227", "21229", "21229", "21229", "21242", "21244", "21247", "21252",
                "21254", "21257", "21259", "23124", "23302", "25069" };
        string[] setir_9_b = {
                "15910000000001301100", "15910000000001302100", "15910000000001402100",
                "15910000000001500030", "15910000000001500060", "15910000000001500100",
                "15910000000001502100", "20910000000001400030", "20910000003001200060",
                "20910000003001400030", "20910000003001400100", "20910000000001400100",
                "20910000000001500030", "20910000003001500030", "20910000000001500100",
                "20910000003001500100", "20910000005001100002", "20910000005001100025",
                "20910000005001100050", "20910000005001200002", "20910000005001100100",
                "20910000005001200100", "20910000005001400002", "20910000005001400025",
                "20910000005001400050", "20910000005001500002", "20910000005001400100",
                "20910000005001500100", "20910000000001501100", "21910000000001400030",
                "21910000000001400060", "21910000000001400100", "21910000000001401100",
                "21910000000001500030", "21910000000001500060", "21910000000001500100",
                "21910000000001501001", "21910000000001501100", "21911000000001400030",
                "21911000000001400060", "21911000000001400100", "21911000000001401030",
                "21911000000001401100", "21911000000001402030", "21911000000001500030",
                "21911000000001500060", "21911000000001500100", "21911000000001501030",
                "21911000000001501100", "21911000000001502100", "21912000000001400030",
                "21912000000001400100", "21912000000001500030", "21912000000001500100",
                "21913000000001401100", "21913000000001501100", "23910000000001501100"};
        
        string[] setir_10_a = {
                "35090", "35100", "35100", "35100", "35938", "38030", "38040", "38130", "38140",
                "38090", "38190", "38939", "38943", "38949", "39010", "39020", "39020", "39070",
                "39080", "39080", "39939", "39940", "39940", "39949", "39949", "40030", "40040",
                "40045", "40050", "40055", "40060", "40061", "40065", "40070", "40075", "40080",
                "40081", "40090", "40130", "40130", "40130", "40140", "40140", "40140", "40145",
                "40145", "40150", "40150", "40160", "40160", "40160", "40165", "40165", "40170",
                "40170", "40175", "40175", "40175", "40180", "40190", "40190", "40190", "40932",
                "40933", "40935", "40939", "40942", "40942", "40943", "40943", "40944", "40944",
                "40945", "40945", "40946", "40946", "40949", "40949"};
        string[] setir_10_b = {
                "41040", "41045", "41050", "41055",
                "41932", "41933","41942", "41943"};

        string[] setir_10_fizik = {
                "41010", "41011", "41015", "41016", "41020",  "41021", 
                "41025", "41026", "41930", "41931",  "41940", "41941",
                "41942",};

        // string[] setir_11a = { "11010", "11020" }; siyahida yoxdur
        string[] A9 = { "28"};
        string[] A9_istisna = { "28110","28111","28120","28121","28130","28131" };
        string[] A10 = { "27012" };
        string[] A10a = { "27012" };
        string[] A10b = { "27012","27013" };
        string[] A10b_27013 = { "27013" };
        //A11 bos
        string[] A12_qisa = {"11112","15022", "15027", "15028",
                "15212", "15214", "15222", "15224", "15227",
                "15622", "15624", "15772", "11712", "20362", "20392", "20394", "20632", "20634", "20652",
                "20654", "20662", "20664", "20682", "20764", "21072", "21074", "21077", "21079",
                "21112", "21114", "21117", "21119", "21127", "21129", "21142",
                "21144", "21212", "21214", "21217", "21219", "21222", "21224", "21227",
                "21227", "21229", "21242", "21244", "21247", "21249","21252", "21254",
                "21257", "21259", "23124", "23302","23911", "24010", "25010", "25011", "25019",
                "25020", "25021", "25029", "25052", "25059",
                "25069", "25079", "25089", "25100", "25101", "25103",
                "25109", "25110", "25119","25120", "25121", "25122", "25123", "25129", "25139", "25159",
                "25270", "25280", "28110", "28111", "28120", "28121", "28130", "28131" };
        string[] A12_eht = { "159", "209", "219", "239", "259" };
        string[] A12_eht_faizler = { "13", "14", "15" };
        string[] A12a = {"11112","15022", "15027", "15028", "15212",
                "15214", "15772", "11712", "15222", "15222", "15224", "15224", "15227", "15227", "15622", "15624",
                "20392", "20394", "20632", "20634", "20652", "20654", "20662", "20664", "20682", "20764",
                "21072", "21074", "21077", "21079", "21112", "21114", "21117", "21119", "21127", "21127",
                "21129", "21129", "21142", "21144", "21212", "21214", "21217", "21219", "21222", "21224",
                "21227", "21227", "21227", "21229", "21229", "21229", "21242", "21244", "21247", "21249","21252",
                "21254", "21257", "21259", "23124", "23302", "25069"  };
        string[] A12b = { "159","209", "219", "239" };
        string[] A12b_istisna = { "10","11", "12", "16" };
        string[] A12c = {"209","219","239"};
        string[] A12c_uzun = { "15910000000001301100","15910000000001302100","15910000000001500100","23911000000001100100" };
        string[] A12c_yanasma = { "11", "12", "16"};
        string[] B1a = {
                "35090", "35100", "35100", "35100", "35938", "38030", "38040", "38130", "38140",
                "38090", "38190", "38939", "38943", "38949", "39010", "39020", "39020", "39070",
                "39080", "39080", "39939", "39940", "39940", "39949", "39949", "40030", "40040",
                "40045", "40050", "40055", "40060", "40061", "40065", "40070", "40075", "40080",
                "40081", "40090", "40130", "40130", "40130", "40140", "40140", "40140", "40145",
                "40145", "40150", "40150", "40160", "40160", "40160", "40165", "40165", "40170",
                "40170", "40175", "40175", "40175", "40180", "40190", "40190", "40190", "40932",
                "40933", "40935", "40939", "40942", "40942", "40943", "40943", "40944", "40944",
                "40945", "40945", "40946", "40946", "40949", "40949" };
        string[] B1b = {
                "41010", "41011", "41015", "41016", "41020", "41020", "41021", "41021", "41025",
                "41026", "41930", "41931", "41932", "41933", "41940", "41941",
                "41942", "41943"};
        string[] B1c = {
                "41040", "41045","41050", "41055"};
        string[] B2b = { "41110", "41112", "41115", "41117", "41120" };
        string[] B4a = { "35020", "35940" };
        string[] B4b = { "35015", "35025",};

        string[] B5a = { "35226", "49025" };
        string[] B5a_1 = {"49025" };
        string[] B8 = { "35770" };
        string[] B10 = { "35026", "35770", "35772", "35227", "41122", "44010", "44510", "45010", "45011",
                "45013", "45019", "45020", "45021", "45021", "45021", "45021", "45021", "45023",
                "45023", "45023", "45023", "45029", "45029", "45029", "45029", "45029", "45050",
                "45079", "45080", "45080", "45089", "45089", "45089", "45089", "45100", "45101",
                "45102", "45103", "45105", "45110", "45150", "45159","45160", "45270", "45280" };
        string[] B10a = { "35772", "35227", "41122" };
        string[] B12 = { "50020", "50060", "50110", "50120", "50130" };
        string[] B12_eht = {"209","219" };
        string[] B12a = { "50020" };
        string[] B13 = { "50130" };
        string[] setir_006 = { "27013000000001000006" };
        string[] setir_67 = { "11010000010000200000", "11010000040000200000", "11020010010000200000", "11020010030000200000", "11020010040000200000", "11020020010000200000", "11020020030000200000", "15020010000004600001", "15020020000004600001" };
        string[] setir_060_uzun = new string[]
    {
                "15910000000001200060",
                "15910000000001500060",
                "20910000000001100060",
                "20910000000001200060",
                "20910000003001200060",
                "21910000000001100060",
                "21910000000001200060",
                "21910000000001400060",
                "21910000000001500060",
                "21910000000001600060",
                "21911000000001100060",
                "21911000000001101060",
                "21911000000001200060",
                "21911000000001201060",
                "21911000000001400060",
                "21911000000001500060",
                "21912000000001100060",
                "21912000000001200060"
    };
        string[] likvid_risk_1 = new string[]
{
    "10010",
    "10020",
    "10050",
    "10060",
    "11010",
    "11020",
    "14010",
    "14014",
    "14030",
    "14034",
    "15020",
    "15025"
};
        string[] likvid_risk_2 = new string[]
                {
                            "35026", "49025", "41011", "41026","40061","44510","35770","35772","411","412"
                };
        string[] likvid_risk_4 = { "10", "11", "14" };
        string[] likvid_risk_4_elave = { "15770", "15025","11710" };
        string[] likvid_risk_4_istisna = { "11010000040000200000", "11010000010000200000", "11020020010000200000" };
        string[] likvid_risk_5 = { "38", "39", "40", "41", "35020", "35025" };
        string[] likvid_risk_5_bk = { "99530" };
        
        string[] setir_001 = new string[]
    {
                "15910000000001001001",
                "15910000000001002001",
                "15910000000001100001",
                "15910000000001102001",
                "15910000000001200001",
                "15910000000001301001",
                "15910000000001302001",
                "15910000000001400001",
                "15910000000001402001",
                "15910000000001500001",
                "15911000000001000001",
                "15911000000001001001",
                "15911000000001003001",
                "15911000000001002001",
                "15911000000001004001",
                "15911000000001005001",
                "15911000000001105001",
                "15911000000001102001",
                "15911000000001301001",
                "15911000000001302001",
                "15911000000001402001",
                "15911000000001405001",
                "15918000000001100001",
                "15918000000001400001",
                "15918000000001500001",
                "15918000003001600001",
                "20910000000001000001",
                "20910000000001100001",
                "20910000003001100001",
                "20910000000001200001",
                "20910000003001200001",
                "20910000000001400001",
                "20910000003001400001",
                "20910000005001100001",
                "20910000005001200001",
                "20910000005001400001",
                "20910000005001500001",
                "20910000000001500001",
                "20910000003001500001",
                "20910000003001600001",
                "21910000000001000001",
                "21910000005001100001",
                "21910000000001100001",
                "21910000000001101001",
                "21910000005001200001",
                "21910000000001200001",
                "21910000000001201001",
                "21910000005001400001",
                "21910000000001400001",
                "21910000000001401001",
                "21910000005001500001",
                "21910000000001500001",
                "21910000000001501001",
                "21910000000001600001",
                "21910000003001600001",
                "21911000000001000001",
                "21911000000001100001",
                "21911000000001101001",
                "21911000000001102001",
                "21911000000001200001",
                "21911000000001201001",
                "21911000000001400001",
                "21911000000001401001",
                "21911000000001402001",
                "21911000000001500001",
                "21911000000001501001",
                "21911000000001600001",
                "21911000003001600001",
                "21912000000001000001",
                "21912000000001100001",
                "21912000000001200001",
                "21912000000001400001",
                "21912000000001500001",
                "21913000000001100001",
                "21913000000001400001",
                "23911000000001400001",
                "23911000000009300001"
    };
        string[] setir_005 = new string[]
    {
                "20910000000001100005",
                "20910000001001100005",
                "20910000003001100005",
                "20910000000001200005",
                "20910000000001400005",
                "20910000001001400005",
                "20910000003001200005",
                "20910000003001400005",
                "20910000000001500005",
                "21910000000001100005",
                "21910000000001200005",
                "21910000000001400005",
                "21910000000001500005",
                "21910000000001600005",
                "21911000000001100005",
                "21911000000001101005",
                "21911000000001102005",
                "21911000000001200005",
                "21911000000001201005",
                "21911000000001400005",
                "21911000000001401005",
                "21911000000001402005",
                "21911000000001500005",
                "21911000000001501005",
                "21912000000001100005",
                "21912000000001101005",
                "21912000000001200005",
                "21912000000001201005",
                "21912000000001400005",
                "21912000000001401005",
                "21912000000001500005",
                "21912000000001501005"
    };
        
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt;

        private void Excel_daily_comment_Yeni()
        {

            //duz olan
            #region sql_kodlar

            
            string daily_report = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "CASE WHEN SUBSTR(ar.licsch, 0, 3) IN ('159','209','219','239','259') THEN SUBSTR(ar.licsch, 16, 2) " +
               "ELSE SUBSTR(ar.licsch, 6, 2) END AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string daily_report_bk = "select t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish," +
                "t.ostatok_ish*ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6)  kurs " +
                "from odb.arh_saldo_vbls t where t.date_oper =TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99531,99550,99300,99301) and t.ostatok_ish<>0";

            string daily_report_kataloq = "select al.date_oper,al.licschkre,substr(al.licschkre,6,2), tk.code,tk.name,al.summa,al.summa_19," +
                "(al.summa+al.summa_19)*ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) kurs " +
                    "from arh_licschkre al,tipkre tk where al.tipkredita = tk.code and (al.date_close is null or al.date_close>TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')) " +
                    "and al.date_oper=TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy')";

            string daily_report1 = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "CASE WHEN SUBSTR(ar.licsch, 0, 3) IN ('159','209','219','239','259') THEN SUBSTR(ar.licsch, 16, 2) " +
               "ELSE SUBSTR(ar.licsch, 6, 2) END AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string daily_report_bk1 = "select t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish," +
                "t.ostatok_ish*ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6)  kurs " +
                "from odb.arh_saldo_vbls t where t.date_oper =TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99531,99550,99300,99301) and t.ostatok_ish<>0";

            string daily_report_kataloq1 = "select al.date_oper,al.licschkre,substr(al.licschkre,6,2), tk.code,tk.name,al.summa,al.summa_19," +
                "(al.summa+al.summa_19)*ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) kurs " +
                    "from arh_licschkre al,tipkre tk where al.tipkredita = tk.code and (al.date_close is null or al.date_close>TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')) " +
                    "and al.date_oper=TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')";

            string gun_erzinde_odenisler = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                "'odenisler'gun_erzinde,round(sum(d.summa_v_nacval), 2) meb, " +
                "case when l.tipkredita = 2 then 'fiziki' " +
                "when l.tipkredita in (1, 3) then 'sahibkar' end tip " +
                ", substr(l.licschkre, 6, 2)val," +
                "case when "+
                "l.date_restructure is null then 'bos' else 'dolu' end rest " +
                "    from arh_licschkre l, arh_dd d,odb.balschkli b " +
                "where d.kredit in (l.licschkre, l.licsch_19) and d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy')   " +
                "and l.date_oper = d.date_oper and substr(d.debet,1,5)= b.balsch and substr(d.debet,10,6)= substr(d.kredit, 10, 6) " +
                "and d.ssk = l.subschkre group by case when d.date_oper = to_date('" + txtdtdunen.Text + "', 'dd/mm/yyyy') then 'dunen' else 'bugun' end,l.tipkredita,substr(l.licschkre, 6, 2),l.date_restructure";
            string qaliqlar_30_90 = "SELECT case when m.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                    " case when " +
                    " m.tipkredita = 1 then 'huquqi' " +
                    " when m.tipkredita = 2 then 'fiziki' " +
                    " else 'sahibkar' end tip, " +
                    " odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, " +
                    " ((m.summa * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)) + (m.summa_19 * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)))  qal , " +
                    " case when m.summa_19 > 0 then 'vk' end gecikme,case when m.date_restructure is not null then 'rest' end restur,substr(m.licschkre, 6, 2) val " +
                    " from view_nacpogprokre_all x, arh_licschkre m where " +
                    " x.date_oper = m.date_oper " +
                    " and x.licschpkre = m.licschpkre and x.subschkre = m.subschkre " +
                    " and m.date_oper between to_date('"+txtdtdunen.Text+ "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and m.date_close is null " +
                    " and x.licschpkre = m.licschpkre and x.subschkre = m.subschkre " +
                    " and m.date_close is null";
            string qaliqlar_kr_tip = " select case when m.date_oper=to_date('"+txtdtdunen.Text+"','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                    " case when " +
                     " (m.tipkredita = 1 or m.tipkredita = 3) and m.index_otrasli != '01902'  then 'biznes' " +
                     " when m.tipkredita = 2 and m.index_otrasli != '01902' then 'fiziki' " +
                     " when m.index_otrasli = '01902' then 'dasinmaz' " +
                     " end tip, sum((m.summa * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)) + (m.summa_19 * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)))  qal,  " +
                     " substr(m.licschkre, 6, 2) val " +
                     " from arh_licschkre m " +
                     " where m.date_oper between to_date('"+txtdtdunen.Text+"', 'dd-mm-yyyy') and to_date('"+txtdtbugun.Text+"','dd-mm-yyyy') and m.date_close is null " +
                     " group by case when m.date_oper = to_date('"+txtdtdunen.Text+"', 'dd/mm/yyyy') then 'dunen' else 'bugun' end,  " +
                     " m.tipkredita,substr(m.licschkre, 6, 2),m.index_otrasli";

            string medaxiller = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar," +
                                "CASE    " +
                    "WHEN r.yurik = 1 then 'huquqi' " +
                    "WHEN r.predprinimatel = 1  THEN 'sahibkar' "+
                    "WHEN r.fizik = 1  THEN 'fiziki' "+
                    "ELSE 'unknown' END AS tip,round(sum(d.summa_v_nacval) , 2) meb,substr(d.kredit,6,2) val,count(d.kredit) " +
                    "      from arh_dd d, regnom r,odb.balschkli b,licsch l " +
                  " where d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') " +
                  "and substr(d.kredit,1,5)= b.balsch and" +
                  " d.kredit=l.licsch and l.registrac_nomer=r.regnom " +
                  "and substr(d.debet,1,5) not in ('66220', '86220') " +
                  " group by case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end," +
                  "CASE " +
                  "WHEN r.yurik = 1 then 'huquqi' " +
                    "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                    "WHEN r.fizik = 1  THEN 'fiziki' " +
                  "  ELSE 'unknown' END,substr(d.kredit,6,2)";
            string mexaricler = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                   "CASE " +
                   "WHEN r.yurik = 1 then 'huquqi' " +
                   "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                   "WHEN r.fizik = 1  THEN 'fiziki' " +
                   "ELSE 'unknown' END AS tip,round(sum(d.summa_v_nacval), 2) meb,substr(d.kredit, 6, 2) val " +
                   "from arh_dd d, regnom r,odb.balschkli b, licsch l " +
                   "where d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and substr(d.debet,1,5)= b.balsch " +
                   "and d.debet = l.licsch and l.registrac_nomer = r.regnom " +
                   "and substr(d.kredit,1,5) not in ('66220', '86220') " +
                   "group by case when d.date_oper = to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') then 'dunen' else 'bugun' end ,  " +
                   "CASE " +
                   "WHEN r.yurik = 1 then 'huquqi' " +
                   "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                   "WHEN r.fizik = 1  THEN 'fiziki' " +
                   "ELSE 'unknown' END,substr(d.kredit, 6, 2)";
            string ver_kr_lar = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd-mm-yyyy') then 'dunen' else 'bugun' end tar, " +
                    "case when l.tipkredita = 2 then 'fiziki' " +
                    "when l.tipkredita in (1, 3) then 'sahibkar' end tip " +
                    ", sum(d.summa_v_nacval) meb,substr(d.kredit, 6, 2) val, " +
                    "case when " +
                    "l.date_restructure is null then 'bos' else 'dolu' end rest " +
                    "from arh_dd d,regnom r, arh_licschkre l,odb.balschkli b " +
                    "where d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and d.debet = l.licschkre " +
                    "and substr(d.kredit,10,6)= substr(l.licschkre, 10, 6) " +
                    "and l.date_oper = d.date_oper and substr(d.kredit,1,5)= b.balsch and d.ssd = l.subschkre and l.date_close is null " +
                    "and substr(d.debet,10,6)= r.regnom " +
                    "group by case when d.date_oper = to_date('" + txtdtdunen.Text + "', 'dd/mm/yyyy') then 'dunen' else 'bugun' end, " +
                    "case when l.tipkredita = 2 then 'fiziki' " +
                    "when l.tipkredita in (1, 3) then 'sahibkar' end ,substr(d.kredit, 6, 2),l.date_restructure";
            #endregion
            DataTable _dt_daily_report = new DataTable();
            DataTable _dt_daily_report_bk = new DataTable();
            DataTable _dt_daily_report_kataloq = new DataTable();

            DataTable _dt_daily_report1 = new DataTable();
            DataTable _dt_daily_report_bk1 = new DataTable();
            DataTable _dt_daily_report_kataloq1 = new DataTable();

            DataTable _dt_odenisler = new DataTable();
            DataTable _dt_qali_gunler = new DataTable();
            DataTable _dt_qaliq_tip = new DataTable();
            DataTable _dt_medaxil = new DataTable();
            DataTable _dt_mexaric = new DataTable();
            DataTable _dt_verilmis = new DataTable();

            using (OracleConnection connection = new OracleConnection(cl.con))
            {
                using (OracleCommand command = new OracleCommand(daily_report, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report);
                }

                using (OracleCommand command = new OracleCommand(daily_report_bk, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_bk);
                }

                using (OracleCommand command = new OracleCommand(daily_report_kataloq, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_kataloq);
                }

                using (OracleCommand command = new OracleCommand(daily_report1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report1);
                }

                using (OracleCommand command = new OracleCommand(daily_report_bk1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_bk1);
                }

                using (OracleCommand command = new OracleCommand(daily_report_kataloq1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_kataloq1);
                }
                using (OracleCommand command = new OracleCommand(gun_erzinde_odenisler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_odenisler);
                }

                using (OracleCommand command = new OracleCommand(qaliqlar_30_90, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qali_gunler);
                }

                using (OracleCommand command = new OracleCommand(qaliqlar_kr_tip, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qaliq_tip);
                }
                using (OracleCommand command = new OracleCommand(medaxiller, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_medaxil);
                }

                using (OracleCommand command = new OracleCommand(mexaricler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_mexaric);
                }

                using (OracleCommand command = new OracleCommand(ver_kr_lar, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_verilmis);
                }
                connection.Close();
            }
            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txtdtbugun.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", ""); ;
            string baseFileName = "CUR.v02.1124d"+yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "Daily_report_comments_Yeni_.xlsm");
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);

            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Daily-Report"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["comments"];
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets["Daily_Credit_Deposit"];
                //******************Setir A1
                //******************Setir B11a
                var total_a1 = new Dictionary<string, decimal>();
                var types_a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a1)
                {
                    var Setir_a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A1.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a1[type] = Setir_a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E8"].Value = total_a1["00"];
                worksheet1.Cells["F8"].Value = total_a1["01"];
                worksheet1.Cells["G8"].Value = total_a1["02"];
                worksheet1.Cells["I8"].Value = total_a1["03"];
                worksheet1.Cells["K8"].Value = total_a1["04"];
                worksheet1.Cells["O8"].Value = total_a1["05"];

                var Setir_a1_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A1.Contains(row.Field<string>(1).Substring(0, 3)))
                                        .ToList();
                decimal total_a1_1_c = Setir_a1_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c8"].Value = total_a1_1_c;

                //******************Setir A2
                var total_a2 = new Dictionary<string, decimal>();
                var types_a2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2)
                {
                    var Setir_a2 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2[type] = Setir_a2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E9"].Value = total_a2["00"];
                worksheet1.Cells["F9"].Value = total_a2["01"];
                worksheet1.Cells["G9"].Value = total_a2["02"];
                worksheet1.Cells["I9"].Value = total_a2["03"];
                worksheet1.Cells["K9"].Value = total_a2["04"];
                worksheet1.Cells["O9"].Value = total_a2["05"];

                var Setir_a2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2.Contains(row.Field<string>(1).Substring(0, 3)))
                                        .ToList();
                decimal total_a2_1_c = Setir_a2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c9"].Value = total_a2_1_c;
                //******************Setir A2a

                var total_a2a = new Dictionary<string, decimal>();
                var types_a2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2a)
                {
                    var Setir_a2a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2a.Contains(row.Field<string>(1)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2a[type] = Setir_a2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E10"].Value = total_a2a["00"];
                worksheet1.Cells["F10"].Value = total_a2a["01"];
                worksheet1.Cells["G10"].Value = total_a2a["02"];
                worksheet1.Cells["I10"].Value = total_a2a["03"];
                worksheet1.Cells["K10"].Value = total_a2a["04"];
                worksheet1.Cells["O10"].Value = total_a2a["05"];
                var Setir_a2a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2a.Contains(row.Field<string>(1)))
                                        .ToList();
                decimal total_a2_1 = Setir_a2a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c10"].Value = total_a2_1;

                //******************Setir A2b
                var total_a2b = new Dictionary<string, decimal>();
                var types_a2b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2b)
                {
                    var Setir_a2b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2b.Contains(row.Field<string>(1)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2b[type] = Setir_a2b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E11"].Value = total_a2b["00"];
                worksheet1.Cells["F11"].Value = total_a2b["01"];
                worksheet1.Cells["G11"].Value = total_a2b["02"];
                worksheet1.Cells["I11"].Value = total_a2b["03"];
                worksheet1.Cells["K11"].Value = total_a2b["04"];
                worksheet1.Cells["O11"].Value = total_a2b["05"];
                var Setir_a2b_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2b.Contains(row.Field<string>(1)) )
                                        .ToList();
                decimal total_a2b_1_c = Setir_a2b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c11"].Value = total_a2b_1_c;
                //******************Setir A3a
                var total_a3a = new Dictionary<string, decimal>();
                var types_a3a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a)
                {
                    var Setir_a3a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a.Contains(row.Field<string>(1).Substring(0, 3)) && A3a_qeydno.Contains(row.Field<string>(1).Substring(9, 6)) 
                                        && row.Field<string>(2) == type &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a3a[type] = Setir_a3a.Sum(row => row.Field<decimal>(3)) / 1000;

                    string str = "15910000000001200100";

                    // Son iki rakamı almak için:
                    string lastTwoDigits = str.Substring(14, 2);
                }
                worksheet1.Cells["E13"].Value = -total_a3a["00"];
                worksheet1.Cells["F13"].Value = -total_a3a["01"];
                worksheet1.Cells["G13"].Value = -total_a3a["02"];
                worksheet1.Cells["I13"].Value = -total_a3a["03"];
                worksheet1.Cells["K13"].Value = -total_a3a["04"];
                worksheet1.Cells["O13"].Value = -total_a3a["05"];
                var Setir_a31_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a.Contains(row.Field<string>(1).Substring(0, 3)) && 
                                        A3a_qeydno.Contains(row.Field<string>(1).Substring(9, 6))&&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_3a_1_c = Setir_a31_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c13"].Value = -total_3a_1_c;

                //******************Setir A3a1
                var total_a3a1 = new Dictionary<string, decimal>();
                var types_a3a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a1)
                {
                    var Setir_a3a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a3a1[type] = Setir_a3a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E14"].Value = total_a3a1["00"];
                worksheet1.Cells["F14"].Value = total_a3a1["01"];
                worksheet1.Cells["G14"].Value = total_a3a1["02"];
                worksheet1.Cells["I14"].Value = total_a3a1["03"];
                worksheet1.Cells["K14"].Value = total_a3a1["04"];
                worksheet1.Cells["O14"].Value = total_a3a1["05"];
                var Setir_a3a1_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_1.Contains(row.Field<string>(1).Substring(0, 5)) )
                                        .ToList();
                decimal total_3a1_1_c = Setir_a3a1_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c14"].Value = total_3a1_1_c;

                //******************Setir A3a1a
                var total_a3a1a = new Dictionary<string, decimal>();
                var types_a3a1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a1a)
                {//&& row.Field<string>(2) == type
                    var Setir_a3a1a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                        && A3a_1a_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
               && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a3a1a[type] = Setir_a3a1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E15"].Value = -total_a3a1a["00"];
                worksheet1.Cells["F15"].Value = -total_a3a1a["01"];
                worksheet1.Cells["G15"].Value = -total_a3a1a["02"];
                worksheet1.Cells["I15"].Value = -total_a3a1a["03"];
                worksheet1.Cells["K15"].Value = -total_a3a1a["04"];
                worksheet1.Cells["O15"].Value = -total_a3a1a["05"];
                var Setir_a3a1a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_1a.Contains(row.Field<string>(1).Substring(0, 5))
                                        && A3a_1a_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
               && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_3a1a_1_c = Setir_a3a1a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C15"].Value = -total_3a1a_1_c;
                //******************Setir A3a2
                var total_a3a2 = new Dictionary<string, decimal>();
                var types_a3a2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a2)
                {
                    var Setir_a3a2 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_2.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                    )
                                        .ToList();
                    total_a3a2[type] = Setir_a3a2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E17"].Value = total_a3a2["00"];
                worksheet1.Cells["F17"].Value = total_a3a2["01"];
                worksheet1.Cells["G17"].Value = total_a3a2["02"];
                worksheet1.Cells["I17"].Value = total_a3a2["03"];
                worksheet1.Cells["K17"].Value = total_a3a2["04"];
                worksheet1.Cells["O17"].Value = total_a3a2["05"];
                var Setir_a3a2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_2.Contains(row.Field<string>(1).Substring(0, 5)) 
                                    )
                                        .ToList();
                decimal total_3a2_1_c = Setir_a3a2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c17"].Value = total_3a2_1_c;

                //******************Setir A3a2a
                var total_a3a2a = new Dictionary<string, decimal>();
                var types_a3a2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a2a)
                {
                    var Setir_a3a2a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_2a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                        && A3a_2a_yanasma.Contains(row.Field<string>(1).Substring(13, 2)))
                                        .ToList();
                    total_a3a2a[type] = Setir_a3a2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E18"].Value = -total_a3a2a["00"];
                worksheet1.Cells["F18"].Value = -total_a3a2a["01"];
                worksheet1.Cells["G18"].Value = -total_a3a2a["02"];
                worksheet1.Cells["I18"].Value = -total_a3a2a["03"];
                worksheet1.Cells["K18"].Value = -total_a3a2a["04"];
                worksheet1.Cells["O18"].Value = -total_a3a2a["05"];
                var Setir_a3a2a_1 = _dt_daily_report1.AsEnumerable()
                                         .Where(row => A3a_2a.Contains(row.Field<string>(1).Substring(0, 5)) 
                                         && A3a_2a_yanasma.Contains(row.Field<string>(1).Substring(13, 2)))
                                         .ToList();
                decimal total_3a2a_1_c = Setir_a3a2a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c18"].Value = -total_3a2a_1_c;

                
                //******************Setir A4_1a
                var total_a4_1a = new Dictionary<string, decimal>();
                var types_a4_1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1a)
                {
                    var Setir_a4_1a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1a[type] = Setir_a4_1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E23"].Value = total_a4_1a["00"];
                worksheet1.Cells["F23"].Value = total_a4_1a["01"];
                worksheet1.Cells["G23"].Value = total_a4_1a["02"];
                worksheet1.Cells["I23"].Value = total_a4_1a["03"];
                worksheet1.Cells["K23"].Value = total_a4_1a["04"];
                worksheet1.Cells["O23"].Value = total_a4_1a["05"];
                var Setir_a4_1a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A4_1a.Contains(row.Field<string>(1).Substring(0, 5)) )
                                        .ToList();
                decimal total_a4_1a_1 = Setir_a4_1a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c23"].Value = total_a4_1a_1;
                //******************Setir A4_1a1
                var total_a4_1a1 = new Dictionary<string, decimal>();
                var types_a4_1a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1a1)
                {
                    var Setir_a4_1a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1a1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1a1[type] = Setir_a4_1a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E24"].Value = total_a4_1a1["00"];
                worksheet1.Cells["F24"].Value = total_a4_1a1["01"];
                worksheet1.Cells["G24"].Value = total_a4_1a1["02"];
                worksheet1.Cells["I24"].Value = total_a4_1a1["03"];
                worksheet1.Cells["K24"].Value = total_a4_1a1["04"];
                worksheet1.Cells["O24"].Value = total_a4_1a1["05"];

                var Setir_a4_1a1_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1a1.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1a1_c = Setir_a4_1a1_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C24"].Value = total_a4_1a1_c;

                //******************Setir A4_1b
                var total_a4_1b = new Dictionary<string, decimal>();
                var types_a4_1b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1b)
                {
                    var Setir_a4_1b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1b[type] = Setir_a4_1b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E25"].Value = total_a4_1b["00"];
                worksheet1.Cells["F25"].Value = total_a4_1b["01"];
                worksheet1.Cells["G25"].Value = total_a4_1b["02"];
                worksheet1.Cells["I25"].Value = total_a4_1b["03"];
                worksheet1.Cells["K25"].Value = total_a4_1b["04"];
                worksheet1.Cells["O25"].Value = total_a4_1b["05"];

                var Setir_a4_1b_d = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1b_d = Setir_a4_1b_d.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c25"].Value = total_a4_1b_d;

                //******************Setir A4_1b
                var total_a4_1b_1 = new Dictionary<string, decimal>();
                var types_a4_1b_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1b_1)
                {
                    var Setir_a4_1b_1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1b1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1b_1[type] = Setir_a4_1b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E26"].Value = total_a4_1b_1["00"];
                worksheet1.Cells["F26"].Value = total_a4_1b_1["01"];
                worksheet1.Cells["G26"].Value = total_a4_1b_1["02"];
                worksheet1.Cells["I26"].Value = total_a4_1b_1["03"];
                worksheet1.Cells["K26"].Value = total_a4_1b_1["04"];
                worksheet1.Cells["O26"].Value = total_a4_1b_1["05"];

                var Setir_a4_1b_1_d = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1b1.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1b_1_d = Setir_a4_1b_1_d.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c26"].Value = total_a4_1b_d;

                //******************Setir A5
                var total_a5 = new Dictionary<string, decimal>();
                var types_a5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a5)
                {
                    var Setir_a5 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A5.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a5[type] = Setir_a5.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E40"].Value = total_a5["00"];
                worksheet1.Cells["F40"].Value = total_a5["01"];
                worksheet1.Cells["G40"].Value = total_a5["02"];
                worksheet1.Cells["I40"].Value = total_a5["03"];
                worksheet1.Cells["K40"].Value = total_a5["04"];
                worksheet1.Cells["O40"].Value = total_a5["05"];

                var Setir_a5_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A5.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a5_c = Setir_a5_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C40"].Value = total_a5_c;

                //******************Setir A6
                var total_a6 = new Dictionary<string, decimal>();
                var types_a6 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a6)
                {
                    var Setir_a6 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A6.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a6[type] = Setir_a6.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E41"].Value = total_a6["00"];
                worksheet1.Cells["F41"].Value = total_a6["01"];
                worksheet1.Cells["G41"].Value = total_a6["02"];
                worksheet1.Cells["I41"].Value = total_a6["03"];
                worksheet1.Cells["K41"].Value = total_a6["04"];
                worksheet1.Cells["O41"].Value = total_a6["05"];

                var Setir_a6_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A6.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a6_1 = Setir_a6_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c41"].Value = total_a6_1;

                //******************Setir A6c
                var total_a6c = new Dictionary<string, decimal>();
                var types_a6c = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a6c)
                {
                    var Setir_a6c = _dt_daily_report.AsEnumerable()
                                        .Where(row => A6c.Contains(row.Field<string>(1)) && row.Field<string>(2) == type
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a6c[type] = Setir_a6c.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E44"].Value = -total_a6c["00"];
                worksheet1.Cells["F44"].Value = -total_a6c["01"];
                worksheet1.Cells["G44"].Value = -total_a6c["02"];
                worksheet1.Cells["I44"].Value = -total_a6c["03"];
                worksheet1.Cells["K44"].Value = -total_a6c["04"];
                worksheet1.Cells["O44"].Value = -total_a6c["05"];

                var Setir_a6c_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A6c.Contains(row.Field<string>(1))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a6c_c = Setir_a6c_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c44"].Value = -total_a6c_c;

                //******************Setir A7
               
                var total_a7 = new Dictionary<string, decimal>();
                var types_a7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a7)
                {
                    var Setir_a7 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A7.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a7[type] = Setir_a7.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E45"].Value = total_a7["00"];
                worksheet1.Cells["F45"].Value = total_a7["01"];
                worksheet1.Cells["G45"].Value = total_a7["02"];
                worksheet1.Cells["I45"].Value = total_a7["03"];
                worksheet1.Cells["K45"].Value = total_a7["04"];
                worksheet1.Cells["O45"].Value = total_a7["05"];

                var Setir_a7_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A7.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a7_1 = Setir_a7_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C45"].Value = total_a7_1;
                //******************Setir A8.1
                var total_a8_1 = new Dictionary<string, decimal>();
                var types_a8_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_1)
                {
                    var Setir_a8_1 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "biznes" 
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_1[type] = Setir_a8_1.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E48"].Value = total_a8_1["00"];
                worksheet1.Cells["F48"].Value = total_a8_1["01"];
                worksheet1.Cells["G48"].Value = total_a8_1["02"];
                worksheet1.Cells["I48"].Value = total_a8_1["03"];
                worksheet1.Cells["K48"].Value = total_a8_1["04"];
                worksheet1.Cells["O48"].Value = total_a8_1["05"];

                var Setir_a8_1_1 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "biznes"))
                                    .ToList();
                decimal total_a8_1_1 = Setir_a8_1_1.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C48"].Value = total_a8_1_1;

                //******************Setir A8.2
                var total_a8_2 = new Dictionary<string, decimal>();
                var types_a8_2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_2)
                {
                    var Setir_a8_2 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki"
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_2[type] = Setir_a8_2.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E49"].Value = total_a8_2["00"];
                worksheet1.Cells["F49"].Value = total_a8_2["01"];
                worksheet1.Cells["G49"].Value = total_a8_2["02"];
                worksheet1.Cells["I49"].Value = total_a8_2["03"];
                worksheet1.Cells["K49"].Value = total_a8_2["04"];
                worksheet1.Cells["O49"].Value = total_a8_2["05"];

                var Setir_a8_1_2 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki"))
                                    .ToList();
                decimal total_a8_1_2 = Setir_a8_1_2.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C49"].Value = total_a8_1_2;


                //******************Setir A8.3
                var total_a8_3 = new Dictionary<string, decimal>();
                var types_a8_3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_3)
                {
                    var Setir_a8_3 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "dasinmaz"
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_3[type] = Setir_a8_3.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E50"].Value = total_a8_3["00"];
                worksheet1.Cells["F50"].Value = total_a8_3["01"];
                worksheet1.Cells["G50"].Value = total_a8_3["02"];
                worksheet1.Cells["I50"].Value = total_a8_3["03"];
                worksheet1.Cells["K50"].Value = total_a8_3["04"];
                worksheet1.Cells["O50"].Value = total_a8_3["05"];

                var Setir_a8_1_3 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "dasinmaz"))
                                    .ToList();
                decimal total_a8_1_3 = Setir_a8_1_3.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C50"].Value = total_a8_1_3;


                //******************Setir A8b
                var total_a8b = new Dictionary<string, decimal>();
                var types_a8b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8b)
                {
                    var Setir_a8b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A8b.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type
                                        && A8b_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    

                    total_a8b[type] = Setir_a8b.Sum(row => row.Field<decimal>(3)) / 1000;

                    
                }
                worksheet1.Cells["E51"].Value = -total_a8b["00"];
                worksheet1.Cells["F51"].Value = -total_a8b["01"];
                worksheet1.Cells["G51"].Value = -total_a8b["02"];
                worksheet1.Cells["I51"].Value = -total_a8b["03"];
                worksheet1.Cells["K51"].Value = -total_a8b["04"];
                worksheet1.Cells["O51"].Value = -total_a8b["05"];

                var Setir_a8b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A8b.Contains(row.Field<string>(1).Substring(0, 3))
                                     && A8b_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a8b_c = Setir_a8b_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C51"].Value = -total_a8b_c;

                //******************Setir A9

                var total_a9 = new Dictionary<string, decimal>();
                var types_a9 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a9)
                {
                    var Setir_a9 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A9.Contains(row.Field<string>(1).Substring(0, 2)) && !A9_istisna.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a9[type] = Setir_a9.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E52"].Value = total_a9["00"];
                worksheet1.Cells["F52"].Value = total_a9["01"];
                worksheet1.Cells["G52"].Value = total_a9["02"];
                worksheet1.Cells["I52"].Value = total_a9["03"];
                worksheet1.Cells["K52"].Value = total_a9["04"];
                worksheet1.Cells["O52"].Value = total_a9["05"];

                var Setir_a9_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A9.Contains(row.Field<string>(1).Substring(0, 2)) && !A9_istisna.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a9_c = Setir_a9_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C52"].Value = total_a9_c;

                //******************Setir A10
                var total_a10 = new Dictionary<string, decimal>();
                var types_a10 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a10)
                {
                    var Setir_a10 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a10[type] = Setir_a10.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E53"].Value = total_a10["00"];
                worksheet1.Cells["F53"].Value = total_a10["01"];
                worksheet1.Cells["G53"].Value = total_a10["02"];
                worksheet1.Cells["I53"].Value = total_a10["03"];
                worksheet1.Cells["K53"].Value = total_a10["04"];
                worksheet1.Cells["O53"].Value = total_a10["05"];

                var Setir_a10_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10_c = Setir_a10_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C53"].Value = total_a10_c;


                //******************Setir A10a
                var total_a10a = new Dictionary<string, decimal>();
                var types_a10a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a10a)
                {
                    var Setir_a10a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A10a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a10a[type] = Setir_a10a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E54"].Value = total_a10a["00"];
                worksheet1.Cells["F54"].Value = total_a10a["01"];
                worksheet1.Cells["G54"].Value = total_a10a["02"];
                worksheet1.Cells["I54"].Value = total_a10a["03"];
                worksheet1.Cells["K54"].Value = total_a10a["04"];
                worksheet1.Cells["O54"].Value = total_a10a["05"];

                var Setir_a10a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10a_c = Setir_a10a_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C54"].Value = total_a10a_c;


                //******************Setir A10b

                var Setir_a10b_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                //var Setir_a10b_c_27013 = _dt_daily_report1.AsEnumerable()
                //                    .Where(row => A10b_27013.Contains(row.Field<string>(1).Substring(0, 5)))
                //                    .ToList();
                decimal total_a10b_c = Setir_a10b_1.Sum(row => row.Field<decimal>(3)) / 1000;

                //var Setir_a10b_e_27012 = _dt_daily_report.AsEnumerable()
                //                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                //                    .ToList();
                var Setir_a10b = _dt_daily_report.AsEnumerable()
                                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10b_e = Setir_a10b.Sum(row => row.Field<decimal>(3))  / 1000;

                worksheet1.Cells["C55"].Value = total_a10b_c;
                worksheet1.Cells["E55"].Value = total_a10b_e;

                //******************Setir A12

                var Setir_a12_c_qisa = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                

                var Setir_a12_c_uzun = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal top1u = Setir_a12_c_uzun.Sum(row => row.Field<decimal>(3))/1000;
                decimal top1q = Setir_a12_c_qisa.Sum(row => row.Field<decimal>(3))/1000;
                decimal total_a12_c = (Setir_a12_c_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_c_uzun.Sum(row => row.Field<decimal>(3))) / 1000; ;

                


                var Setir_a12_e_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                
                var Setir_a12_e_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "00"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_e = (Setir_a12_e_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_e_uzun.Sum(row => row.Field<decimal>(3))) / 1000;
                
                var Setir_a12_f_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                var Setir_a12_f_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "01"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_f = (Setir_a12_f_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_f_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_g_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                var Setir_a12_g_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "02"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_g = (Setir_a12_g_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_g_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_i_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                var Setir_a12_i_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "03"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_i = (Setir_a12_i_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_i_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_k_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                var Setir_a12_k_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "04"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_k = (Setir_a12_k_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_k_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_o_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                var Setir_a12_o_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "05"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_o = (Setir_a12_o_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_o_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                worksheet1.Cells["C57"].Value = total_a12_c;
                worksheet1.Cells["E57"].Value = total_a12_e;
                worksheet1.Cells["F57"].Value = total_a12_f;
                worksheet1.Cells["G57"].Value = total_a12_g;
                worksheet1.Cells["I57"].Value = total_a12_i;
                worksheet1.Cells["K57"].Value = total_a12_k;
                worksheet1.Cells["O57"].Value = total_a12_o;

                //******************Setir A12a

                var Setir_a12a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a12a_c = Setir_a12a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_a12a_e = Setir_a12a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_a12a_f = Setir_a12a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_a12a_g = Setir_a12a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_a12a_i = Setir_a12a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_a12a_k = Setir_a12a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_a12a_o = Setir_a12a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C58"].Value = total_a12a_c;
                worksheet1.Cells["E58"].Value = total_a12a_e;
                worksheet1.Cells["F58"].Value = total_a12a_f;
                worksheet1.Cells["G58"].Value = total_a12a_g;
                worksheet1.Cells["I58"].Value = total_a12a_i;
                worksheet1.Cells["K58"].Value = total_a12a_k;
                worksheet1.Cells["O58"].Value = total_a12a_o;

                //******************Setir A12b
                //A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2)))

                var total_a12b = new Dictionary<string, decimal>();
                var types_a12b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a12b)
                {
                    var Setir_a12b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2))
                                        && row.Field<string>(2) == type)
                                        .ToList();
                    total_a12b[type] = Setir_a12b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E59"].Value = -total_a12b["00"];
                worksheet1.Cells["F59"].Value = -total_a12b["01"];
                worksheet1.Cells["G59"].Value = -total_a12b["02"];
                worksheet1.Cells["I59"].Value = -total_a12b["03"];
                worksheet1.Cells["K59"].Value = -total_a12b["04"];
                worksheet1.Cells["O59"].Value = -total_a12b["05"];

                var Setir_a12b_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2)))
                                    .ToList();
                decimal total_a12b_1 = Setir_a12b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C59"].Value = -total_a12b_1;


                DataTable filteredDataTable = _dt_daily_report1.Clone();
                foreach (var row in Setir_a12_c_qisa)
                {
                    filteredDataTable.Rows.Add(row.ItemArray);
                }

                // DataGridView'e yeni DataTable'ı atayarak güncelle
                dataGridView1.DataSource = filteredDataTable;

                //******************Setir 12b
                var total_a12c = new Dictionary<string, decimal>();
                var types_a12c = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a12c)
                {
                    var Setir_a12c = _dt_daily_report.AsEnumerable()
                                        .Where(row => ((A12c.Contains(row.Field<string>(1).Substring(0, 3)) && !A12c_yanasma.Contains(row.Field<string>(1).Substring(13, 2)))|| A12c_uzun.Contains(row.Field<string>(1)))
                                        && row.Field<string>(2) == type
                                          &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a12c[type] = Setir_a12c.Sum(row => row.Field<decimal>(3)) / 1000;

                    
                }
                worksheet1.Cells["E60"].Value = -total_a12c["00"] + total_a12b["00"]; //- total_a3a["00"] - total_a6c["00"] - total_a8b["00"] - total_a12b["00"];
                worksheet1.Cells["F60"].Value = -total_a12c["01"] + total_a12b["01"];// - total_a3a["01"] - total_a6c["01"] - total_a8b["01"] - total_a12b["01"];
                worksheet1.Cells["G60"].Value = -total_a12c["02"] + total_a12b["02"];// - total_a3a["02"] - total_a6c["02"] - total_a8b["02"] - total_a12b["02"];
                worksheet1.Cells["I60"].Value = -total_a12c["03"] + total_a12b["03"]; //- total_a3a["03"] - total_a6c["03"] - total_a8b["03"] - total_a12b["03"];
                worksheet1.Cells["K60"].Value = -total_a12c["04"] + total_a12b["04"]; //- total_a3a["04"] - total_a6c["04"] - total_a8b["04"] - total_a12b["04"];
                worksheet1.Cells["O60"].Value = -total_a12c["05"] + total_a12b["05"];// - total_a3a["05"] - total_a6c["05"] - total_a8b["05"] - total_a12b["05"];
                var Setir_a12c_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => ((A12c.Contains(row.Field<string>(1).Substring(0, 3)) && !A12c_yanasma.Contains(row.Field<string>(1).Substring(13, 2))) || A12c_uzun.Contains(row.Field<string>(1)))
                                        &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_12c_1_c = Setir_a12c_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c60"].Value = -total_12c_1_c + total_a12b_1;// - total_a12b_1; //- total_3a_1_c - total_a6c_c - total_a8b_c - total_a12b_1;
                

                //******************Setir B1a

                var Setir_b1a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1a_c = Setir_b1a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1a_e = Setir_b1a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1a_f = Setir_b1a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1a_g = Setir_b1a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1a_i = Setir_b1a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1a_k = Setir_b1a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1a_o = Setir_b1a_o.Sum(row => row.Field<decimal>(3)) / 1000;



                worksheet1.Cells["C64"].Value = -total_b1a_c;
                worksheet1.Cells["E64"].Value = -total_b1a_e;
                worksheet1.Cells["F64"].Value = -total_b1a_f;
                worksheet1.Cells["G64"].Value = -total_b1a_g;
                worksheet1.Cells["I64"].Value = -total_b1a_i;
                worksheet1.Cells["K64"].Value = -total_b1a_k;
                worksheet1.Cells["O64"].Value = -total_b1a_o;

                //******************Setir B1b

                var Setir_b1b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1b_c = Setir_b1b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1b_e = Setir_b1b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1b_f = Setir_b1b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1b_g = Setir_b1b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1b_i = Setir_b1b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1b_k = Setir_b1b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1b_o = Setir_b1b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C65"].Value = -total_b1b_c;
                worksheet1.Cells["E65"].Value = -total_b1b_e;
                worksheet1.Cells["F65"].Value = -total_b1b_f;
                worksheet1.Cells["G65"].Value = -total_b1b_g;
                worksheet1.Cells["I65"].Value = -total_b1b_i;
                worksheet1.Cells["K65"].Value = -total_b1b_k;
                worksheet1.Cells["O65"].Value = -total_b1b_o;

                //******************Setir B1c

                var Setir_b1c_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1c_c = Setir_b1c_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1c_e = Setir_b1c_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1c_f = Setir_b1c_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1c_g = Setir_b1c_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1c_i = Setir_b1c_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1c_k = Setir_b1c_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1c_o = Setir_b1c_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C66"].Value = -total_b1c_c;
                worksheet1.Cells["E66"].Value = -total_b1c_e;
                worksheet1.Cells["F66"].Value = -total_b1c_f;
                worksheet1.Cells["G66"].Value = -total_b1c_g;
                worksheet1.Cells["I66"].Value = -total_b1c_i;
                worksheet1.Cells["K66"].Value = -total_b1c_k;
                worksheet1.Cells["O66"].Value = -total_b1c_o;

                //******************Setir B2b

                var Setir_b2b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b2b_c = Setir_b2b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b2b_e = Setir_b2b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b2b_f = Setir_b2b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b2b_g = Setir_b2b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b2b_i = Setir_b2b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b2b_k = Setir_b2b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b2b_o = Setir_b2b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C69"].Value = total_b2b_c;
                worksheet1.Cells["E69"].Value = total_b2b_e;
                worksheet1.Cells["F69"].Value = total_b2b_f;
                worksheet1.Cells["G69"].Value = total_b2b_g;
                worksheet1.Cells["I69"].Value = total_b2b_i;
                worksheet1.Cells["K69"].Value = total_b2b_k;
                worksheet1.Cells["O69"].Value = total_b2b_o;

                //******************Setir B4a

                var Setir_b4a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b4a_c = Setir_b4a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b4a_e = Setir_b4a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b4a_f = Setir_b4a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b4a_g = Setir_b4a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b4a_i = Setir_b4a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b4a_k = Setir_b4a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b4a_o = Setir_b4a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C72"].Value = -total_b4a_c;
                worksheet1.Cells["E72"].Value = -total_b4a_e;
                worksheet1.Cells["F72"].Value = -total_b4a_f;
                worksheet1.Cells["G72"].Value = -total_b4a_g;
                worksheet1.Cells["I72"].Value = -total_b4a_i;
                worksheet1.Cells["K72"].Value = -total_b4a_k;
                worksheet1.Cells["O72"].Value = -total_b4a_o;

                //******************Setir B4b

                var Setir_b4b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b4b_c = Setir_b4b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b4b_e = Setir_b4b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b4b_f = Setir_b4b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b4b_g = Setir_b4b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b4b_i = Setir_b4b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b4b_k = Setir_b4b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b4b_o = Setir_b4b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C73"].Value = -total_b4b_c;
                worksheet1.Cells["E73"].Value = -total_b4b_e;
                worksheet1.Cells["F73"].Value = -total_b4b_f;
                worksheet1.Cells["G73"].Value = -total_b4b_g;
                worksheet1.Cells["I73"].Value = -total_b4b_i;
                worksheet1.Cells["K73"].Value = -total_b4b_k;
                worksheet1.Cells["O73"].Value = -total_b4b_o;

                //******************Setir B5a

                var Setir_b5a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b5a_c = Setir_b5a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b5a_e = Setir_b5a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b5a_f = Setir_b5a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b5a_g = Setir_b5a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b5a_i = Setir_b5a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b5a_k = Setir_b5a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b5a_o = Setir_b5a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C75"].Value = -total_b5a_c;
                worksheet1.Cells["E75"].Value = -total_b5a_e;
                worksheet1.Cells["F75"].Value = -total_b5a_f;
                worksheet1.Cells["G75"].Value = -total_b5a_g;
                worksheet1.Cells["I75"].Value = -total_b5a_i;
                worksheet1.Cells["K75"].Value = -total_b5a_k;
                worksheet1.Cells["O75"].Value = -total_b5a_o;

                //******************Setir B5a1

                var Setir_b5a1_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b5a1_c = Setir_b5a1_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b5a1_e = Setir_b5a1_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b5a1_f = Setir_b5a1_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b5a1_g = Setir_b5a1_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b5a1_i = Setir_b5a1_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b5a1_k = Setir_b5a1_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b5a1_o = Setir_b5a1_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C76"].Value = -total_b5a1_c;
                worksheet1.Cells["E76"].Value = -total_b5a1_e;
                worksheet1.Cells["F76"].Value = -total_b5a1_f;
                worksheet1.Cells["G76"].Value = -total_b5a1_g;
                worksheet1.Cells["I76"].Value = -total_b5a1_i;
                worksheet1.Cells["K76"].Value = -total_b5a1_k;
                worksheet1.Cells["O76"].Value = -total_b5a1_o;

                //******************Setir B8

                var Setir_b8_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b8_c = Setir_b8_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b8_e = Setir_b8_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b8_f = Setir_b8_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b8_g = Setir_b8_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b8_i = Setir_b8_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b8_k = Setir_b8_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b8_o = Setir_b8_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C85"].Value = total_b8_c;
                worksheet1.Cells["E85"].Value = total_b8_e;
                worksheet1.Cells["F85"].Value = total_b8_f;
                worksheet1.Cells["G85"].Value = total_b8_g;
                worksheet1.Cells["I85"].Value = total_b8_i;
                worksheet1.Cells["K85"].Value = total_b8_k;
                worksheet1.Cells["O85"].Value = total_b8_o;

                //******************Setir B10

                var Setir_b10_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b10_c = Setir_b10_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b10_e = Setir_b10_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b10_f = Setir_b10_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b10_g = Setir_b10_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b10_i = Setir_b10_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b10_k = Setir_b10_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b10_o = Setir_b10_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C88"].Value = -total_b10_c;
                worksheet1.Cells["E88"].Value = -total_b10_e;
                worksheet1.Cells["F88"].Value = -total_b10_f;
                worksheet1.Cells["G88"].Value = -total_b10_g;
                worksheet1.Cells["I88"].Value = -total_b10_i;
                worksheet1.Cells["K88"].Value = -total_b10_k;
                worksheet1.Cells["O88"].Value = -total_b10_o;

                //******************Setir B10a
                var total_b10a = new Dictionary<string, decimal>();
                var types_b10a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b10a)
                {
                    var Setir_b10a = _dt_daily_report.AsEnumerable()
                                        .Where(row => B10a.Contains(row.Field<string>(1).Substring(0, 5)) 
                                        && row.Field<string>(2) == type)
                                        .ToList();
                    total_b10a[type] = Setir_b10a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E89"].Value = total_b10a["00"];
                worksheet1.Cells["F89"].Value = total_b10a["01"];
                worksheet1.Cells["G89"].Value = total_b10a["02"];
                worksheet1.Cells["I89"].Value = total_b10a["03"];
                worksheet1.Cells["K89"].Value = total_b10a["04"];
                worksheet1.Cells["O89"].Value = total_b10a["05"];

                var Setir_b10a_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B10a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b10a_1 = Setir_b10a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C89"].Value = total_b10a_1;

                //******************Setir B12
                var total_b12 = new Dictionary<string, decimal>();
                var total_b12_eh = new Dictionary<string, decimal>();
                var types_b12 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b12)
                {
                    var Setir_b12 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>(B12.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00"))
                                        .ToList();
                    total_b12[type] = Setir_b12.Sum(row => row.Field<decimal>(3)) / 1000;
                    decimal cem= Setir_b12.Sum(row => row.Field<decimal>(3)) / 1000;
                    var Setir_b12_eh = _dt_daily_report.AsEnumerable()
                                        .Where(row => (B12.Contains(row.Field<string>(1).Substring(0, 5)) ||
                                       (B12_eht.Contains(row.Field<string>(1).Substring(0, 3)) &&
                                       Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) < 25
                                       && Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) != 100)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b12_eh[type] = Setir_b12_eh.Sum(row => row.Field<decimal>(3)) / 1000;
                    decimal cem1 = Setir_b12_eh.Sum(row => row.Field<decimal>(3)) / 1000;
                }


                worksheet1.Cells["E92"].Value = -total_b12_eh["00"];
                worksheet1.Cells["F92"].Value = -total_b12_eh["01"];
                worksheet1.Cells["G92"].Value = -total_b12_eh["02"];
                worksheet1.Cells["I92"].Value = -total_b12_eh["03"];
                worksheet1.Cells["K92"].Value = -total_b12_eh["04"];
                worksheet1.Cells["O92"].Value = -total_b12_eh["05"];
                var Setir_b12_1 = _dt_daily_report1.AsEnumerable()
                                        .Where (row => (B12.Contains(row.Field<string>(1).Substring(0, 5)) ||
                                        (B12_eht.Contains(row.Field<string>(1).Substring(0, 3)) &&
                                        Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) < 25
                                        && Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) != 100)))
                                        .ToList();
                decimal total_b12_1 = Setir_b12_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C92"].Value = -total_b12_1;

                //******************Setir B12a
                var total_b12a = new Dictionary<string, decimal>();
                var types_b12a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b12a)
                {
                    var Setir_b12a = _dt_daily_report.AsEnumerable()
                                        .Where(row => B12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b12a[type] = Setir_b12a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E93"].Value = -total_b12a["00"];
                worksheet1.Cells["F93"].Value = -total_b12a["01"];
                worksheet1.Cells["G93"].Value = -total_b12a["02"];
                worksheet1.Cells["I93"].Value = -total_b12a["03"];
                worksheet1.Cells["K93"].Value = -total_b12a["04"];
                worksheet1.Cells["O93"].Value = -total_b12a["05"];
                var Setir_b12a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => B12a.Contains(row.Field<string>(1).Substring(0, 5)) )
                                        .ToList();
                decimal total_b12a_1 = Setir_b12a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C93"].Value = -total_b12a_1;
                //******************Setir B12
                var total_b13 = new Dictionary<string, decimal>();
                var types_b13 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b13)
                {
                    var Setir_b13 = _dt_daily_report.AsEnumerable()
                                        .Where(row => B13.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b13[type] = Setir_b13.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E94"].Value = -total_b13["00"];
                worksheet1.Cells["F94"].Value = -total_b13["01"];
                worksheet1.Cells["G94"].Value = -total_b13["02"];
                worksheet1.Cells["I94"].Value = -total_b13["03"];
                worksheet1.Cells["K94"].Value = -total_b13["04"];
                worksheet1.Cells["O94"].Value = -total_b13["05"];
                var Setir_b13_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => B13.Contains(row.Field<string>(1).Substring(0, 5)) )
                                        .ToList();
                decimal total_b13_1 = Setir_b13_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C94"].Value = -total_b13_1;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 1
                var total_bo1 = new Dictionary<string, decimal>();
                var types_bo1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo1)
                {
                    var Setir_bo1 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99550" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo1[type] = Setir_bo1.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E99"].Value = total_bo1["00"];
                worksheet1.Cells["F99"].Value = total_bo1["01"];
                worksheet1.Cells["G99"].Value = total_bo1["02"];
                worksheet1.Cells["I99"].Value = total_bo1["03"];
                worksheet1.Cells["K99"].Value = total_bo1["04"];
                worksheet1.Cells["O99"].Value = total_bo1["05"];
                var Setir_1_bo1 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99550" )
                                    .ToList();
                decimal total_1_bo1 = Setir_1_bo1.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C99"].Value = total_1_bo1;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 2
                var total_bo2 = new Dictionary<string, decimal>();
                var types_bo2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo2)
                {
                    var Setir_bo2 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99530" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo2[type] = Setir_bo2.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E100"].Value = total_bo2["00"];
                worksheet1.Cells["F100"].Value = total_bo2["01"];
                worksheet1.Cells["G100"].Value = total_bo2["02"];
                worksheet1.Cells["I100"].Value = total_bo2["03"];
                worksheet1.Cells["K100"].Value = total_bo2["04"];
                worksheet1.Cells["O100"].Value = total_bo2["05"];
                var Setir_1_bo2 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99530")
                                    .ToList();
                decimal total_1_bo2 = Setir_1_bo2.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C100"].Value = total_1_bo2;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 3
                var total_bo3 = new Dictionary<string, decimal>();
                var types_bo3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo3)
                {
                    var Setir_bo3 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99531" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo3[type] = Setir_bo3.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E101"].Value = total_bo3["00"];
                worksheet1.Cells["F101"].Value = total_bo3["01"];
                worksheet1.Cells["G101"].Value = total_bo3["02"];
                worksheet1.Cells["I101"].Value = total_bo3["03"];
                worksheet1.Cells["K101"].Value = total_bo3["04"];
                worksheet1.Cells["O101"].Value = total_bo3["05"];
                var Setir_1_bo3 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99531")
                                    .ToList();
                decimal total_1_bo3 = Setir_1_bo3.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C101"].Value = total_1_bo3;

                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 7
                var total_bo7 = new Dictionary<string, decimal>();
                var types_bo7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo7)
                {
                    var Setir_bo7 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99531" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo7[type] = Setir_bo7.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E105"].Value = total_bo7["00"];
                worksheet1.Cells["F105"].Value = total_bo7["01"];
                worksheet1.Cells["G105"].Value = total_bo7["02"];
                worksheet1.Cells["I105"].Value = total_bo7["03"];
                worksheet1.Cells["K105"].Value = total_bo7["04"];
                worksheet1.Cells["O105"].Value = total_bo7["05"];
                var Setir_1_bo7 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99531")
                                    .ToList();
                decimal total_1_bo7 = Setir_1_bo7.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C105"].Value = total_1_bo7;

                //******************Setir Risklər barədə məlumatlar 1a
                var total_risk_1a = new Dictionary<string, decimal>();
                var types_risk_1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_1a)
                {
                    var Setir_risk_1a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" && 
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar") 
                                        && row.Field<string>(4) == "vk" && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_1a[type] = Setir_risk_1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E118"].Value = total_risk_1a["00"];
                worksheet1.Cells["F118"].Value = total_risk_1a["01"];
                worksheet1.Cells["G118"].Value = total_risk_1a["02"];
                worksheet1.Cells["I118"].Value = total_risk_1a["03"];
                worksheet1.Cells["K118"].Value = total_risk_1a["04"];
                worksheet1.Cells["O118"].Value = total_risk_1a["05"];
                var Setir_1_risk_1a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(4) == "vk" )
                                        .ToList();
                decimal total_1_risk_1a = Setir_1_risk_1a.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C118"].Value = total_1_risk_1a;

                //******************Setir Risklər barədə məlumatlar 1b
                var total_risk_1b = new Dictionary<string, decimal>();
                var types_risk_1b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_1b)
                {
                    var Setir_risk_1b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki" 
                                        && row.Field<string>(4) == "vk" && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_1b[type] = Setir_risk_1b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E119"].Value = total_risk_1b["00"];
                worksheet1.Cells["F119"].Value = total_risk_1b["01"];
                worksheet1.Cells["G119"].Value = total_risk_1b["02"];
                worksheet1.Cells["I119"].Value = total_risk_1b["03"];
                worksheet1.Cells["K119"].Value = total_risk_1b["04"];
                worksheet1.Cells["O119"].Value = total_risk_1b["05"];
                var Setir_1_risk_1b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(4) == "vk")
                                        .ToList();
                decimal total_1_risk_1b = Setir_1_risk_1b.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C119"].Value = total_1_risk_1b;

                //******************Setir Risklər barədə məlumatlar 2a
                var total_risk_2a = new Dictionary<string, decimal>();
                var types_risk_2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_2a)
                {
                    var Setir_risk_2a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_2a[type] = Setir_risk_2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E122"].Value = total_risk_2a["00"];
                worksheet1.Cells["F122"].Value = total_risk_2a["01"];
                worksheet1.Cells["G122"].Value = total_risk_2a["02"];
                worksheet1.Cells["I122"].Value = total_risk_2a["03"];
                worksheet1.Cells["K122"].Value = total_risk_2a["04"];
                worksheet1.Cells["O122"].Value = total_risk_2a["05"];
                var Setir_1_risk_2a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_2a = Setir_1_risk_2a.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C122"].Value = total_1_risk_2a;

                //******************Setir Risklər barədə məlumatlar 2b
                var total_risk_2b = new Dictionary<string, decimal>();
                var types_risk_2b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_2b)
                {
                    var Setir_risk_2b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki" 
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_2b[type] = Setir_risk_2b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E123"].Value = total_risk_2b["00"];
                worksheet1.Cells["F123"].Value = total_risk_2b["01"];
                worksheet1.Cells["G123"].Value = total_risk_2b["02"];
                worksheet1.Cells["I123"].Value = total_risk_2b["03"];
                worksheet1.Cells["K123"].Value = total_risk_2b["04"];
                worksheet1.Cells["O123"].Value = total_risk_2b["05"];
                var Setir_1_risk_2b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_2b = Setir_1_risk_2b.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C123"].Value = total_1_risk_2b;

                //******************Setir Risklər barədə məlumatlar 3a
                var total_risk_3a = new Dictionary<string, decimal>();
                var types_risk_3a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_3a)
                {
                    var Setir_risk_3a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_3a[type] = Setir_risk_3a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E126"].Value = total_risk_3a["00"];
                worksheet1.Cells["F126"].Value = total_risk_3a["01"];
                worksheet1.Cells["F126"].Value = total_risk_3a["02"];
                worksheet1.Cells["G126"].Value = total_risk_3a["03"];
                worksheet1.Cells["K126"].Value = total_risk_3a["04"];
                worksheet1.Cells["O126"].Value = total_risk_3a["05"];
                var Setir_1_risk_3a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_3a = Setir_1_risk_3a.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C126"].Value = total_1_risk_3a;

                //******************Setir Risklər barədə məlumatlar 3b
                var total_risk_3b = new Dictionary<string, decimal>();
                var types_risk_3b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_3b)
                {
                    var Setir_risk_3b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_3b[type] = Setir_risk_3b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E127"].Value = total_risk_3b["00"];
                worksheet1.Cells["F127"].Value = total_risk_3b["01"];
                worksheet1.Cells["G127"].Value = total_risk_3b["02"];
                worksheet1.Cells["I127"].Value = total_risk_3b["03"];
                worksheet1.Cells["K127"].Value = total_risk_3b["04"];
                worksheet1.Cells["O127"].Value = total_risk_3b["05"];
                var Setir_1_risk_3b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_3b = Setir_1_risk_3b.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C127"].Value = total_1_risk_3b;

                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 4
                var total_bo4 = new Dictionary<string, decimal>();
                var types_bo4 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo4)
                {
                    var Setir_bo4 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99300" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo4[type] = Setir_bo4.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E129"].Value = total_bo4["00"];
                worksheet1.Cells["F129"].Value = total_bo4["01"];
                worksheet1.Cells["G129"].Value = total_bo4["02"];
                worksheet1.Cells["I129"].Value = total_bo4["03"];
                worksheet1.Cells["K129"].Value = total_bo4["04"];
                worksheet1.Cells["O129"].Value = total_bo4["05"];

                var Setir_1_bo4 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99300")
                                    .ToList();
                decimal total_1_bo4 = Setir_1_bo4.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C129"].Value = total_1_bo4;


                //******************Setir Likvildik riskləri  1
                var total_lr1 = new Dictionary<string, decimal>();
                var types_lr1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr1)
                {
                    var Setir_lr1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => likvid_risk_1.Contains(row.Field<string>(1).Substring(0, 5))
                              && row.Field<string>(2) == type)
                                        .ToList();
                    //total_lr1[type] = Setir_lr1.Sum(row => row.Field<decimal>(3)) / 1000;
                    total_lr1[type] = Setir_lr1.Sum(row =>
                    {
                        string substringValue = row.Field<string>(1).Substring(0, 5);

                        // Eğer substringValue "15020" veya "15025" ise, meblağı %25 azalt
                        decimal discountFactor = (substringValue == "15020" || substringValue == "15025") ? 0.75m : 1.0m;

                        return row.Field<decimal>(3) * discountFactor / 1000;
                    });
                }
                worksheet1.Cells["E134"].Value = total_lr1["00"];
                worksheet1.Cells["F134"].Value = total_lr1["01"];
                worksheet1.Cells["G134"].Value = total_lr1["02"];
                worksheet1.Cells["I134"].Value = total_lr1["03"];
                worksheet1.Cells["K134"].Value = total_lr1["04"];
                worksheet1.Cells["O134"].Value = total_lr1["05"];

                var sh_67_00 = _dt_daily_report1.AsEnumerable()
                .Where(row => likvid_risk_1.Contains(row.Field<string>(1).Substring(0, 5)))
                    .ToList();

                decimal total_67_d0 = sh_67_00.Sum(row =>
                {
                    string substringValue = row.Field<string>(1).Substring(0, 5);

                    // Eğer substringValue "15020" veya "15025" ise, meblağı %25 azalt
                    decimal discountFactor = (substringValue == "15020" || substringValue == "15025") ? 0.75m : 1.0m;

                    return row.Field<decimal>(3) * discountFactor / 1000;
                });
                worksheet1.Cells["C134"].Value = total_67_d0;

                var total_lr2 = new Dictionary<string, decimal>();
                var types_lr2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr2)
                {
                    var Setir_lr2 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    (row.Field<string>(1).Substring(0, 1) == "3" || row.Field<string>(1).Substring(0, 1) == "4")
                    && row.Field<string>(2).Substring(0, 2) == type
                    && (!likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 5))) && !likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 3))
                )
                .ToList();
                    total_bo4[type] = Setir_lr2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E135"].Value = -total_bo4["00"];
                worksheet1.Cells["F135"].Value = -total_bo4["01"];
                worksheet1.Cells["G135"].Value = -total_bo4["02"];
                worksheet1.Cells["I135"].Value = -total_bo4["03"];
                worksheet1.Cells["K135"].Value = -total_bo4["04"];
                worksheet1.Cells["O135"].Value = -total_bo4["05"];

                var Setir_lr2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    (row.Field<string>(1).Substring(0, 1) == "3" || row.Field<string>(1).Substring(0, 1) == "4")
                    
                    && (!likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 5))) && !likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 3))
                )
                .ToList();
                decimal total_lr2_1 = Setir_lr2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C135"].Value = -total_lr2_1;

                

                var total_lr4 = new Dictionary<string, decimal>();
                var types_lr4 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr4)
                {
                    var Setir_lr4 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    likvid_risk_4.Contains(row.Field<string>(1).Substring(0, 2)) && !likvid_risk_4_istisna.Contains(row.Field<string>(1))
                    && row.Field<string>(2).Substring(0, 2) == type)
                .ToList();
                    total_lr4[type] = Setir_lr4.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E137"].Value = total_lr4["00"];
                worksheet1.Cells["F137"].Value = total_lr4["01"];
                worksheet1.Cells["G137"].Value = total_lr4["02"];
                worksheet1.Cells["I137"].Value = total_lr4["03"];
                worksheet1.Cells["K137"].Value = total_lr4["04"];
                worksheet1.Cells["O137"].Value = total_lr4["05"];

                var Setir_lr2_4 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    likvid_risk_4.Contains(row.Field<string>(1).Substring(0, 2)) && !likvid_risk_4_istisna.Contains(row.Field<string>(1)))
                .ToList();
                decimal total_lr2_4 = Setir_lr2_4.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C137"].Value = total_lr2_4;

                var total_lr5 = new Dictionary<string, decimal>();
                var types_lr5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr5)
                {
                    var Setir_lr5 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 2)) || likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 5)))
                    && row.Field<string>(2).Substring(0, 2) == type)
                .ToList();
                    total_lr5[type] = Setir_lr5.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E138"].Value = -(total_lr5["00"]+ total_bo2["00"]);
                worksheet1.Cells["F138"].Value = -(total_lr5["01"]+ total_bo2["01"]);
                worksheet1.Cells["G138"].Value = -(total_lr5["02"]+ total_bo2["02"]);
                worksheet1.Cells["I138"].Value = -(total_lr5["03"]+ total_bo2["03"]);
                worksheet1.Cells["K138"].Value = -(total_lr5["04"]+ total_bo2["04"]);
                worksheet1.Cells["O138"].Value = -(total_lr5["05"]+ total_bo2["05"]);

                var Setir_lr2_5 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 2)) || likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 5))))
                .ToList();
                decimal total_lr2_5 = Setir_lr2_5.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C138"].Value = -(total_lr2_5+ total_1_bo2);
                decimal t1 = total_lr2_5;
                decimal t2 = total_1_bo2;
                //Setir 136
                if (total_lr2_5 != 0)
                {
                    worksheet1.Cells["C136"].Value = -total_lr2_4 / total_lr2_5;
                }
                else
                {
                    worksheet1.Cells["C136"].Value = 0; // veya başka bir değer
                }

                if (total_lr5["00"] != 0)
                {
                    worksheet1.Cells["E136"].Value = -total_lr4["00"] / total_lr5["00"];
                }
                else
                {
                    worksheet1.Cells["E136"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["01"] != 0)
                {
                    worksheet1.Cells["F136"].Value = -total_lr4["01"] / total_lr5["01"];
                }
                else
                {
                    worksheet1.Cells["F136"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["02"] != 0)
                {
                    worksheet1.Cells["G136"].Value = -total_lr4["02"] / total_lr5["02"];
                }
                else
                {
                    worksheet1.Cells["G136"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["03"] != 0)
                {
                    worksheet1.Cells["I136"].Value = -total_lr4["03"] / total_lr5["03"];
                }
                else
                {
                    worksheet1.Cells["I136"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["04"] != 0)
                {
                    worksheet1.Cells["K136"].Value = -total_lr4["04"] / total_lr5["04"];
                }
                else
                {
                    worksheet1.Cells["K136"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["05"] != 0)
                {
                    worksheet1.Cells["O136"].Value = -total_lr4["05"] / total_lr5["05"];
                }
                else
                {
                    worksheet1.Cells["O136"].Value = 0; // veya başka bir değer
                }

                //IV Hissə – Balans maddələri üzrə dəyişikliklərə dair əlavə məlumatlar*******************

                var total_IV_hisse_1 = new Dictionary<string, decimal>();
                var types_IV_hisse_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_1)
                {
                    var Setir_IV_hisse_1 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    
                    total_IV_hisse_1[type] = Setir_IV_hisse_1.Sum(row => row.Field<decimal>(2)) / 1000;                  
                }
                worksheet1.Cells["E144"].Value = total_IV_hisse_1["00"];
                worksheet1.Cells["F144"].Value = total_IV_hisse_1["01"];
                worksheet1.Cells["G144"].Value = total_IV_hisse_1["02"];
                worksheet1.Cells["I144"].Value = total_IV_hisse_1["03"];
                worksheet1.Cells["K144"].Value = total_IV_hisse_1["04"];
                worksheet1.Cells["O144"].Value = total_IV_hisse_1["05"];

                var Setir_IV_hisse_1_1 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_1 = Setir_IV_hisse_1_1.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["c144"].Value = total_IV_hisse_1_1;

                var total_IV_hisse_3 = new Dictionary<string, decimal>();
                var types_IV_hisse_3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr2)
                {
                    var Setir_IV_hisse_3 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_3[type] = Setir_IV_hisse_3.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E146"].Value = total_IV_hisse_3["00"];
                worksheet1.Cells["F146"].Value = total_IV_hisse_3["01"];
                worksheet1.Cells["G146"].Value = total_IV_hisse_3["02"];
                worksheet1.Cells["I146"].Value = total_IV_hisse_3["03"];
                worksheet1.Cells["K146"].Value = total_IV_hisse_3["04"];
                worksheet1.Cells["O146"].Value = total_IV_hisse_3["05"];

                var Setir_IV_hisse_1_3 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_3 = Setir_IV_hisse_1_3.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["c146"].Value = total_IV_hisse_1_3;

                var total_IV_hisse_5 = new Dictionary<string, decimal>();
                decimal sayH = 0;
                var types_IV_hisse_5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_5)
                {
                    var Setir_IV_hisse_5 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_5[type] = Setir_IV_hisse_5.Sum(row => row.Field<decimal>(2)) / 1000;
                    sayH = Setir_IV_hisse_5.Sum(row => row.Field<decimal>(4));
                }
                worksheet1.Cells["E148"].Value = total_IV_hisse_5["00"];
                worksheet1.Cells["F148"].Value = total_IV_hisse_5["01"];
                worksheet1.Cells["G148"].Value = total_IV_hisse_5["02"];
                worksheet1.Cells["I148"].Value = total_IV_hisse_5["03"];
                worksheet1.Cells["K148"].Value = total_IV_hisse_5["04"];
                worksheet1.Cells["O148"].Value = total_IV_hisse_5["05"];

                worksheet3.Cells["C29"].Value = total_IV_hisse_5;
                worksheet3.Cells["C29"].Value = total_IV_hisse_5;
                worksheet3.Cells["D30"].Value = sayH;
                worksheet3.Cells["D30"].Value = sayH;

                var Setir_IV_hisse_1_5 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_IV_hisse_1_5 = Setir_IV_hisse_1_5.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C148"].Value = total_IV_hisse_1_5;

                var total_IV_hisse_7 = new Dictionary<string, decimal>();
                var types_IV_hisse_7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_7)
                {
                    var Setir_IV_hisse_7 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_7[type] = Setir_IV_hisse_7.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E150"].Value = total_IV_hisse_7["00"];
                worksheet1.Cells["F150"].Value = total_IV_hisse_7["01"];
                worksheet1.Cells["G150"].Value = total_IV_hisse_7["02"];
                worksheet1.Cells["I150"].Value = total_IV_hisse_7["03"];
                worksheet1.Cells["K150"].Value = total_IV_hisse_7["04"];
                worksheet1.Cells["O150"].Value = total_IV_hisse_7["05"];

                var Setir_IV_hisse_1_7 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_IV_hisse_1_7 = Setir_IV_hisse_1_7.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C150"].Value = total_IV_hisse_1_7;

                var total_IV_hisse_9 = new Dictionary<string, decimal>();
                decimal sayS = 0;
                var types_IV_hisse_9 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_9)
                {
                    var Setir_IV_hisse_9 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_9[type] = Setir_IV_hisse_9.Sum(row => row.Field<decimal>(2)) / 1000;
                    sayS = Setir_IV_hisse_9.Sum(row => row.Field<decimal>(4));
                }
                worksheet1.Cells["E152"].Value = total_IV_hisse_9["00"];
                worksheet1.Cells["F152"].Value = total_IV_hisse_9["01"];
                worksheet1.Cells["G152"].Value = total_IV_hisse_9["02"];
                worksheet1.Cells["I152"].Value = total_IV_hisse_9["03"];
                worksheet1.Cells["K152"].Value = total_IV_hisse_9["04"];
                worksheet1.Cells["O152"].Value = total_IV_hisse_9["05"];

                worksheet3.Cells["C34"].Value = total_IV_hisse_9;
                worksheet3.Cells["C34"].Value = total_IV_hisse_9;
                worksheet3.Cells["D35"].Value = sayS;
                worksheet3.Cells["D35"].Value = sayS;

                var Setir_IV_hisse_1_9 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_IV_hisse_1_9 = Setir_IV_hisse_1_9.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C152"].Value = total_IV_hisse_1_9;

                var total_IV_hisse_10 = new Dictionary<string, decimal>();
                var types_IV_hisse_10 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_9)
                {
                    var Setir_IV_hisse_10 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_10[type] = Setir_IV_hisse_10.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E153"].Value = total_IV_hisse_10["00"];
                worksheet1.Cells["F153"].Value = total_IV_hisse_10["01"];
                worksheet1.Cells["G153"].Value = total_IV_hisse_10["02"];
                worksheet1.Cells["I153"].Value = total_IV_hisse_10["03"];
                worksheet1.Cells["K153"].Value = total_IV_hisse_10["04"];
                worksheet1.Cells["O153"].Value = total_IV_hisse_10["05"];

                var Setir_IV_hisse_1_10 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_IV_hisse_1_10 = Setir_IV_hisse_1_10.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C153"].Value = total_IV_hisse_1_10;


                var total_IV_hisse_11a = new Dictionary<string, decimal>();
                var types_IV_hisse_11a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_11a)
                {
                    var Setir_IV_hisse_11a = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type && row.Field<string>(3)== "bos")
                .ToList();
                    total_IV_hisse_11a[type] = Setir_IV_hisse_11a.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E155"].Value = total_IV_hisse_11a["00"];
                worksheet1.Cells["F155"].Value = total_IV_hisse_11a["01"];
                worksheet1.Cells["G155"].Value = total_IV_hisse_11a["02"];
                worksheet1.Cells["I155"].Value = total_IV_hisse_11a["03"];
                worksheet1.Cells["K155"].Value = total_IV_hisse_11a["04"];
                worksheet1.Cells["O155"].Value = total_IV_hisse_11a["05"];

                var Setir_IV_hisse_1_11a = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar" && row.Field<string>(3) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_11a = Setir_IV_hisse_1_11a.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C155"].Value = total_IV_hisse_1_11a;


                var total_IV_hisse_11b = new Dictionary<string, decimal>();
                var types_IV_hisse_11b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_11b)
                {
                    var Setir_IV_hisse_11b = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type && row.Field<string>(3) == "bos")
                .ToList();
                    total_IV_hisse_11b[type] = Setir_IV_hisse_11b.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E156"].Value = total_IV_hisse_11b["00"];
                worksheet1.Cells["F156"].Value = total_IV_hisse_11b["01"];
                worksheet1.Cells["G156"].Value = total_IV_hisse_11b["02"];
                worksheet1.Cells["I156"].Value = total_IV_hisse_11b["03"];
                worksheet1.Cells["K156"].Value = total_IV_hisse_11b["04"];
                worksheet1.Cells["O156"].Value = total_IV_hisse_11b["05"];

                var Setir_IV_hisse_1_11b = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar" && row.Field<string>(3) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_11b = Setir_IV_hisse_1_11b.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C156"].Value = total_IV_hisse_1_11b;


                var total_IV_hisse_12a = new Dictionary<string, decimal>();
                var types_IV_hisse_12a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_12a)
                {
                    var Setir_IV_hisse_12a = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(3) == "sahibkar")
                    && row.Field<string>(4).Substring(0, 2) == type && row.Field<string>(5) == "bos")
                .ToList();
                    total_IV_hisse_12a[type] = Setir_IV_hisse_12a.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E159"].Value = total_IV_hisse_12a["00"];
                worksheet1.Cells["F159"].Value = total_IV_hisse_12a["01"];
                worksheet1.Cells["G159"].Value = total_IV_hisse_12a["02"];
                worksheet1.Cells["I159"].Value = total_IV_hisse_12a["03"];
                worksheet1.Cells["K159"].Value = total_IV_hisse_12a["04"];
                worksheet1.Cells["O159"].Value = total_IV_hisse_12a["05"];

                var Setir_IV_hisse_1_12a = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(3) == "sahibkar" && row.Field<string>(5) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_12a = Setir_IV_hisse_1_12a.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C159"].Value = total_IV_hisse_1_12a;


                var total_IV_hisse_12b = new Dictionary<string, decimal>();
                var types_IV_hisse_12b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_12b)
                {
                    var Setir_IV_hisse_12b = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(3) == "fiziki")
                    && row.Field<string>(4).Substring(0, 2) == type && row.Field<string>(5) == "bos")
                .ToList();
                    total_IV_hisse_12b[type] = Setir_IV_hisse_12b.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E160"].Value = total_IV_hisse_12b["00"];
                worksheet1.Cells["F160"].Value = total_IV_hisse_12b["01"];
                worksheet1.Cells["G160"].Value = total_IV_hisse_12b["02"];
                worksheet1.Cells["I160"].Value = total_IV_hisse_12b["03"];
                worksheet1.Cells["K160"].Value = total_IV_hisse_12b["04"];
                worksheet1.Cells["O160"].Value = total_IV_hisse_12b["05"];

                var Setir_IV_hisse_1_12b = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(3) == "fiziki" && row.Field<string>(5) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_12b = Setir_IV_hisse_1_12b.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C160"].Value = total_IV_hisse_1_12b;

                worksheet1.Cells["C139"].Value = Math.Round(Convert.ToDecimal(lcrcemd) * 100, 2);
                worksheet1.Cells["D139"].Value = Math.Round(Convert.ToDecimal(lcrcem) * 100, 2);
                worksheet1.Cells["E139"].Value = Math.Round(Convert.ToDecimal(lcrazn) * 100, 2); 
                worksheet1.Cells["F139"].Value = Math.Round(Convert.ToDecimal(lcrval) * 100, 2);
                //worksheet1.Cells["D137"].Value = Math.Round(Convert.ToDecimal(lcr4), 2);
                //worksheet1.Cells["D138"].Value = Math.Round(Convert.ToDecimal(lcr5), 2);

                //Daily_Credit_Deposit*******************

                var total_Setir_DM1_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_1_depo = total_Setir_DM1_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM1_depo = total_Setir_DM1_depo.Sum(row => row.Field<decimal>(4));


                worksheet3.Cells["C24"].Value = total_IV_hisse_1_1_depo;
                worksheet3.Cells["C25"].Value = total_IV_hisse_1_1_depo;
                worksheet3.Cells["D24"].Value = say_DM1_depo;
                worksheet3.Cells["D25"].Value = say_DM1_depo;

                var Setir_DM2_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_Setir_DM2_depo = Setir_DM2_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM2_depo = Setir_DM2_depo.Sum(row => row.Field<decimal>(4));

                worksheet3.Cells["C29"].Value = total_Setir_DM2_depo;
                worksheet3.Cells["C30"].Value = total_Setir_DM2_depo;
                worksheet3.Cells["D29"].Value = say_DM2_depo;
                worksheet3.Cells["D30"].Value = say_DM2_depo;

                var Setir_DM3_1_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_Setir_DM3_1_depo = Setir_DM3_1_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM3_1_depo = Setir_DM3_1_depo.Sum(row => row.Field<decimal>(4));

                worksheet3.Cells["C34"].Value = total_Setir_DM3_1_depo;
                worksheet3.Cells["C35"].Value = total_Setir_DM3_1_depo;
                worksheet3.Cells["D34"].Value = say_DM3_1_depo;
                worksheet3.Cells["D35"].Value = say_DM3_1_depo;

                worksheet1.Cells["A3"].Value = "Bank Melli İran Bakı filialı";
                worksheet1.Cells["A2"].Value = "Tarix:" + txtdtbugun.Text;


                filePath = System.IO.Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }
        private void Excel_daily_comment_Yeni_son()
        {

            string connectionString = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
            //duz olan
            string daily_report = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "CASE WHEN SUBSTR(ar.licsch, 0, 3) IN ('159','209','219','239','259') THEN SUBSTR(ar.licsch, 16, 2) " +
               "ELSE SUBSTR(ar.licsch, 6, 2) END AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string daily_report_bk = "select t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish," +
                "t.ostatok_ish*ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6)  kurs " +
                "from odb.arh_saldo_vbls t where t.date_oper =TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99531,99550,99300,99301) and t.ostatok_ish<>0";

            string daily_report_kataloq = "select al.date_oper,al.licschkre,substr(al.licschkre,6,2), tk.code,tk.name,al.summa,al.summa_19," +
                "(al.summa+al.summa_19)*ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) kurs " +
                    "from arh_licschkre al,tipkre tk where al.tipkredita = tk.code and (al.date_close is null or al.date_close>TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')) " +
                    "and al.date_oper=TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy')";

            string daily_report1 = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "CASE WHEN SUBSTR(ar.licsch, 0, 3) IN ('159','209','219','239','259') THEN SUBSTR(ar.licsch, 16, 2) " +
               "ELSE SUBSTR(ar.licsch, 6, 2) END AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string daily_report_bk1 = "select t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish," +
                "t.ostatok_ish*ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6)  kurs " +
                "from odb.arh_saldo_vbls t where t.date_oper =TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99531,99550,99300,99301) and t.ostatok_ish<>0";

            string daily_report_kataloq1 = "select al.date_oper,al.licschkre,substr(al.licschkre,6,2), tk.code,tk.name,al.summa,al.summa_19," +
                "(al.summa+al.summa_19)*ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) ekv," +
                "ROUND(odb.func_get_kurval(substr(al.licschkre,6,2),al.date_oper),6) kurs " +
                    "from arh_licschkre al,tipkre tk where al.tipkredita = tk.code and (al.date_close is null or al.date_close>TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')) " +
                    "and al.date_oper=TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')";

            string gun_erzinde_odenisler = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                "'odenisler'gun_erzinde,round(sum(d.summa_v_nacval), 2) meb, " +
                "case when l.tipkredita = 2 then 'fiziki' " +
                "when l.tipkredita in (1, 3) then 'sahibkar' end tip " +
                ", substr(l.licschkre, 6, 2)val," +
                "case when " +
                "l.date_restructure is null then 'bos' else 'dolu' end rest " +
                "    from arh_licschkre l, arh_dd d,odb.balschkli b " +
                "where d.kredit in (l.licschkre, l.licsch_19) and d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy')   " +
                "and l.date_oper = d.date_oper and substr(d.debet,1,5)= b.balsch and substr(d.debet,10,6)= substr(d.kredit, 10, 6) " +
                "and d.ssk = l.subschkre group by case when d.date_oper = to_date('" + txtdtdunen.Text + "', 'dd/mm/yyyy') then 'dunen' else 'bugun' end,l.tipkredita,substr(l.licschkre, 6, 2),l.date_restructure";
            string qaliqlar_30_90 = "SELECT case when m.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                    " case when " +
                    " m.tipkredita = 1 then 'huquqi' " +
                    " when m.tipkredita = 2 then 'fiziki' " +
                    " else 'sahibkar' end tip, " +
                    " odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, " +
                    " ((m.summa * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)) + (m.summa_19 * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)))  qal , " +
                    " case when m.summa_19 > 0 then 'vk' end gecikme,case when m.date_restructure is not null then 'rest' end restur,substr(m.licschkre, 6, 2) val " +
                    " from view_nacpogprokre_all x, arh_licschkre m where " +
                    " x.date_oper = m.date_oper " +
                    " and x.licschpkre = m.licschpkre and x.subschkre = m.subschkre " +
                    " and m.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and m.date_close is null " +
                    " and x.licschpkre = m.licschpkre and x.subschkre = m.subschkre " +
                    " and m.date_close is null";
            string qaliqlar_kr_tip = " select case when m.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                    " case when " +
                     " (m.tipkredita = 1 or m.tipkredita = 3) and m.index_otrasli != '01902'  then 'biznes' " +
                     " when m.tipkredita = 2 and m.index_otrasli != '01902' then 'fiziki' " +
                     " when m.index_otrasli = '01902' then 'dasinmaz' " +
                     " end tip, sum((m.summa * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)) + (m.summa_19 * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)))  qal,  " +
                     " substr(m.licschkre, 6, 2) val " +
                     " from arh_licschkre m " +
                     " where m.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and m.date_close is null " +
                     " group by case when m.date_oper = to_date('" + txtdtdunen.Text + "', 'dd/mm/yyyy') then 'dunen' else 'bugun' end,  " +
                     " m.tipkredita,substr(m.licschkre, 6, 2),m.index_otrasli";

            string medaxiller = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar," +
                                "CASE    " +
                    "WHEN r.yurik = 1 then 'huquqi' " +
                    "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                    "WHEN r.fizik = 1  THEN 'fiziki' " +
                    "ELSE 'unknown' END AS tip,round(sum(d.summa_v_nacval) , 2) meb,substr(d.kredit,6,2) val,count(d.kredit) " +
                    "      from arh_dd d, regnom r,odb.balschkli b,licsch l " +
                  " where d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') " +
                  "and substr(d.kredit,1,5)= b.balsch and" +
                  " d.kredit=l.licsch and l.registrac_nomer=r.regnom " +
                  "and substr(d.debet,1,5) not in ('66220', '86220') " +
                  " group by case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end," +
                  "CASE " +
                  "WHEN r.yurik = 1 then 'huquqi' " +
                    "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                    "WHEN r.fizik = 1  THEN 'fiziki' " +
                  "  ELSE 'unknown' END,substr(d.kredit,6,2)";
            string mexaricler = "select case when d.date_oper=to_date('" + txtdtdunen.Text + "','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                   "CASE " +
                   "WHEN r.yurik = 1 then 'huquqi' " +
                   "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                   "WHEN r.fizik = 1  THEN 'fiziki' " +
                   "ELSE 'unknown' END AS tip,round(sum(d.summa_v_nacval), 2) meb,substr(d.kredit, 6, 2) val " +
                   "from arh_dd d, regnom r,odb.balschkli b, licsch l " +
                   "where d.date_oper between to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') and to_date('" + txtdtbugun.Text + "','dd-mm-yyyy') and substr(d.debet,1,5)= b.balsch " +
                   "and d.debet = l.licsch and l.registrac_nomer = r.regnom " +
                   "and substr(d.kredit,1,5) not in ('66220', '86220') " +
                   "group by case when d.date_oper = to_date('" + txtdtdunen.Text + "', 'dd-mm-yyyy') then 'dunen' else 'bugun' end ,  " +
                   "CASE " +
                   "WHEN r.yurik = 1 then 'huquqi' " +
                   "WHEN r.predprinimatel = 1  THEN 'sahibkar' " +
                   "WHEN r.fizik = 1  THEN 'fiziki' " +
                   "ELSE 'unknown' END,substr(d.kredit, 6, 2)";
            string ver_kr_lar = "select case when m.date_oper=to_date('"+txtdtdunen.Text+"','dd/mm/yyyy') then 'dunen' else 'bugun' end tar, " +
                        "  case when  " +
                        " (m.tipkredita = 1 or m.tipkredita = 3) and m.index_otrasli!='01902'  then 'biznes'  " +
                        "  when m.tipkredita = 2 and m.index_otrasli!='01902' then 'fiziki' " +
                        "  when  m.index_otrasli='01902' then 'dasinmaz' " +
                        "  end tip, sum((m.summa * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)) + (m.summa_19 * ROUND(odb.func_get_kurval(substr(m.licschkre, 6, 2), m.date_oper), 6)))  qal,  " +
                        "  substr(m.licschkre, 6, 2) val  " +
                        "  from arh_licschkre m  " +
                        "  where m.date_oper between to_date('"+txtdtdunen.Text+"', 'dd-mm-yyyy') and to_date('"+txtdtbugun.Text+"','dd-mm-yyyy') and m.date_close is null  " +
                        "  group by case when m.date_oper = to_date('"+txtdtdunen.Text+"', 'dd/mm/yyyy') then 'dunen' else 'bugun' end,  " +
                        "  m.tipkredita,substr(m.licschkre, 6, 2),m.index_otrasli";

            DataTable _dt_daily_report = new DataTable();
            DataTable _dt_daily_report_bk = new DataTable();
            DataTable _dt_daily_report_kataloq = new DataTable();

            DataTable _dt_daily_report1 = new DataTable();
            DataTable _dt_daily_report_bk1 = new DataTable();
            DataTable _dt_daily_report_kataloq1 = new DataTable();

            DataTable _dt_odenisler = new DataTable();
            DataTable _dt_qali_gunler = new DataTable();
            DataTable _dt_qaliq_tip = new DataTable();
            DataTable _dt_medaxil = new DataTable();
            DataTable _dt_mexaric = new DataTable();
            DataTable _dt_verilmis = new DataTable();

            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand command = new OracleCommand(daily_report, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report);
                }

                using (OracleCommand command = new OracleCommand(daily_report_bk, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_bk);
                }

                using (OracleCommand command = new OracleCommand(daily_report_kataloq, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_kataloq);
                }

                using (OracleCommand command = new OracleCommand(daily_report1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report1);
                }

                using (OracleCommand command = new OracleCommand(daily_report_bk1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_bk1);
                }

                using (OracleCommand command = new OracleCommand(daily_report_kataloq1, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_daily_report_kataloq1);
                }
                using (OracleCommand command = new OracleCommand(gun_erzinde_odenisler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_odenisler);
                }

                using (OracleCommand command = new OracleCommand(qaliqlar_30_90, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qali_gunler);
                }

                using (OracleCommand command = new OracleCommand(qaliqlar_kr_tip, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_qaliq_tip);
                }
                using (OracleCommand command = new OracleCommand(medaxiller, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_medaxil);
                }

                using (OracleCommand command = new OracleCommand(mexaricler, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_mexaric);
                }

                using (OracleCommand command = new OracleCommand(ver_kr_lar, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(_dt_verilmis);
                }
                connection.Close();
            }
            string dosyayolu = @"C:\BMI_\huqui_sorgu";
            string textBoxText = txtdtbugun.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", ""); ;
            string baseFileName = "CUR.v02.1124d" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = @"C:\BMI_\Daily_report_comments_Yeni_.xlsm";
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);

            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }
            //"15020",
            FileInfo templateFile = new FileInfo(templateFilePath);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Daily-Report"];
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets["comments"];
                ExcelWorksheet worksheet3 = package.Workbook.Worksheets["Daily_Credit_Deposit"];
                //******************Setir A1
                //******************Setir B11a
                var total_a1 = new Dictionary<string, decimal>();
                var types_a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a1)
                {
                    var Setir_a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A1.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a1[type] = Setir_a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E9"].Value = total_a1["00"];
                worksheet1.Cells["F9"].Value = total_a1["01"];
                worksheet1.Cells["G9"].Value = total_a1["02"];
                worksheet1.Cells["I9"].Value = total_a1["03"];
                worksheet1.Cells["K9"].Value = total_a1["04"];
                worksheet1.Cells["O9"].Value = total_a1["05"];

                var Setir_a1_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A1.Contains(row.Field<string>(1).Substring(0, 3)))
                                        .ToList();
                decimal total_a1_1_c = Setir_a1_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c9"].Value = total_a1_1_c;

                //******************Setir A2
                var total_a2 = new Dictionary<string, decimal>();
                var types_a2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2)
                {
                    var Setir_a2 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2[type] = Setir_a2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E10"].Value = total_a2["00"];
                worksheet1.Cells["F10"].Value = total_a2["01"];
                worksheet1.Cells["G10"].Value = total_a2["02"];
                worksheet1.Cells["I10"].Value = total_a2["03"];
                worksheet1.Cells["K10"].Value = total_a2["04"];
                worksheet1.Cells["O10"].Value = total_a2["05"];

                var Setir_a2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2.Contains(row.Field<string>(1).Substring(0, 3)))
                                        .ToList();
                decimal total_a2_1_c = Setir_a2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c10"].Value = total_a2_1_c;
                //******************Setir A2a

                var total_a2a = new Dictionary<string, decimal>();
                var types_a2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2a)
                {
                    var Setir_a2a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2a.Contains(row.Field<string>(1)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2a[type] = Setir_a2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E11"].Value = total_a2a["00"];
                worksheet1.Cells["F11"].Value = total_a2a["01"];
                worksheet1.Cells["G11"].Value = total_a2a["02"];
                worksheet1.Cells["I11"].Value = total_a2a["03"];
                worksheet1.Cells["K11"].Value = total_a2a["04"];
                worksheet1.Cells["O11"].Value = total_a2a["05"];
                var Setir_a2a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2a.Contains(row.Field<string>(1)))
                                        .ToList();
                decimal total_a2_1 = Setir_a2a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c11"].Value = total_a2_1;

                //******************Setir A2b
                var total_a2b = new Dictionary<string, decimal>();
                var types_a2b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a2b)
                {
                    var Setir_a2b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A2b.Contains(row.Field<string>(1)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a2b[type] = Setir_a2b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E12"].Value = total_a2b["00"];
                worksheet1.Cells["F12"].Value = total_a2b["01"];
                worksheet1.Cells["G12"].Value = total_a2b["02"];
                worksheet1.Cells["I12"].Value = total_a2b["03"];
                worksheet1.Cells["K12"].Value = total_a2b["04"];
                worksheet1.Cells["O12"].Value = total_a2b["05"];
                var Setir_a2b_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A2b.Contains(row.Field<string>(1)))
                                        .ToList();
                decimal total_a2b_1_c = Setir_a2b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c12"].Value = total_a2b_1_c;
                //******************Setir A3a
                var total_a3a = new Dictionary<string, decimal>();
                var types_a3a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a)
                {
                    var Setir_a3a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a.Contains(row.Field<string>(1).Substring(0, 3)) && A3a_qeydno.Contains(row.Field<string>(1).Substring(9, 6))
                                        && row.Field<string>(2) == type &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a3a[type] = Setir_a3a.Sum(row => row.Field<decimal>(3)) / 1000;

                    string str = "15910000000001200100";

                    // Son iki rakamı almak için:
                    string lastTwoDigits = str.Substring(14, 2);
                }
                worksheet1.Cells["E14"].Value = -total_a3a["00"];
                worksheet1.Cells["F14"].Value = -total_a3a["01"];
                worksheet1.Cells["G14"].Value = -total_a3a["02"];
                worksheet1.Cells["I14"].Value = -total_a3a["03"];
                worksheet1.Cells["K14"].Value = -total_a3a["04"];
                worksheet1.Cells["O14"].Value = -total_a3a["05"];
                var Setir_a31_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a.Contains(row.Field<string>(1).Substring(0, 3)) &&
                                        A3a_qeydno.Contains(row.Field<string>(1).Substring(9, 6)) &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_3a_1_c = Setir_a31_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c14"].Value = -total_3a_1_c;

                //******************Setir A3a1
                var total_a3a1 = new Dictionary<string, decimal>();
                var types_a3a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a1)
                {
                    var Setir_a3a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a3a1[type] = Setir_a3a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E15"].Value = total_a3a1["00"];
                worksheet1.Cells["F15"].Value = total_a3a1["01"];
                worksheet1.Cells["G15"].Value = total_a3a1["02"];
                worksheet1.Cells["I15"].Value = total_a3a1["03"];
                worksheet1.Cells["K15"].Value = total_a3a1["04"];
                worksheet1.Cells["O15"].Value = total_a3a1["05"];
                var Setir_a3a1_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_1.Contains(row.Field<string>(1).Substring(0, 5)))
                                        .ToList();
                decimal total_3a1_1_c = Setir_a3a1_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c15"].Value = total_3a1_1_c;

                //******************Setir A3a1a
                var total_a3a1a = new Dictionary<string, decimal>();
                var types_a3a1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a1a)
                {//&& row.Field<string>(2) == type
                    var Setir_a3a1a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                        && A3a_1a_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
               && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a3a1a[type] = Setir_a3a1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E16"].Value = -total_a3a1a["00"];
                worksheet1.Cells["F16"].Value = -total_a3a1a["01"];
                worksheet1.Cells["G16"].Value = -total_a3a1a["02"];
                worksheet1.Cells["I16"].Value = -total_a3a1a["03"];
                worksheet1.Cells["K16"].Value = -total_a3a1a["04"];
                worksheet1.Cells["O16"].Value = -total_a3a1a["05"];
                var Setir_a3a1a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_1a.Contains(row.Field<string>(1).Substring(0, 5))
                                        && A3a_1a_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
               && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_3a1a_1_c = Setir_a3a1a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C16"].Value = -total_3a1a_1_c;
                //******************Setir A3a2
                var total_a3a2 = new Dictionary<string, decimal>();
                var types_a3a2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a2)
                {
                    var Setir_a3a2 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_2.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                    )
                                        .ToList();
                    total_a3a2[type] = Setir_a3a2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E18"].Value = total_a3a2["00"];
                worksheet1.Cells["F18"].Value = total_a3a2["01"];
                worksheet1.Cells["G18"].Value = total_a3a2["02"];
                worksheet1.Cells["I18"].Value = total_a3a2["03"];
                worksheet1.Cells["K18"].Value = total_a3a2["04"];
                worksheet1.Cells["O18"].Value = total_a3a2["05"];
                var Setir_a3a2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A3a_2.Contains(row.Field<string>(1).Substring(0, 5))
                                    )
                                        .ToList();
                decimal total_3a2_1_c = Setir_a3a2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c18"].Value = total_3a2_1_c;

                //******************Setir A3a2a
                var total_a3a2a = new Dictionary<string, decimal>();
                var types_a3a2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a3a2a)
                {
                    var Setir_a3a2a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A3a_2a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type
                                        && A3a_2a_yanasma.Contains(row.Field<string>(1).Substring(13, 2)))
                                        .ToList();
                    total_a3a2a[type] = Setir_a3a2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E19"].Value = -total_a3a2a["00"];
                worksheet1.Cells["F19"].Value = -total_a3a2a["01"];
                worksheet1.Cells["G19"].Value = -total_a3a2a["02"];
                worksheet1.Cells["I19"].Value = -total_a3a2a["03"];
                worksheet1.Cells["K19"].Value = -total_a3a2a["04"];
                worksheet1.Cells["O19"].Value = -total_a3a2a["05"];
                var Setir_a3a2a_1 = _dt_daily_report1.AsEnumerable()
                                         .Where(row => A3a_2a.Contains(row.Field<string>(1).Substring(0, 5))
                                         && A3a_2a_yanasma.Contains(row.Field<string>(1).Substring(13, 2)))
                                         .ToList();
                decimal total_3a2a_1_c = Setir_a3a2a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c19"].Value = -total_3a2a_1_c;


                //******************Setir A4_1a
                var total_a4_1a = new Dictionary<string, decimal>();
                var types_a4_1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1a)
                {
                    var Setir_a4_1a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1a[type] = Setir_a4_1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E24"].Value = total_a4_1a["00"];
                worksheet1.Cells["F24"].Value = total_a4_1a["01"];
                worksheet1.Cells["G24"].Value = total_a4_1a["02"];
                worksheet1.Cells["I24"].Value = total_a4_1a["03"];
                worksheet1.Cells["K24"].Value = total_a4_1a["04"];
                worksheet1.Cells["O24"].Value = total_a4_1a["05"];
                var Setir_a4_1a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => A4_1a.Contains(row.Field<string>(1).Substring(0, 5)))
                                        .ToList();
                decimal total_a4_1a_1 = Setir_a4_1a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c24"].Value = total_a4_1a_1;
                //******************Setir A4_1a1
                var total_a4_1a1 = new Dictionary<string, decimal>();
                var types_a4_1a1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1a1)
                {
                    var Setir_a4_1a1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1a1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1a1[type] = Setir_a4_1a1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E25"].Value = total_a4_1a1["00"];
                worksheet1.Cells["F25"].Value = total_a4_1a1["01"];
                worksheet1.Cells["G25"].Value = total_a4_1a1["02"];
                worksheet1.Cells["I25"].Value = total_a4_1a1["03"];
                worksheet1.Cells["K25"].Value = total_a4_1a1["04"];
                worksheet1.Cells["O25"].Value = total_a4_1a1["05"];

                var Setir_a4_1a1_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1a1.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1a1_c = Setir_a4_1a1_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C25"].Value = total_a4_1a1_c;

                //******************Setir A4_1b
                var total_a4_1b = new Dictionary<string, decimal>();
                var types_a4_1b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1b)
                {
                    var Setir_a4_1b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1b[type] = Setir_a4_1b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E26"].Value = total_a4_1b["00"];
                worksheet1.Cells["F26"].Value = total_a4_1b["01"];
                worksheet1.Cells["G26"].Value = total_a4_1b["02"];
                worksheet1.Cells["I26"].Value = total_a4_1b["03"];
                worksheet1.Cells["K26"].Value = total_a4_1b["04"];
                worksheet1.Cells["O26"].Value = total_a4_1b["05"];

                var Setir_a4_1b_d = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1b_d = Setir_a4_1b_d.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c26"].Value = total_a4_1b_d;

                //******************Setir A4_1b
                var total_a4_1b_1 = new Dictionary<string, decimal>();
                var types_a4_1b_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a4_1b_1)
                {
                    var Setir_a4_1b_1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A4_1b1.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a4_1b_1[type] = Setir_a4_1b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E27"].Value = total_a4_1b_1["00"];
                worksheet1.Cells["F27"].Value = total_a4_1b_1["01"];
                worksheet1.Cells["G27"].Value = total_a4_1b_1["02"];
                worksheet1.Cells["I27"].Value = total_a4_1b_1["03"];
                worksheet1.Cells["K27"].Value = total_a4_1b_1["04"];
                worksheet1.Cells["O27"].Value = total_a4_1b_1["05"];

                var Setir_a4_1b_1_d = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A4_1b1.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a4_1b_1_d = Setir_a4_1b_1_d.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c27"].Value = total_a4_1b_d;

                //******************Setir A5
                var total_a5 = new Dictionary<string, decimal>();
                var types_a5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a5)
                {
                    var Setir_a5 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A5.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a5[type] = Setir_a5.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E41"].Value = total_a5["00"];
                worksheet1.Cells["F41"].Value = total_a5["01"];
                worksheet1.Cells["G41"].Value = total_a5["02"];
                worksheet1.Cells["I41"].Value = total_a5["03"];
                worksheet1.Cells["K41"].Value = total_a5["04"];
                worksheet1.Cells["O41"].Value = total_a5["05"];

                var Setir_a5_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A5.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a5_c = Setir_a5_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C41"].Value = total_a5_c;

                //******************Setir A6
                var total_a6 = new Dictionary<string, decimal>();
                var types_a6 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a6)
                {
                    var Setir_a6 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A6.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a6[type] = Setir_a6.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E42"].Value = total_a6["00"];
                worksheet1.Cells["F42"].Value = total_a6["01"];
                worksheet1.Cells["G42"].Value = total_a6["02"];
                worksheet1.Cells["I42"].Value = total_a6["03"];
                worksheet1.Cells["K42"].Value = total_a6["04"];
                worksheet1.Cells["O42"].Value = total_a6["05"];

                var Setir_a6_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A6.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a6_1 = Setir_a6_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c42"].Value = total_a6_1;

                //******************Setir A6a
                var total_a6a = new Dictionary<string, decimal>();
                var types_a6a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a6a)
                {
                    var Setir_a6a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A6a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a6a[type] = Setir_a6a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E43"].Value = total_a6a["00"];
                worksheet1.Cells["F43"].Value = total_a6a["01"];
                worksheet1.Cells["G43"].Value = total_a6a["02"];
                worksheet1.Cells["I43"].Value = total_a6a["03"];
                worksheet1.Cells["K43"].Value = total_a6a["04"];
                worksheet1.Cells["O43"].Value = total_a6a["05"];

                var Setir_a6a_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A6a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a6a_1 = Setir_a6a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c43"].Value = total_a6a_1;

                //******************Setir A6c
                var total_a6c = new Dictionary<string, decimal>();
                var types_a6c = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a6c)
                {
                    var Setir_a6c = _dt_daily_report.AsEnumerable()
                                        .Where(row => A6c.Contains(row.Field<string>(1)) && row.Field<string>(2) == type
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a6c[type] = Setir_a6c.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E45"].Value = -total_a6c["00"];
                worksheet1.Cells["F45"].Value = -total_a6c["01"];
                worksheet1.Cells["G45"].Value = -total_a6c["02"];
                worksheet1.Cells["I45"].Value = -total_a6c["03"];
                worksheet1.Cells["K45"].Value = -total_a6c["04"];
                worksheet1.Cells["O45"].Value = -total_a6c["05"];

                var Setir_a6c_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A6c.Contains(row.Field<string>(1))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a6c_c = Setir_a6c_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c45"].Value = -total_a6c_c;

                //******************Setir A7

                var total_a7 = new Dictionary<string, decimal>();
                var types_a7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a7)
                {
                    var Setir_a7 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A7.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a7[type] = Setir_a7.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E46"].Value = total_a7["00"];
                worksheet1.Cells["F46"].Value = total_a7["01"];
                worksheet1.Cells["G46"].Value = total_a7["02"];
                worksheet1.Cells["I46"].Value = total_a7["03"];
                worksheet1.Cells["K46"].Value = total_a7["04"];
                worksheet1.Cells["O46"].Value = total_a7["05"];

                var Setir_a7_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A7.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a7_1 = Setir_a7_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C46"].Value = total_a7_1;
                //******************Setir A8.1
                var total_a8_1 = new Dictionary<string, decimal>();
                var types_a8_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_1)
                {
                    var Setir_a8_1 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "biznes"
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_1[type] = Setir_a8_1.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E49"].Value = total_a8_1["00"];
                worksheet1.Cells["F49"].Value = total_a8_1["01"];
                worksheet1.Cells["G49"].Value = total_a8_1["02"];
                worksheet1.Cells["I49"].Value = total_a8_1["03"];
                worksheet1.Cells["K49"].Value = total_a8_1["04"];
                worksheet1.Cells["O49"].Value = total_a8_1["05"];

                var Setir_a8_1_1 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "biznes"))
                                    .ToList();
                decimal total_a8_1_1 = Setir_a8_1_1.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C49"].Value = total_a8_1_1;

                //******************Setir A8.2
                var total_a8_2 = new Dictionary<string, decimal>();
                var types_a8_2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_2)
                {
                    var Setir_a8_2 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki"
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_2[type] = Setir_a8_2.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E50"].Value = total_a8_2["00"];
                worksheet1.Cells["F50"].Value = total_a8_2["01"];
                worksheet1.Cells["G50"].Value = total_a8_2["02"];
                worksheet1.Cells["I50"].Value = total_a8_2["03"];
                worksheet1.Cells["K50"].Value = total_a8_2["04"];
                worksheet1.Cells["O50"].Value = total_a8_2["05"];

                var Setir_a8_1_2 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki"))
                                    .ToList();
                decimal total_a8_1_2 = Setir_a8_1_2.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C50"].Value = total_a8_1_2;

                //******************Setir A8.3
                var total_a8_3 = new Dictionary<string, decimal>();
                var types_a8_3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8_3)
                {
                    var Setir_a8_3 = _dt_qaliq_tip.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "dasinmaz"
                                         && row.Field<string>(3) == type))
                                        .ToList();
                    total_a8_3[type] = Setir_a8_3.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E51"].Value = total_a8_3["00"];
                worksheet1.Cells["F51"].Value = total_a8_3["01"];
                worksheet1.Cells["G51"].Value = total_a8_3["02"];
                worksheet1.Cells["I51"].Value = total_a8_3["03"];
                worksheet1.Cells["K51"].Value = total_a8_3["04"];
                worksheet1.Cells["O51"].Value = total_a8_3["05"];

                var Setir_a8_1_3 = _dt_qaliq_tip.AsEnumerable()
                                    .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "dasinmaz"))
                                    .ToList();
                decimal total_a8_1_3 = Setir_a8_1_3.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C51"].Value = total_a8_1_3;

                //******************Setir A8b
                var total_a8b = new Dictionary<string, decimal>();
                var types_a8b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a8b)
                {
                    var Setir_a8b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A8b.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == type
                                        && A8b_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();


                    total_a8b[type] = Setir_a8b.Sum(row => row.Field<decimal>(3)) / 1000;


                }
                worksheet1.Cells["E52"].Value = -total_a8b["00"];
                worksheet1.Cells["F52"].Value = -total_a8b["01"];
                worksheet1.Cells["G52"].Value = -total_a8b["02"];
                worksheet1.Cells["I52"].Value = -total_a8b["03"];
                worksheet1.Cells["K52"].Value = -total_a8b["04"];
                worksheet1.Cells["O52"].Value = -total_a8b["05"];

                var Setir_a8b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A8b.Contains(row.Field<string>(1).Substring(0, 3))
                                     && A8b_yanasma.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a8b_c = Setir_a8b_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C52"].Value = -total_a8b_c;

                //******************Setir A9

                var total_a9 = new Dictionary<string, decimal>();
                var types_a9 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a9)
                {
                    var Setir_a9 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A9.Contains(row.Field<string>(1).Substring(0, 2)) && !A9_istisna.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a9[type] = Setir_a9.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E53"].Value = total_a9["00"];
                worksheet1.Cells["F53"].Value = total_a9["01"];
                worksheet1.Cells["G53"].Value = total_a9["02"];
                worksheet1.Cells["I53"].Value = total_a9["03"];
                worksheet1.Cells["K53"].Value = total_a9["04"];
                worksheet1.Cells["O53"].Value = total_a9["05"];

                var Setir_a9_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A9.Contains(row.Field<string>(1).Substring(0, 2)) && !A9_istisna.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a9_c = Setir_a9_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C53"].Value = total_a9_c;

                //******************Setir A10
                var total_a10 = new Dictionary<string, decimal>();
                var types_a10 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a10)
                {
                    var Setir_a10 = _dt_daily_report.AsEnumerable()
                                        .Where(row => A10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a10[type] = Setir_a10.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E54"].Value = total_a10["00"];
                worksheet1.Cells["F54"].Value = total_a10["01"];
                worksheet1.Cells["G54"].Value = total_a10["02"];
                worksheet1.Cells["I54"].Value = total_a10["03"];
                worksheet1.Cells["K54"].Value = total_a10["04"];
                worksheet1.Cells["O54"].Value = total_a10["05"];

                var Setir_a10_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10_c = Setir_a10_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C54"].Value = total_a10_c;


                //******************Setir A10a
                var total_a10a = new Dictionary<string, decimal>();
                var types_a10a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a10a)
                {
                    var Setir_a10a = _dt_daily_report.AsEnumerable()
                                        .Where(row => A10a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_a10a[type] = Setir_a10a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E55"].Value = total_a10a["00"];
                worksheet1.Cells["F55"].Value = total_a10a["01"];
                worksheet1.Cells["G55"].Value = total_a10a["02"];
                worksheet1.Cells["I55"].Value = total_a10a["03"];
                worksheet1.Cells["K55"].Value = total_a10a["04"];
                worksheet1.Cells["O55"].Value = total_a10a["05"];

                var Setir_a10a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10a_c = Setir_a10a_c.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C55"].Value = total_a10a_c;


                //******************Setir A10b

                var Setir_a10b_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                //var Setir_a10b_c_27013 = _dt_daily_report1.AsEnumerable()
                //                    .Where(row => A10b_27013.Contains(row.Field<string>(1).Substring(0, 5)))
                //                    .ToList();
                decimal total_a10b_c = Setir_a10b_1.Sum(row => row.Field<decimal>(3)) / 1000;

                //var Setir_a10b_e_27012 = _dt_daily_report.AsEnumerable()
                //                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                //                    .ToList();
                var Setir_a10b = _dt_daily_report.AsEnumerable()
                                    .Where(row => A10b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a10b_e = Setir_a10b.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C56"].Value = total_a10b_c;
                worksheet1.Cells["E56"].Value = total_a10b_e;

                //******************Setir A12

                var Setir_a12_c_qisa = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();


                var Setir_a12_c_uzun = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2))
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal top1u = Setir_a12_c_uzun.Sum(row => row.Field<decimal>(3)) / 1000;
                decimal top1q = Setir_a12_c_qisa.Sum(row => row.Field<decimal>(3)) / 1000;
                decimal total_a12_c = (Setir_a12_c_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_c_uzun.Sum(row => row.Field<decimal>(3))) / 1000; ;




                var Setir_a12_e_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();

                var Setir_a12_e_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "00"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_e = (Setir_a12_e_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_e_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_f_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                var Setir_a12_f_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "01"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_f = (Setir_a12_f_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_f_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_g_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                var Setir_a12_g_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "02"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_g = (Setir_a12_g_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_g_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_i_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                var Setir_a12_i_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "03"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_i = (Setir_a12_i_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_i_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_k_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                var Setir_a12_k_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "04"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_k = (Setir_a12_k_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_k_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                var Setir_a12_o_qisa = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_qisa.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                var Setir_a12_o_uzun = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12_eht.Contains(row.Field<string>(1).Substring(0, 3)) && A12_eht_faizler.Contains(row.Field<string>(1).Substring(13, 2)) && row.Field<string>(2) == "05"
                                    && (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                    .ToList();
                decimal total_a12_o = (Setir_a12_o_qisa.Sum(row => row.Field<decimal>(3)) + Setir_a12_o_uzun.Sum(row => row.Field<decimal>(3))) / 1000;

                worksheet1.Cells["C58"].Value = total_a12_c;
                worksheet1.Cells["E58"].Value = total_a12_e;
                worksheet1.Cells["F58"].Value = total_a12_f;
                worksheet1.Cells["G58"].Value = total_a12_g;
                worksheet1.Cells["I58"].Value = total_a12_i;
                worksheet1.Cells["K58"].Value = total_a12_k;
                worksheet1.Cells["O58"].Value = total_a12_o;

                //******************Setir A12a

                var Setir_a12a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_a12a_c = Setir_a12a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_a12a_e = Setir_a12a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_a12a_f = Setir_a12a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_a12a_g = Setir_a12a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_a12a_i = Setir_a12a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_a12a_k = Setir_a12a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_a12a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => A12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_a12a_o = Setir_a12a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C59"].Value = total_a12a_c;
                worksheet1.Cells["E59"].Value = total_a12a_e;
                worksheet1.Cells["F59"].Value = total_a12a_f;
                worksheet1.Cells["G59"].Value = total_a12a_g;
                worksheet1.Cells["I59"].Value = total_a12a_i;
                worksheet1.Cells["K59"].Value = total_a12a_k;
                worksheet1.Cells["O59"].Value = total_a12a_o;

                //******************Setir A12b
                //A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2)))

                var total_a12b = new Dictionary<string, decimal>();
                var types_a12b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a12b)
                {
                    var Setir_a12b = _dt_daily_report.AsEnumerable()
                                        .Where(row => A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2))
                                        && row.Field<string>(2) == type)
                                        .ToList();
                    total_a12b[type] = Setir_a12b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E60"].Value = -total_a12b["00"];
                worksheet1.Cells["F60"].Value = -total_a12b["01"];
                worksheet1.Cells["G60"].Value = -total_a12b["02"];
                worksheet1.Cells["I60"].Value = -total_a12b["03"];
                worksheet1.Cells["K60"].Value = -total_a12b["04"];
                worksheet1.Cells["O60"].Value = -total_a12b["05"];

                var Setir_a12b_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => A12b.Contains(row.Field<string>(1).Substring(0, 3)) && !A12b_istisna.Contains(row.Field<string>(1).Substring(13, 2)))
                                    .ToList();
                decimal total_a12b_1 = Setir_a12b_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C60"].Value = -total_a12b_1;


                DataTable filteredDataTable = _dt_daily_report1.Clone();
                foreach (var row in Setir_a12_c_qisa)
                {
                    filteredDataTable.Rows.Add(row.ItemArray);
                }

                // DataGridView'e yeni DataTable'ı atayarak güncelle
                dataGridView1.DataSource = filteredDataTable;

                //******************Setir 12b
                var total_a12c = new Dictionary<string, decimal>();
                var types_a12c = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_a12c)
                {
                    var Setir_a12c = _dt_daily_report.AsEnumerable()
                                        .Where(row => ((A12c.Contains(row.Field<string>(1).Substring(0, 3)) && !A12c_yanasma.Contains(row.Field<string>(1).Substring(13, 2))) || A12c_uzun.Contains(row.Field<string>(1)))
                                        && row.Field<string>(2) == type
                                          &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                    total_a12c[type] = Setir_a12c.Sum(row => row.Field<decimal>(3)) / 1000;


                }
                worksheet1.Cells["E61"].Value = -total_a12c["00"] + total_a12b["00"]; //- total_a3a["00"] - total_a6c["00"] - total_a8b["00"] - total_a12b["00"];
                worksheet1.Cells["F61"].Value = -total_a12c["01"] + total_a12b["01"];// - total_a3a["01"] - total_a6c["01"] - total_a8b["01"] - total_a12b["01"];
                worksheet1.Cells["G61"].Value = -total_a12c["02"] + total_a12b["02"];// - total_a3a["02"] - total_a6c["02"] - total_a8b["02"] - total_a12b["02"];
                worksheet1.Cells["I61"].Value = -total_a12c["03"] + total_a12b["03"]; //- total_a3a["03"] - total_a6c["03"] - total_a8b["03"] - total_a12b["03"];
                worksheet1.Cells["K61"].Value = -total_a12c["04"] + total_a12b["04"]; //- total_a3a["04"] - total_a6c["04"] - total_a8b["04"] - total_a12b["04"];
                worksheet1.Cells["O61"].Value = -total_a12c["05"] + total_a12b["05"];// - total_a3a["05"] - total_a6c["05"] - total_a8b["05"] - total_a12b["05"];
                var Setir_a12c_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => ((A12c.Contains(row.Field<string>(1).Substring(0, 3)) && !A12c_yanasma.Contains(row.Field<string>(1).Substring(13, 2))) || A12c_uzun.Contains(row.Field<string>(1)))
                                        &&
                                        (Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) >= 25
                                        || Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) == 100))
                                        .ToList();
                decimal total_12c_1_c = Setir_a12c_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["c61"].Value = -total_12c_1_c + total_a12b_1;// - total_a12b_1; //- total_3a_1_c - total_a6c_c - total_a8b_c - total_a12b_1;


                //******************Setir B1a

                var Setir_b1a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1a_c = Setir_b1a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1a_e = Setir_b1a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1a_f = Setir_b1a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1a_g = Setir_b1a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1a_i = Setir_b1a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1a_k = Setir_b1a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1a_o = Setir_b1a_o.Sum(row => row.Field<decimal>(3)) / 1000;



                worksheet1.Cells["C65"].Value = -total_b1a_c;
                worksheet1.Cells["E65"].Value = -total_b1a_e;
                worksheet1.Cells["F65"].Value = -total_b1a_f;
                worksheet1.Cells["G65"].Value = -total_b1a_g;
                worksheet1.Cells["I65"].Value = -total_b1a_i;
                worksheet1.Cells["K65"].Value = -total_b1a_k;
                worksheet1.Cells["O65"].Value = -total_b1a_o;

                //******************Setir B1b

                var Setir_b1b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1b_c = Setir_b1b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1b_e = Setir_b1b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1b_f = Setir_b1b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1b_g = Setir_b1b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1b_i = Setir_b1b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1b_k = Setir_b1b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1b_o = Setir_b1b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C66"].Value = -total_b1b_c;
                worksheet1.Cells["E66"].Value = -total_b1b_e;
                worksheet1.Cells["F66"].Value = -total_b1b_f;
                worksheet1.Cells["G66"].Value = -total_b1b_g;
                worksheet1.Cells["I66"].Value = -total_b1b_i;
                worksheet1.Cells["K66"].Value = -total_b1b_k;
                worksheet1.Cells["O66"].Value = -total_b1b_o;

                //******************Setir B1c

                var Setir_b1c_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b1c_c = Setir_b1c_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b1c_e = Setir_b1c_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b1c_f = Setir_b1c_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b1c_g = Setir_b1c_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b1c_i = Setir_b1c_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b1c_k = Setir_b1c_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b1c_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B1c.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b1c_o = Setir_b1c_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C67"].Value = -total_b1c_c;
                worksheet1.Cells["E67"].Value = -total_b1c_e;
                worksheet1.Cells["F67"].Value = -total_b1c_f;
                worksheet1.Cells["G67"].Value = -total_b1c_g;
                worksheet1.Cells["I67"].Value = -total_b1c_i;
                worksheet1.Cells["K67"].Value = -total_b1c_k;
                worksheet1.Cells["O67"].Value = -total_b1c_o;

                //******************Setir B2b

                var Setir_b2b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b2b_c = Setir_b2b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b2b_e = Setir_b2b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b2b_f = Setir_b2b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b2b_g = Setir_b2b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b2b_i = Setir_b2b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b2b_k = Setir_b2b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b2b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B2b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b2b_o = Setir_b2b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C70"].Value = total_b2b_c;
                worksheet1.Cells["E70"].Value = total_b2b_e;
                worksheet1.Cells["F70"].Value = total_b2b_f;
                worksheet1.Cells["G70"].Value = total_b2b_g;
                worksheet1.Cells["I70"].Value = total_b2b_i;
                worksheet1.Cells["K70"].Value = total_b2b_k;
                worksheet1.Cells["O70"].Value = total_b2b_o;

                //******************Setir B4a

                var Setir_b4a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b4a_c = Setir_b4a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b4a_e = Setir_b4a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b4a_f = Setir_b4a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b4a_g = Setir_b4a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b4a_i = Setir_b4a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b4a_k = Setir_b4a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b4a_o = Setir_b4a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C73"].Value = -total_b4a_c;
                worksheet1.Cells["E73"].Value = -total_b4a_e;
                worksheet1.Cells["F73"].Value = -total_b4a_f;
                worksheet1.Cells["G73"].Value = -total_b4a_g;
                worksheet1.Cells["I73"].Value = -total_b4a_i;
                worksheet1.Cells["K73"].Value = -total_b4a_k;
                worksheet1.Cells["O73"].Value = -total_b4a_o;

                //******************Setir B4b

                var Setir_b4b_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b4b_c = Setir_b4b_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b4b_e = Setir_b4b_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b4b_f = Setir_b4b_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b4b_g = Setir_b4b_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b4b_i = Setir_b4b_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b4b_k = Setir_b4b_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b4b_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B4b.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b4b_o = Setir_b4b_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C74"].Value = -total_b4b_c;
                worksheet1.Cells["E74"].Value = -total_b4b_e;
                worksheet1.Cells["F74"].Value = -total_b4b_f;
                worksheet1.Cells["G74"].Value = -total_b4b_g;
                worksheet1.Cells["I74"].Value = -total_b4b_i;
                worksheet1.Cells["K74"].Value = -total_b4b_k;
                worksheet1.Cells["O74"].Value = -total_b4b_o;

                //******************Setir B5a

                var Setir_b5a_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b5a_c = Setir_b5a_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b5a_e = Setir_b5a_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b5a_f = Setir_b5a_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b5a_g = Setir_b5a_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b5a_i = Setir_b5a_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b5a_k = Setir_b5a_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b5a_o = Setir_b5a_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C76"].Value = -total_b5a_c;
                worksheet1.Cells["E76"].Value = -total_b5a_e;
                worksheet1.Cells["F76"].Value = -total_b5a_f;
                worksheet1.Cells["G76"].Value = -total_b5a_g;
                worksheet1.Cells["I76"].Value = -total_b5a_i;
                worksheet1.Cells["K76"].Value = -total_b5a_k;
                worksheet1.Cells["O76"].Value = -total_b5a_o;

                //******************Setir B5a1

                var Setir_b5a1_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b5a1_c = Setir_b5a1_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b5a1_e = Setir_b5a1_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b5a1_f = Setir_b5a1_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b5a1_g = Setir_b5a1_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b5a1_i = Setir_b5a1_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b5a1_k = Setir_b5a1_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b5a1_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B5a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b5a1_o = Setir_b5a1_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C77"].Value = -total_b5a1_c;
                worksheet1.Cells["E77"].Value = -total_b5a1_e;
                worksheet1.Cells["F77"].Value = -total_b5a1_f;
                worksheet1.Cells["G77"].Value = -total_b5a1_g;
                worksheet1.Cells["I77"].Value = -total_b5a1_i;
                worksheet1.Cells["K77"].Value = -total_b5a1_k;
                worksheet1.Cells["O77"].Value = -total_b5a1_o;

                //******************Setir B8

                var Setir_b8_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b8_c = Setir_b8_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b8_e = Setir_b8_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b8_f = Setir_b8_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b8_g = Setir_b8_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b8_i = Setir_b8_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b8_k = Setir_b8_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b8_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B8.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b8_o = Setir_b8_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C86"].Value = total_b8_c;
                worksheet1.Cells["E86"].Value = total_b8_e;
                worksheet1.Cells["F86"].Value = total_b8_f;
                worksheet1.Cells["G86"].Value = total_b8_g;
                worksheet1.Cells["I86"].Value = total_b8_i;
                worksheet1.Cells["K86"].Value = total_b8_k;
                worksheet1.Cells["O86"].Value = total_b8_o;

                //******************Setir B10

                var Setir_b10_c = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b10_c = Setir_b10_c.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_e = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_b10_e = Setir_b10_e.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_f = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "01")
                                    .ToList();
                decimal total_b10_f = Setir_b10_f.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_g = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "02")
                                    .ToList();
                decimal total_b10_g = Setir_b10_g.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_i = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "03")
                                    .ToList();
                decimal total_b10_i = Setir_b10_i.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_k = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "04")
                                    .ToList();
                decimal total_b10_k = Setir_b10_k.Sum(row => row.Field<decimal>(3)) / 1000;

                var Setir_b10_o = _dt_daily_report.AsEnumerable()
                                    .Where(row => B10.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "05")
                                    .ToList();
                decimal total_b10_o = Setir_b10_o.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C89"].Value = -total_b10_c;
                worksheet1.Cells["E89"].Value = -total_b10_e;
                worksheet1.Cells["F89"].Value = -total_b10_f;
                worksheet1.Cells["G89"].Value = -total_b10_g;
                worksheet1.Cells["I89"].Value = -total_b10_i;
                worksheet1.Cells["K89"].Value = -total_b10_k;
                worksheet1.Cells["O89"].Value = -total_b10_o;

                //******************Setir B10a
                var total_b10a = new Dictionary<string, decimal>();
                var types_b10a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b10a)
                {
                    var Setir_b10a = _dt_daily_report.AsEnumerable()
                                        .Where(row => B10a.Contains(row.Field<string>(1).Substring(0, 5))
                                        && row.Field<string>(2) == type)
                                        .ToList();
                    total_b10a[type] = Setir_b10a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E90"].Value = total_b10a["00"];
                worksheet1.Cells["F90"].Value = total_b10a["01"];
                worksheet1.Cells["G90"].Value = total_b10a["02"];
                worksheet1.Cells["I90"].Value = total_b10a["03"];
                worksheet1.Cells["K90"].Value = total_b10a["04"];
                worksheet1.Cells["O90"].Value = total_b10a["05"];

                var Setir_b10a_1 = _dt_daily_report1.AsEnumerable()
                                    .Where(row => B10a.Contains(row.Field<string>(1).Substring(0, 5)))
                                    .ToList();
                decimal total_b10a_1 = Setir_b10a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C90"].Value = total_b10a_1;

                //******************Setir B12
                var total_b12 = new Dictionary<string, decimal>();
                var total_b12_eh = new Dictionary<string, decimal>();
                var types_b12 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b12)
                {
                    var Setir_b12 = _dt_daily_report.AsEnumerable()
                                        .Where(row => (B12.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00"))
                                        .ToList();
                    total_b12[type] = Setir_b12.Sum(row => row.Field<decimal>(3)) / 1000;
                    decimal cem = Setir_b12.Sum(row => row.Field<decimal>(3)) / 1000;
                    var Setir_b12_eh = _dt_daily_report.AsEnumerable()
                                        .Where(row => (B12.Contains(row.Field<string>(1).Substring(0, 5)) ||
                                       (B12_eht.Contains(row.Field<string>(1).Substring(0, 3)) &&
                                       Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) < 25
                                       && Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) != 100)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b12_eh[type] = Setir_b12_eh.Sum(row => row.Field<decimal>(3)) / 1000;
                    decimal cem1 = Setir_b12_eh.Sum(row => row.Field<decimal>(3)) / 1000;
                }


                worksheet1.Cells["E93"].Value = -total_b12_eh["00"];
                worksheet1.Cells["F93"].Value = -total_b12_eh["01"];
                worksheet1.Cells["G93"].Value = -total_b12_eh["02"];
                worksheet1.Cells["I93"].Value = -total_b12_eh["03"];
                worksheet1.Cells["K93"].Value = -total_b12_eh["04"];
                worksheet1.Cells["O93"].Value = -total_b12_eh["05"];
                var Setir_b12_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => (B12.Contains(row.Field<string>(1).Substring(0, 5)) ||
                                       (B12_eht.Contains(row.Field<string>(1).Substring(0, 3)) &&
                                       Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 2))) < 25
                                       && Convert.ToInt32(row.Field<string>(1).Substring(Math.Max(0, row.Field<string>(1).Length - 3))) != 100)))
                                        .ToList();
                decimal total_b12_1 = Setir_b12_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C93"].Value = -total_b12_1;

                //******************Setir B12a
                var total_b12a = new Dictionary<string, decimal>();
                var types_b12a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b12a)
                {
                    var Setir_b12a = _dt_daily_report.AsEnumerable()
                                        .Where(row => B12a.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b12a[type] = Setir_b12a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E94"].Value = -total_b12a["00"];
                worksheet1.Cells["F94"].Value = -total_b12a["01"];
                worksheet1.Cells["G94"].Value = -total_b12a["02"];
                worksheet1.Cells["I94"].Value = -total_b12a["03"];
                worksheet1.Cells["K94"].Value = -total_b12a["04"];
                worksheet1.Cells["O94"].Value = -total_b12a["05"];
                var Setir_b12a_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => B12a.Contains(row.Field<string>(1).Substring(0, 5)))
                                        .ToList();
                decimal total_b12a_1 = Setir_b12a_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C94"].Value = -total_b12a_1;
                //******************Setir B12
                var total_b13 = new Dictionary<string, decimal>();
                var types_b13 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_b13)
                {
                    var Setir_b13 = _dt_daily_report.AsEnumerable()
                                        .Where(row => B13.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == type)
                                        .ToList();
                    total_b13[type] = Setir_b13.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E95"].Value = -total_b13["00"];
                worksheet1.Cells["F95"].Value = -total_b13["01"];
                worksheet1.Cells["G95"].Value = -total_b13["02"];
                worksheet1.Cells["I95"].Value = -total_b13["03"];
                worksheet1.Cells["K95"].Value = -total_b13["04"];
                worksheet1.Cells["O95"].Value = -total_b13["05"];
                var Setir_b13_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row => B13.Contains(row.Field<string>(1).Substring(0, 5)))
                                        .ToList();
                decimal total_b13_1 = Setir_b13_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C95"].Value = -total_b13_1;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 1
                var total_bo1 = new Dictionary<string, decimal>();
                var types_bo1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo1)
                {
                    var Setir_bo1 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99550" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo1[type] = Setir_bo1.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E100"].Value = total_bo1["00"];
                worksheet1.Cells["F100"].Value = total_bo1["01"];
                worksheet1.Cells["G100"].Value = total_bo1["02"];
                worksheet1.Cells["I100"].Value = total_bo1["03"];
                worksheet1.Cells["K100"].Value = total_bo1["04"];
                worksheet1.Cells["O100"].Value = total_bo1["05"];
                var Setir_1_bo1 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99550")
                                    .ToList();
                decimal total_1_bo1 = Setir_1_bo1.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C100"].Value = total_1_bo1;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 2
                var total_bo2 = new Dictionary<string, decimal>();
                var types_bo2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo2)
                {
                    var Setir_bo2 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99530" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo2[type] = Setir_bo2.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E101"].Value = total_bo2["00"];
                worksheet1.Cells["F101"].Value = total_bo2["01"];
                worksheet1.Cells["G101"].Value = total_bo2["02"];
                worksheet1.Cells["I101"].Value = total_bo2["03"];
                worksheet1.Cells["K101"].Value = total_bo2["04"];
                worksheet1.Cells["O101"].Value = total_bo2["05"];
                var Setir_1_bo2 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99530")
                                    .ToList();
                decimal total_1_bo2 = Setir_1_bo2.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C101"].Value = total_1_bo2;
                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 3
                var total_bo3 = new Dictionary<string, decimal>();
                var types_bo3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo3)
                {
                    var Setir_bo3 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99531" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo3[type] = Setir_bo3.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E102"].Value = total_bo3["00"];
                worksheet1.Cells["F102"].Value = total_bo3["01"];
                worksheet1.Cells["G102"].Value = total_bo3["02"];
                worksheet1.Cells["I102"].Value = total_bo3["03"];
                worksheet1.Cells["K102"].Value = total_bo3["04"];
                worksheet1.Cells["O102"].Value = total_bo3["05"];
                var Setir_1_bo3 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99531")
                                    .ToList();
                decimal total_1_bo3 = Setir_1_bo3.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C102"].Value = total_1_bo3;

                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 7
                var total_bo7 = new Dictionary<string, decimal>();
                var types_bo7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo7)
                {
                    var Setir_bo7 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99531" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo7[type] = Setir_bo7.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E106"].Value = total_bo7["00"];
                worksheet1.Cells["F106"].Value = total_bo7["01"];
                worksheet1.Cells["G106"].Value = total_bo7["02"];
                worksheet1.Cells["I106"].Value = total_bo7["03"];
                worksheet1.Cells["K106"].Value = total_bo7["04"];
                worksheet1.Cells["O106"].Value = total_bo7["05"];
                var Setir_1_bo7 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99531")
                                    .ToList();
                decimal total_1_bo7 = Setir_1_bo7.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C106"].Value = total_1_bo7;

                //******************Setir Risklər barədə məlumatlar 1a
                var total_risk_1a = new Dictionary<string, decimal>();
                var types_risk_1a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_1a)
                {
                    var Setir_risk_1a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(4) == "vk" && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_1a[type] = Setir_risk_1a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E119"].Value = total_risk_1a["00"];
                worksheet1.Cells["F119"].Value = total_risk_1a["01"];
                worksheet1.Cells["G119"].Value = total_risk_1a["02"];
                worksheet1.Cells["I119"].Value = total_risk_1a["03"];
                worksheet1.Cells["K119"].Value = total_risk_1a["04"];
                worksheet1.Cells["O119"].Value = total_risk_1a["05"];
                var Setir_1_risk_1a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(4) == "vk")
                                        .ToList();
                decimal total_1_risk_1a = Setir_1_risk_1a.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C119"].Value = total_1_risk_1a;

                //******************Setir Risklər barədə məlumatlar 1b
                var total_risk_1b = new Dictionary<string, decimal>();
                var types_risk_1b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_1b)
                {
                    var Setir_risk_1b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(4) == "vk" && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_1b[type] = Setir_risk_1b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E120"].Value = total_risk_1b["00"];
                worksheet1.Cells["F120"].Value = total_risk_1b["01"];
                worksheet1.Cells["G120"].Value = total_risk_1b["02"];
                worksheet1.Cells["I120"].Value = total_risk_1b["03"];
                worksheet1.Cells["K120"].Value = total_risk_1b["04"];
                worksheet1.Cells["O120"].Value = total_risk_1b["05"];
                var Setir_1_risk_1b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(4) == "vk")
                                        .ToList();
                decimal total_1_risk_1b = Setir_1_risk_1b.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C120"].Value = total_1_risk_1b;

                //******************Setir Risklər barədə məlumatlar 2a
                var total_risk_2a = new Dictionary<string, decimal>();
                var types_risk_2a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_2a)
                {
                    var Setir_risk_2a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_2a[type] = Setir_risk_2a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E123"].Value = total_risk_2a["00"];
                worksheet1.Cells["F123"].Value = total_risk_2a["01"];
                worksheet1.Cells["G123"].Value = total_risk_2a["02"];
                worksheet1.Cells["I123"].Value = total_risk_2a["03"];
                worksheet1.Cells["K123"].Value = total_risk_2a["04"];
                worksheet1.Cells["O123"].Value = total_risk_2a["05"];
                var Setir_1_risk_2a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_2a = Setir_1_risk_2a.Sum(row => row.Field<decimal>(3)) / 1000;

                worksheet1.Cells["C123"].Value = total_1_risk_2a;

                //******************Setir Risklər barədə məlumatlar 2b
                var total_risk_2b = new Dictionary<string, decimal>();
                var types_risk_2b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_2b)
                {
                    var Setir_risk_2b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_2b[type] = Setir_risk_2b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E124"].Value = total_risk_2b["00"];
                worksheet1.Cells["F124"].Value = total_risk_2b["01"];
                worksheet1.Cells["G124"].Value = total_risk_2b["02"];
                worksheet1.Cells["I124"].Value = total_risk_2b["03"];
                worksheet1.Cells["K124"].Value = total_risk_2b["04"];
                worksheet1.Cells["O124"].Value = total_risk_2b["05"];
                var Setir_1_risk_2b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_2b = Setir_1_risk_2b.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C124"].Value = total_1_risk_2b;

                //******************Setir Risklər barədə məlumatlar 3a
                var total_risk_3a = new Dictionary<string, decimal>();
                var types_risk_3a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_3a)
                {
                    var Setir_risk_3a = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_3a[type] = Setir_risk_3a.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E127"].Value = total_risk_3a["00"];
                worksheet1.Cells["F127"].Value = total_risk_3a["01"];
                worksheet1.Cells["F127"].Value = total_risk_3a["02"];
                worksheet1.Cells["G127"].Value = total_risk_3a["03"];
                worksheet1.Cells["K127"].Value = total_risk_3a["04"];
                worksheet1.Cells["O127"].Value = total_risk_3a["05"];
                var Setir_1_risk_3a = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        (row.Field<string>(1) == "huquqi" || row.Field<string>(1) == "sahibkar")
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_3a = Setir_1_risk_3a.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C127"].Value = total_1_risk_3a;

                //******************Setir Risklər barədə məlumatlar 3b
                var total_risk_3b = new Dictionary<string, decimal>();
                var types_risk_3b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_risk_3b)
                {
                    var Setir_risk_3b = _dt_qali_gunler.AsEnumerable()
                                        .Where(row => row.Field<string>(0) == "bugun" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90 && row.Field<string>(6) == type)
                                        .ToList();
                    total_risk_3b[type] = Setir_risk_3b.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E128"].Value = total_risk_3b["00"];
                worksheet1.Cells["F128"].Value = total_risk_3b["01"];
                worksheet1.Cells["G128"].Value = total_risk_3b["02"];
                worksheet1.Cells["I128"].Value = total_risk_3b["03"];
                worksheet1.Cells["K128"].Value = total_risk_3b["04"];
                worksheet1.Cells["O128"].Value = total_risk_3b["05"];
                var Setir_1_risk_3b = _dt_qali_gunler.AsEnumerable()
                                    .Where(row => row.Field<string>(0) == "dunen" &&
                                        row.Field<string>(1) == "fiziki"
                                        && row.Field<string>(5) == "rest"
                                        && row.Field<decimal>(2) > 90)
                                        .ToList();
                decimal total_1_risk_3b = Setir_1_risk_3b.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C128"].Value = total_1_risk_3b;

                //******************Setir BALANSARXASI ÖHDƏLİKLƏR 4
                var total_bo4 = new Dictionary<string, decimal>();
                var types_bo4 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_bo4)
                {
                    var Setir_bo4 = _dt_daily_report_bk.AsEnumerable()
                                        .Where(row => row.Field<string>(1) == "99300" && row.Field<string>(3) == type)
                                        .ToList();
                    total_bo4[type] = Setir_bo4.Sum(row => row.Field<decimal>(6)) / 1000;
                }
                worksheet1.Cells["E130"].Value = total_bo4["00"];
                worksheet1.Cells["F130"].Value = total_bo4["01"];
                worksheet1.Cells["G130"].Value = total_bo4["02"];
                worksheet1.Cells["I130"].Value = total_bo4["03"];
                worksheet1.Cells["K130"].Value = total_bo4["04"];
                worksheet1.Cells["O130"].Value = total_bo4["05"];

                var Setir_1_bo4 = _dt_daily_report_bk1.AsEnumerable()
                                    .Where(row => row.Field<string>(1) == "99300")
                                    .ToList();
                decimal total_1_bo4 = Setir_1_bo4.Sum(row => row.Field<decimal>(6)) / 1000;
                worksheet1.Cells["C130"].Value = total_1_bo4;


                //******************Setir Likvildik riskləri  1
                var total_lr1 = new Dictionary<string, decimal>();
                var types_lr1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr1)
                {
                    var Setir_lr1 = _dt_daily_report.AsEnumerable()
                                        .Where(row => likvid_risk_1.Contains(row.Field<string>(1).Substring(0, 5))
                              && row.Field<string>(2) == type)
                                        .ToList();
                    //total_lr1[type] = Setir_lr1.Sum(row => row.Field<decimal>(3)) / 1000;
                    total_lr1[type] = Setir_lr1.Sum(row =>
                    {
                        string substringValue = row.Field<string>(1).Substring(0, 5);

                        // Eğer substringValue "15020" veya "15025" ise, meblağı %25 azalt
                        decimal discountFactor = (substringValue == "15020" || substringValue == "15025") ? 0.75m : 1.0m;

                        return row.Field<decimal>(3) * discountFactor / 1000;
                    });
                }
                worksheet1.Cells["E135"].Value = total_lr1["00"];
                worksheet1.Cells["F135"].Value = total_lr1["01"];
                worksheet1.Cells["G135"].Value = total_lr1["02"];
                worksheet1.Cells["I135"].Value = total_lr1["03"];
                worksheet1.Cells["K135"].Value = total_lr1["04"];
                worksheet1.Cells["O135"].Value = total_lr1["05"];

                var sh_67_00 = _dt_daily_report1.AsEnumerable()
                .Where(row => likvid_risk_1.Contains(row.Field<string>(1).Substring(0, 5)))
                    .ToList();

                decimal total_67_d0 = sh_67_00.Sum(row =>
                {
                    string substringValue = row.Field<string>(1).Substring(0, 5);

                    // Eğer substringValue "15020" veya "15025" ise, meblağı %25 azalt
                    decimal discountFactor = (substringValue == "15020" || substringValue == "15025") ? 0.75m : 1.0m;

                    return row.Field<decimal>(3) * discountFactor / 1000;
                });
                worksheet1.Cells["C135"].Value = total_67_d0;

                var total_lr2 = new Dictionary<string, decimal>();
                var types_lr2 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr2)
                {
                    var Setir_lr2 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    (row.Field<string>(1).Substring(0, 1) == "3" || row.Field<string>(1).Substring(0, 1) == "4")
                    && row.Field<string>(2).Substring(0, 2) == type
                    && (!likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 5))) && !likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 3))
                )
                .ToList();
                    total_bo4[type] = Setir_lr2.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E136"].Value = -total_bo4["00"];
                worksheet1.Cells["F136"].Value = -total_bo4["01"];
                worksheet1.Cells["G136"].Value = -total_bo4["02"];
                worksheet1.Cells["I136"].Value = -total_bo4["03"];
                worksheet1.Cells["K136"].Value = -total_bo4["04"];
                worksheet1.Cells["O136"].Value = -total_bo4["05"];

                var Setir_lr2_1 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    (row.Field<string>(1).Substring(0, 1) == "3" || row.Field<string>(1).Substring(0, 1) == "4")

                    && (!likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 5))) && !likvid_risk_2.Contains(row.Field<string>(1).Substring(0, 3))
                )
                .ToList();
                decimal total_lr2_1 = Setir_lr2_1.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C136"].Value = -total_lr2_1;



                var total_lr4 = new Dictionary<string, decimal>();
                var types_lr4 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr4)
                {
                    var Setir_lr4 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_4_elave.Contains(row.Field<string>(1).Substring(0, 5)) || likvid_risk_4.Contains(row.Field<string>(1).Substring(0, 2))) && !likvid_risk_4_istisna.Contains(row.Field<string>(1))
                    && row.Field<string>(2).Substring(0, 2) == type)
                .ToList();
                    //total_lr4[type] = Setir_lr4.Sum(row => row.Field<decimal>(3)) / 1000;
                    total_lr4[type] = Setir_lr4.Sum(row =>
                    {
                        // İlk 5 simvol "15025"-ə bərabərdirsə, dəyəri 25% azaldırıq
                        if (row.Field<string>(1).Substring(0, 5) == "15025")
                        {
                            return row.Field<decimal>(3) * 0.75m; // 25% azaldılmış məbləğ
                        }
                        else
                        {
                            return row.Field<decimal>(3); // Qalan şərtlərdə məbləğ olduğu kimi
                        }
                    }) / 1000;
                }
               
                worksheet1.Cells["E138"].Value = total_lr4["00"];
                worksheet1.Cells["F138"].Value = total_lr4["01"];
                worksheet1.Cells["G138"].Value = total_lr4["02"];
                worksheet1.Cells["I138"].Value = total_lr4["03"];
                worksheet1.Cells["K138"].Value = total_lr4["04"];
                worksheet1.Cells["O138"].Value = total_lr4["05"];

                var Setir_lr2_4 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_4_elave.Contains(row.Field<string>(1).Substring(0, 5)) 
                    || likvid_risk_4.Contains(row.Field<string>(1).Substring(0, 2))) 
                    && !likvid_risk_4_istisna.Contains(row.Field<string>(1))).ToList();

                decimal total_lr2_4 = Setir_lr2_4.Sum(row =>
                {
                    // İlk 5 simvol "15025"-ə bərabərdirsə, dəyəri 25% azaldırıq
                    if (row.Field<string>(1).Substring(0, 5) == "15025")
                    {
                        return row.Field<decimal>(3) * 0.75m; // 25% azaldılmış məbləğ
                    }
                    else
                    {
                        return row.Field<decimal>(3); // Qalan şərtlərdə məbləğ olduğu kimi
                    }
                }) / 1000;
                worksheet1.Cells["C138"].Value = total_lr2_4;

                var total_lr5 = new Dictionary<string, decimal>();
                var types_lr5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr5)
                {
                    var Setir_lr5 = _dt_daily_report.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 2)) || likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 5)))
                    && row.Field<string>(2).Substring(0, 2) == type)
                .ToList();
                    total_lr5[type] = Setir_lr5.Sum(row => row.Field<decimal>(3)) / 1000;
                }
                worksheet1.Cells["E139"].Value = -(total_lr5["00"] + total_bo2["00"]);
                worksheet1.Cells["F139"].Value = -(total_lr5["01"] + total_bo2["01"]);
                worksheet1.Cells["G139"].Value = -(total_lr5["02"] + total_bo2["02"]);
                worksheet1.Cells["I139"].Value = -(total_lr5["03"] + total_bo2["03"]);
                worksheet1.Cells["K139"].Value = -(total_lr5["04"] + total_bo2["04"]);
                worksheet1.Cells["O139"].Value = -(total_lr5["05"] + total_bo2["05"]);

                var Setir_lr2_5 = _dt_daily_report1.AsEnumerable()
                                        .Where(row =>
                    (likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 2)) || likvid_risk_5.Contains(row.Field<string>(1).Substring(0, 5))))
                .ToList();
                decimal total_lr2_5 = Setir_lr2_5.Sum(row => row.Field<decimal>(3)) / 1000;
                worksheet1.Cells["C139"].Value = -(total_lr2_5 + total_1_bo2);
                decimal t1 = total_lr2_5;
                decimal t2 = total_1_bo2;
                //Setir 136
                if (total_lr2_5 != 0)
                {
                    worksheet1.Cells["C137"].Value = -total_lr2_4 / total_lr2_5;
                }
                else
                {
                    worksheet1.Cells["C137"].Value = 0; // veya başka bir değer
                }

                if (total_lr5["00"] != 0)
                {
                    worksheet1.Cells["E137"].Value = -total_lr4["00"] / total_lr5["00"];
                }
                else
                {
                    worksheet1.Cells["E137"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["01"] != 0)
                {
                    worksheet1.Cells["F137"].Value = -total_lr4["01"] / total_lr5["01"];
                }
                else
                {
                    worksheet1.Cells["F137"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["02"] != 0)
                {
                    worksheet1.Cells["G137"].Value = -total_lr4["02"] / total_lr5["02"];
                }
                else
                {
                    worksheet1.Cells["G137"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["03"] != 0)
                {
                    worksheet1.Cells["I137"].Value = -total_lr4["03"] / total_lr5["03"];
                }
                else
                {
                    worksheet1.Cells["I137"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["04"] != 0)
                {
                    worksheet1.Cells["K137"].Value = -total_lr4["04"] / total_lr5["04"];
                }
                else
                {
                    worksheet1.Cells["K137"].Value = 0; // veya başka bir değer
                }
                if (total_lr5["05"] != 0)
                {
                    worksheet1.Cells["O137"].Value = -total_lr4["05"] / total_lr5["05"];
                }
                else
                {
                    worksheet1.Cells["O137"].Value = 0; // veya başka bir değer
                }

                //IV Hissə – Balans maddələri üzrə dəyişikliklərə dair əlavə məlumatlar*******************

                var total_IV_hisse_1 = new Dictionary<string, decimal>();
                var types_IV_hisse_1 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_1)
                {
                    var Setir_IV_hisse_1 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();

                    total_IV_hisse_1[type] = Setir_IV_hisse_1.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E145"].Value = total_IV_hisse_1["00"];
                worksheet1.Cells["F145"].Value = total_IV_hisse_1["01"];
                worksheet1.Cells["G145"].Value = total_IV_hisse_1["02"];
                worksheet1.Cells["I145"].Value = total_IV_hisse_1["03"];
                worksheet1.Cells["K145"].Value = total_IV_hisse_1["04"];
                worksheet1.Cells["O145"].Value = total_IV_hisse_1["05"];

                var Setir_IV_hisse_1_1 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_1 = Setir_IV_hisse_1_1.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["c145"].Value = total_IV_hisse_1_1;

                var total_IV_hisse_3 = new Dictionary<string, decimal>();
                var types_IV_hisse_3 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_lr2)
                {
                    var Setir_IV_hisse_3 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_3[type] = Setir_IV_hisse_3.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E147"].Value = total_IV_hisse_3["00"];
                worksheet1.Cells["F147"].Value = total_IV_hisse_3["01"];
                worksheet1.Cells["G147"].Value = total_IV_hisse_3["02"];
                worksheet1.Cells["I147"].Value = total_IV_hisse_3["03"];
                worksheet1.Cells["K147"].Value = total_IV_hisse_3["04"];
                worksheet1.Cells["O147"].Value = total_IV_hisse_3["05"];

                var Setir_IV_hisse_1_3 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_3 = Setir_IV_hisse_1_3.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["c147"].Value = total_IV_hisse_1_3;

                var total_IV_hisse_5 = new Dictionary<string, decimal>();
                decimal sayH = 0;
                var types_IV_hisse_5 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_5)
                {
                    var Setir_IV_hisse_5 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_5[type] = Setir_IV_hisse_5.Sum(row => row.Field<decimal>(2)) / 1000;
                    sayH = Setir_IV_hisse_5.Sum(row => row.Field<decimal>(4));
                }
                worksheet1.Cells["E149"].Value = total_IV_hisse_5["00"];
                worksheet1.Cells["F149"].Value = total_IV_hisse_5["01"];
                worksheet1.Cells["G149"].Value = total_IV_hisse_5["02"];
                worksheet1.Cells["I149"].Value = total_IV_hisse_5["03"];
                worksheet1.Cells["K149"].Value = total_IV_hisse_5["04"];
                worksheet1.Cells["O149"].Value = total_IV_hisse_5["05"];

                worksheet3.Cells["C30"].Value = total_IV_hisse_5;
                worksheet3.Cells["C30"].Value = total_IV_hisse_5;
                worksheet3.Cells["D31"].Value = sayH;
                worksheet3.Cells["D31"].Value = sayH;

                var Setir_IV_hisse_1_5 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_IV_hisse_1_5 = Setir_IV_hisse_1_5.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C149"].Value = total_IV_hisse_1_5;

                var total_IV_hisse_7 = new Dictionary<string, decimal>();
                var types_IV_hisse_7 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_7)
                {
                    var Setir_IV_hisse_7 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_7[type] = Setir_IV_hisse_7.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E151"].Value = total_IV_hisse_7["00"];
                worksheet1.Cells["F151"].Value = total_IV_hisse_7["01"];
                worksheet1.Cells["G151"].Value = total_IV_hisse_7["02"];
                worksheet1.Cells["I151"].Value = total_IV_hisse_7["03"];
                worksheet1.Cells["K151"].Value = total_IV_hisse_7["04"];
                worksheet1.Cells["O151"].Value = total_IV_hisse_7["05"];

                var Setir_IV_hisse_1_7 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_IV_hisse_1_7 = Setir_IV_hisse_1_7.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C151"].Value = total_IV_hisse_1_7;

                var total_IV_hisse_9 = new Dictionary<string, decimal>();
                decimal sayS = 0;
                var types_IV_hisse_9 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_9)
                {
                    var Setir_IV_hisse_9 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_9[type] = Setir_IV_hisse_9.Sum(row => row.Field<decimal>(2)) / 1000;
                    sayS = Setir_IV_hisse_9.Sum(row => row.Field<decimal>(4));
                }
                worksheet1.Cells["E153"].Value = total_IV_hisse_9["00"];
                worksheet1.Cells["F153"].Value = total_IV_hisse_9["01"];
                worksheet1.Cells["G153"].Value = total_IV_hisse_9["02"];
                worksheet1.Cells["I153"].Value = total_IV_hisse_9["03"];
                worksheet1.Cells["K153"].Value = total_IV_hisse_9["04"];
                worksheet1.Cells["O153"].Value = total_IV_hisse_9["05"];

                worksheet3.Cells["C35"].Value = total_IV_hisse_9;
                worksheet3.Cells["C35"].Value = total_IV_hisse_9;
                worksheet3.Cells["D36"].Value = sayS;
                worksheet3.Cells["D36"].Value = sayS;

                var Setir_IV_hisse_1_9 = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_IV_hisse_1_9 = Setir_IV_hisse_1_9.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C153"].Value = total_IV_hisse_1_9;

                var total_IV_hisse_10 = new Dictionary<string, decimal>();
                var types_IV_hisse_10 = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_9)
                {
                    var Setir_IV_hisse_10 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type)
                .ToList();
                    total_IV_hisse_10[type] = Setir_IV_hisse_10.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E154"].Value = total_IV_hisse_10["00"];
                worksheet1.Cells["F154"].Value = total_IV_hisse_10["01"];
                worksheet1.Cells["G154"].Value = total_IV_hisse_10["02"];
                worksheet1.Cells["I154"].Value = total_IV_hisse_10["03"];
                worksheet1.Cells["K154"].Value = total_IV_hisse_10["04"];
                worksheet1.Cells["O154"].Value = total_IV_hisse_10["05"];

                var Setir_IV_hisse_1_10 = _dt_mexaric.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_IV_hisse_1_10 = Setir_IV_hisse_1_10.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C154"].Value = total_IV_hisse_1_10;


                var total_IV_hisse_11a = new Dictionary<string, decimal>();
                var types_IV_hisse_11a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_11a)
                {
                    var Setir_IV_hisse_11a = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    && row.Field<string>(3).Substring(0, 2) == type && row.Field<string>(3) == "bos")
                .ToList();
                    total_IV_hisse_11a[type] = Setir_IV_hisse_11a.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E156"].Value = total_IV_hisse_11a["00"];
                worksheet1.Cells["F156"].Value = total_IV_hisse_11a["01"];
                worksheet1.Cells["G156"].Value = total_IV_hisse_11a["02"];
                worksheet1.Cells["I156"].Value = total_IV_hisse_11a["03"];
                worksheet1.Cells["K156"].Value = total_IV_hisse_11a["04"];
                worksheet1.Cells["O156"].Value = total_IV_hisse_11a["05"];

                var Setir_IV_hisse_1_11a = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar" && row.Field<string>(3) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_11a = Setir_IV_hisse_1_11a.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C156"].Value = total_IV_hisse_1_11a;


                var total_IV_hisse_11b = new Dictionary<string, decimal>();
                var types_IV_hisse_11b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_11b)
                {
                    var Setir_IV_hisse_11b = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    && row.Field<string>(3).Substring(0, 2) == type && row.Field<string>(3) == "bos")
                .ToList();
                    total_IV_hisse_11b[type] = Setir_IV_hisse_11b.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E157"].Value = total_IV_hisse_11b["00"];
                worksheet1.Cells["F157"].Value = total_IV_hisse_11b["01"];
                worksheet1.Cells["G157"].Value = total_IV_hisse_11b["02"];
                worksheet1.Cells["I157"].Value = total_IV_hisse_11b["03"];
                worksheet1.Cells["K157"].Value = total_IV_hisse_11b["04"];
                worksheet1.Cells["O157"].Value = total_IV_hisse_11b["05"];

                var Setir_IV_hisse_1_11b = _dt_verilmis.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(1) == "sahibkar" && row.Field<string>(3) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_11b = Setir_IV_hisse_1_11b.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C157"].Value = total_IV_hisse_1_11b;


                var total_IV_hisse_12a = new Dictionary<string, decimal>();
                var types_IV_hisse_12a = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_12a)
                {
                    var Setir_IV_hisse_12a = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(3) == "sahibkar")
                    && row.Field<string>(4).Substring(0, 2) == type && row.Field<string>(5) == "bos")
                .ToList();
                    total_IV_hisse_12a[type] = Setir_IV_hisse_12a.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E160"].Value = total_IV_hisse_12a["00"];
                worksheet1.Cells["F160"].Value = total_IV_hisse_12a["01"];
                worksheet1.Cells["G160"].Value = total_IV_hisse_12a["02"];
                worksheet1.Cells["I160"].Value = total_IV_hisse_12a["03"];
                worksheet1.Cells["K160"].Value = total_IV_hisse_12a["04"];
                worksheet1.Cells["O160"].Value = total_IV_hisse_12a["05"];

                var Setir_IV_hisse_1_12a = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(3) == "sahibkar" && row.Field<string>(5) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_12a = Setir_IV_hisse_1_12a.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C160"].Value = total_IV_hisse_1_12a;


                var total_IV_hisse_12b = new Dictionary<string, decimal>();
                var types_IV_hisse_12b = new string[] { "00", "01", "02", "03", "04", "05" };
                foreach (var type in types_IV_hisse_12b)
                {
                    var Setir_IV_hisse_12b = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(3) == "fiziki")
                    && row.Field<string>(4).Substring(0, 2) == type && row.Field<string>(5) == "bos")
                .ToList();
                    total_IV_hisse_12b[type] = Setir_IV_hisse_12b.Sum(row => row.Field<decimal>(2)) / 1000;
                }
                worksheet1.Cells["E161"].Value = total_IV_hisse_12b["00"];
                worksheet1.Cells["F161"].Value = total_IV_hisse_12b["01"];
                worksheet1.Cells["G161"].Value = total_IV_hisse_12b["02"];
                worksheet1.Cells["I161"].Value = total_IV_hisse_12b["03"];
                worksheet1.Cells["K161"].Value = total_IV_hisse_12b["04"];
                worksheet1.Cells["O161"].Value = total_IV_hisse_12b["05"];

                var Setir_IV_hisse_1_12b = _dt_odenisler.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "dunen" && row.Field<string>(3) == "fiziki" && row.Field<string>(5) == "bos")
                    )
                .ToList();
                decimal total_IV_hisse_1_12b = Setir_IV_hisse_1_12b.Sum(row => row.Field<decimal>(2)) / 1000;
                worksheet1.Cells["C161"].Value = total_IV_hisse_1_12b;

                worksheet1.Cells["C140"].Value = Math.Round(Convert.ToDecimal(lcrcemd) * 100, 2);
                worksheet1.Cells["D140"].Value = Math.Round(Convert.ToDecimal(lcrcem) * 100, 2);
                worksheet1.Cells["E140"].Value = Math.Round(Convert.ToDecimal(lcrazn) * 100, 2);
                worksheet1.Cells["F140"].Value = Math.Round(Convert.ToDecimal(lcrval) * 100, 2);
                //worksheet1.Cells["D140"].Value = Math.Round(Convert.ToDecimal(lcr4), 2);
                //worksheet1.Cells["D140"].Value = Math.Round(Convert.ToDecimal(lcr5), 2);

                //Daily_Credit_Deposit*******************

                var total_Setir_DM1_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "fiziki")
                    )
                .ToList();
                decimal total_IV_hisse_1_1_depo = total_Setir_DM1_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM1_depo = total_Setir_DM1_depo.Sum(row => row.Field<decimal>(4));


                worksheet3.Cells["C25"].Value = total_IV_hisse_1_1_depo;
                worksheet3.Cells["C26"].Value = total_IV_hisse_1_1_depo;
                worksheet3.Cells["D25"].Value = say_DM1_depo;
                worksheet3.Cells["D26"].Value = say_DM1_depo;

                var Setir_DM2_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "huquqi")
                    )
                .ToList();
                decimal total_Setir_DM2_depo = Setir_DM2_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM2_depo = Setir_DM2_depo.Sum(row => row.Field<decimal>(4));

                worksheet3.Cells["C30"].Value = total_Setir_DM2_depo;
                worksheet3.Cells["C31"].Value = total_Setir_DM2_depo;
                worksheet3.Cells["D30"].Value = say_DM2_depo;
                worksheet3.Cells["D31"].Value = say_DM2_depo;

                var Setir_DM3_1_depo = _dt_medaxil.AsEnumerable()
                                        .Where(row => (row.Field<string>(0) == "bugun" && row.Field<string>(1) == "sahibkar")
                    )
                .ToList();
                decimal total_Setir_DM3_1_depo = Setir_DM3_1_depo.Sum(row => row.Field<decimal>(2)) / 1000;
                decimal say_DM3_1_depo = Setir_DM3_1_depo.Sum(row => row.Field<decimal>(4));

                worksheet3.Cells["C35"].Value = total_Setir_DM3_1_depo;
                worksheet3.Cells["C36"].Value = total_Setir_DM3_1_depo;
                worksheet3.Cells["D35"].Value = say_DM3_1_depo;
                worksheet3.Cells["D36"].Value = say_DM3_1_depo;

                worksheet1.Cells["A4"].Value = "Bank Melli İran Bakı filialı";
                worksheet1.Cells["A3"].Value = "Tarix:" + txtdtbugun.Text;
                #endregion

                filePath = System.IO.Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                System.Diagnostics.Process.Start(filePath);
            }
        }
        #region LCR dan melumata gore
        private void excel()
        {


            DataTable dt_L2 = new DataTable();
            DataTable dt_proqnoz = new DataTable();
            DataTable dt_xett = new DataTable();
            DataTable dt_likvid = new DataTable();
            DataTable dt_qiym_kag = new DataTable();
            DataTable dt_akk_qarant = new DataTable();
            DataTable dt_99531 = new DataTable();

            #region sql_kodlar


            string LCR = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "SUBSTR(ar.licsch, 6, 2) AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string Proqnoz = "select distinct n.licschpkre,n.val,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez,n.gec_gun," +
                "n.meb*ROUND(odb.func_get_kurval(substr(n.licschpkre,6,2),TO_DATE('" + txtdtbugun.Text + "', 'DD-MM-YYYY')),6) ekv,n.faiz,n.mud," +
                "'' ay,asz.odenis,n.tip from " +
      "(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, substr(m.licschpkre, 6, 2) val, m.meb, m.faiz, m.mud, m.tip, count(*) kol from " +
      "(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19, s.setmininterestreserves min_rez, " +
      "odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, t.summakre meb, t.procstavkre faiz, t.srok mud, t.tipkredita tip " +
      "from view_nacpogprokre_all x, odb.arh_licschkre t, odb.srokpogprockre s, tipkre g " +
      "where t.tipkredita = g.code and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null " +
      "and x.date_oper = to_date('" + txtdtbugun.Text + "', 'dd/mm/yyyy') and x.date_oper = t.date_oper " +
      "order by(x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)), x.date_oper asc) m " +
      "group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, m.meb, m.faiz, m.mud, m.tip " +
      "order by m.gec_gun) n, (SELECT distinct substr(ar.kredit, 10, 6) qeyd, ar.ssk, sum(ar.summa_v_nacval) odenis " +
      "                    FROM regnom rr, licschkre l " +
      "                    JOIN arh_dd ar ON EXTRACT(MONTH FROM ar.date_oper) = EXTRACT(MONTH FROM TO_DATE('" + txtdtbugun.Text + "', 'DD-MM-YYYY')) " +
      "                    AND EXTRACT(YEAR FROM ar.date_oper) = EXTRACT(YEAR FROM TO_DATE('" + txtdtbugun.Text + "', 'DD-MM-YYYY')) " +
      "                    WHERE substr(l.licschkre, 10, 6) = rr.regnom and(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
      "                    AND ar.ssk = l.subschkre AND " +
      "                    ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
      "                     group by substr(ar.kredit, 10, 6),ar.ssk) asz where substr(n.licschpkre, 10, 6) = asz.qeyd(+)and n.sk = asz.ssk(+)";

            string daily_report_bk_xett = "select distinct t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish, " +
                "t.ostatok_ish* ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv, " +
                "ROUND(odb.func_get_kurval(substr(t.licsch, 6, 2), t.date_oper), 6)  kurs ,ar.date_planclose," +
                "odb.tar_ferq360(ar.date_planclose,t.date_oper) gun_ferqi," +
                "ar.tipkredita " +
                "from odb.arh_saldo_vbls t, arh_licschkre ar,tipkre g where ar.tipkredita = g.code " +
                "and t.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99540,99550,99531) and t.ostatok_ish > 0 " +
                "and substr(t.licsch,10,6)= substr(ar.licschkre, 10, 6) and t.ssls = ar.subschkre and t.date_oper=ar.date_oper and t.licsch is not null";

            string likvid = "select l.licsch,substr(l.licsch,6,2),round(sum(l.saldo_ish_nacval/1000),2) from odb.arh_saldo_ls l " +
                   "where l.date_oper=TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') " +
                   " and substr(l.licsch,1,5) in ('15770','11710') " +
                   " group by l.licsch";

            string qiymetli_kag = "select substr(t.licsch_cb,0,5)hes,substr(t.licsch_cb,6,2) val,t.subsch_cb,t.summa_cb," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licsch_cb,6,2),t.date_oper),6) ekv,t.diskont," +
                "odb.tar_ferq360(t.date_planclose,t.date_oper) gun_ferqi from odb.arh_licsch_cb t " +
                "where t.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy') and t.summa > 0 order by t.licsch_cb,t.subsch_cb";

            string akk_qarant = "select t.vbs bk,substr(t.licschgar,6,2)val," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licschgar,6,2),t.date_oper),6) ekv ," +
               " case when t.date_prolong is null then odb.tar_ferq360(t.date_planclose, t.date_oper) " +
               "else odb.tar_ferq360(t.date_prolong, t.date_oper) end gun_ferqi " +
               " from arh_licschgar t where t.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy')and t.summa > 0";

            string birterefli_99531 = "select * from odb.arh_saldo_vbls t where t.ostatok_ish > 0 and t.vbs='99531' " +
                "and t.date_oper = TO_DATE('" + txtdtbugun.Text + "', 'dd/mm/yyyy')";
            #endregion

            using (OracleConnection connection = new OracleConnection(cl.con))

            {
                using (OracleCommand command = new OracleCommand(LCR, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_L2);
                }
                using (OracleCommand command = new OracleCommand(Proqnoz, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_proqnoz);
                    foreach (DataRow row in dt_proqnoz.Rows)
                    {
                        double meb = Convert.ToDouble(row[7]);
                        double faizOranı = Convert.ToDouble(row[8]);
                        int vadeMüddeti = Convert.ToInt32(row[9]) / 30;

                        double aylıkÖdeme = Math.Round(CalculateMonthlyPayment(meb, vadeMüddeti, faizOranı), 2);
                        row[10] = aylıkÖdeme;
                    }



                }
                using (OracleCommand command = new OracleCommand(daily_report_bk_xett, connection))
                {
                    dt_xett.Clear();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_xett);
                    connection.Close();

                    //gridControl1.DataSource = dt_xett;
                }

                using (OracleCommand command = new OracleCommand(likvid, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_likvid);

                }

                using (OracleCommand command = new OracleCommand(qiymetli_kag, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_qiym_kag);

                }

                using (OracleCommand command = new OracleCommand(akk_qarant, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_akk_qarant);
                }
                using (OracleCommand command = new OracleCommand(birterefli_99531, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_99531);
                    connection.Close();
                }

            }

            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txtdtbugun.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "LCR_1124m" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "LCR_.xlsm");
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);
            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }

            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar


                ExcelWorksheet wsL1 = package.Workbook.Worksheets["L1"];
                ExcelWorksheet wsL2 = package.Workbook.Worksheets["L2"];
                ExcelWorksheet wsL3_A = package.Workbook.Worksheets["L3 (A)"];
                ExcelWorksheet wsL3_B = package.Workbook.Worksheets["L3 (B)"];
                ExcelWorksheet L4 = package.Workbook.Worksheets["L4"];

                string[] L2_c15 = { "100" };
                string[] L2_c16 = { "110" };
                string[] L2_c16_ist = { "11010000010000200000", "11020020010000200000" };
                string[] L2_c17 = { "14010", "14014", "14030", "14034" };
                string[] L2_f16 = { "11010000040000200000" };

                string[] L3A_c21 = { "410", "419" };
                string[] L3A_f24 = { "411", "412" };
                string[] L3A_c36 = { "35015", "35020", "35025", "35026" };
                string[] L3A_c37 = { "49025" };
                string[] L3A_c39 = { "35090", "35190" };
                string[] L3A_c42 = { "35938" };
                string[] L3A_c42_qisa = { "38", "39", "40" };

                //L2
                var sh_L2_c15 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c15 = sh_L2_c15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_d15 = dt_L2.AsEnumerable()
                .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                .ToList();
                decimal total_sh_L2_d15 = sh_L2_d15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c16 = dt_L2.AsEnumerable()
                    .Where(row => L2_c16.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00"
                     && !L2_c16_ist.Contains(row.Field<string>(1)))
                    .ToList();
                decimal total_sh_L2_c16 = sh_L2_c16.Sum(row => row.Field<decimal>(3)) / 1000;


                var sh_L2_d16 = dt_L2.AsEnumerable()
                .Where(row => L2_c16.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00"
                 && !L2_c16_ist.Contains(row.Field<string>(1)))
                .ToList();
                decimal total_sh_L2_d16 = sh_L2_d16.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c17 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c17.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c17 = sh_L2_c17.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_f16 = dt_L2.AsEnumerable()
                                    .Where(row => L2_f16.Contains(row.Field<string>(1)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_f16 = sh_L2_f16.Sum(row => row.Field<decimal>(3)) / 1000;

                //L3A

                var sh_L3A_c21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c21 = sh_L3A_c21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d21 = sh_L3A_d21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_f24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_f24 = sh_L3A_f24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d24 = sh_L3A_d24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c36 = sh_L3A_c36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d36 = sh_L3A_d36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c37 = sh_L3A_c37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d37 = sh_L3A_d37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c39 = sh_L3A_c39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d39 = sh_L3A_d39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c42 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                                    .ToList();
                decimal total_sh_L3A_c42 = sh_L3A_c42.Sum(row => row.Field<decimal>(3)) / 1000;

                DataTable filteredDataTable = dt_L2.Clone(); // İlk tablonun şemasını kopyala
                //foreach (var row in sh_L3A_c42)
                //{
                //    filteredDataTable.Rows.Add(row.ItemArray);
                //}

                //// DataGridView'e yeni DataTable'ı atayarak güncelle
                //gridControl1.DataSource = filteredDataTable;

                var sh_L3A_d42 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                .ToList();

                decimal total_sh_L3A_d42 = sh_L3A_d42.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_c43 = sh_L3A_c43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_d43 = sh_L3A_d43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_c44 = sh_L3A_c44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_d44 = sh_L3A_d44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c45 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                                    .ToList();
                decimal total_sh_L3A_c45 = sh_L3A_c45.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d45 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                .ToList();

                decimal total_sh_L3A_d45 = sh_L3A_d45.Sum(row => row.Field<decimal>(3)) / 1000;

                //L3 B PROQNOZ

                var sh_L3B_c19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) <= 90 && row.Field<string>(1) == "00")
                      .ToList();

                decimal total_sh_L3B_c19 = sh_L3B_c19.Sum(row => Convert.ToDecimal(row.Field<string>(10))) / 1000;

                var sh_L3B_c19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) > 90 && row.Field<string>(1) == "00")
                      .ToList();

                //double total_sh_L3B_c19_90_cox = sh_L3B_c19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;
                decimal total_sh_L3B_c19_90_cox = sh_L3B_c19
                .Where(row => row.Field<decimal?>(11) != null)
                .Sum(row => row.Field<decimal>(11)) / 1000;


                var sh_L3B_d19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) <= 90 && row.Field<string>(1) != "00")
                      .ToList();

                decimal total_sh_L3B_d19 = sh_L3B_d19.Sum(row => Convert.ToDecimal(row.Field<string>(10))) / 1000;

                var sh_L3B_d19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) > 90 && row.Field<string>(1) != "00")
                      .ToList();

                //double total_sh_L3B_d19_90_cox = sh_L3B_d19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;

                decimal total_sh_L3B_d19_90_cox = sh_L3B_d19
                .Where(row => row.Field<decimal?>(11) != null)
                .Sum(row => row.Field<decimal>(11)) / 1000;

                //QIYMETLI KAGIZLAR

                var sh_L3B_c33 = dt_likvid.AsEnumerable()
                     .Where(row => row.Field<string>(1) == "00")
                     .ToList();
                decimal total_sh_L3B_c33 = sh_L3B_c33.Sum(row => row.Field<decimal>(2));

                //var sh_L3B_c49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) == "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_c49 = sh_L3B_c49.Sum(row => row.Field<decimal>(4)) / 1000;

                //var sh_L3B_d49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) != "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_d49 = sh_L3B_d49.Sum(row => row.Field<decimal>(4)) / 1000;

                //Akkreditiv 30 gune
                var sh_L3A_c76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_c76_30 = sh_L3A_c76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_d76_30 = sh_L3A_d76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Akkreditiv 30 gunden cox
                var sh_L3A_f76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_f76_90 = sh_L3A_f76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_g76_90 = sh_L3A_g76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gune
                var sh_L3A_c77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_c77_30 = sh_L3A_c77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_d77_30 = sh_L3A_d77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gunden cox
                var sh_L3A_f77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_30 = sh_L3A_f77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_30 = sh_L3A_g77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_f77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_90 = sh_L3A_f77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_90 = sh_L3A_g77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //XETLER 30 79
                var sh_L3A_c79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                        && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                double total_sh_L3A_c79_30 = sh_L3A_c79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                var sh_L3A_d79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                    .ToList();


                double total_sh_L3A_d79_30 = sh_L3A_d79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                //XETLER 30 80
                var sh_L3A_c80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                      && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_c80_30 = sh_L3A_c80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_d80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_d80_30 = sh_L3A_d80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 79
                var sh_L3A_f79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) == "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_f79_30 = sh_L3A_f79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) != "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_g79_30 = sh_L3A_g79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 80
                var sh_L3A_f80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_f80_30 = sh_L3A_f80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_g80_30 = sh_L3A_g80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //Birtərəfli ləğv etmə
                var sh_L3A_f84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) == "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_f84 = sh_L3A_f84.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_g84 = sh_L3A_g84.Sum(row => row.Field<decimal>(6)) / 1000;

                wsL1.Cells[9, 3].Value = txtdtbugun.Text;

                wsL2.Cells[15, 3].Value = total_sh_L2_c15;
                wsL2.Cells[15, 4].Value = total_sh_L2_d15;

                wsL2.Cells[16, 3].Value = total_sh_L2_c16;
                wsL2.Cells[16, 4].Value = total_sh_L2_d16;
                wsL2.Cells[16, 6].Value = total_sh_L2_f16;

                wsL2.Cells[17, 3].Value = total_sh_L2_c17;

                wsL3_A.Cells[21, 3].Value = -total_sh_L3A_c21;
                wsL3_A.Cells[21, 4].Value = -total_sh_L3A_d21;

                wsL3_A.Cells[24, 3].Value = -total_sh_L3A_f24;
                wsL3_A.Cells[24, 4].Value = -total_sh_L3A_f24;

                wsL3_A.Cells[36, 3].Value = -total_sh_L3A_c36;
                wsL3_A.Cells[36, 4].Value = -total_sh_L3A_d36;

                wsL3_A.Cells[37, 3].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[37, 4].Value = -total_sh_L3A_d37;

                wsL3_A.Cells[39, 3].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[39, 4].Value = -total_sh_L3A_d39;

                wsL3_A.Cells[42, 3].Value = -total_sh_L3A_c42;
                wsL3_A.Cells[42, 4].Value = -total_sh_L3A_d42;

                wsL3_A.Cells[43, 3].Value = -total_sh_L3A_c43;
                wsL3_A.Cells[43, 4].Value = -total_sh_L3A_d43;

                wsL3_A.Cells[44, 3].Value = -total_sh_L3A_c44;
                wsL3_A.Cells[44, 4].Value = -total_sh_L3A_d44;

                wsL3_A.Cells[45, 3].Value = -total_sh_L3A_c45;
                wsL3_A.Cells[45, 4].Value = -total_sh_L3A_d45;

                wsL3_B.Cells[19, 3].Value = total_sh_L3B_c19 + total_sh_L3B_c19_90_cox;
                wsL3_B.Cells[19, 4].Value = total_sh_L3B_d19;

                wsL3_A.Cells[79, 3].Value = total_sh_L3A_c79_30;
                wsL3_A.Cells[79, 4].Value = total_sh_L3A_d79_30;

                wsL3_A.Cells[80, 3].Value = total_sh_L3A_c80_30;
                wsL3_A.Cells[80, 4].Value = total_sh_L3A_d80_30;

                wsL3_A.Cells[79, 6].Value = total_sh_L3A_f79_30;
                wsL3_A.Cells[79, 7].Value = total_sh_L3A_g79_30;

                wsL3_A.Cells[80, 6].Value = total_sh_L3A_f80_30;
                wsL3_A.Cells[80, 7].Value = total_sh_L3A_g80_30;

                wsL3_B.Cells[37, 3].Value = total_sh_L3B_c33;//hazirlamaq
                wsL3_B.Cells[37, 4].Value = 0;

                //wsL3_B.Cells[49, 3].Value = total_sh_L3B_c49;
                //wsL3_B.Cells[49, 4].Value = total_sh_L3B_d49;

                wsL3_A.Cells[76, 3].Value = total_sh_L3A_c76_30;
                wsL3_A.Cells[76, 4].Value = total_sh_L3A_d76_30;

                wsL3_A.Cells[76, 6].Value = total_sh_L3A_f76_90;
                wsL3_A.Cells[76, 7].Value = total_sh_L3A_f76_90;

                wsL3_A.Cells[77, 3].Value = total_sh_L3A_c77_30;
                wsL3_A.Cells[77, 4].Value = total_sh_L3A_d77_30;

                wsL3_A.Cells[77, 6].Value = total_sh_L3A_f77_90;
                wsL3_A.Cells[77, 7].Value = total_sh_L3A_g77_90;

                wsL3_A.Cells[84, 6].Value = total_sh_L3A_f84;
                wsL3_A.Cells[84, 7].Value = total_sh_L3A_g84;

                object formula_A6 = L4.Cells["D7"].Formula;
                L4.Calculate();
                lcrcem = L4.Cells["C6"].Value.ToString();
                lcrazn = L4.Cells["D6"].Value.ToString();
                lcrval = L4.Cells["E6"].Value.ToString();
                lcr4 = L4.Cells["C7"].Value.ToString();
                lcr5 = L4.Cells["C10"].Value.ToString();

                #endregion

                filePath = System.IO.Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                //System.Diagnostics.Process.Start(filePath);
            }
        }
        #endregion


        private void txtdtdunen_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                button1.Focus();
                e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
            }
        }

        private void txtdtbugun_KeyDown(object sender, KeyEventArgs e)
        {
                if (e.KeyCode == Keys.Enter)
                {
                    // Enter tuşuna basıldığında ikinci TextBox'a odaklan
                    txtdtdunen.Focus();
                    e.SuppressKeyPress = true; // Enter tuşunun varsayılan işlevini engelle
                }
        }

        private void txtdtdunen_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txtdtdunen.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txtdtdunen.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }

        private void excel_dunen()
        {


            DataTable dt_L2 = new DataTable();
            DataTable dt_proqnoz = new DataTable();
            DataTable dt_xett = new DataTable();
            DataTable dt_likvid = new DataTable();
            DataTable dt_qiym_kag = new DataTable();
            DataTable dt_akk_qarant = new DataTable();
            DataTable dt_99531 = new DataTable();

            #region sql_kodlar

            
            string LCR = "SELECT ar.date_oper AS tarix, ar.licsch AS hesab, " +
               "SUBSTR(ar.licsch, 6, 2) AS valyuta, " +
               "ar.saldo_ish_nacval AS qaliq " +
               "FROM odb.arh_saldo_ls ar, licsch ch " +
               "WHERE ar.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') " +
               "AND ch.licsch = ar.licsch " +
               "AND (ch.date_close_licsch IS NULL OR ar.date_oper <= ch.date_close_licsch)";

            string Proqnoz = "select distinct n.licschpkre,n.val,n.sk,n.procstavrez,n.procstavrez_19,n.min_rez,n.gec_gun," +
                "n.meb*ROUND(odb.func_get_kurval(substr(n.licschpkre,6,2),TO_DATE('" + txtdtdunen.Text + "', 'DD-MM-YYYY')),6) ekv,n.faiz,n.mud," +
                "'' ay,asz.odenis,n.tip from " +
      "(select m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, substr(m.licschpkre, 6, 2) val, m.meb, m.faiz, m.mud, m.tip, count(*) kol from " +
      "(select x.date_oper, x.licschpkre, x.subschkre sk, t.procstavrez, t.procstavrez_19, s.setmininterestreserves min_rez, " +
      "odb.tar_ferq360(x.date_oper, nvl(x.lastoverduedate, x.date_oper)) gec_gun, t.summakre meb, t.procstavkre faiz, t.srok mud, t.tipkredita tip " +
      "from view_nacpogprokre_all x, odb.arh_licschkre t, odb.srokpogprockre s, tipkre g " +
      "where t.tipkredita = g.code and x.licschpkre = t.licschpkre and x.subschkre = t.subschkre and t.licschkre = s.licschkre and x.subschkre = s.subschkre and t.date_close is null " +
      "and x.date_oper = to_date('" + txtdtdunen.Text + "', 'dd/mm/yyyy') and x.date_oper = t.date_oper " +
      "order by(x.date_oper - nvl(func_get_overdue_min_date(x.lastoverduedate_ish, x.lastoverduedate, x.lodinterest_ish), x.date_oper)), x.date_oper asc) m " +
      "group by m.licschpkre, m.sk, m.procstavrez, m.procstavrez_19, m.min_rez, m.gec_gun, m.meb, m.faiz, m.mud, m.tip " +
      "order by m.gec_gun) n, (SELECT distinct substr(ar.kredit, 10, 6) qeyd, ar.ssk, sum(ar.summa_v_nacval) odenis " +
      "                    FROM regnom rr, licschkre l " +
      "                    JOIN arh_dd ar ON EXTRACT(MONTH FROM ar.date_oper) = EXTRACT(MONTH FROM TO_DATE('" + txtdtdunen.Text + "', 'DD-MM-YYYY')) " +
      "                    AND EXTRACT(YEAR FROM ar.date_oper) = EXTRACT(YEAR FROM TO_DATE('" + txtdtdunen.Text + "', 'DD-MM-YYYY')) " +
      "                    WHERE substr(l.licschkre, 10, 6) = rr.regnom and(ar.kredit = l.licschkre OR ar.kredit = l.licschpkre OR ar.kredit = l.licsch_19 OR ar.kredit = l.licschppkre) " +
      "                    AND ar.ssk = l.subschkre AND " +
      "                    ((substr(ar.debet, 0, 1) in (3, 4) and ar.kredit in (l.licschkre, l.licschpkre, l.licsch_19, l.licschppkre))) " +
      "                     group by substr(ar.kredit, 10, 6),ar.ssk) asz where substr(n.licschpkre, 10, 6) = asz.qeyd(+)and n.sk = asz.ssk(+)";

            string daily_report_bk_xett = "select distinct t.date_oper tarix,t.vbs,t.licsch,substr(t.licsch,6,2),t.ssls,t.ostatok_ish, " +
                "t.ostatok_ish* ROUND(odb.func_get_kurval(substr(t.licsch,6,2),t.date_oper),6) ekv, " +
                "ROUND(odb.func_get_kurval(substr(t.licsch, 6, 2), t.date_oper), 6)  kurs ,ar.date_planclose," +
                "odb.tar_ferq360(ar.date_planclose,t.date_oper) gun_ferqi," +
                "ar.tipkredita " +
                "from odb.arh_saldo_vbls t, arh_licschkre ar,tipkre g where ar.tipkredita = g.code " +
                "and t.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') and t.vbs in (99530,99540,99550,99531) and t.ostatok_ish > 0 " +
                "and substr(t.licsch,10,6)= substr(ar.licschkre, 10, 6) and t.ssls = ar.subschkre and t.date_oper=ar.date_oper and t.licsch is not null";

            string likvid = "select l.licsch,substr(l.licsch,6,2),round(sum(l.saldo_ish_nacval/1000),2) from odb.arh_saldo_ls l " +
                   "where l.date_oper=TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') " +
                   " and substr(l.licsch,1,5) in ('15770','11710') " +
                   " group by l.licsch";

            string qiymetli_kag = "select substr(t.licsch_cb,0,5)hes,substr(t.licsch_cb,6,2) val,t.subsch_cb,t.summa_cb," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licsch_cb,6,2),t.date_oper),6) ekv,t.diskont," +
                "odb.tar_ferq360(t.date_planclose,t.date_oper) gun_ferqi from odb.arh_licsch_cb t " +
                "where t.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy') and t.summa > 0 order by t.licsch_cb,t.subsch_cb";

            string akk_qarant = "select t.vbs bk,substr(t.licschgar,6,2)val," +
                "t.summa * ROUND(odb.func_get_kurval(substr(t.licschgar,6,2),t.date_oper),6) ekv ," +
               " case when t.date_prolong is null then odb.tar_ferq360(t.date_planclose, t.date_oper) " +
               "else odb.tar_ferq360(t.date_prolong, t.date_oper) end gun_ferqi " +
               " from arh_licschgar t where t.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')and t.summa > 0";

            string birterefli_99531 = "select * from odb.arh_saldo_vbls t where t.ostatok_ish > 0 and t.vbs='99531' " +
                "and t.date_oper = TO_DATE('" + txtdtdunen.Text + "', 'dd/mm/yyyy')";

            #endregion
            using (OracleConnection connection = new OracleConnection(cl.con))

            {
                using (OracleCommand command = new OracleCommand(LCR, connection))
                {
                    connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_L2);
                }
                using (OracleCommand command = new OracleCommand(Proqnoz, connection))
                {
                    //connection.Open();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_proqnoz);
                    foreach (DataRow row in dt_proqnoz.Rows)
                    {
                        double meb = Convert.ToDouble(row[7]);
                        double faizOranı = Convert.ToDouble(row[8]);
                        int vadeMüddeti = Convert.ToInt32(row[9]) / 30;

                        double aylıkÖdeme = Math.Round(CalculateMonthlyPayment(meb, vadeMüddeti, faizOranı), 2);
                        row[10] = aylıkÖdeme;
                    }



                }
                using (OracleCommand command = new OracleCommand(daily_report_bk_xett, connection))
                {
                    dt_xett.Clear();
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_xett);
                    connection.Close();

                    //gridControl1.DataSource = dt_xett;
                }

                using (OracleCommand command = new OracleCommand(likvid, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_likvid);

                }

                using (OracleCommand command = new OracleCommand(qiymetli_kag, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_qiym_kag);

                }

                using (OracleCommand command = new OracleCommand(akk_qarant, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_akk_qarant);
                }
                using (OracleCommand command = new OracleCommand(birterefli_99531, connection))
                {
                    OracleDataAdapter adapter = new OracleDataAdapter(command);
                    adapter.Fill(dt_99531);
                    connection.Close();
                }

            }

            string dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");
            string textBoxText = txtdtbugun.Text; // TextBox'tan alınan metni sakla
            string yeniMetin = textBoxText.Replace("-", "");
            string baseFileName = "LCR_1124m" + yeniMetin; // Temel dosya adı
            string fileName = baseFileName + ".xlsm";
            string templateFilePath = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Muhasibat", "Exceller", "LCR_.xlsm");
            string filePath = System.IO.Path.Combine(dosyayolu, fileName);
            if (File.Exists(System.IO.Path.Combine(dosyayolu, fileName)))
            {
                int fileCounter = 1;
                while (File.Exists(System.IO.Path.Combine(dosyayolu, $"{baseFileName} - {fileCounter}.xlsm")))
                {
                    fileCounter++;
                }
                fileName = $"{baseFileName} - {fileCounter}.xlsm";
            }

            FileInfo templateFile = new FileInfo(templateFilePath);
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                #region excel_kodlar

                
                ExcelWorksheet wsL1 = package.Workbook.Worksheets["L1"];
                ExcelWorksheet wsL2 = package.Workbook.Worksheets["L2"];
                ExcelWorksheet wsL3_A = package.Workbook.Worksheets["L3 (A)"];
                ExcelWorksheet wsL3_B = package.Workbook.Worksheets["L3 (B)"];
                ExcelWorksheet L4 = package.Workbook.Worksheets["L4"];

                string[] L2_c15 = { "100" };
                string[] L2_c16 = { "110" };
                string[] L2_c16_ist = { "11010000010000200000", "11020020010000200000" };
                string[] L2_c17 = { "14010", "14014", "14030", "14034" };
                string[] L2_f16 = { "11010000040000200000" };

                string[] L3A_c21 = { "410", "419" };
                string[] L3A_f24 = { "411", "412" };
                string[] L3A_c36 = { "35015", "35020", "35025", "35026" };
                string[] L3A_c37 = { "49025" };
                string[] L3A_c39 = { "35090", "35190" };
                string[] L3A_c42 = { "35938" };
                string[] L3A_c42_qisa = { "38", "39", "40" };

                //L2
                var sh_L2_c15 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c15 = sh_L2_c15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_d15 = dt_L2.AsEnumerable()
                .Where(row => L2_c15.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                .ToList();
                decimal total_sh_L2_d15 = sh_L2_d15.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c16 = dt_L2.AsEnumerable()
                    .Where(row => L2_c16.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00"
                     && !L2_c16_ist.Contains(row.Field<string>(1)))
                    .ToList();
                decimal total_sh_L2_c16 = sh_L2_c16.Sum(row => row.Field<decimal>(3)) / 1000;


                var sh_L2_d16 = dt_L2.AsEnumerable()
                .Where(row => L2_c16.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00"
                 && !L2_c16_ist.Contains(row.Field<string>(1)))
                .ToList();
                decimal total_sh_L2_d16 = sh_L2_d16.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_c17 = dt_L2.AsEnumerable()
                                    .Where(row => L2_c17.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_c17 = sh_L2_c17.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L2_f16 = dt_L2.AsEnumerable()
                                    .Where(row => L2_f16.Contains(row.Field<string>(1)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L2_f16 = sh_L2_f16.Sum(row => row.Field<decimal>(3)) / 1000;
                
                //L3A

                var sh_L3A_c21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c21 = sh_L3A_c21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d21 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c21.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d21 = sh_L3A_d21.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_f24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_f24 = sh_L3A_f24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d24 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_f24.Contains(row.Field<string>(1).Substring(0, 3)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d24 = sh_L3A_d24.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c36 = sh_L3A_c36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d36 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c36.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d36 = sh_L3A_d36.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c37 = sh_L3A_c37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d37 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c37.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d37 = sh_L3A_d37.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) == "00")
                                    .ToList();
                decimal total_sh_L3A_c39 = sh_L3A_c39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d39 = dt_L2.AsEnumerable()
                                    .Where(row => L3A_c39.Contains(row.Field<string>(1).Substring(0, 5)) && row.Field<string>(2) != "00")
                                    .ToList();
                decimal total_sh_L3A_d39 = sh_L3A_d39.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c42 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                                    .ToList();
                decimal total_sh_L3A_c42 = sh_L3A_c42.Sum(row => row.Field<decimal>(3)) / 1000;

                DataTable filteredDataTable = dt_L2.Clone(); // İlk tablonun şemasını kopyala
                //foreach (var row in sh_L3A_c42)
                //{
                //    filteredDataTable.Rows.Add(row.ItemArray);
                //}

                //// DataGridView'e yeni DataTable'ı atayarak güncelle
                //gridControl1.DataSource = filteredDataTable;

                var sh_L3A_d42 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) <= 100000)
                .ToList();

                decimal total_sh_L3A_d42 = sh_L3A_d42.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_c43 = sh_L3A_c43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d43 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 100001 && Math.Abs(row.Field<decimal>(3)) <= 500000)
                .ToList();
                decimal total_sh_L3A_d43 = sh_L3A_d43.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_c44 = sh_L3A_c44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d44 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) >= 500001 && Math.Abs(row.Field<decimal>(3)) <= 1000000)
                .ToList();
                decimal total_sh_L3A_d44 = sh_L3A_d44.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_c45 = dt_L2.AsEnumerable()
                                    .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                                     && row.Field<string>(2) == "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                                    .ToList();
                decimal total_sh_L3A_c45 = sh_L3A_c45.Sum(row => row.Field<decimal>(3)) / 1000;

                var sh_L3A_d45 = dt_L2.AsEnumerable()
                .Where(row => (L3A_c42.Contains(row.Field<string>(1).Substring(0, 5)) || L3A_c42_qisa.Contains(row.Field<string>(1).Substring(0, 2)))
                 && row.Field<string>(2) != "00" && Math.Abs(row.Field<decimal>(3)) > 1000000)
                .ToList();

                decimal total_sh_L3A_d45 = sh_L3A_d45.Sum(row => row.Field<decimal>(3)) / 1000;

                //L3 B PROQNOZ

                var sh_L3B_c19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) <= 90 && row.Field<string>(1) == "00")
                      .ToList();

                decimal total_sh_L3B_c19 = sh_L3B_c19.Sum(row => Convert.ToDecimal(row.Field<string>(10))) / 1000;

                var sh_L3B_c19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) > 90 && row.Field<string>(1) == "00")
                      .ToList();

                //double total_sh_L3B_c19_90_cox = sh_L3B_c19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;
                decimal total_sh_L3B_c19_90_cox = sh_L3B_c19
                .Where(row => row.Field<decimal?>(11) != null)
                .Sum(row => row.Field<decimal>(11)) / 1000;


                var sh_L3B_d19 = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) <= 90 && row.Field<string>(1) != "00")
                      .ToList();

                decimal total_sh_L3B_d19 = sh_L3B_d19.Sum(row => Convert.ToDecimal(row.Field<string>(10))) / 1000;

                var sh_L3B_d19_90_cox = dt_proqnoz.AsEnumerable()
                      .Where(row => row.Field<decimal>(6) > 90 && row.Field<string>(1) != "00")
                      .ToList();

                //double total_sh_L3B_d19_90_cox = sh_L3B_d19.Sum(row => Convert.ToDouble(row.Field<string>(11))) / 1000;

                decimal total_sh_L3B_d19_90_cox = sh_L3B_d19
                .Where(row => row.Field<decimal?>(11) != null)
                .Sum(row => row.Field<decimal>(11)) / 1000;

                //QIYMETLI KAGIZLAR

                var sh_L3B_c33 = dt_likvid.AsEnumerable()
                     .Where(row => row.Field<string>(1) == "00")
                     .ToList();
                decimal total_sh_L3B_c33 = sh_L3B_c33.Sum(row => row.Field<decimal>(2));

                //var sh_L3B_c49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) == "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_c49 = sh_L3B_c49.Sum(row => row.Field<decimal>(4)) / 1000;

                //var sh_L3B_d49 = dt_qiym_kag.AsEnumerable()
                //     .Where(row => Convert.ToInt32(row.Field<decimal>(6)) <= 30 && row.Field<string>(1) != "00"
                //      && (row.Field<string>(0) == "14010" || row.Field<string>(0) == "14030"))
                //     .ToList();
                //decimal total_sh_L3B_d49 = sh_L3B_d49.Sum(row => row.Field<decimal>(4)) / 1000;

                //Akkreditiv 30 gune
                var sh_L3A_c76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_c76_30 = sh_L3A_c76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d76_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_d76_30 = sh_L3A_d76_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Akkreditiv 30 gunden cox
                var sh_L3A_f76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_f76_90 = sh_L3A_f76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g76_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99540")
                      .ToList();

                decimal total_sh_L3A_g76_90 = sh_L3A_g76_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gune
                var sh_L3A_c77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_c77_30 = sh_L3A_c77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_d77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) <= 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_d77_30 = sh_L3A_d77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                //Qarantiya 30 gunden cox
                var sh_L3A_f77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_30 = sh_L3A_f77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_30 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_30 = sh_L3A_g77_30.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_f77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) == "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_f77_90 = sh_L3A_f77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                var sh_L3A_g77_90 = dt_akk_qarant.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(3)) > 30 && row.Field<string>(1) != "00"
                        && row.Field<string>(0) == "99550")
                      .ToList();

                decimal total_sh_L3A_g77_90 = sh_L3A_g77_90.Sum(row => row.Field<decimal>(2)) / 1000;

                //XETLER 30 79
                var sh_L3A_c79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                        && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                double total_sh_L3A_c79_30 = sh_L3A_c79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                var sh_L3A_d79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                    .ToList();


                double total_sh_L3A_d79_30 = sh_L3A_d79_30.Sum(row => Convert.ToDouble(row.Field<decimal>(6))) / 1000;

                //XETLER 30 80
                var sh_L3A_c80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                      && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_c80_30 = sh_L3A_c80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_d80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_d80_30 = sh_L3A_d80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 79
                var sh_L3A_f79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) == "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_f79_30 = sh_L3A_f79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g79_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) > 30 && row.Field<string>(3) != "00"
                      && row.Field<decimal>(10) == 2 && row.Field<string>(1) == "99530")
                      .ToList();

                decimal total_sh_L3A_g79_30 = sh_L3A_g79_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //XETLER 90 80
                var sh_L3A_f80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) == "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_f80_30 = sh_L3A_f80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g80_30 = dt_xett.AsEnumerable()
                      .Where(row => Convert.ToInt32(row.Field<decimal>(9)) <= 30 && row.Field<string>(3) != "00"
                       && row.Field<string>(1) == "99530" && (row.Field<decimal>(10) == 1 || row.Field<decimal>(10) == 3))
                      .ToList();

                decimal total_sh_L3A_g80_30 = sh_L3A_g80_30.Sum(row => row.Field<decimal>(6)) / 1000;

                //Birtərəfli ləğv etmə
                var sh_L3A_f84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) == "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_f84 = sh_L3A_f84.Sum(row => row.Field<decimal>(6)) / 1000;

                var sh_L3A_g84 = dt_xett.AsEnumerable()
                      .Where(row => row.Field<string>(3) != "00"
                        && row.Field<string>(1) == "99531")
                      .ToList();

                decimal total_sh_L3A_g84 = sh_L3A_g84.Sum(row => row.Field<decimal>(6)) / 1000;

                wsL1.Cells[9, 3].Value = txtdtbugun.Text;

                wsL2.Cells[15, 3].Value = total_sh_L2_c15;
                wsL2.Cells[15, 4].Value = total_sh_L2_d15;

                wsL2.Cells[16, 3].Value = total_sh_L2_c16;
                wsL2.Cells[16, 4].Value = total_sh_L2_d16;
                wsL2.Cells[16, 6].Value = total_sh_L2_f16;

                wsL2.Cells[17, 3].Value = total_sh_L2_c17;

                wsL3_A.Cells[21, 3].Value = -total_sh_L3A_c21;
                wsL3_A.Cells[21, 4].Value = -total_sh_L3A_d21;

                wsL3_A.Cells[24, 3].Value = -total_sh_L3A_f24;
                wsL3_A.Cells[24, 4].Value = -total_sh_L3A_f24;

                wsL3_A.Cells[36, 3].Value = -total_sh_L3A_c36;
                wsL3_A.Cells[36, 4].Value = -total_sh_L3A_d36;

                wsL3_A.Cells[37, 3].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[37, 4].Value = -total_sh_L3A_d37;

                wsL3_A.Cells[39, 3].Value = -total_sh_L3A_c37;
                wsL3_A.Cells[39, 4].Value = -total_sh_L3A_d39;

                wsL3_A.Cells[42, 3].Value = -total_sh_L3A_c42;
                wsL3_A.Cells[42, 4].Value = -total_sh_L3A_d42;

                wsL3_A.Cells[43, 3].Value = -total_sh_L3A_c43;
                wsL3_A.Cells[43, 4].Value = -total_sh_L3A_d43;

                wsL3_A.Cells[44, 3].Value = -total_sh_L3A_c44;
                wsL3_A.Cells[44, 4].Value = -total_sh_L3A_d44;

                wsL3_A.Cells[45, 3].Value = -total_sh_L3A_c45;
                wsL3_A.Cells[45, 4].Value = -total_sh_L3A_d45;

                wsL3_B.Cells[19, 3].Value = total_sh_L3B_c19 + total_sh_L3B_c19_90_cox;
                wsL3_B.Cells[19, 4].Value = total_sh_L3B_d19;

                wsL3_A.Cells[79, 3].Value = total_sh_L3A_c79_30;
                wsL3_A.Cells[79, 4].Value = total_sh_L3A_d79_30;

                wsL3_A.Cells[80, 3].Value = total_sh_L3A_c80_30;
                wsL3_A.Cells[80, 4].Value = total_sh_L3A_d80_30;

                wsL3_A.Cells[79, 6].Value = total_sh_L3A_f79_30;
                wsL3_A.Cells[79, 7].Value = total_sh_L3A_g79_30;

                wsL3_A.Cells[80, 6].Value = total_sh_L3A_f80_30;
                wsL3_A.Cells[80, 7].Value = total_sh_L3A_g80_30;

                wsL3_B.Cells[37, 3].Value = total_sh_L3B_c33;//hazirlamaq
                wsL3_B.Cells[37, 4].Value = 0;

                //wsL3_B.Cells[49, 3].Value = total_sh_L3B_c49;
                //wsL3_B.Cells[49, 4].Value = total_sh_L3B_d49;

                wsL3_A.Cells[76, 3].Value = total_sh_L3A_c76_30;
                wsL3_A.Cells[76, 4].Value = total_sh_L3A_d76_30;

                wsL3_A.Cells[76, 6].Value = total_sh_L3A_f76_90;
                wsL3_A.Cells[76, 7].Value = total_sh_L3A_f76_90;

                wsL3_A.Cells[77, 3].Value = total_sh_L3A_c77_30;
                wsL3_A.Cells[77, 4].Value = total_sh_L3A_d77_30;

                wsL3_A.Cells[77, 6].Value = total_sh_L3A_f77_90;
                wsL3_A.Cells[77, 7].Value = total_sh_L3A_g77_90;

                wsL3_A.Cells[84, 6].Value = total_sh_L3A_f84;
                wsL3_A.Cells[84, 7].Value = total_sh_L3A_g84;

                object formula_A6 = L4.Cells["C6"].Formula;
                L4.Calculate();
                lcrcemd = L4.Cells["C6"].Value.ToString();

                button1.Text = "Sorğu";

                #endregion
                filePath = System.IO.Path.Combine(dosyayolu, fileName);
                package.SaveAs(new FileInfo(filePath)); // Excel dosyasını kaydet
                //System.Diagnostics.Process.Start(filePath);
            }
        }
        public static double CalculateMonthlyPayment(double presentValue, double financingPeriod, double interestRatePerYear)
        {
            double monthlyInterestRate = interestRatePerYear / 1200;
            double x = Math.Pow(1 + monthlyInterestRate, financingPeriod);
            double monthlyPayment = (presentValue * monthlyInterestRate * x) / (x - 1);
            return monthlyPayment;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "Hazırlanır...";
            excel();
            excel_dunen();
            Excel_daily_comment_Yeni_son();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string girilenTarih = txtdtbugun.Text;

            // Eğer girilen tarih boşsa veya uzunluğu 8 değilse veya geçerli bir tarih değilse işlemi yapma
            if (!string.IsNullOrEmpty(girilenTarih) && girilenTarih.Length == 8)
            {
                // Tarihi "ddMMyyyy" formatından "dd-MM-yyyy" formatına çevirin
                DateTime tarih;
                if (DateTime.TryParseExact(girilenTarih, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out tarih))
                {
                    string yeniFormatliTarih = tarih.ToString("dd-MM-yyyy");
                    // TextBox2'ye yeni formatlı tarihi yazın
                    txtdtbugun.Text = yeniFormatliTarih;
                }
                else
                {
                    // Geçersiz tarih girişi uyarısı verebilirsiniz
                    //MessageBox.Show("Geçersiz tarih girişi.");
                    //textBox2.Clear(); // Hatalı girişi temizle
                }
            }
        }
    }
}
