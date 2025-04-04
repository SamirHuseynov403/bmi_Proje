using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Oracle.ManagedDataAccess.Client;

namespace BMI
{
    public partial class frmsmssiyahi : DevExpress.XtraEditors.XtraForm
    {
        public frmsmssiyahi()
        {
            InitializeComponent();
        }
        string ad;
        OracleConnection con = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
        OracleCommand komyazdir;
        OracleCommand komutmurraz;
        DataTable muracietler = new DataTable();
        DataTable muracietlerzamin = new DataTable();
        public void tamkataloq()
        {

            try
            {
                gridControl1.DataSource = null;
                gridmuracietler.Columns.Clear();
                 muracietler.Clear();
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int iltarix = Convert.ToInt32(tarixIl);
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
            OracleDataAdapter isleme = new OracleDataAdapter("select distinct z1.ad,z1.sk,z1.hs,z1.kr,z1.qal,z1.vkes," +
                "z2.vkfaiz,z1.it01,z1.tel,z1.mob,z1.sorgu from(select distinct r.name_regnom ad, lk.licschkre hs,lk.licschpkre hp," +
                "lk.subschkre sk, lk.summakre kr, lk.summa qal, lk.summa_19 vkes, lk.date_open,sr.item_01 it01,r.telefon tel," +
                " r.mobilniy mob,(case length(trim(translate(r.mobilniy,'(-)',' '))) when 9 then 'Duz mobil' when 10 then 'Duz mobil' else 'Ferqli mobil' end) as sorgu from odb.licschkre lk, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr," +
                " creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 " +
                "and  lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code " +
                "and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre " +
                "and lk.subschkre = ci.subschkre and lk.date_close is null ) z1,(select x.date_oper, x.licschpkre," +
                " x.subschkre, x.nacprospro_ish - x.pogprospro_ish vkfaiz, x.nacpro_ish - x.pogpro_ish ferq " +
                "from view_nacpogprokre_all x where x.date_oper = to_date(sysdate)) z2 where z1.hp = z2.licschpkre " +
                "and z1.sk = z2.subschkre and(z1.vkes > 0 or z2.vkfaiz > 0)", con);
            isleme.Fill(muracietler);
                gridControl1.DataSource = muracietler;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

            gridmuracietler.Columns[0].Caption = "Adı";
            gridmuracietler.Columns[0].Width = 200;
            gridmuracietler.Columns[1].Caption = "Ssuda";
            gridmuracietler.Columns[1].Width = 60;
            gridmuracietler.Columns[2].Caption = "Hesab";
            gridmuracietler.Columns[2].Width = 100;
            gridmuracietler.Columns[2].Visible = false;
            gridmuracietler.Columns[3].Caption = "Kr məbləğ";
            gridmuracietler.Columns[3].Width = 100;
            gridmuracietler.Columns[4].Caption = "Qalıq";
            gridmuracietler.Columns[4].Width = 100;
            gridmuracietler.Columns[5].Caption = "Vk qalıq";
            gridmuracietler.Columns[5].Width = 100;
            gridmuracietler.Columns[6].Caption = "Vk % qalıq";
            gridmuracietler.Columns[6].Width = 100;
            gridmuracietler.Columns[7].Caption = "İtem 1";
            gridmuracietler.Columns[7].Width = 100;
            gridmuracietler.Columns[8].Caption = "Telefonlar";
            gridmuracietler.Columns[8].Width = 350;
            gridmuracietler.Columns[9].Caption = "Mobil";
            gridmuracietler.Columns[9].Width = 150;
            gridmuracietler.Columns[10].Caption = "Sorğu";
            gridmuracietler.Columns[10].Width = 100;


            }
            catch (Exception)
            {


            }
            finally { };


        }
        public void tamkataloqzamin()
        {
            gridControl1.DataSource = null;
            gridmuracietler.Columns.Clear();
            muracietlerzamin.Clear();

            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            
            OracleDataAdapter isleme = new OracleDataAdapter("select distinct z1.ad,z1.sk,z1.hs,z1.it01,z1.zamad,z1.zamtel,z1.sorgu from(select distinct r.name_regnom ad, lk.licschkre hs,lk.licschpkre hp," +
                "lk.subschkre sk, lk.summakre kr, lk.summa qal, lk.summa_19 vkes, lk.date_open,sr.item_01 it01,r.telefon tel," +
                " trim(translate(r.mobilniy,'-',' ')) mob,(case length(trim(translate(g.telefon,'(-)',' '))) when 9 then 'Duz mobil' when 10 then 'Duz mobil' else 'Ferqli mobil' end) as sorgu,g.guarantee_name zamad,trim(translate(g.telefon,'(-)',' ')) zamtel from odb.licschkre lk,odb.creditinfoguarantee g, regnom r, kuratorkredita kk, tipzal g, tipkre tk, srokpogprockre sr," +
                " creditinfo ci where  substr(lk.licschkre, 10, 6) = r.regnom and length(lk.licschkre) = 20 " +
                "and  lk.kurator = kk.code(+) and lk.tipzaloga = g.code and lk.tipkredita = tk.code " +
                "and lk.licschkre = sr.licschkre and lk.subschkre = sr.subschkre and lk.licschkre = ci.licschkre " +
                "and lk.subschkre = ci.subschkre and lk.date_close is null and lk.licschkre = g.licschkre  and lk.subschkre = g.subschkre ) z1,(select x.date_oper, x.licschpkre," +
                " x.subschkre, x.nacprospro_ish - x.pogprospro_ish vkfaiz, x.nacpro_ish - x.pogpro_ish ferq " +
                "from view_nacpogprokre_all x where x.date_oper = to_date(sysdate)) z2 where z1.hp = z2.licschpkre " +
                "and z1.sk = z2.subschkre and(z1.vkes > 0 or z2.vkfaiz > 0)", con);
            isleme.Fill(muracietlerzamin);
            gridControl1.DataSource = muracietlerzamin;
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            gridmuracietler.Columns[0].Caption = "Adı";
            gridmuracietler.Columns[0].Width = 300;
            gridmuracietler.Columns[1].Caption = "Ssuda";
            gridmuracietler.Columns[1].Width = 100;
            gridmuracietler.Columns[2].Caption = "Hesab";
            gridmuracietler.Columns[2].Width = 250;
            gridmuracietler.Columns[3].Caption = "İtem 1";
            gridmuracietler.Columns[3].Width = 150;
            gridmuracietler.Columns[4].Caption = "Zamin";
            gridmuracietler.Columns[4].Width = 300;
            gridmuracietler.Columns[5].Caption = "Mobil";
            gridmuracietler.Columns[5].Width = 150;


            // }
            //catch (Exception)
            //{


            //}
            //finally { };


        }
        private void exceleatsms()
        {

            try
            {
                System.Data.DataRow row = gridmuracietler.GetDataRow(gridmuracietler.FocusedRowHandle);
                string sorg = row[10].ToString(); 
                saveFileDialog1.FileName = "SMS (" + ad + " " + sorg + " " + DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year + ")";
                saveFileDialog1.Filter = "XLS Dosyaları (*.xls)|*.xls";

                saveFileDialog1.InitialDirectory = @"\\fs\12345\Personal\Anar_Is\smsler";

                //eğer saveFileDiaolog1 açıldığında Evet’e tıklanırsa

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    DevExpress.XtraPrinting.XlsExportOptions _Options = new DevExpress.XtraPrinting.XlsExportOptions();

                    _Options.SheetName = "Smsler(" + DateTime.Now.Day + "." + DateTime.Now.Month + "." + DateTime.Now.Year + ")";

                    gridmuracietler.ExportToXls(saveFileDialog1.FileName, _Options);

                    if (MessageBox.Show("Yüklədiyiniz excel faylını açmaq istəyirsiniz?", "Excel dosyası", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {

                        //Kaydedilen Excel Dosyasını açar.

                        System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                    }

                }
            }
            catch (Exception)
            {


            }
            finally { }
        }

        private void txtaxtar_Click(object sender, EventArgs e)
        {
            ad = "borcalan";
            tamkataloq();
            btnexcel.Enabled = true;
        }

        private void btnexcel_Click(object sender, EventArgs e)
        {
            exceleatsms();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            ad = "zamin";
            tamkataloqzamin();
            btnexcel.Enabled = true;
        }
    }
}