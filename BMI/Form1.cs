using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using Kredit_isler;
using BMI.Muhasibat;
using BMI.Esas;
using BMI.Emeliyyat;
using BMI.Emek_haqqi_ve_davamiyyet.Classlar;

namespace BMI
{
    public partial class Form1 : Form
    {
       public Onlayn frmana10;
        Form1 fgh;
        public Form1()
        {
            InitializeComponent();
        }
        cl_isciler cl = new cl_isciler();
        public string icraci_kod { get; set; }
        public Logiin frmgirish1;
        public Oraclebaglanti orabag;
        public void ChangeKeyboardLangENG()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage); InputLanguage.CurrentInputLanguage = l;
        }
        private void dYPArayışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DYP dypar = new DYP();
            dypar.icraci_kod = tlsplblAdi.Text;
            dypar.Show();
        }
        string tamad;
        private void pulKöçürməsiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Pul_Kocurmesi plk = new Pul_Kocurmesi();
            plk.icraci_kod = tlsplblAdi.Text;
            plk.ShowDialog();
            tamad = lblTamad.Text;
        }
        private void gedənHəvaləToolStripMenuItem_Click(object sender, EventArgs e)
        {
            geden_hevale gdnhevale = new geden_hevale();
            gdnhevale.icraci_kod = tlsplblAdi.Text;
            gdnhevale.ShowDialog();
        }
        private void gələnHəvaləToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Gelen_hevale ghev = new Gelen_hevale();
            ghev.icraci_kod = tlsplblAdi.Text;
            ghev.ShowDialog();
        }
        private void daxilOlanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            daxil_olan_mektub frm3 = new daxil_olan_mektub();
            frm3.icraci_kod = tlsplblAdi.Text;
            frm3.Show();
        }
        private void xaricOlanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Xarix_olan_mektub x_mkt = new Xarix_olan_mektub();
            x_mkt.icraci_kod = tlsplblAdi.Text;
            x_mkt.ShowDialog();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            lblTamad.Visible = true;
            Logiin frmg = new Logiin();
            frmg.frmana = this;
            frmg.ShowDialog();
        }
        private void metkeyboardLangRUS()
        {
            throw new NotImplementedException();
        }
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        private void müraciətlərÜmumiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmmuracietler frmes = new frmmuracietler();
            frmes.Dock = DockStyle.Fill;
            frmes.TopLevel = false;
            frmes.FormBorderStyle = FormBorderStyle.None;
            panelesas.Controls.Add(frmes);
            frmes.icraci_kod = tlsplblAdi.Text;
            frmes.Show();
        }
        private void kreditOnlaynMüraciətlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string tamad = lblTamad.Text;
            Onlayn only = new Onlayn();
            only.Show();
            only.label7.Text = tamad;
        }
        private void lblTamad_Click(object sender, EventArgs e)
        {

        }
        private void tələbəKöçürməsiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Telebe tlbform = new Telebe();
            tlbform.Show();
            
        }
        private void əməliyyatDPToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void kreditMeqaviləsiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Anakredit ankr = new Anakredit();
            Avtomobil avto = new Avtomobil();
            Zamin zaminnn = new Zamin();
            ankr.label12.Text = lblTamad.Text;
            ankr.icraciaditam = lblTamad.Text;
            zaminnn.label19.Text = lblTamad.Text;
            ankr.icraci_kod = tlsplblAdi.Text;
            avto.icraci_kod = tlsplblAdi.Text;
            ankr.Show();
            Form1 frm1 = new Form1();
            Zamin zmn = new Zamin();
        }
        private void kreditToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void kreditSilinməlriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit_silinme krsil = new Kredit_silinme();
            krsil.Show();
        }
        private void excelAtToolStripMenuItem_Click(object sender, EventArgs e)
        {
             AMLexcel amlex=new AMLexcel();
             amlex.ShowDialog();
        }
        private void zaminDəyişməToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Elave_Zamin zmnelave = new Elave_Zamin();
            zmnelave.icraci_kod = tlsplblAdi.Text;
            zmnelave.ShowDialog();
        }
        private void mənzilVaxtUzadılmaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            İpoteka_vaxt_uzadilma zmnelave = new İpoteka_vaxt_uzadilma();
            zmnelave.icraci_kod = tlsplblAdi.Text;
            zmnelave.ShowDialog();
        }
        private void verilmişKreditlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Verilmis_kr verkr = new Verilmis_kr();
            verkr.ShowDialog();

        }
        private void verilmişKrArasıMəlumatlarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void nömrəÜzrəSorğuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tel_axraris tlaxtar = new Tel_axraris();
            tlaxtar.ShowDialog();
        }
        private void əlaqəliŞəxslərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AML_elaqeli_sexs amlsex = new AML_elaqeli_sexs();
            amlsex.ShowDialog();
        }
        private void hesabQalıqlarıÜzrəMəlumatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hesab_qaliq hsbq = new Hesab_qaliq();
            hsbq.ShowDialog();
        }
        private void kataloqToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void saipaGirovdanÇıxmaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmcarsmektubcix frm1 = new frmcarsmektubcix();
            frm1.icraci_kod = tlsplblAdi.Text;
            frm1.Show();
        }
        private void zaminəTəmizlikArayışıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            zaminarayis zmn = new zaminarayis();
            zmn.icraci_kod = tlsplblAdi.Text;
            zmn.ShowDialog();
        }
        private void mühasibatToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void iddiaƏrizəsiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmİddiaerizesi frmid = new frmİddiaerizesi();
            frmid.icraci_kod = tlsplblAdi.Text;
            frmid.Show();
        }
        private void borcalanTemizlikArayışıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmborcalantemizlik frmborc = new frmborcalantemizlik();
            frmborc.icraci_kod = tlsplblAdi.Text;
            frmborc.ShowDialog();
        }
        private void kreditFaizlərinVəMRTHesablanmasıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void istehlakKreditlərÜzrəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_kredit_faizleri frmfaiz = new frm_kredit_faizleri();
            frmfaiz.icraci_kod = tlsplblAdi.Text;
            frmfaiz.Show();
        }
        private void zaminDəyşməToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmehtiyatlarin_cedveli frmfaiz = new frmehtiyatlarin_cedveli();
            //frmfaiz.icraci_kod = tlsplblAdi.Text;
            frmfaiz.Show();
        }
        private void atlasCarsGriovlarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAtlas_Cars_girov frmcars = new frmAtlas_Cars_girov();
            frmcars.icraci_kod = tlsplblAdi.Text;
            frmcars.Show();
        }
        private void sMSSiyahıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmsmssiyahi frm = new frmsmssiyahi();
            frm.ShowDialog();
        }
        private void hesabÜzrəSorğuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmhesabsorgu frm = new frmhesabsorgu();
            frm.Show();
        }
        private void balansdaQeydiyyatNoSorğuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmbalanssorgu frm = new frmbalanssorgu();
            frm.Show();
        }
        private void ödənişlərVəFərqliSilinmələrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void sorguToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Aktivlerin_tesnifi.frmduz_sehv_ehtiyatlar frm = new Aktivlerin_tesnifi.frmduz_sehv_ehtiyatlar();
            frm.Show();
        }
        private void hüquqiŞəxsHesabatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_AML_huquqi_hesab_imza frm = new frm_AML_huquqi_hesab_imza();
            frm.Show();
        }
        private void kredtÜmumiSorğuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_kredit_umumi_sorgular frm = new frm_kredit_umumi_sorgular();
            frm.Show();
        }
        private void şirkətIləBağlıHesabQalıqlarıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        private void mühasibatToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }
        //s1
        private void əlaqəliDepozitlərinCəmiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_elaqeli_dep_cem frm = new Muhasibat.frm_elaqeli_dep_cem();
            frm.Show();
        }
        //s2
        private void dailyReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_Daily_Report frm = new Muhasibat.frm_Daily_Report();
            frm.label2.Visible = false;
            frm.textBox3.Visible = false;
            frm.button1.Text = "Daily Report";
            frm.button1.Location = new System.Drawing.Point(53, 70);
            frm.textBox2.Location = new System.Drawing.Point(59, 40);
            frm.Show();
            
        }
        //s3
        private void lCRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.LCR frm = new Muhasibat.LCR();
            frm.Show();
        }
        private void restruktruzasiyaOlunmuşKreditlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_rest_yaxs frm = new frm_rest_yaxs();
            frm.Show();
        }
        //s10
        private void iranHesabatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_iran_hesabat frm = new Muhasibat.frm_iran_hesabat();
            frm.Show();
        }
        //s4
        private void dailyCommentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_Daily_Report frm = new Muhasibat.frm_Daily_Report();
            frm.Text = "Daily comments";
            frm.label2.Visible = true;
            frm.textBox3.Visible = true;
            frm.button1.Text = "Daily comments";
            frm.Show();
        }
        //s5
        private void tələbliVəMüddətliDepozitlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_Dep frm = new Muhasibat.frm_Dep();
            frm.Show();
        }
        private void kreditPortfeliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit.frm_portfel frm = new Kredit.frm_portfel();
            frm.Text = "Kredit növləri üzrə";
            frm.lbld_yolu.Text = "tip";

            frm.Show();
        }
        private void işSektorlarıÜzrəVerilmişKreditVəQalıqlarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit.frm_portfel frm = new Kredit.frm_portfel();
            frm.Text = "Kredit girovları üzrə";
            frm.lbld_yolu.Text = "zaloq";
            frm.Show();
        }
        private void kreditPortfeliIşSektorÜzrəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit.frm_portfel frm = new Kredit.frm_portfel();
            frm.Text = "Kredit iş sektorları üzrə";
            frm.lbld_yolu.Text = "sektor";
            frm.Show();
        }
        private void kreditPortfeliQalıqlarÜzrəToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit.frm_portfel frm = new Kredit.frm_portfel();
            frm.Text = "Kredit qalıqları üzrə";
            frm.lbld_yolu.Text = "qaliq";
            
            frm.Show();
        }
        private void kataloqToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmehtiyyatlar hsbq = new frmehtiyyatlar();
            hsbq.Show();
        }
        private void ödənişlərVəFərqliSilinmələrToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frm_odenis_ferqlisilinme frm = new frm_odenis_ferqlisilinme();
            frm.Show();
        }
        private void aktivCariHesabMəlumatıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Emeliyyat.frm_cari_hesab_melumati frm = new Emeliyyat.frm_cari_hesab_melumati();
            frm.Show();
        }
        private void kreditiBağlıOlubBkMəbləğiQalanlarVəYaUyğunsuzOlanlarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Kredit.frm_bk_qalanlar frm = new Kredit.frm_bk_qalanlar();
            frm.Show();
                    

        }
        //s6
        private void dailyReportYeniToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_report_comments_Yeni frm = new Muhasibat.frm_report_comments_Yeni();
            frm.Show();
        }
        //s7
        private void nXVSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_valyuta_alis_satis frm = new Muhasibat.frm_valyuta_alis_satis();
            frm.Show();
        }
        private void gündəlikLimitiKeçmişMüştərilərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AML.frm_mebleg_uzre_sorgu_limit_asma frm = new AML.frm_mebleg_uzre_sorgu_limit_asma();
            frm.Show();
        }
        private void hesablarınRiskQruplariÜzrəYenilənməSorğusuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AML.frm_risk_yenilenme frm = new AML.frm_risk_yenilenme();
            frm.Show();
        }
        private void başİdarəAylıqHesabatSaySorğusuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AML.Aml_bas_idare_ayliq_hesabat frm = new AML.Aml_bas_idare_ayliq_hesabat();
            frm.Show();
        }
        //s8
        private void xBBISToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_XBBIS xb=new Muhasibat.frm_XBBIS();
            xb.Show();
        }
        //s9
        private void aDD1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_ADD1 xb = new Muhasibat.frm_ADD1();
            xb.Show();
        }
        private void panelesas_Paint(object sender, PaintEventArgs e)
        {

        }
        private void setupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (icraci_kod == "45")
            {
                frm_setup setupForm = new frm_setup(this);
                setupForm.Show();
            }
            else
            {
                MessageBox.Show("Bu forma girişiniz yoxdur", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
        }
        private void dailyReportToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_Daily_Report frm = new Muhasibat.frm_Daily_Report();
            frm.label2.Visible = false;
            frm.textBox3.Visible = false;
            frm.button1.Text = "Daily Report";
            frm.button1.Location = new System.Drawing.Point(53, 70);
            frm.textBox2.Location = new System.Drawing.Point(59, 40);
            frm.Show();
        }
        private void lCRToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.LCR frm = new Muhasibat.LCR();
            frm.Show();
        }
        private void dailyCommentsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //System.Threading.Thread.Sleep(1000); // 1 saniyə gözləyək
            Muhasibat.frm_Daily_Report frm = new Muhasibat.frm_Daily_Report();
            frm.Text = "Daily comments";
            frm.label2.Visible = true;
            frm.textBox3.Visible = true;
            frm.button1.Text = "Daily comments";
            frm.Show();
        }
        private void tələbliVəMüddətliDepozitlərToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_Dep frm = new Muhasibat.frm_Dep();
            frm.Show();
        }
        private void dailyReportYeniToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_report_comments_Yeni frm = new Muhasibat.frm_report_comments_Yeni();
            frm.Show();
        }
        private void nXVSToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_valyuta_alis_satis frm = new Muhasibat.frm_valyuta_alis_satis();
            frm.Show();
        }
        private void aDD1ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_ADD1 xb = new Muhasibat.frm_ADD1();
            xb.Show();
        }
        private void iranMonthlyReportKreditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_iran_hesabat frm = new Muhasibat.frm_iran_hesabat();
            frm.Show();
        }
        private void əlaqəliDepozitlərinCəmiToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_elaqeli_dep_cem frm = new Muhasibat.frm_elaqeli_dep_cem();
            frm.Show();
        }
        private void xBBISToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_XBBIS xb = new Muhasibat.frm_XBBIS();
            xb.Show();
        }
        private void verilmişödənilmişVəHesablanmışFaizlərToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_kredit_umumi_sorgular frm=new frm_kredit_umumi_sorgular();  
            frm.Show();
        }
        private void theCorrespondentAccountsNostroToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sorgular.frm_Currency_accounts_Nostro fr = new Sorgular.frm_Currency_accounts_Nostro();
            fr.Show();
        }

        private void rezidentVəQeyriRezidentHesabQalıqlarıToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Muhasibat.frm_reziden_ve_qeyri_rezident fr=new frm_reziden_ve_qeyri_rezident();
            fr.Show();
        }

        private void əməkHaqqıHesablanmasıVəDavamiyyətToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frm_Main frm=new frm_Main();
            cl.icraci_ad = lblTamad.Text;
            frm.Show();
        }

        private void exchangeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ManagerApproval.frmExchangeTesdiq frm=new ManagerApproval.frmExchangeTesdiq();
            frm.icraci_ad = lblTamad.Text;
            frm.ShowDialog();
        }

        private void exchangeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Kassa.frmExchange frm=new Kassa.frmExchange();
            frm.icraci_ad = lblTamad.Text;
            frm.icraci_kod= tlsplblAdi.Text;
            frm.ShowDialog();
        }

        private void uTimeMasterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UTimeMaster.frmUTimemasterMain fr=new UTimeMaster.frmUTimemasterMain();
            fr.ShowDialog();
        }

        private void benefisiarMülkiyyətçiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void benefisiarMülkiyyətçiToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            AML.frm_Benefisiar_hesabat fr = new AML.frm_Benefisiar_hesabat();
            fr.ShowDialog();
        }
    }
}
