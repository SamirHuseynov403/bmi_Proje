using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using Word = Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Globalization;
using System.Data.OracleClient;
using DevExpress.XtraEditors;

namespace BMI
{
    public partial class frmbircixaris : Form
    {
        public frmbircixaris()
        {
            InitializeComponent();
        }

        private void cixar_melumati_gonder()
        {
            frmmenzil fmenzil = (frmmenzil)Application.OpenForms["frmmenzil"];
            //fmenzil.zam1adi = this.txbzam1adi.Text;
        }
        private void BtnGuncelle_Click(object sender, EventArgs e)
        {

        }

        private void btngonder1_Click(object sender, EventArgs e)
        {

        }
        private void YazıBoyutunuDeğiştir(Control.ControlCollection controls, float yeniBoyut)
        {
            foreach (Control control in controls)
            {
                if (control is TextEdit || control is LabelControl) // TextEdit veya LabelControl için kontrol
                {
                    if (control is TextEdit)
                    {
                        TextEdit textEdit = control as TextEdit;
                        textEdit.Properties.Appearance.Font = new Font(textEdit.Properties.Appearance.Font.FontFamily, yeniBoyut);
                    }
                    //else if (control is LabelControl)
                    //{
                    //    LabelControl labelControl = control as LabelControl;
                    //    labelControl.Appearance.Font = new Font(labelControl.Appearance.Font.FontFamily, yeniBoyut);
                    //}
                }

                // Eğer iç içe bir kontrol yapısı (örneğin, PanelControl içindeki kontroller) varsa, bu yapılarda da işlem yapılmalıdır.
                if (control.HasChildren)
                {
                    YazıBoyutunuDeğiştir(control.Controls, yeniBoyut);
                }
            }
        }
        void clasagonder()
        {
            csSablontoplusu cssb = new csSablontoplusu();
            cssb.C_seriyaNo= txtmserno.Text;
            cssb.C_obadi = txthuqadi.Text;
            cssb.C_umsahe = txtsahe.Text;
            cssb.C_yasayis = txtyasayis.Text;
            cssb.C_yardimci = txtyardimci.Text;
            cssb.C_reyesno = txtreysno.Text;
            cssb.C_qeydno = txtqeydno.Text;
            cssb.C_tarix = dategtarixi.Text;
            cssb.C_otaqsay = txtotaq.Text;
            cssb.C_kitabno = txtkitabno.Text;
            cssb.C_vereqno = txtvereqno.Text;
            cssb.C_girovunvan = txts1girovunvan.Text;
            cssb.C_sahibi = txtmsahibadi.Text;
            cssb.zaminmuqn1 = txts1zamno.Text;

        }

        private void frmbircixaris_Load(object sender, EventArgs e)
        {
            //float yeniBoyut = 11.0f; // Yeni yazı boyutunu ayarlayın
            //YazıBoyutunuDeğiştir(this.Controls, yeniBoyut);
        }
    }
}
