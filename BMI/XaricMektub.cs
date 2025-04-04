using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Wordekocur = Microsoft.Office.Interop.Word;
using System.IO;

namespace BMI
{
    public partial class XaricMektub : Form
    {
        public XaricMektub()
        {
            InitializeComponent();
        }
        string filetype = string.Empty;
        public Xarix_olan_mektub frmMktbxrc;
        public string icraci_kod = string.Empty;


        private void txbgedtarix_TextChanged(object sender, EventArgs e)
        {

        }

        private void XaricMektub_Load(object sender, EventArgs e)
        {
            //textBox2.Text = Istifadeciler.icracikodu.ToString();//burda icraci kodunu texte yazdirdim
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1();


            if (lblSgldml.Text == "insert")
            {
                if (cmbGonderyer.Text.Length > 0 && cmbGisamezmun.Text.Length > 0)
                {
                    DialogResult dr = MessageBox.Show("Məktubun qeydiyyata alınsın ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        string metin = rctbMektubmetn.Text;
                        string metinWithEscape = metin.Replace("'", "''");

                        frmMktbxrc.mktb.Xmgonderilenyer = cmbGonderyer.Text.TrimStart();
                        frmMktbxrc.mktb.Xaricmektbtarix = dtpTarix.Text;
                        frmMktbxrc.mktb.Xmgisamezmun = cmbGisamezmun.Text.TrimStart();
                        frmMktbxrc.mktb.icracikodu = icraci_kod;
                        frmMktbxrc.mktb.Xmmektubmetn = metinWithEscape;
                        frmMktbxrc.mktb.Xaricmektinsert();

                        btnTemizle.PerformClick();
                        frmMktbxrc.btnYenile.PerformClick();
                    }
                }
                else
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            else if (lblSgldml.Text == "update")
            {
                if (cmbGonderyer.Text.Length > 0 && cmbGisamezmun.Text.Length > 0)
                {
                    DialogResult dr = MessageBox.Show("Məktuba düzəliş edilsin ?", "Sual", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dr == DialogResult.Yes)
                    {
                        string metin = rctbMektubmetn.Text;
                        string metinWithEscape = metin.Replace("'", "''");

                        frmMktbxrc.mktb.Xaricmektbtarix = dtpTarix.Text;
                        frmMktbxrc.mktb.Xmgonderilenyer = cmbGonderyer.Text.TrimStart();
                        frmMktbxrc.mktb.Xmgisamezmun = cmbGisamezmun.Text.TrimStart();
                        frmMktbxrc.mktb.Xmmektubmetn = metinWithEscape;
                        frmMktbxrc.mktb.Xaricmektupdate();

                        frmMktbxrc.btnYenile.PerformClick();
                    }
                }
                else
                {
                    MessageBox.Show("Məlumatlar tam doldurulmayıb !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            opnflYol.Title = "Sənədi seçin...";
            opnflYol.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            opnflYol.Multiselect = false;
            //openFileDialog1.CheckFileExists = false;
            opnflYol.RestoreDirectory = true;
            opnflYol.Filter = "Dəstəklənən bütün fayllar|*.docx;*.doc;*.rtf;*.txt| Word faylları(*.doc, *.docx) |*.docx; *.doc| Rich text(*.rtf)|*.rtf| Notepad(*.txt)|*.txt";
            opnflYol.FilterIndex = 1;
            if (opnflYol.ShowDialog() == DialogResult.OK)
            {
                txtYol.Text = opnflYol.FileName;
                rctbMektubmetn.Clear();
                filetype = Path.GetExtension(txtYol.Text);

                if (filetype == ".doc" || filetype == ".docx")
                {
                    Wordekocur.Application WordApp = new Wordekocur.Application();
                    object File = txtYol.Text;
                    object nullobject = System.Reflection.Missing.Value;
                    WordApp.DisplayAlerts = Wordekocur.WdAlertLevel.wdAlertsNone;

                    Wordekocur._Document docs = WordApp.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);

                    docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();

                    this.rctbMektubmetn.Paste();
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    WordApp.Quit(ref nullobject, ref nullobject, ref nullobject);
                }
                else if (filetype == ".rtf")
                {
                    FileStream fStream = new FileStream(txtYol.Text, FileMode.Open, FileAccess.Read);

                    rctbMektubmetn.LoadFile(fStream, RichTextBoxStreamType.RichText);

                }
                else if (filetype == ".txt")
                {
                    FileStream fStream = new FileStream(txtYol.Text, FileMode.Open, FileAccess.Read);

                    rctbMektubmetn.LoadFile(fStream, RichTextBoxStreamType.PlainText);
                }
            }
        }

        private void btnTemizle_Click(object sender, EventArgs e)
        {
            rctbMektubmetn.Clear();
            txtYol.Clear();
            cmbGisamezmun.Text = string.Empty;
            cmbGonderyer.Text = string.Empty;
        }
    }
}
