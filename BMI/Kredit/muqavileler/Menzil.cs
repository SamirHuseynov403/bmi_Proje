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
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Interop.Excel;

namespace BMI
{
    public partial class Menzil : Form
    {
        private readonly string templatefilname;
        public Menzil()
        {
            InitializeComponent();
        }
        private void word_at()
        {
            Girov_melumatlari girov = new Girov_melumatlari();
            var adi = txbmenzilBorcalan.Text;
            var passport = txbmenzilPasport.Text;
            var pasvertarix = dateTimePicker1.Text;
            var Orqan = cboxOrqan.Text;
            var Unvan = txbmenzilUnvan.Text;
            var CariHesab = txbmenzilCariHesab.Text;
            var Valyuta = cboxValyuta.Text;
            var Serencam = txbmenzilSerencam.Text;
            var MuqNo = txbmenzilMuqno.Text;
            var Muqtarix = dateTimePicker2.Text;
            var Teyinat = cboxTeyinat.Text;
            var Olke = cboxOlke.Text;
            var mebleg = txbmenzilMebleg.Text;
            var muddet = txbmenzilMuddet.Text;
            var faiz = txbmenzilFaiz.Text;
            var vkfaiz = txbmenzilVKFaiz.Text;
            var ayliq = txbmenzilAyliq.Text;
            var fifd = txbmenzilFİFD.Text;
            //cixaris melumatlari
            var seriya = girov.txbgirovSeriya.Text;
            var huquqadi = girov.cboxgirovObyekt.Text;
            var Girunvan = girov.txbgirovUnvan.Text;
            var reyno = girov.txbgirovReyestr.Text;
            var umumisahe = girov.txbgirovUmumisahe.Text;
            var yassahe = girov.txbgirovYasSahe.Text;
            var yarsahe = girov.txbgirovYarSahe.Text;
            var otaq = girov.txbgirovOtaq.Text;
            var qeydtarix = girov.dateTimePicker2.Text;
            var qeydno = girov.txbgirovQeydNo.Text;
            var kitabno = girov.txbgirovKitabNo.Text;
            var vereqno = girov.txbgirovVereqNo.Text;

            // TODO: Word Export
            var wordapp = new Word.Application();
            wordapp.Visible = false;

            try
            {

                var wordDocument = wordapp.Documents.Open(templatefilname);
                ReplaceWordStub("{adi}", adi, wordDocument);
                ReplaceWordStub("{passport}", passport, wordDocument);
                ReplaceWordStub("{pasvertarix}", pasvertarix, wordDocument);
                ReplaceWordStub("{Orqan}", Orqan, wordDocument);
                ReplaceWordStub("{Unvan}", Unvan, wordDocument);
                ReplaceWordStub("{CariHesab}", CariHesab, wordDocument);
                ReplaceWordStub("{Valyuta}", Valyuta, wordDocument);
                ReplaceWordStub("{Serencam}", Serencam, wordDocument);
                ReplaceWordStub("{MuqNo}", MuqNo, wordDocument);
                ReplaceWordStub("{Muqtarix}", Muqtarix, wordDocument);
                ReplaceWordStub("{Teyinat}", Teyinat, wordDocument);
                ReplaceWordStub("{Olke}", Olke, wordDocument);
                ReplaceWordStub("{mebleg}", mebleg, wordDocument);
                ReplaceWordStub("{muddet}", muddet, wordDocument);
                ReplaceWordStub("{faiz}", faiz, wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz, wordDocument);
                ReplaceWordStub("{vkfaiz}", vkfaiz, wordDocument);
                ReplaceWordStub("{ayliq}", ayliq, wordDocument);
                ReplaceWordStub("{fifd}", fifd, wordDocument);
                //cixaris melumatlari
                ReplaceWordStub("{seriya}", seriya, wordDocument);
                ReplaceWordStub("{huquqadi}", huquqadi, wordDocument);
                ReplaceWordStub("{Girunvan}", Girunvan, wordDocument);
                ReplaceWordStub("{reyno}", reyno, wordDocument);
                ReplaceWordStub("{umumisahe}", umumisahe, wordDocument);
                ReplaceWordStub("{yassahe}", yassahe, wordDocument);
                ReplaceWordStub("{yarsahe}", yarsahe, wordDocument);
                ReplaceWordStub("{otaq}", otaq, wordDocument);
                ReplaceWordStub("{qeydtarix}", qeydtarix, wordDocument);
                ReplaceWordStub("{qeydno}", qeydno, wordDocument);
                ReplaceWordStub("{kitabno}", kitabno, wordDocument);
                ReplaceWordStub("{vereqno}", vereqno, wordDocument);






                //wordapp.Visible = true;
                ////wordDocument.SaveAs(@"C:\result.docx");
                //wordDocument.SaveAs(@"C:\FormPK\FormPK\bin\Debug\Erizeler\adi.docx");
                ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler\Erize1.docx
                ////C:\Pul Kocutrme\Pul Kocutrme\bin\Debug\Erizeler
            }
            catch
            {
                //MessageBox.Show("Sehv oldu");
            }

        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document WordDocument)
        {
            var range = WordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        private void Menzil_Load(object sender, EventArgs e)
        {
            Menzil_sahibi msah = new Menzil_sahibi();
            if (cboxMenzilsahibi.Text== "Fərqli şəxsin adına")
            {
                msah.ShowDialog();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Zaminler zmnlar = new Zaminler();
            zmnlar.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Girov_melumatlari gml = new Girov_melumatlari();
            gml.ShowDialog();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
