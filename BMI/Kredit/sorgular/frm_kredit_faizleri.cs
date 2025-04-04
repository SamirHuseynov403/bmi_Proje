using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using wordeaktar = Microsoft.Office.Interop.Word;
using Excel.FinancialFunctions;

namespace BMI
{
    public partial class frm_kredit_faizleri : Form
    {
        public frm_kredit_faizleri()
        {
            InitializeComponent();
        }
        public string icraci_kod { get; set; }
        public static double calcPayment(double presentValue, double financingPeriod, double interestRatePerYear)
        {
            double a, b, x;
            double monthlyPayment;
            a = (1 + interestRatePerYear / 1200);
            b = financingPeriod;
            x = Math.Pow(a, b);
            x = 1 / x;
            x = 1 - x;
            monthlyPayment = (presentValue) * (interestRatePerYear / 1200) / x;
            return (monthlyPayment);
        }
        private void ayliqi_tap()
        {

            try
            {
                if (kat_faiz.Text == "" || Kat_kredit.Text == "" || Kat_muddet.Text == "" || txbgelir.Text == "")
                {
                    Kat_ayliq.Text = "";
                }
                else
                {
                    double pv = Convert.ToDouble(Kat_kredit.Text);
                    double intrate = Convert.ToDouble(kat_faiz.Text);
                    double period = Convert.ToDouble(Kat_muddet.Text);
                    double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                    string str = pmt.ToString("f2");
                    Kat_ayliq.Text = str;
                }
            }
            catch (Exception)
            {

            }
            finally { }
        }
        private void BGN()
        {
            try
            {
                double borc, gelir, bgn;
                txbBGN.BackColor = Color.White;
                txbBGN.ForeColor = Color.Black;
                if (txbgelir.Text == "" || txbborc.Text == "" || Kat_ayliq.Text == "")
                {

                }
                else
                {
                    gelir = Convert.ToDouble(txbgelir.Text);
                    borc = Convert.ToDouble(txbborc.Text) + Convert.ToDouble(Kat_ayliq.Text);
                    bgn = Math.Round(borc / gelir * 100, 2);
                    txbBGN.Text = bgn.ToString();
                    if (bgn > 45)
                    {
                        txbBGN.BackColor = Color.Yellow;
                    }
                    if (bgn > 70)
                    {
                        txbBGN.BackColor = Color.Red;
                        txbBGN.ForeColor = Color.Blue;
                    }
                }

            }
            catch (Exception)
            {

            }
            finally { }

        }
        private void faizhesabla()
        {
            try
            {
                if (comboBox2.Text == "Adi hesablama")
                {

                }
                else
                {
                    if (Kat_muddet.Text == "" || Kat_kredit.Text == "" || txbborc.Text == "")
                    {

                    }
                    else
                    {
                        int faizitap, muddet = Convert.ToInt32(Kat_muddet.Text);
                        double anafaiz = 12, ayliq;
                        if (comboBox2.Text == "İlk müştəri")
                        {
                            anafaiz = anafaiz + 0.5;
                        }
                        if (muddet > 12 & muddet <= 24)
                        {
                            anafaiz = anafaiz + 0.5;
                        }
                        if (muddet > 24 & muddet <= 36)
                        {
                            anafaiz = anafaiz + 1;
                        }
                        if (muddet > 36 & muddet <= 48)
                        {
                            anafaiz = anafaiz + 1.5;
                        }
                        if (muddet > 48 & muddet <= 60)
                        {
                            anafaiz = anafaiz + 2;
                        }
                        if (Kat_kredit.Text == "" || Kat_muddet.Text == "" || txbgelir.Text == "" || txbborc.Text == "")
                        {
                            Kat_ayliq.Text = "";
                        }
                        else
                        {
                            double borc, gelir, bgn;
                            double pv = Convert.ToDouble(Kat_kredit.Text);
                            double intrate = anafaiz;
                            double period = Convert.ToDouble(Kat_muddet.Text);
                            double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                            string str = pmt.ToString("f2");

                            gelir = Convert.ToDouble(txbgelir.Text);
                            borc = Convert.ToDouble(txbborc.Text) + Convert.ToDouble(str);
                            bgn = Math.Round(borc / gelir * 100, 2);
                            if (bgn > 45)
                            {
                                anafaiz = anafaiz + 1.5;
                            }
                            kat_faiz.Text = anafaiz.ToString();
                        }

                    }
                }
            }
            catch (Exception)
            {

            }
            finally { }

        }
        private void faiz_neticeleri()
        {

            try
            {
                if (comboBox2.Text == "Adi hesablama")
                {

                }
                else
                {


                    double ilkfaiz = 12, muddet = Convert.ToDouble(Kat_muddet.Text), bgn = 0, tekrar = 0;
                    if (comboBox2.Text == "İlk müştəri")
                    {
                        tekrar = 0.5;
                        txbtekrar.Text = tekrar.ToString();
                    }
                    if (comboBox2.Text == "Təkrar müştəri")
                    {
                        tekrar = 0;
                        txbtekrar.Text = tekrar.ToString();
                    }

                    if (Convert.ToDouble(txbBGN.Text) <= 45)
                    {
                        bgn = 0;
                        txbehtiyyat.Text = bgn.ToString();
                    }
                    if (Convert.ToDouble(txbBGN.Text) > 45)
                    {
                        bgn = 1.5;
                        txbehtiyyat.Text = bgn.ToString();
                    }
                    if (muddet <= 12)
                    {
                        muddet = 0;
                        txbfaizneticelerimuddet.Text = muddet.ToString();
                    }
                    if (muddet > 12 & muddet <= 24)
                    {
                        muddet = 0.5;
                        txbfaizneticelerimuddet.Text = muddet.ToString();
                    }
                    if (muddet > 24 & muddet <= 36)
                    {
                        muddet = 1;
                        txbfaizneticelerimuddet.Text = muddet.ToString();
                    }
                    if (muddet > 36 & muddet <= 48)
                    {
                        muddet = 1.5;
                        txbfaizneticelerimuddet.Text = muddet.ToString();
                    }
                    if (muddet > 48)
                    {
                        muddet = 2;
                        txbfaizneticelerimuddet.Text = muddet.ToString();
                    }

                    ilkfaiz = ilkfaiz + muddet + bgn + tekrar;
                    txbfaizneticelerinetice.Text = ilkfaiz.ToString();
                }
            }
            catch (Exception)
            {

            }
            finally { }

        }
        DataTable tablo = new DataTable();
        private void wordeat()
        {
            int ratingxali = 0;
            double fazixali = 12;
            int satirsay = 12, sutunsayi = 2;
            int satirsay1 = 7, sutunsayi1 = 2;
            object omising = System.Reflection.Missing.Value;
            object oendifdoc = "\\endofdoc";
            wordeaktar.Application wordap = new wordeaktar.Application();
            wordeaktar.Document wordoc = wordap.Documents.Add(ref omising, ref omising, ref omising, ref omising);
            wordap.Visible = true;
            object orng = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;
            orng = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;
            wordeaktar.Paragraph paraqraf;
            paraqraf = wordoc.Content.Paragraphs.Add(ref orng);
            paraqraf.Range.Text = "Müştəri reytinqi";
            paraqraf.Range.ParagraphFormat.Alignment = wordeaktar.WdParagraphAlignment.wdAlignParagraphCenter;
            paraqraf.Range.Font.Size = 14;
            //paraqraf.Range.Font.Bold = 1;

            paraqraf.Range.Font.Shadow = 0;
            paraqraf.Format.SpaceAfter = 10;
            paraqraf.Range.InsertParagraphAfter();

            wordeaktar.Range wrdrng = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;
            wordeaktar.Table tablo = wordoc.Tables.Add(wrdrng, satirsay, sutunsayi, ref omising, ref omising);

            tablo.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;
            tablo.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;

            tablo.Rows[1].Cells[1].Range.Text = "Kredit stajı";

            tablo.Rows[2].Cells[1].Range.Text = "Kredit stajının keyfiyyəti";
            tablo.Rows[3].Cells[1].Range.Text = "Rəsmi iş yeri";
            tablo.Rows[4].Cells[1].Range.Text = "Ailə vəziyyəti";
            tablo.Rows[5].Cells[1].Range.Text = "Ehtiyyat dərəcəsi";
            tablo.Rows[6].Cells[1].Range.Text = "Dövlət sektoru";
            tablo.Rows[7].Cells[1].Range.Text = "İşçi təqdimatı";
            tablo.Rows[8].Cells[1].Range.Text = "Kredit müddəti";
            tablo.Rows[9].Cells[1].Range.Text = "Təkrar müştəri";
            tablo.Rows[10].Cells[1].Range.Text = "Müştərinin yaşı";
            tablo.Rows[11].Cells[1].Range.Text = "Təminat";
            tablo.Rows[12].Cells[1].Range.Text = "Rating";
            tablo.Rows[12].Cells[1].Range.Font.Size = 12;
            tablo.Rows[12].Cells[1].Range.Font.Bold = 1;

            if (cmbstaj.Text == "1")
            {
                tablo.Rows[1].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbstaj.Text == "25")
            {
                tablo.Rows[1].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            if (cmbgec.Text == "0")
            {
                tablo.Rows[2].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            else if (cmbgec.Text == "90")
            {
                tablo.Rows[2].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            tablo.Rows[3].Cells[2].Range.Text = "1";
            ratingxali = ratingxali + 1;
            if (cmbaile.Text == "Subay")
            {
                tablo.Rows[4].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbaile.Text == "Evli")
            {
                tablo.Rows[4].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            if (Convert.ToDouble(txbBGN.Text) > 45)
            {
                tablo.Rows[5].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;

            }
            else if (Convert.ToDouble(txbBGN.Text) < 45)
            {
                tablo.Rows[5].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;

            }
            if (cmbisi.Text == "Özəl")
            {
                tablo.Rows[6].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbisi.Text == "Dövlət")
            {
                tablo.Rows[6].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            if (cmbisci.Text == "Xeyr")
            {
                tablo.Rows[7].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbisci.Text == "Bəli")
            {
                tablo.Rows[7].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            if (Convert.ToDouble(Kat_muddet.Text) <= 18)
            {
                tablo.Rows[8].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;

            }
            else if (Convert.ToDouble(Kat_muddet.Text) > 18)
            {
                tablo.Rows[8].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }

            //tablo.Rows[9].Cells[2].Range.Text = cm;
            if (comboBox2.Text == "İlk müştəri")
            {
                tablo.Rows[9].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;

            }
            else if (comboBox2.Text == "Təkrar müştəri")
            {
                tablo.Rows[9].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;

            }
            if (cmbyas.Text == "1")
            {
                tablo.Rows[10].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbyas.Text == "25")
            {
                tablo.Rows[10].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            if (cmbgirov.Text == "Zaminlik")
            {
                tablo.Rows[11].Cells[2].Range.Text = "0";
                ratingxali = ratingxali + 0;
            }
            else if (cmbgirov.Text == "Digər")
            {
                tablo.Rows[11].Cells[2].Range.Text = "1";
                ratingxali = ratingxali + 1;
            }
            tablo.Rows[12].Cells[2].Range.Text = ratingxali.ToString();
            tablo.Rows[12].Range.Font.Bold = 1;
            tablo.Rows[12].Range.Font.Size = 12;



            object orng1 = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;
            orng1 = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;
            wordeaktar.Paragraph paraqraf1;
            paraqraf1 = wordoc.Content.Paragraphs.Add(ref orng1);
            paraqraf1.Range.Text = "Kredit faizinin hesablanması";
            paraqraf1.Range.ParagraphFormat.Alignment = wordeaktar.WdParagraphAlignment.wdAlignParagraphCenter;

            paraqraf1.Range.Font.Size = 14;

            paraqraf1.Range.Font.Shadow = 0;
            paraqraf1.Format.SpaceAfter = 10;
            paraqraf1.Range.InsertParagraphAfter();
            wordeaktar.Range wrdrng1 = wordoc.Bookmarks.get_Item(ref oendifdoc).Range;

            wordeaktar.Table tablo1 = wordoc.Tables.Add(wrdrng1, satirsay1, sutunsayi1, ref omising, ref omising);
            tablo1.Borders.InsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;
            tablo1.Borders.OutsideLineStyle = wordeaktar.WdLineStyle.wdLineStyleSingle;
            tablo1.Rows[1].Cells[1].Range.Text = "Məlumat";
            tablo1.Rows[1].Range.Font.Bold = 1;
            tablo1.Rows[1].Cells[2].Range.Text = "Cavab";
            tablo1.Rows[1].Cells[2].Range.Font.Bold = 1;
            tablo1.Rows[2].Cells[1].Range.Text = "Minimal faiz";
            tablo1.Rows[3].Cells[1].Range.Text = "DTİ";
            tablo1.Rows[4].Cells[1].Range.Text = "Müddət (aylarla)";
            tablo1.Rows[5].Cells[1].Range.Text = "Təkrar müştəri";
            tablo1.Rows[6].Cells[1].Range.Text = "Kreditin faiz dərəcəsi";
            tablo1.Rows[7].Cells[1].Range.Text = "İcraçı";
            tablo1.Rows[7].Cells[1].Range.Font.Bold = 1;
            tablo1.Rows[2].Cells[2].Range.Text = "12%";

            tablo1.Rows[3].Cells[2].Range.Text = txbBGN.Text;

            tablo1.Rows[4].Cells[2].Range.Text = Kat_muddet.Text;
            if (comboBox2.Text == "İlk müştəri")
            {
                tablo1.Rows[5].Cells[2].Range.Text = "Xeyr"; ;
                ratingxali = ratingxali + 0;
            }
            else if (comboBox2.Text == "Təkrar müştəri")
            {
                tablo1.Rows[5].Cells[2].Range.Text = "Bəli";
                ratingxali = ratingxali + 1;
            }

            tablo1.Rows[5].Range.Font.Bold = 1;
            tablo1.Rows[5].Range.Font.Size = 12;
            tablo1.Rows[6].Cells[2].Range.Text = txbfaizneticelerinetice.Text;

        }
        private void maasvergicix()
        {
            if (textBox3.Text == "" || Convert.ToDouble(textBox3.Text) < 261)
            {
                textBox2.Text = "";
            }
            else
            {


                double maas = Convert.ToDouble(textBox3.Text);
                double minyas = Convert.ToDouble(textBox1.Text);
                double temizmaas;
                if (maas < 2500)
                {
                    temizmaas = maas - (maas - minyas) * 14 / 100;

                    textBox2.Text = temizmaas.ToString();
                    if (textBox3.Text == "")
                    {
                        textBox2.Text = "";
                    }
                }
                else if (maas >= 2500)
                {

                    temizmaas = maas - ((maas - 2500) * 25 / 100) - 350;
                    textBox2.Text = temizmaas.ToString();
                    if (textBox3.Text == "")
                    {
                        textBox2.Text = "";
                    }
                }
                else if (maas < 261)
                {
                    textBox2.Text = "";
                }
            }


        }

        private void Kat_kredit_TextChanged(object sender, EventArgs e)
        {
            if (Kat_muddet.Text == "" || kat_faiz.Text == "" || Kat_muddet.Text == "")
            {

            }
            else if (comboBox2.Text == "Adi hesablama")
            {
                double pv = Convert.ToDouble(Kat_kredit.Text);
                double intrate = Convert.ToDouble(kat_faiz.Text);
                double period = Convert.ToDouble(Kat_muddet.Text);
                double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                string str = pmt.ToString("f2");
                Kat_ayliq.Text = str;
            }
            else
            {
                ayliqi_tap();
                faizhesabla();
                BGN();
                faiz_neticeleri();
                FIFD();
            }

        }

        private void kat_faiz_TextChanged(object sender, EventArgs e)
        {
            if (Kat_muddet.Text == "" || kat_faiz.Text == "" || Kat_muddet.Text == "")
            {

            }
            else if (comboBox2.Text == "Adi hesablama")
            {
                double pv = Convert.ToDouble(Kat_kredit.Text);
                double intrate = Convert.ToDouble(kat_faiz.Text);
                double period = Convert.ToDouble(Kat_muddet.Text);
                double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                string str = pmt.ToString("f2");
                Kat_ayliq.Text = str;
            }
            else
            {
                ayliqi_tap();
                faizhesabla();
                BGN();
                faiz_neticeleri();
                FIFD1();
            }

        }

        private void Kat_muddet_TextChanged(object sender, EventArgs e)
        {
            if (Kat_muddet.Text == "" || kat_faiz.Text == "" || Kat_muddet.Text == "")
            {

            }
            else if (comboBox2.Text == "Adi hesablama")
            {
                double pv = Convert.ToDouble(Kat_kredit.Text);
                double intrate = Convert.ToDouble(kat_faiz.Text);
                double period = Convert.ToDouble(Kat_muddet.Text);
                double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                string str = pmt.ToString("f2");
                Kat_ayliq.Text = str;
            }
            else
            {
                ayliqi_tap();
                faizhesabla();
                BGN();
                faiz_neticeleri();
                FIFD();
            }

        }

        private void Kat_ayliq_TextChanged(object sender, EventArgs e)
        {
            if (Kat_muddet.Text == "" || kat_faiz.Text == "" || Kat_muddet.Text == "")
            {

            }
            else if (comboBox2.Text == "Adi hesablama")
            {
                double pv = Convert.ToDouble(Kat_kredit.Text);
                double intrate = Convert.ToDouble(kat_faiz.Text);
                double period = Convert.ToDouble(Kat_muddet.Text);
                double pmt = Convert.ToDouble(calcPayment(pv, period, intrate));
                string str = pmt.ToString("f2");
                Kat_ayliq.Text = str;
            }
            else
            {
                ayliqi_tap();
                faizhesabla();
                BGN();
                faiz_neticeleri();
                FIFD();
            }

        }

        private void txbgelir_TextChanged(object sender, EventArgs e)
        {
            BGN();
            faizhesabla();
            faiz_neticeleri();
        }

        private void txbborc_TextChanged(object sender, EventArgs e)
        {
            BGN();
            ayliqi_tap();
            faizhesabla();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "Adi hesablama")
            {
                kat_faiz.Enabled = true;
                kat_faiz.Text = "";
            }
            if (comboBox2.Text == "İlk müştəri")
            {
                kat_faiz.Text = "";
                kat_faiz.Enabled = false;
                faizhesabla();
                faiz_neticeleri();


            }
            if (comboBox2.Text == "Təkrar müştəri")
            {
                kat_faiz.Text = "";
                kat_faiz.Enabled = false;
                faizhesabla();
                faiz_neticeleri();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            wordeat();
            FIFD();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            maasvergicix();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            maasvergicix();
        }
        private void FIFD()
        {
            double muddet =Convert.ToDouble( Kat_muddet.Text);
            double xhaq = Convert.ToDouble(Kat_kredit.Text) * 1/100;
            double mebleg = Convert.ToDouble(Kat_kredit.Text) * -1;
            double kr = mebleg+xhaq, ay = Convert.ToDouble(Kat_ayliq.Text);


            double[] values1 = new double[] { kr, ay };
            double[] values2 = new double[] { kr, ay, ay, };
            double[] values3 = new double[] { kr, ay, ay, ay, };
            double[] values4 = new double[] { kr, ay, ay, ay, ay, };
            double[] values5 = new double[] { kr, ay, ay, ay, ay, ay, };
            double[] values6 = new double[] { kr, ay, ay, ay, ay, ay, ay, };
            double[] values7 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, };
            double[] values8 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values9 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values10 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values11 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values12 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values13 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values14 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values15 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values16 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values17 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values18 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values19 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay };
            double[] values20 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values21 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values22 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values23 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values24 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values25 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values26 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values27 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values28 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values29 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values30 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values31 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values32 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values33 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values34 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values35 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values36 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay };
            double[] values37 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values38 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values39 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values40 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values41 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values42 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values43 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values44 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values45 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values46 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values47 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values48 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values49 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values50 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values51 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values52 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values53 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values54 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values55 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values56 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values57 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values58 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values59 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double[] values60 = new double[] { kr, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, ay, };
            double FIFD;
            if (muddet==1)
            {
                 FIFD= Math.Round(Financial.Irr(values1) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 2)
            {
                FIFD = Math.Round(Financial.Irr(values2) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 3)
            {
                FIFD = Math.Round(Financial.Irr(values3) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 4)
            {
                FIFD = Math.Round(Financial.Irr(values4) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 5)
            {
                FIFD = Math.Round(Financial.Irr(values5) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 6)
            {
                FIFD = Math.Round(Financial.Irr(values6) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 7)
            {
                FIFD = Math.Round(Financial.Irr(values7) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 8)
            {
                FIFD = Math.Round(Financial.Irr(values8) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 9)
            {
                FIFD = Math.Round(Financial.Irr(values9) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 10)
            {
                FIFD = Math.Round(Financial.Irr(values10) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 11)
            {
                FIFD = Math.Round(Financial.Irr(values11) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 12)
            {
                FIFD = Math.Round(Financial.Irr(values12) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 13)
            {
                FIFD = Math.Round(Financial.Irr(values13) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 14)
            {
                FIFD = Math.Round(Financial.Irr(values14) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 15)
            {
                FIFD = Math.Round(Financial.Irr(values15) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 16)
            {
                FIFD = Math.Round(Financial.Irr(values16) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 17)
            {
                FIFD = Math.Round(Financial.Irr(values17) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 18)
            {
                FIFD = Math.Round(Financial.Irr(values18) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 19)
            {
                FIFD = Math.Round(Financial.Irr(values19) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 20)
            {
                FIFD = Math.Round(Financial.Irr(values20) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 21)
            {
                FIFD = Math.Round(Financial.Irr(values21) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 22)
            {
                FIFD = Math.Round(Financial.Irr(values22) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 23)
            {
                FIFD = Math.Round(Financial.Irr(values23) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 24)
            {
                FIFD = Math.Round(Financial.Irr(values24) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 25)
            {
                FIFD = Math.Round(Financial.Irr(values25) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 26)
            {
                FIFD = Math.Round(Financial.Irr(values26) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 27)
            {
                FIFD = Math.Round(Financial.Irr(values27) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 28)
            {
                FIFD = Math.Round(Financial.Irr(values28) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 29)
            {
                FIFD = Math.Round(Financial.Irr(values29) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 30)
            {
                FIFD = Math.Round(Financial.Irr(values30) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 31)
            {
                FIFD = Math.Round(Financial.Irr(values31) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 32)
            {
                FIFD = Math.Round(Financial.Irr(values32) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 33)
            {
                FIFD = Math.Round(Financial.Irr(values33) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 34)
            {
                FIFD = Math.Round(Financial.Irr(values34) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 35)
            {
                FIFD = Math.Round(Financial.Irr(values35) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 36)
            {
                FIFD = Math.Round(Financial.Irr(values36) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 37)
            {
                FIFD = Math.Round(Financial.Irr(values37) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 38)
            {
                FIFD = Math.Round(Financial.Irr(values38) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 39)
            {
                FIFD = Math.Round(Financial.Irr(values39) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 40)
            {
                FIFD = Math.Round(Financial.Irr(values40) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 41)
            {
                FIFD = Math.Round(Financial.Irr(values41) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 42)
            {
                FIFD = Math.Round(Financial.Irr(values42) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 43)
            {
                FIFD = Math.Round(Financial.Irr(values43) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 44)
            {
                FIFD = Math.Round(Financial.Irr(values44) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 45)
            {
                FIFD = Math.Round(Financial.Irr(values45) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 46)
            {
                FIFD = Math.Round(Financial.Irr(values46) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 47)
            {
                FIFD = Math.Round(Financial.Irr(values47) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 48)
            {
                FIFD = Math.Round(Financial.Irr(values48) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 49)
            {
                FIFD = Math.Round(Financial.Irr(values49) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 50)
            {
                FIFD = Math.Round(Financial.Irr(values51) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 52)
            {
                FIFD = Math.Round(Financial.Irr(values52) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 53)
            {
                FIFD = Math.Round(Financial.Irr(values53) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 54)
            {
                FIFD = Math.Round(Financial.Irr(values54) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 55)
            {
                FIFD = Math.Round(Financial.Irr(values55) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 56)
            {
                FIFD = Math.Round(Financial.Irr(values56) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 57)
            {
                FIFD = Math.Round(Financial.Irr(values57) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 58)
            {
                FIFD = Math.Round(Financial.Irr(values58) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            else if (muddet == 59)
            {
                FIFD = Math.Round(Financial.Irr(values59) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();

            }
            else if (muddet == 60)
            {
                FIFD = Math.Round(Financial.Irr(values60) * 12 * 100, 3);
                double sayi = Math.Round(FIFD, 2);
                txbnetice.Text = sayi.ToString();
            }
            


            //double FIFD = Math.Round(Financial.Irr(values36) * 12 * 100, 3);
            
            

        }
        private void FIFD1()
        {
            double muddet = Convert.ToDouble(Kat_muddet.Text);
            double xhaq = Convert.ToDouble(Kat_kredit.Text) * 1 / 100;
            double mebleg = Convert.ToDouble(Kat_kredit.Text) * -1;
            double kr = mebleg + xhaq;
            double ay = Convert.ToDouble(Kat_ayliq.Text);

            // muddet sayına uyğun values massivini yarat
            double[] values = new double[(int)muddet + 1];
            values[0] = kr;  // İlk dəyər kreditin dəyəri olur
            for (int i = 1; i <= muddet; i++)
            {
                values[i] = ay;  // Sonrakı aylıq dəyərlər
            }

            // IRR hesablamaq
            double FIFD = Math.Round(Financial.Irr(values) * 12 * 100, 3);
            txbnetice.Text = FIFD.ToString("F2");
        }

        private void labelControl8_Click(object sender, EventArgs e)
        {

        }

        private void txbnetice_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
