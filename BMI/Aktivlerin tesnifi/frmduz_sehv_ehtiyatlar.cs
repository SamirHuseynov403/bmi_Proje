using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace BMI.Aktivlerin_tesnifi
{
    public partial class frmduz_sehv_ehtiyatlar : Form
    {
        public frmduz_sehv_ehtiyatlar()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string dosyaYolu = "C:\\Users\\Samir\\Desktop\\Aktivler yeni qaydalar.txt"; // Dosya yolu
            string arananMetin = "q_"+textBox1.Text; // Aradığınız metin

            string bulunanSatir = null; // Bulunan satırı saklamak için bir değişken

            foreach (string satir in File.ReadLines(dosyaYolu, Encoding.Default))
            {
                if (satir.Contains(arananMetin))
                {
                    // "_q" bulunan metinleri boşlukla değiştir
                    string duzeltilmisSatir = satir.Replace("_q", " ");

                    bulunanSatir = duzeltilmisSatir; // Aradığınız metni içeren tam satırı saklayın
                    break; // İlk eşleşmeyi bulduktan sonra döngüyü durdurun
                }
            }

            if (bulunanSatir != null)
            {
                // Bulunan satırı ekrana yazdırabilirsiniz
                richTextBox1.Text = bulunanSatir;
            }
            else
            {
                // Aradığınız metni içeren satırı bulamazsanız bir mesaj gösterebilirsiniz
                richTextBox1.Text = "Heç bir nəticə tapılmadı";
            }
        }
    }
    }
    

