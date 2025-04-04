using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;
//using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using BMI.Muhasibat;
using System.Drawing;
using DevExpress.Drawing.TextFormatter.Internal;

namespace BMI
{

    class Aletler
    {
        private string baglanti = "Data Source=BMI;User ID=FOXPRO;Password=pass";
        cl_yanasmalar cl = new cl_yanasmalar();
        public string sorgular = "huqui_sorgu";
        //Layihənin əsas qovluğunu qaytarır
        //public static string Layiheanaqovluq()
        //{
        //    //diqqet c# acanda 3 seviye olani aktiv ele setup edende ise komente al
        //    // Tətbiqin işlədiyi yer
        //    string islekqovluq = AppDomain.CurrentDomain.BaseDirectory;

        //    // 3 səviyyə yuxarı çıxaraq əsas layihə qovluğuna çatırıq
        //    //return Directory.GetParent(islekqovluq).Parent.Parent.Parent.FullName;
        //    if (Debugger.IsAttached)
        //    {
        //        //string islekqovluq = AppDomain.CurrentDomain.BaseDirectory;
        //        return Directory.GetParent(islekqovluq).Parent.Parent.Parent.FullName;
        //    }
        //    else
        //    {
        //        // Quraşdırılmış versiya üçün
        //        return AppDomain.CurrentDomain.BaseDirectory;
        //    }
        //}
        //public static string Layiheanaqovluq()
        //{
        //    string islekqovluq = AppDomain.CurrentDomain.BaseDirectory;

        //    // 3 səviyyə yuxarı çıxarkən null olub-olmadığını yoxla
        //    DirectoryInfo dir = Directory.GetParent(islekqovluq);
        //    if (dir?.Parent?.Parent?.Parent != null)
        //    {
        //        return dir.Parent.Parent.Parent.FullName;
        //    }

        //    return islekqovluq; // Əgər mümkün deyilsə, sadəcə işlək qovluğu qaytar
        //}
        public static string Layiheanaqovluq()//novbeti versiya 31-01-2025
        {
            string islekqovluq = AppDomain.CurrentDomain.BaseDirectory;

            if (Debugger.IsAttached)
            {
                // Layihə içində işlədilirsə, 3 səviyyə yuxarı qalx
                return Directory.GetParent(islekqovluq).Parent.Parent.Parent.FullName;
            }
            else
            {
                // Quraşdırılmış versiyada "bin" qovluğunu nəzərə al
                if (islekqovluq.Contains("bin"))
                {
                    return Directory.GetParent(islekqovluq).FullName; // Bin qovluğunu çıxar
                }
                return islekqovluq;
            }
        }
        public static string TarixinSonReqemineSufiksElaveEt(string date)
        {
            if (DateTime.TryParseExact(date, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                int lastDigit = parsedDate.Year % 10; // Tarixin ilin sonuncu rəqəmi

                if (lastDigit == 0)
                {
                    int lastTwoDigits = parsedDate.Year % 100; // Tarixin ilin son iki rəqəmi
                    switch (lastTwoDigits)
                    {
                        case 10: return "-cu";
                        case 20: return "-ci";
                        case 30: return "-cı";
                        case 40: return "-cı";
                        case 50: return "-ci";
                        case 60: return "-cı";
                        case 70: return "-ci";
                        case 80: return "-ci";
                        case 90: return "-cu";
                        default: return "Qeyri-düzgün il";
                    }
                }

                switch (lastDigit)
                {
                    case 1: return   "-ci";
                    case 2: return   "-ci";
                    case 3: return   "-cü";
                    case 4: return   "-cü";
                    case 5: return   "-ci";
                    case 6: return   "-cı";
                    case 7: return   "-ci";
                    case 8: return   "-ci";
                    case 9: return   "-cu";
                    default: return "Qeyri-düzgün il";
                }
            }
            else
            {
                return "Tarix formatı düzgün deyil";
            }
        }
        // Default ikon
        public static Icon DefaultIcon { get; set; }
        // İkona resursdan dəyər təyin edin (layihənin resurslarından bir ikon seçin)
        static Aletler()
        {
            DefaultIcon = Properties.Resources.logo1; // Resursdakı ikon
        }
        public static string TarixiSozeCevir(string date)
        {
            if (DateTime.TryParseExact(date, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                string day = parsedDate.ToString("dd", CultureInfo.InvariantCulture);
                string month = parsedDate.ToString("MMMM", CultureInfo.CreateSpecificCulture("az-Latn-AZ")); // Ayı Azərbaycan dilində almaq üçün
                string year = parsedDate.ToString("yyyy", CultureInfo.InvariantCulture);

                string suffix = TarixinSonReqemineSufiksElaveEt(date);

                return $"{day} {month} {year}{suffix}";
            }
            else
            {
                return "Tarix formatı düzgün deyil";
            }
        }
        public void CreateWordDocument(string templatePath, string outputPath, Dictionary<string, string> replacements)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            // Şablon sənədini açın
            var wordDocument = wordApp.Documents.Open(templatePath);

            // Placeholder-ləri dəyişin
            foreach (var item in replacements)
            {
                ReplaceWordStub(item.Key, item.Value, wordDocument);
            }
            wordApp.Visible = true;
            // Sənədi saxlayın
            wordDocument.SaveAs(outputPath);

            // Sənədi və tətbiqi bağlayın
            //wordDocument.Close();
            //wordApp.Quit();

            Console.WriteLine($"Sənəd uğurla yaradıldı: {outputPath}");
        }

        
        private void ReplaceWordStub(string placeholder, string replacement, Document document)
        {
            foreach (Range range in document.StoryRanges)
            {
                range.Find.Execute(FindText: placeholder, ReplaceWith: replacement, Replace: WdReplace.wdReplaceAll);
            }
        }
        public string GedenMektubNo(string tableName, string columnName, string ilColumnName)
        {
            try
            {
                string tarixIl = DateTime.Now.Date.Year.ToString();
                int mek_no_arti = 0;
                string query = $@"
                SELECT MAX(-TO_NUMBER(SUBSTR(t.{columnName}, 5, 5))) mn 
                FROM {tableName} t 
                WHERE t.{ilColumnName} = '{tarixIl}'";

                using (OracleConnection Orcon = new OracleConnection(baglanti))
                {
                    Orcon.Open();
                    using (OracleCommand cmd = new OracleCommand(query, Orcon))
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read() && dr[0] != DBNull.Value)
                        {
                            int id = int.Parse(dr[0].ToString());
                            mek_no_arti = id + 1;
                        }
                        else
                        {
                            mek_no_arti = 1;
                        }
                    }
                }

                return $"{tarixIl}-{mek_no_arti}";
            }
            catch (Exception ex)
            {
                throw new Exception("Error while generating mek_no: " + ex.Message);
            }
        }
        public  string YaziyaCevir(decimal tutar, string valyuta)
        {
            string[] birler = { "", " bir", " iki", " üç", " dörd", " beş", " altı", " yeddi", " səkkiz", " doqquz" };
            string[] onlar = { "", " on", " iyirmi", " otuz", " qırx", " əlli", " altmış", " yetmiş", " səksən", " doxsan" };
            string[] binler = { "", " min", " milyon", " milyard", " trilyon" };

            if (tutar == 0)
                return "sıfır";

            string sTutar = tutar.ToString("F2").Replace('.', ',');
            string tamHisse = sTutar.Split(',')[0];
            string kicikHisse = sTutar.Split(',')[1];

            string sonuc = ConvertTamHisse(tamHisse, birler, onlar, binler);

            // Tam hissə üçün uyğun valyuta yazılır
            sonuc += $" {ValyutaTamAd(valyuta)}";

            if (kicikHisse != "00")
            {
                sonuc += " " + ConvertKicikHisse(kicikHisse, birler, onlar) + $" {ValyutaKicikAd(valyuta)}";
            }

            return sonuc.Trim();
        }
        private  string ValyutaTamAd(string valyuta)
        {
            switch (valyuta)
            {
                case "AZN": return "manat";
                case "USD": return "dollar";
                case "AVRO": return "avro";
                default: return "";
            }
        }
        private  string ValyutaKicikAd(string valyuta)
        {
            switch (valyuta)
            {
                case "AZN": return "qəpik";
                case "USD": return "sent";
                case "AVRO": return "sent";
                default: return "";
            }
        }
        private  string ConvertTamHisse(string tamHisse, string[] birler, string[] onlar, string[] binler)
        {
            tamHisse = tamHisse.PadLeft(15, '0');
            string sonuc = "";

            for (int i = 0; i < tamHisse.Length; i += 3)
            {
                string ucReqem = tamHisse.Substring(i, 3);
                if (ucReqem != "000")
                {
                    sonuc += ConvertUcReqem(ucReqem, birler, onlar) + binler[(tamHisse.Length - i) / 3 - 1] + " ";
                }
            }
            return sonuc;
        }
        private  string ConvertUcReqem(string ucReqem, string[] birler, string[] onlar)
        {
            string sonuc = "";

            if (ucReqem[0] != '0')
                sonuc += (ucReqem[0] == '1' ? " yüz" : birler[ucReqem[0] - '0'] + " yüz");

            sonuc += onlar[ucReqem[1] - '0'];
            sonuc += birler[ucReqem[2] - '0'];

            return sonuc;
        }
        private  string ConvertKicikHisse(string kicikHisse, string[] birler, string[] onlar)
        {
            return onlar[kicikHisse[0] - '0'] + birler[kicikHisse[1] - '0'];
        }
        public void GedenMektubaElaveEt(string gonyer,string comboBoxText, string avtNo, string icraciKod)
        {
            string gun = DateTime.Now.Day.ToString();
            string ay = DateTime.Now.Month.ToString();
            string il = DateTime.Now.Year.ToString();
            string dateString = DateTime.Now.ToString("dd-MM-yyyy");

            string mezm = "";
            if (comboBoxText == "Girovdan çıxma")
            {
                mezm = " avto gir azad Cars";
            }
            else if (comboBoxText == "Texpasport dəyişmə")
            {
                mezm = " avto texpasport dəyişmə Cars";
            }

            string tammezmun = avtNo + mezm;

            using (OracleConnection orcon = new OracleConnection(baglanti))
            {
                orcon.Open();
                string query = "INSERT INTO odb.xaric_mektub (gon_yer, tarix, qisa_mez, icraci, il) " +
                               "VALUES (:gon_yer, TO_DATE(:tarix, 'dd-MM-yyyy'), :qisa_mez, :icraci, :il)";

                using (OracleCommand orcom = new OracleCommand(query, orcon))
                {
                    orcom.Parameters.Add("gon_yer", OracleDbType.Varchar2).Value = gonyer;
                    orcom.Parameters.Add("tarix", OracleDbType.Varchar2).Value = dateString;
                    orcom.Parameters.Add("qisa_mez", OracleDbType.Varchar2).Value = tammezmun;
                    orcom.Parameters.Add("icraci", OracleDbType.Varchar2).Value = icraciKod;
                    orcom.Parameters.Add("il", OracleDbType.Int32).Value = int.Parse(il);

                    orcom.ExecuteNonQuery();
                }
            }
        }
    }
}
