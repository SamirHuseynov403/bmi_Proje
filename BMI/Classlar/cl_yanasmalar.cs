using System;
using System.Collections.Generic;
//using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI.Muhasibat
{
    class cl_yanasmalar
    {
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public System.Data.DataTable dt_L2;
        public System.Data.DataTable dt_L3_A;
        public System.Data.DataTable dt_L3_B;
        public string baglanti = "Data Source=C:\\BMI_\\db_report.db;Version=3;";
        public string con = "DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass";
        public string con_odb = "Data Source=BMI;User ID=odb;Password=Lah$2021!";
        public string connStringUTimemaster = "Host=172.16.0.5;Port=7496;Username=postgres;Password=123456;Database=biotime";
        public string dosyayolu;
        public string textBoxText; // TextBox'tan alınan metni sakla
        public string yeniMetin;
        public string baseFileName; // Temel dosya adı
        public string fileName;
        public string templateFilePath;
        public string filePath;
        public int icraci;
    }

}
