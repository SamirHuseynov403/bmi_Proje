using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BMI.Emek_haqqi_ve_davamiyyet.Classlar
{
    public class cl_isciler
    {
        public string icraci_kod { get; set; }
        public string icraci_ad { get; set; }
        public class Isciler
        {
            public string Regnum { get; set; }
            public string SAA { get; set; }
            public string Pincode { get; set; }
            public string SSN { get; set; }
            public string CariHesab { get; set; }
            public string Departament { get; set; }
            public string Vezife { get; set; }
            public DateTime DaxilOlmaTarixi { get; set; }
            public string MezuniyyetGS { get; set; }
            public decimal EmekHaqqi { get; set; }
            public string MezuniyyetQaliq { get; set; }
            public string Staj { get; set; }

        }
        public class KassaKurs
        {
            public DateTime? Tarix { get; set; }
            public string Valyuta { get; set; }
            public decimal? NAlis { get; set; }
            public decimal? NSatis { get; set; }
            public decimal? QNAlis { get; set; }
            public decimal? QNSatis { get; set; }
            public string Icraci { get; set; }
            public string Approval { get; set; }
            public DateTime? ApprovalDate { get; set; }
            public string ApprovedBy { get; set; }
            public string Mesaj { get; set; }
            public int Countday { get; set; }
        }
        public class InputOutputWorkers
        {
            public string Regnom { get; set; }
            public string ASA { get; set; }
            public DateTime WorkDay { get; set; }
            public TimeSpan CheckInTime { get; set; }
            public TimeSpan CheckOutTime { get; set; }
            public TimeSpan DelayTime { get; set; }

        }
    }
}
