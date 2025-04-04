using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BMI.Muhasibat;
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;
using System.Windows.Forms;
using DevExpress.Xpo.DB;
using DocumentFormat.OpenXml.Office.Word;
using Npgsql;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Filtering.Templates;

namespace BMI.Emek_haqqi_ve_davamiyyet.Classlar
{
    public class cl_Database
    {
        cl_yanasmalar cl = new cl_yanasmalar();

        public string today = DateTime.Now.ToString("dd-MM-yyyy");
        public void InsertIsciler(string saa, string pincode, string regnum, string ssn, string cariHesab,
                                  string department, string vezife, DateTime daxilOlmaTarixi)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "INSERT INTO odb.bmi_eh_isciler (SAA, PINCODE, REGNUM, SSN, CARI_HESAB, " +
                               "DEPARTAMENT, VEZIFE, DAXIL_OLMA_TARIXI) " +
                               "VALUES (:saa, :pincode, :regnum, :ssn, :cariHesab, " +
                               ":department, :vezife, TO_DATE(:daxilOlmaTarixi, 'DD-MM-YYYY'))";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":saa", OracleDbType.Varchar2).Value = saa;
                    cmd.Parameters.Add(":pincode", OracleDbType.Varchar2).Value = pincode;
                    cmd.Parameters.Add(":regnum", OracleDbType.Varchar2).Value = regnum;
                    cmd.Parameters.Add(":ssn", OracleDbType.Varchar2).Value = ssn;
                    cmd.Parameters.Add(":cariHesab", OracleDbType.Varchar2).Value = cariHesab;
                    cmd.Parameters.Add(":department", OracleDbType.Varchar2).Value = department;
                    cmd.Parameters.Add(":vezife", OracleDbType.Varchar2).Value = vezife;
                    //cmd.Parameters.Add(":daxilOlmaTarixi", OracleDbType.Date).Value = daxilOlmaTarixi;
                    cmd.Parameters.Add(":daxilOlmaTarixi", OracleDbType.Varchar2).Value = daxilOlmaTarixi.ToString("dd-MM-yyyy");

                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void InsertMezuniyyetler(string regnum, string mez_type, string emr_no, string saa, string emr_tar,
                                  string mez_gun, string mez_start, string mez_end,string evezedici,string ICRACI)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "INSERT INTO odb.bmi_eh_davamiyyet (REGNUM, MEZ_TYPE, EMR_NO, AD, EMR_TAR, " +
                               "MEZ_GUN, CREATED_AT, MEZ_START,MEZ_END,EVEZEDICI,ICRACI,ICRA_TAR) " +
                               "VALUES (:regnum, :mez_type, :emr_no, :ad, TO_DATE(:emr_tar, 'DD-MM-YYYY'), " +
                               ":mez_gun, TO_DATE(:creadet_tar, 'DD-MM-YYYY'), TO_DATE(:mez_start," +
                               " 'DD-MM-YYYY'),TO_DATE(:mez_end, 'DD-MM-YYYY'),:EVEZEDICI,:ICRACI,:ICRA_TAR)";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":regnum", OracleDbType.Varchar2).Value = regnum;
                    cmd.Parameters.Add(":mez_type", OracleDbType.Varchar2).Value = mez_type;
                    cmd.Parameters.Add(":emr_no", OracleDbType.Varchar2).Value = emr_no;
                    cmd.Parameters.Add(":ad", OracleDbType.Varchar2).Value = saa;
                    cmd.Parameters.Add(":emr_tar", OracleDbType.Varchar2).Value = emr_tar;
                    cmd.Parameters.Add(":mez_gun", OracleDbType.Varchar2).Value = mez_gun;
                    cmd.Parameters.Add(":creadet_tar", OracleDbType.Varchar2).Value = today;
                    cmd.Parameters.Add(":mez_start", OracleDbType.Varchar2).Value = mez_start;
                    cmd.Parameters.Add(":mez_end", OracleDbType.Varchar2).Value = mez_end;
                    cmd.Parameters.Add(":EVEZEDICI", OracleDbType.Varchar2).Value = evezedici;
                    cmd.Parameters.Add(":ICRACI", OracleDbType.Varchar2).Value = ICRACI;
                    cmd.Parameters.Add(":ICRA_TAR", OracleDbType.Varchar2).Value = DateTime.Now;

                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }
        //public void InsertMezuniyyetHesablama(string regnum, string emr_no, string emr_tar, string mez_type,
        //                          string mez_gun, decimal mez_meb, string odenis_sorgu, string icraci)
        //{
        //    string mez_start, mez_end;
        //    using (OracleConnection conn=new OracleConnection(cl.con_odb))
        //    {
        //        conn.Open();
        //        string query = @"INSERT INTO bmi_eh_mezuniyyet_cedveli (regnum, emr_no, emr_tar,mez_type, mez_gun, mez_meb, odenis_sorgu, icraci, icra_tar)
        //                        VALUES (:regnum,:emr_no,:emr_tar,:mez_type,:mez_gun,:mez_meb,:odenis_sorgu,:icraci,:icra_tar ";

        //        using (OracleCommand cmd = new OracleCommand(query, conn))
        //        {
        //            cmd.Parameters.Add("regnum", OracleDbType.Varchar2).Value=regnum;
        //            cmd.Parameters.Add("emr_no", OracleDbType.Varchar2).Value = emr_no;
        //            cmd.Parameters.Add("emr_tar", OracleDbType.Varchar2).Value = emr_tar;
        //            cmd.Parameters.Add("mez_type", OracleDbType.Varchar2).Value = mez_type;
        //            cmd.Parameters.Add("mez_gun", OracleDbType.Varchar2).Value = mez_gun;
        //            cmd.Parameters.Add("mez_meb", OracleDbType.Decimal).Value = mez_meb;
        //            cmd.Parameters.Add("odenis_sorgu", OracleDbType.Varchar2).Value = odenis_sorgu;
        //            cmd.Parameters.Add("icraci", OracleDbType.Varchar2).Value = icraci;
        //            cmd.Parameters.Add("icra_tar", OracleDbType.Varchar2).Value = DateTime.Now;

        //            cmd.ExecuteNonQuery();
        //        }
        //        conn.Close();   
        //    }
        //}

        public void InsertMezuniyyetHesablama(string regnum, string emr_no, string mez_start, string mez_end, decimal mez_meb, string odenis_sorgu, string icraci)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();

                // SQL sorğusu - hər ay neçə gün düşdüyünü tapır
                string query = @"
            WITH AyGunleri AS (
                SELECT 
                    TO_CHAR(TO_DATE(:mez_start, 'YYYY-MM-DD'), 'MM-YYYY') AS Ay,
                    EXTRACT(DAY FROM LAST_DAY(TO_DATE(:mez_start, 'YYYY-MM-DD'))) 
                    - EXTRACT(DAY FROM TO_DATE(:mez_start, 'YYYY-MM-DD')) + 1 AS Gunler
                FROM DUAL
                UNION ALL
                SELECT 
                    TO_CHAR(TO_DATE(:mez_end, 'YYYY-MM-DD'), 'MM-YYYY') AS Ay,
                    TO_DATE(:mez_end, 'YYYY-MM-DD') - TO_DATE(:mez_start, 'YYYY-MM-DD') + 1 AS Gunler
                FROM DUAL
            )
            SELECT * FROM AyGunleri";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add("mez_start", OracleDbType.Varchar2).Value = mez_start;
                    cmd.Parameters.Add("mez_end", OracleDbType.Varchar2).Value = mez_end;

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string emr_tar = reader["Ay"].ToString(); // Ay-İl formatında
                            int mez_gun = Convert.ToInt32(reader["Gunler"]); // Həmin ayın gün sayı

                            // Hər ay üçün INSERT əməliyyatı
                            using (OracleCommand insertCmd = new OracleCommand(@"
                        INSERT INTO bmi_eh_mezuniyyet_cedveli (regnum, emr_no, emr_tar, mez_gun, mez_meb, odenis_sorgu, icraci, icra_tar)
                        VALUES (:regnum, :emr_no, :emr_tar, :mez_gun, :mez_meb, :odenis_sorgu, :icraci, :icra_tar)", conn))
                            {
                                insertCmd.Parameters.Add("regnum", OracleDbType.Varchar2).Value = regnum;
                                insertCmd.Parameters.Add("emr_no", OracleDbType.Varchar2).Value = emr_no;
                                insertCmd.Parameters.Add("emr_tar", OracleDbType.Varchar2).Value = emr_tar;
                                insertCmd.Parameters.Add("mez_gun", OracleDbType.Int32).Value = mez_gun;
                                insertCmd.Parameters.Add("mez_meb", OracleDbType.Decimal).Value = mez_meb;
                                insertCmd.Parameters.Add("odenis_sorgu", OracleDbType.Varchar2).Value = odenis_sorgu;
                                insertCmd.Parameters.Add("icraci", OracleDbType.Varchar2).Value = icraci;
                                insertCmd.Parameters.Add("icra_tar", OracleDbType.Date).Value = DateTime.Now; // Bugünkü tarix

                                insertCmd.ExecuteNonQuery();
                            }
                        }
                    }
                }

                conn.Close();
            }
        }

        public void InsertMaasElave(string cari_il_ay, string saa, string cari_hesab, decimal avans, decimal mukafat,
                            decimal elave_eh, decimal mezuniyyet_haqqi, decimal xestelik_vereqesi,
                            decimal emr07, decimal komp_odenisi, decimal hediye, decimal msss,
                            decimal mad98_2_1, decimal mad98_2_3, decimal hyh)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "INSERT INTO bmi_eh_maas_elaveler " +
                               "(cari_il_ay, SAA, Cari_Hesab, Avans, Mukafat, elave_eh, Mezuniyyet_Haqqi, " +
                               "Xestelik_Vereqesi,Emr_Say_07, Komp_odenisi, Hediyye, MSSS, " +
                               "mad_98_2_1, mad_98_2_3, HYH) " +
                               "VALUES (:cari_il_ay, :saa, :cari_hesab, :avans, :mukafat, :elave_eh, " +
                               ":mezuniyyet_haqqi, :xestelik_vereqesi, :emr07, :komp_odenisi, :hediye, :msss, " +
                               ":mad98_2_1, :mad98_2_3, :hyh)";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":cari_il_ay", OracleDbType.Varchar2).Value = cari_il_ay;
                    cmd.Parameters.Add(":saa", OracleDbType.Varchar2).Value = saa;
                    cmd.Parameters.Add(":cari_hesab", OracleDbType.Varchar2).Value = cari_hesab;
                    cmd.Parameters.Add(":avans", OracleDbType.Decimal).Value = avans;
                    cmd.Parameters.Add(":mukafat", OracleDbType.Decimal).Value = mukafat;
                    cmd.Parameters.Add(":elave_eh", OracleDbType.Decimal).Value = elave_eh;
                    cmd.Parameters.Add(":mezuniyyet_haqqi", OracleDbType.Decimal).Value = mezuniyyet_haqqi;
                    cmd.Parameters.Add(":xestelik_vereqesi", OracleDbType.Decimal).Value = xestelik_vereqesi;
                    cmd.Parameters.Add(":emr07", OracleDbType.Decimal).Value = emr07;
                    cmd.Parameters.Add(":komp_odenisi", OracleDbType.Decimal).Value = komp_odenisi;
                    cmd.Parameters.Add(":hediye", OracleDbType.Decimal).Value = hediye;
                    cmd.Parameters.Add(":msss", OracleDbType.Decimal).Value = msss;
                    cmd.Parameters.Add(":mad98_2_1", OracleDbType.Decimal).Value = mad98_2_1;
                    cmd.Parameters.Add(":mad98_2_3", OracleDbType.Decimal).Value = mad98_2_3;
                    cmd.Parameters.Add(":hyh", OracleDbType.Decimal).Value = hyh;

                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void InsertExchange(string tarix, string valyuta, object n_alis, object n_satis,
                    object qn_alis, object qn_satis, string icraci, string status, string mesaj, int insertGroupNo)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                
                // INSERT əməliyyatı INSERT_GROUP_NO ilə birlikdə
                string query = "INSERT INTO bmi_kassa_kurs " +
                               "(TARIX, VALYUTA, N_ALIS, N_SATIS, QN_ALIS, QN_SATIS, " +
                               "ICRACI, APPROVAL, INSERT_GROUP_NO) " +
                               "VALUES (TO_DATE(:tarix, 'DD-MM-YYYY'), :valyuta, :n_alis, :n_satis, :qn_alis, :qn_satis, " +
                               ":icraci, :approval, :insertGroupNo)";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":tarix", OracleDbType.Varchar2).Value = tarix;
                    cmd.Parameters.Add(":valyuta", OracleDbType.Varchar2).Value = valyuta;
                    cmd.Parameters.Add(":n_alis", OracleDbType.Decimal).Value = n_alis ?? DBNull.Value;
                    cmd.Parameters.Add(":n_satis", OracleDbType.Decimal).Value = n_satis ?? DBNull.Value;
                    cmd.Parameters.Add(":qn_alis", OracleDbType.Decimal).Value = qn_alis ?? DBNull.Value;
                    cmd.Parameters.Add(":qn_satis", OracleDbType.Decimal).Value = qn_satis ?? DBNull.Value;
                    cmd.Parameters.Add(":icraci", OracleDbType.Varchar2).Value = icraci;
                    cmd.Parameters.Add(":approval", OracleDbType.Varchar2).Value = "gözləmədə";
                    cmd.Parameters.Add(":insertGroupNo", OracleDbType.Int32).Value = insertGroupNo;
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void UpdateExchange(string tarix, string valyuta, object n_alis, object n_satis,
                           object qn_alis, object qn_satis, string icraci, string status, string mesaj)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "UPDATE bmi_kassa_kurs SET " +
                               "N_ALIS = :n_alis, " +
                               "N_SATIS = :n_satis, " +
                               "QN_ALIS = :qn_alis, " +
                               "QN_SATIS = :qn_satis, " +
                               "ICRACI = :icraci, " +
                               "APPROVAL = :status, " +
                               "APPROVAL_DATE = SYSDATE, " + // Yeniləndikdə tarixi əlavə edirik
                               "mesaj = :mesaj " +
                               " WHERE TARIX = TO_DATE(:tarix, 'DD-MM-YYYY') AND VALYUTA = :valyuta and APPROVAL='imtina'";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    // Parametrləri əlavə edirik
                    cmd.Parameters.Add(":n_alis", OracleDbType.Decimal).Value = n_alis ?? DBNull.Value;
                    cmd.Parameters.Add(":n_satis", OracleDbType.Decimal).Value = n_satis ?? DBNull.Value;
                    cmd.Parameters.Add(":qn_alis", OracleDbType.Decimal).Value = qn_alis ?? DBNull.Value;
                    cmd.Parameters.Add(":qn_satis", OracleDbType.Decimal).Value = qn_satis ?? DBNull.Value;
                    cmd.Parameters.Add(":icraci", OracleDbType.Varchar2).Value = icraci;
                    cmd.Parameters.Add(":status", OracleDbType.Varchar2).Value = status;
                    cmd.Parameters.Add(":mesaj", OracleDbType.Varchar2).Value = mesaj;

                    // Şərtə əsasən tarix və valyutanı əlavə edirik
                    cmd.Parameters.Add(":tarix", OracleDbType.Varchar2).Value = tarix;
                    cmd.Parameters.Add(":valyuta", OracleDbType.Varchar2).Value = valyuta;

                    int rowsAffected = cmd.ExecuteNonQuery();

                }
                conn.Close();
            }
        }
        public void InsertIscininYeniEH(string regnum, string saa, string il,string ay_il, decimal eh_new, string created_at)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "INSERT INTO bmi_eh_isci_emekhaqqi_add (REGNUM, SAA, il, eh_new, CREATED_AT,il_ay) " +
                               "VALUES (:regnum, :saa, :il,:il_ay, :eh_new, TO_DATE(:created_at, 'DD-MM-YYYY'))";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    // Parametrləri əlavə edirik
                    cmd.Parameters.Add(":regnum", OracleDbType.Varchar2).Value = regnum;
                    cmd.Parameters.Add(":saa", OracleDbType.Varchar2).Value = saa;
                    cmd.Parameters.Add(":il", OracleDbType.Varchar2).Value = il;
                    cmd.Parameters.Add(":il_ay", OracleDbType.Varchar2).Value = ay_il;
                    cmd.Parameters.Add(":eh_new", OracleDbType.Decimal).Value = eh_new;
                    cmd.Parameters.Add(":created_at", OracleDbType.Varchar2).Value = today;

                    // Sorğunu icra edirik
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }
        public void InsertIscininIllikMez(string regnum, string saa, string il, string mez_new, string created_at)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "INSERT INTO bmi_eh_isci_mezuniyyet_add (REGNUM, SAA, il, mez_new, CREATED_AT) " +
                               "VALUES (:regnum, :saa, :il, :mez_new, TO_DATE(:created_at, 'DD-MM-YYYY'))";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    // Parametrləri əlavə edirik
                    cmd.Parameters.Add(":regnum", OracleDbType.Varchar2).Value = regnum;
                    cmd.Parameters.Add(":saa", OracleDbType.Varchar2).Value = saa;
                    cmd.Parameters.Add(":il", OracleDbType.Varchar2).Value = il;
                    cmd.Parameters.Add(":mez_new", OracleDbType.Varchar2).Value = mez_new;
                    cmd.Parameters.Add(":created_at", OracleDbType.Varchar2).Value = today;

                    // Sorğunu icra edirik
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }
        public List<Isciler> GetIsciler()
        {
            List<Isciler> iscilerList = new List<Isciler>();

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = @"SELECT i.SAA,i.PINCODE,i.REGNUM,i.SSN,i.CARI_HESAB,i.DEPARTAMENT,i.VEZIFE, 
                           TO_CHAR(i.DAXIL_OLMA_TARIXI, 'DD-MM-YYYY') AS DAXIL_OLMA_TARIXI,

                           (SELECT m.mez_new
                            FROM bmi_eh_isci_mezuniyyet_add m 
                            WHERE m.regnum = i.regnum 
                            ORDER BY m.mez_id DESC 
                            FETCH FIRST 1 ROWS ONLY) AS il_mez,

                           (SELECT e.eh_new 
                            FROM bmi_eh_isci_emekhaqqi_add e 
                            WHERE e.regnum = i.regnum 
                            ORDER BY e.eh_id DESC 
                            FETCH FIRST 1 ROWS ONLY) AS eh_emek,

                           (SELECT NVL(SUM(m.mez_new), 0)
                            FROM bmi_eh_isci_mezuniyyet_add m
                            WHERE m.regnum = i.regnum) 
                           - 
                           (SELECT NVL(SUM(d.mez_gun), 0)
                            FROM bmi_eh_davamiyyet d
                            WHERE d.regnum = i.regnum) AS qal_mez,
                            FLOOR(MONTHS_BETWEEN(SYSDATE, i.DAXIL_OLMA_TARIXI) / 12) || ' il ' ||
                                   MOD(FLOOR(MONTHS_BETWEEN(SYSDATE, i.DAXIL_OLMA_TARIXI)), 12) || ' ay' AS staj

                    FROM bmi_eh_isciler i";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Isciler isciler = new Isciler()
                            {
                                SAA = reader["SAA"].ToString(),
                                Pincode = reader["PINCODE"].ToString(),
                                Regnum = reader["REGNUM"].ToString(),
                                SSN = reader["SSN"].ToString(),
                                CariHesab = reader["CARI_HESAB"].ToString(),
                                Departament = reader["DEPARTAMENT"].ToString(),
                                Vezife = reader["VEZIFE"].ToString(),
                                DaxilOlmaTarixi = DateTime.ParseExact(reader["DAXIL_OLMA_TARIXI"].ToString(), "dd-MM-yyyy", null),
                                MezuniyyetGS = reader["il_mez"].ToString(),
                                EmekHaqqi = reader["EH_EMEK"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["EH_EMEK"]),
                                MezuniyyetQaliq = reader["qal_mez"].ToString(),
                                Staj = reader["staj"].ToString()
                            };
                            iscilerList.Add(isciler);
                        }
                    }
                }
                conn.Close();
            }
            return iscilerList;
        }
        public List<KassaKurs> GetKassaKursByDate(string selectedDate)
        {
            List<KassaKurs> kassaKursList = new List<KassaKurs>();

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = @"SELECT k.TARIX, k.VALYUTA, k.N_ALIS, k.N_SATIS, 
                                k.QN_ALIS, k.QN_SATIS, k.ICRACI, 
                                k.APPROVAL, k.APPROVAL_DATE, k.APPROVED_BY 
                         FROM bmi_kassa_kurs k 
                         WHERE k.TARIX = TO_DATE(:selectedDate, 'DD-MM-YYYY')";

                //string query = @"SELECT k.TARIX, k.VALYUTA, k.N_ALIS, k.N_SATIS, 
                //           k.QN_ALIS, k.QN_SATIS, k.ICRACI, 
                //           k.APPROVAL, k.APPROVAL_DATE, k.APPROVED_BY 
                //    FROM bmi_kassa_kurs k 
                //    ORDER BY k.INSERT_GROUP_NO asc";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":selectedDate", OracleDbType.Varchar2).Value = selectedDate;

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            KassaKurs kurs = new KassaKurs()
                            {
                                Tarix = reader["TARIX"] != DBNull.Value ? Convert.ToDateTime(reader["TARIX"]) : (DateTime?)null,
                                Valyuta = reader["VALYUTA"].ToString(),
                                NAlis = reader["N_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_ALIS"]) : (decimal?)null,
                                NSatis = reader["N_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_SATIS"]) : (decimal?)null,
                                QNAlis = reader["QN_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_ALIS"]) : (decimal?)null,
                                QNSatis = reader["QN_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_SATIS"]) : (decimal?)null,
                                Icraci = reader["ICRACI"].ToString(),
                                Approval = reader["APPROVAL"].ToString(),
                                ApprovalDate = reader["APPROVAL_DATE"] != DBNull.Value ? Convert.ToDateTime(reader["APPROVAL_DATE"]) : (DateTime?)null,
                                ApprovedBy = reader["APPROVED_BY"].ToString()
                            };

                            kassaKursList.Add(kurs);
                        }
                    }
                }
                conn.Close();
            }
            return kassaKursList;
        }
        public List<KassaKurs> GetKassaKurs()
        {
            List<KassaKurs> kassaKursList = new List<KassaKurs>();

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = @"SELECT k.TARIX, k.VALYUTA, k.N_ALIS, k.N_SATIS, 
                                k.QN_ALIS, k.QN_SATIS, k.ICRACI, 
                                k.APPROVAL, k.APPROVAL_DATE, k.APPROVED_BY,k.mesaj ,k.insert_group_no
                                FROM bmi_kassa_kurs k order by k.TARIX desc,
                                CASE Valyuta
                                WHEN 'USD' THEN 1
                                WHEN 'AVRO' THEN 2
                                WHEN 'RUB' THEN 3
                                WHEN 'AED' THEN 4
                                WHEN 'IRR' THEN 5
                                ELSE 6
                                END ";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            KassaKurs kurs = new KassaKurs()
                            {
                                Tarix = reader["TARIX"] != DBNull.Value ? Convert.ToDateTime(reader["TARIX"]) : (DateTime?)null,
                                Valyuta = reader["VALYUTA"].ToString(),
                                NAlis = reader["N_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_ALIS"]) : (decimal?)null,
                                NSatis = reader["N_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_SATIS"]) : (decimal?)null,
                                QNAlis = reader["QN_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_ALIS"]) : (decimal?)null,
                                QNSatis = reader["QN_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_SATIS"]) : (decimal?)null,
                                Icraci = reader["ICRACI"].ToString(),
                                Approval = reader["APPROVAL"].ToString(),
                                ApprovalDate = reader["APPROVAL_DATE"] != DBNull.Value ? Convert.ToDateTime(reader["APPROVAL_DATE"]) : (DateTime?)null,
                                ApprovedBy = reader["APPROVED_BY"].ToString(),
                                Mesaj = reader["MESAJ"].ToString(),
                                Countday = reader["insert_group_no"] != DBNull.Value ? Convert.ToInt16(reader["insert_group_no"]): 0
                            };

                            kassaKursList.Add(kurs);
                        }
                    }
                }
                conn.Close();
            }
            return kassaKursList;
        }
        public List<Isciler> GetIDavamiyyet()
        {
            List<Isciler> iscilerList = new List<Isciler>();

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = @"select e.regnum,e.ad,e.mez_type,e.emr_no,e.emr_tar,e.mez_gun,
                                e.created_at,e.mez_start,e.mez_end,e.evezedici,e.approval,e.approval_date,
                                e.approved_by from bmi_eh_davamiyyet e
                                ";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Isciler isciler = new Isciler()
                            {
                                SAA = reader["SAA"].ToString(),
                                Pincode = reader["PINCODE"].ToString(),
                                Regnum = reader["REGNUM"].ToString(),
                                SSN = reader["SSN"].ToString(),
                                CariHesab = reader["CARI_HESAB"].ToString(),
                                Departament = reader["DEPARTAMENT"].ToString(),
                                Vezife = reader["VEZIFE"].ToString(),
                                DaxilOlmaTarixi = DateTime.ParseExact(reader["DAXIL_OLMA_TARIXI"].ToString(), "dd-MM-yyyy", null),
                                MezuniyyetGS = reader["il_mez"].ToString(),
                                EmekHaqqi = reader["EH_EMEK"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["EH_EMEK"]),
                                MezuniyyetQaliq = reader["qal_mez"].ToString(),
                                Staj = reader["staj"].ToString()
                            };
                            iscilerList.Add(isciler);
                        }
                    }
                }
                conn.Close();
            }
            return iscilerList;
        }
    }
}
