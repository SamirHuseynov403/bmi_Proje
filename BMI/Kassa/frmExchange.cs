using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using BMI.Muhasibat;
using DevExpress.CodeParser;
using DevExpress.Drawing.TextFormatter.Internal;
using Microsoft.Office.Interop.Word;
using Oracle.ManagedDataAccess.Client;
using Org.BouncyCastle.Asn1.Cmp;
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;
using Word = Microsoft.Office.Interop.Word;

namespace BMI.Kassa
{
    public partial class frmExchange : Form
    {
        public frmExchange()
        {
            InitializeComponent();
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        cl_Database db = new cl_Database();
        Aletler aletler = new Aletler();
        string qovluqyolu = Aletler.Layiheanaqovluq();
        public string icraci_kod { get; set; }
        public string icraci_ad { get; set; }
        private void GetApprovalStatus(string tarix)
        {
            string query = @"SELECT k.approval, k.mesaj
                        FROM bmi_kassa_kurs k
                        WHERE k.INSERT_GROUP_NO = (
                            SELECT MAX(INSERT_GROUP_NO) FROM bmi_kassa_kurs
                        )";

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                try
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        // Tarixi parametr kimi əlavə edirik
                        cmd.Parameters.Add(":tarix", OracleDbType.Varchar2).Value = tarix;

                        using (OracleDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Nəticələri müvafiq label-lara yazırıq
                                lbl_status.Text = reader["approval"] != DBNull.Value ? reader["approval"].ToString() : "";
                                lbl_Imtina1.Text = reader["mesaj"] != DBNull.Value ? reader["mesaj"].ToString() : "";
                                txtM_Imtina.Text = reader["mesaj"] != DBNull.Value ? reader["mesaj"].ToString() : "";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private bool CheckIfDateExists()
        {
            string query = @"SELECT count(*)
                FROM bmi_kassa_kurs k
                WHERE k.INSERT_GROUP_NO = (
                    SELECT MAX(INSERT_GROUP_NO) FROM bmi_kassa_kurs
                ) and (k.approval='imtina' or k.approval='gözləmədə')";

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                try
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        int count = Convert.ToInt32(cmd.ExecuteScalar());

                        return count > 0; // Əgər 0-dan böyükdürsə, tarix bazada mövcuddur
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }
        private object ParseDecimalOrNull(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return DBNull.Value;

            if (decimal.TryParse(text, out decimal result))
                return result;

            return DBNull.Value; // Yanlış formatda dəyər daxil edilsə də NULL göndərilir
        }
        private void DisableControls(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is TextBox textBox)
                {
                    textBox.ReadOnly = true; // TextBox üçün ReadOnly
                }
                else if (ctrl is RichTextBox richTextBox)
                {
                    richTextBox.ReadOnly = true; // RichTextBox üçün ReadOnly
                }
                else if (ctrl is ComboBox || ctrl is Button)
                {
                    ctrl.Enabled = false; // ComboBox və Button-lar üçün Enabled false
                }

                // Əgər ctrl-in içində başqa kontrollerlər varsa, onları da yoxla
                if (ctrl.HasChildren)
                {
                    DisableControls(ctrl);
                }
                btn_Yenile.Enabled = true;
            }
        }
        private void AccessUser()
        {
            string userRole = UserAccess.GetUserRoleExchang(Convert.ToInt32(icraci_kod));
            if (userRole == "ReadOnly")
            {
                DisableControls(this);
            }
            else if (userRole == "NoAccess")
            {
                MessageBox.Show("Sizin giriş icazəniz yoxdur.", "Bloklanmış giriş", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close(); // Formu bağlayırıq
            }
        }
        private void InsertRecord()
        {
            try
            {
                string tarix = lbl_tarix.Text; // Tarixi əldə edin
                int insertGroupNo = 1;

                using (OracleConnection conn = new OracleConnection(cl.con_odb))
                {
                    conn.Open();
                    string selectQuery = "SELECT NVL(MAX(INSERT_GROUP_NO), 0) + 1 FROM bmi_kassa_kurs";
                    using (OracleCommand selectCmd = new OracleCommand(selectQuery, conn))
                    {
                        object result = selectCmd.ExecuteScalar();
                        insertGroupNo = Convert.ToInt32(result);
                    }
                }

                // USD üçün məlumat
                db.InsertExchange(tarix, "USD",
                           ParseDecimalOrNull(txtNAUSD.Text),
                           ParseDecimalOrNull(txtNSUSD.Text),
                           ParseDecimalOrNull(txtQNAUSD.Text),
                           ParseDecimalOrNull(txtQNSUSD.Text), icraci_ad, "gözləmədə", "", insertGroupNo);

                // AVRO üçün məlumat
                db.InsertExchange(tarix, "AVRO",
                           ParseDecimalOrNull(txtNAAVRO.Text),
                           ParseDecimalOrNull(txtNSAVRO.Text),
                           ParseDecimalOrNull(txtQNAAVRO.Text),
                           ParseDecimalOrNull(txtQNSAVRO.Text), icraci_ad, "gözləmədə", "", insertGroupNo);

                // IRR üçün məlumat
                db.InsertExchange(tarix, "IRR",
                           ParseDecimalOrNull(txtNAIRR.Text),
                           ParseDecimalOrNull(txtNSIRR.Text),
                           ParseDecimalOrNull(txtQNAIRR.Text),
                           ParseDecimalOrNull(txtQNSIRR.Text), icraci_ad, "gözləmədə", "", insertGroupNo);

                // AED üçün məlumat
                db.InsertExchange(tarix, "AED",
                           ParseDecimalOrNull(txtNAAED.Text),
                           ParseDecimalOrNull(txtNSAED.Text),
                           ParseDecimalOrNull(txtQNAAED.Text),
                           ParseDecimalOrNull(txtQNSAED.Text), icraci_ad, "gözləmədə", "", insertGroupNo);

                // RUB üçün məlumat
                db.InsertExchange(tarix, "RUB",
                           ParseDecimalOrNull(txtNARUB.Text),
                           ParseDecimalOrNull(txtNSRUB.Text),
                           ParseDecimalOrNull(txtQNARUB.Text),
                           ParseDecimalOrNull(txtQNSRUB.Text), icraci_ad, "gözləmədə", "", insertGroupNo);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdateRecord()
        {
            try
            {
                string tarix = lbl_tarix.Text; // Tarixi əldə edin

                // USD üçün məlumat
                db.UpdateExchange(tarix, "USD",
                               ParseDecimalOrNull(txtNAUSD.Text),
                               ParseDecimalOrNull(txtNSUSD.Text),
                               ParseDecimalOrNull(txtQNAUSD.Text),
                               ParseDecimalOrNull(txtQNSUSD.Text), icraci_ad, "gözləmədə", "");

                // AVRO üçün məlumat
                db.UpdateExchange(tarix, "AVRO",
                               ParseDecimalOrNull(txtNAAVRO.Text),
                               ParseDecimalOrNull(txtNSAVRO.Text),
                               ParseDecimalOrNull(txtQNAAVRO.Text),
                               ParseDecimalOrNull(txtQNSAVRO.Text), icraci_ad, "gözləmədə", "");

                // IRR üçün məlumat
                db.UpdateExchange(tarix, "IRR",
                               ParseDecimalOrNull(txtNAIRR.Text),
                               ParseDecimalOrNull(txtNSIRR.Text),
                               ParseDecimalOrNull(txtQNAIRR.Text),
                               ParseDecimalOrNull(txtQNSIRR.Text), icraci_ad, "gözləmədə", "");

                // AED üçün məlumat
                db.UpdateExchange(tarix, "AED",
                               ParseDecimalOrNull(txtNAAED.Text),
                               ParseDecimalOrNull(txtNSAED.Text),
                               ParseDecimalOrNull(txtQNAAED.Text),
                               ParseDecimalOrNull(txtQNSAED.Text), icraci_ad, "gözləmədə", "");

                // RUB üçün məlumat
                db.UpdateExchange(tarix, "RUB",
                               ParseDecimalOrNull(txtNARUB.Text),
                               ParseDecimalOrNull(txtNSRUB.Text),
                               ParseDecimalOrNull(txtQNARUB.Text),
                               ParseDecimalOrNull(txtQNSRUB.Text), icraci_ad, "gözləmədə", "");

                MessageBox.Show("Məlumat uğurla dəyişdirildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void wordeat()
        {
            
        }
        private void LoadDataByDate(string tarix)
        {
            //string query = "SELECT VALYUTA, N_ALIS, N_SATIS, QN_ALIS, QN_SATIS ,approval,mesaj  " +
            //               "FROM bmi_kassa_kurs WHERE TARIX = TO_DATE(:tarix, 'DD-MM-YYYY')";

            string query = @"SELECT VALYUTA, N_ALIS, N_SATIS, QN_ALIS, QN_SATIS ,approval,mesaj
                            FROM bmi_kassa_kurs k
                            WHERE k.INSERT_GROUP_NO = (
                                SELECT MAX(INSERT_GROUP_NO) FROM bmi_kassa_kurs
                            ) AND TRUNC(k.TARIX) = TRUNC(SYSDATE)";

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                try
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        cmd.Parameters.Add(":tarix", OracleDbType.Varchar2).Value = tarix;

                        using (OracleDataReader reader = cmd.ExecuteReader())
                        {
                            bool dataFound = false;

                            while (reader.Read())
                            {
                                dataFound = true;
                                string valyuta = reader["VALYUTA"].ToString();
                                string n_alis = reader["N_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_ALIS"]).ToString("F4") : "";
                                string n_satis = reader["N_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["N_SATIS"]).ToString("F4") : "";
                                string qn_alis = reader["QN_ALIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_ALIS"]).ToString("F4") : "";
                                string qn_satis = reader["QN_SATIS"] != DBNull.Value ? Convert.ToDecimal(reader["QN_SATIS"]).ToString("F4") : "";

                                string approval = reader["approval"].ToString();
                                string mesaj = reader["mesaj"].ToString();


                                // Formdakı TextBox-ları doldururuq
                                switch (valyuta)
                                {
                                    case "USD":
                                        txtNAUSD.Text = n_alis;
                                        txtNSUSD.Text = n_satis;
                                        txtQNAUSD.Text = qn_alis;
                                        txtQNSUSD.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        lbl_Imtina1.Text = mesaj;
                                        txtM_Imtina.Text = mesaj;
                                        break;

                                    case "AVRO":
                                        txtNAAVRO.Text = n_alis;
                                        txtNSAVRO.Text = n_satis;
                                        txtQNAAVRO.Text = qn_alis;
                                        txtQNSAVRO.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        lbl_Imtina1.Text = mesaj;
                                        txtM_Imtina.Text = mesaj;
                                        break;

                                    case "IRR":
                                        txtNAIRR.Text = n_alis;
                                        txtNSIRR.Text = n_satis;
                                        txtQNAIRR.Text = qn_alis;
                                        txtQNSIRR.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        lbl_Imtina1.Text = mesaj;
                                        txtM_Imtina.Text = mesaj;
                                        break;

                                    case "AED":
                                        txtNAAED.Text = n_alis;
                                        txtNSAED.Text = n_satis;
                                        txtQNAAED.Text = qn_alis;
                                        txtQNSAED.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        lbl_Imtina1.Text = mesaj;
                                        txtM_Imtina.Text = mesaj;
                                        break;

                                    case "RUB":
                                        txtNARUB.Text = n_alis;
                                        txtNSRUB.Text = n_satis;
                                        txtQNARUB.Text = qn_alis;
                                        txtQNSRUB.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        lbl_Imtina1.Text = mesaj;
                                        txtM_Imtina.Text = mesaj;
                                        break;
                                }
                            }

                            //if (!dataFound)
                            //{
                            //    MessageBox.Show("Bu tarix üçün məlumat tapılmadı.", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //}
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void LoadDataExchangeToGrid(string tarix)
        {
            List<KassaKurs> Kurslar = db.GetKassaKursByDate(tarix); // Oracle-dan məlumatları alırıq

            dtg_Exchange.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var kurslar in Kurslar)
            {
                dtg_Exchange.Rows.Add(kurslar.Tarix.HasValue ? kurslar.Tarix.Value.ToString("dd-MM-yyyy") : "", kurslar.Valyuta, kurslar.NAlis, kurslar.NSatis, kurslar.QNAlis, kurslar.QNSatis,
                    kurslar.Icraci, kurslar.Approval, kurslar.ApprovalDate, kurslar.ApprovedBy,kurslar.Countday);
            }
        }
        private void LoadDataExchangeToGridAll()
        {
            List<KassaKurs> Kurslar = db.GetKassaKurs(); // Oracle-dan məlumatları alırıq

            dtg_Exchange.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var kurslar in Kurslar)
            {
                dtg_Exchange.Rows.Add(kurslar.Tarix.HasValue ? kurslar.Tarix.Value.ToString("dd-MM-yyyy") : "",
                    kurslar.Valyuta, kurslar.NAlis.HasValue ? kurslar.NAlis.Value.ToString("F4") : "0.0000",
                    kurslar.NSatis.HasValue ? kurslar.NSatis.Value.ToString("F4") : "0.0000",
                    kurslar.QNAlis .HasValue ? kurslar.QNAlis.Value.ToString("F4") : "0.0000",
                    kurslar.QNSatis.HasValue ? kurslar.QNSatis.Value.ToString("F4") : "0.0000",
                    kurslar.Icraci, kurslar.Approval, kurslar.ApprovalDate, kurslar.ApprovedBy,kurslar.Countday);
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime today = DateTime.Now.Date;
            //btn_sorgu.Enabled = true;
            // Bugünkü və gələcək tarixlər seçilibsə:
            if (e.Start.Date > today.Date)
            {
                MessageBox.Show("Gələcək tarixləri seçmək mümkün deyil!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Seçimi geri qaytarırıq (istifadəçiyə xəbərdarlıq veririk və seçimi keçmiş tarixə qaytarırıq).
                monthCalendar1.SetSelectionRange(today.AddDays(-1), today.AddDays(-1)); // Dünənki tarixi seçili edir
                //btn_sorgu.Enabled = false;
            }
            else
            {
                // MonthCalendar-dan seçilən tarixi al
                DateTime selectedDate = monthCalendar1.SelectionStart;
                string formattedDate = selectedDate.ToString("dd-MM-yyyy"); // Oracle üçün uyğun tarix formatı

                // Oracle sorğusu
                string query = @"
        SELECT c.space_or_star AS gun
        FROM odb.calendar c
        WHERE c.date_oper = TO_DATE(:selectedDate, 'DD-MM-YYYY')";

                using (OracleConnection connection = new OracleConnection(cl.con))
                {
                    try
                    {

                        connection.Open();
                        using (OracleCommand command = new OracleCommand(query, connection))
                        {
                            // Tarixi parametr olaraq ötür
                            command.Parameters.Add(new OracleParameter("selectedDate", formattedDate));

                            // Sorğunu icra et
                            object result = command.ExecuteScalar();

                            if (result != null && result.ToString() == "*")
                            {
                                //btn_sorgu.Enabled = false;
                                // Əgər nəticə "*" isə, qeyri-iş günü mesajını göstər
                                MessageBox.Show("Seçdiyiniz tarix qeyri-iş günüdür!", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            }
                            else
                            {
                                LoadDataExchangeToGrid(formattedDate);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Xəta baş verdikdə mesaj göstər
                        MessageBox.Show($"Xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void ButonSorgu()
        {
            if (CheckIfDateExists())
            {
                btn_Elave.Text = "Düzəliş et";
            }
            else
            {
                btn_Elave.Text = "İcra et";
            }
        }
        private void btn_Elave_Click(object sender, EventArgs e)
        {
            if (CheckIfDateExists())
            {
                UpdateRecord();
            }
            else 
            {
                InsertRecord();
            }
            
            GetApprovalStatus(lbl_tarix.Text);
            LoadDataExchangeToGridAll();
        }
        private void frmExchange_Load(object sender, EventArgs e)
        {
            AccessUser();
            lbl_tarix.Text = DateTime.Now.ToString("dd-MM-yyyy");
            LoadDataByDate(lbl_tarix.Text);
            LoadDataExchangeToGridAll();
            ButonSorgu();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            LoadDataByDate(lbl_tarix.Text);
            LoadDataExchangeToGridAll();
            ButonSorgu();
        }
        private void dtg_Exchange_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            cl.dosyayolu = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis exceller");

            if (e.ColumnIndex == 11 && e.RowIndex >= 0) // Butonun olduğu sütun indeksi
            {
                // Seçilmiş tarixi götürürük
                string selectedcount = dtg_Exchange.Rows[e.RowIndex].Cells[10].Value.ToString();
                string selectedDate = dtg_Exchange.Rows[e.RowIndex].Cells[0].Value.ToString();

                int count = 1;
                var replacements = new Dictionary<string, string>();
                replacements["{tar}"] = Aletler.TarixiSozeCevir(selectedDate);

                // BÜTÜN SƏTİRLƏRİ YOXLAYIRIQ (yalnız startIndex-dən deyil)
                for (int i = 0; i < dtg_Exchange.Rows.Count; i++)
                {
                    if (dtg_Exchange.Rows[i].Cells[0].Value?.ToString() == selectedDate && dtg_Exchange.Rows[i].Cells[10].Value?.ToString() == selectedcount)
                    {
                        for (int j = 1; j < 6; j++)
                        {
                            // Dinamik repl yaradılır
                            string repl = "{v" + count + "}";

                            // replacements-ə məlumat əlavə edilir
                            replacements[repl] = dtg_Exchange.Rows[i].Cells[j].Value?.ToString() ?? "";

                            count++; // Sayğacı artırırıq ki, açarlar unikal olsun
                        }
                    }
                }

                // Faylın adını yoxlayırıq, əgər mövcuddursa yeni ad veririk
                if (File.Exists(System.IO.Path.Combine(cl.dosyayolu, selectedDate)))
                {
                    int fileCounter = 1;
                    while (File.Exists(System.IO.Path.Combine(cl.dosyayolu, $"{cl.baseFileName} - {fileCounter}.doc")))
                    {
                        fileCounter++;
                    }
                    selectedDate = $"{selectedDate} - {fileCounter}.doc";
                }

                // Fayl yollarını təyin edirik
                string kassaWord = System.IO.Path.Combine(qovluqyolu, "Fayllar", "Kassa", "Wordler", "Exchange.doc");
                string outputPath = System.IO.Path.Combine(qovluqyolu, aletler.sorgular, "Yaradilmis wordler", selectedDate);

                // Word sənədini yaradın
                aletler.CreateWordDocument(kassaWord, outputPath, replacements);
            }


        }

    }
    }

