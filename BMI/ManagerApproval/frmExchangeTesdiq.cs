using System;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using BMI.Muhasibat;
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using DevExpress.CodeParser;
using System.Threading.Tasks;
using System.Diagnostics;
using DevExpress.DataProcessing.InMemoryDataProcessor;




namespace BMI.ManagerApproval
{
    public partial class frmExchangeTesdiq : Form
    {
        public frmExchangeTesdiq()
        {
            InitializeComponent();
        }
        cl_Database db = new cl_Database();
        cl_yanasmalar cl = new cl_yanasmalar();

        public string icraci_ad { get; set; }

        private void LoadDataFromExchange()
        {
            string query = "SELECT * FROM bmi_kassa_kurs k WHERE k.approval = 'gözləmədə'";

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            // USD üçün məlumatları doldururuq
                            if (reader["VALYUTA"].ToString() == "USD")
                            {
                                txtNAUSD.Text = reader["N_ALIS"]?.ToString();
                                txtNSUSD.Text = reader["N_SATIS"]?.ToString();
                                txtQNAUSD.Text = reader["QN_ALIS"]?.ToString();
                                txtQNSUSD.Text = reader["QN_SATIS"]?.ToString();
                            }

                            // AVRO üçün məlumatları doldururuq
                            else if (reader["VALYUTA"].ToString() == "AVRO")
                            {
                                txtNAAVRO.Text = reader["N_ALIS"]?.ToString();
                                txtNSAVRO.Text = reader["N_SATIS"]?.ToString();
                                txtQNAAVRO.Text = reader["QN_ALIS"]?.ToString();
                                txtQNSAVRO.Text = reader["QN_SATIS"]?.ToString();
                            }

                            // IRR üçün məlumatları doldururuq
                            else if (reader["VALYUTA"].ToString() == "IRR")
                            {
                                txtNAIRR.Text = reader["N_ALIS"]?.ToString();
                                txtNSIRR.Text = reader["N_SATIS"]?.ToString();
                                txtQNAIRR.Text = reader["QN_ALIS"]?.ToString();
                                txtQNSIRR.Text = reader["QN_SATIS"]?.ToString();
                            }

                            // AED üçün məlumatları doldururuq
                            else if (reader["VALYUTA"].ToString() == "AED")
                            {
                                txtNAAED.Text = reader["N_ALIS"]?.ToString();
                                txtNSAED.Text = reader["N_SATIS"]?.ToString();
                                txtQNAAED.Text = reader["QN_ALIS"]?.ToString();
                                txtQNSAED.Text = reader["QN_SATIS"]?.ToString();
                            }

                            // RUB üçün məlumatları doldururuq
                            else if (reader["VALYUTA"].ToString() == "RUB")
                            {
                                txtNARUB.Text = reader["N_ALIS"]?.ToString();
                                txtNSRUB.Text = reader["N_SATIS"]?.ToString();
                                txtQNARUB.Text = reader["QN_ALIS"]?.ToString();
                                txtQNSRUB.Text = reader["QN_SATIS"]?.ToString();
                            }
                        }
                    }
                }
            }
        }

        private void LoadDataByDate(string tarix)
        {
            string query = "SELECT VALYUTA, N_ALIS, N_SATIS, QN_ALIS, QN_SATIS ,approval "+
                          " FROM bmi_kassa_kurs WHERE TARIX = TO_DATE(:tarix, 'DD-MM-YYYY') and INSERT_GROUP_NO = ( " +
                         "       SELECT MAX(INSERT_GROUP_NO) FROM bmi_kassa_kurs "+
                          "  )";

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
                                string n_alis = reader["N_ALIS"] != DBNull.Value ? reader["N_ALIS"].ToString() : "";
                                string n_satis = reader["N_SATIS"] != DBNull.Value ? reader["N_SATIS"].ToString() : "";
                                string qn_alis = reader["QN_ALIS"] != DBNull.Value ? reader["QN_ALIS"].ToString() : "";
                                string qn_satis = reader["QN_SATIS"] != DBNull.Value ? reader["QN_SATIS"].ToString() : "";
                                string approval = reader["approval"].ToString();


                                // Formdakı TextBox-ları doldururuq
                                switch (valyuta)
                                {
                                    case "USD":
                                        txtNAUSD.Text = n_alis;
                                        txtNSUSD.Text = n_satis;
                                        txtQNAUSD.Text = qn_alis;
                                        txtQNSUSD.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        break;

                                    case "AVRO":
                                        txtNAAVRO.Text = n_alis;
                                        txtNSAVRO.Text = n_satis;
                                        txtQNAAVRO.Text = qn_alis;
                                        txtQNSAVRO.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        break;

                                    case "IRR":
                                        txtNAIRR.Text = n_alis;
                                        txtNSIRR.Text = n_satis;
                                        txtQNAIRR.Text = qn_alis;
                                        txtQNSIRR.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        break;

                                    case "AED":
                                        txtNAAED.Text = n_alis;
                                        txtNSAED.Text = n_satis;
                                        txtQNAAED.Text = qn_alis;
                                        txtQNSAED.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        break;

                                    case "RUB":
                                        txtNARUB.Text = n_alis;
                                        txtNSRUB.Text = n_satis;
                                        txtQNARUB.Text = qn_alis;
                                        txtQNSRUB.Text = qn_satis;
                                        lbl_status.Text = approval;
                                        break;
                                }
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

        private void LoadDataExchangeToGrid(string tarix)
        {
            List<KassaKurs> Kurslar = db.GetKassaKursByDate(tarix); // Oracle-dan məlumatları alırıq

            dtg_Exchange.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var kurslar in Kurslar)
            {
                dtg_Exchange.Rows.Add(kurslar.Tarix.HasValue ? kurslar.Tarix.Value.ToString("dd-MM-yyyy") : "", kurslar.Valyuta, kurslar.NAlis, kurslar.NSatis, kurslar.QNAlis, kurslar.QNSatis,
                    kurslar.Icraci, kurslar.Approval, kurslar.ApprovalDate, kurslar.ApprovedBy, kurslar.Mesaj);
            }
        }

        private void LoadDataExchangeToGridAll()
        {
            List<KassaKurs> Kurslar = db.GetKassaKurs(); // Oracle-dan məlumatları alırıq

            dtg_Exchange.Rows.Clear(); // Köhnə məlumatları silirik

            foreach (var kurslar in Kurslar)
            {
                dtg_Exchange.Rows.Add(kurslar.Tarix.HasValue ? kurslar.Tarix.Value.ToString("dd-MM-yyyy") : "", kurslar.Valyuta, kurslar.NAlis, kurslar.NSatis, kurslar.QNAlis, kurslar.QNSatis,
                    kurslar.Icraci, kurslar.Approval, kurslar.ApprovalDate, kurslar.ApprovedBy,kurslar.Mesaj);
            }
        }

        private void Approval(string approvedByUser)
        {
            string query = "UPDATE bmi_kassa_kurs " +
               "SET APPROVAL = 'təsdiqləndi', " +
               "APPROVAL_DATE = SYSDATE, " +
               "APPROVED_BY = :approved_by " +
               "WHERE APPROVAL = 'gözləmədə'";

            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                try
                {
                    conn.Open();
                    using (OracleCommand cmd = new OracleCommand(query, conn))
                    {
                        cmd.Parameters.Add(":approved_by", OracleDbType.Varchar2).Value = approvedByUser;
                        int rowsAffected = cmd.ExecuteNonQuery(); // Dəyişdirilən sətrlərin sayını alırıq

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Məlumatlar qeyd təsdiqləndi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Təsdiqlənəcək qeyd tapılmadı.", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmExchangeTesdiq_Load(object sender, EventArgs e)
        {
            lbl_tarix.Text = DateTime.Now.ToString("dd-MM-yyyy");
            LoadDataByDate(lbl_tarix.Text);
            LoadDataFromExchange();
            LoadDataExchangeToGridAll();
        }

        private void btn_Elave_Click(object sender, EventArgs e)
        {
            Approval(icraci_ad);
            LoadDataByDate(lbl_tarix.Text);
            LoadDataFromExchange();
            LoadDataExchangeToGridAll();
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

        private void button1_Click(object sender, EventArgs e)
        {
            // Əvvəlcə "gözləmədə" olan kursun olub-olmadığını yoxlayırıq
            if (!CheckPendingCourse())
            {
                MessageBox.Show("Gözləmədə olan kurs tapılmadı və ya yenilənmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // Əgər kurs tapılmadısa, davam etmirik
            }

            // İstifadəçidən imtina səbəbini alırıq
            string imtinaSebebi = Interaction.InputBox("İmtina səbəbini daxil edin:", "İmtina", "");

            if (!string.IsNullOrWhiteSpace(imtinaSebebi))
            {
                // Bazaya yazırıq
                UpdateImtinaSebebi(imtinaSebebi, icraci_ad);

                MessageBox.Show("İmtina səbəbi uğurla qeyd edildi.", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDataFromExchange();
                LoadDataByDate(lbl_tarix.Text);
                LoadDataExchangeToGridAll();
            }
            else
            {
                MessageBox.Show("İmtina səbəbi daxil edilmədi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private bool CheckPendingCourse()
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "SELECT COUNT(*) FROM bmi_kassa_kurs WHERE APPROVAL = 'gözləmədə'";
                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0; // Əgər 1 və ya daha çox "gözləmədə" olan kurs varsa, true qaytarır
                }
            }
        }
        private void UpdateImtinaSebebi(string imtinaSebebi, string approvedByUser)
        {
            using (OracleConnection conn = new OracleConnection(cl.con_odb))
            {
                conn.Open();
                string query = "UPDATE bmi_kassa_kurs " +
                               "SET APPROVAL = 'imtina', " +
                               "APPROVAL_DATE = SYSDATE, " +
                               "APPROVED_BY = :approved_by, " +
                               "MESAJ = :imtinaSebebi " +
                               "WHERE APPROVAL = 'gözləmədə'";

                using (OracleCommand cmd = new OracleCommand(query, conn))
                {
                    cmd.Parameters.Add(":approved_by", OracleDbType.Varchar2).Value = approvedByUser;
                    cmd.Parameters.Add(":imtinaSebebi", OracleDbType.Varchar2).Value = imtinaSebebi;

                    cmd.ExecuteNonQuery(); // Artıq burada rowAffected yoxlaması lazım deyil
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Açılacaq saytın linki
            string url = "https://azn.day.az/az/";

            // Chrome-un tam yolunu qeyd edin
            string chromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";

            try
            {
                // Chrome ilə saytı açır
                Process.Start(chromePath, url);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message);
            }
        }
    }
}
