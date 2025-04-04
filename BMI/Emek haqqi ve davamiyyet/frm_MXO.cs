using BMI.Emek_haqqi_ve_davamiyyet.Classlar;
using BMI.Muhasibat;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Windows.Forms;
using static BMI.Emek_haqqi_ve_davamiyyet.Classlar.cl_isciler;

namespace BMI.Emek_haqqi_ve_davamiyyet
{
    public partial class frm_MXO : Form
    {
        public frm_MXO()
        {
            InitializeComponent();
        }
        cl_isciler cl = new cl_isciler();
        cl_Database db = new cl_Database();
        cl_yanasmalar cly = new cl_yanasmalar();
        private void InsertMezuniyyetler()
        {
            try
            {
                string regnum = txb_QeydNo.Text;
                string mez_type = cmb_MTip.Text;
                string emr_no = txb_EmrNo.Text;
                string saa = txb_SAA.Text;
                string mez_gun = txb_MSay.Text;
                string evezedici = cmb_isci_siyahisi.Text;


                DateTime emr_tar, mez_start, mez_end;

                string[] dateFormats = { "dd-MM-yyyy", "d-MM-yyyy", "dd-M-yyyy", "d-M-yyyy", "yyyy-MM-dd", "MM/dd/yyyy" };

                if (!DateTime.TryParseExact(txb_EmrTar.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out emr_tar))
                {
                    MessageBox.Show("Əmr tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!DateTime.TryParseExact(txb_MDovrStart.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out mez_start))
                {
                    MessageBox.Show("Başlanğıc tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!DateTime.TryParseExact(txb_MDovrEnd.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out mez_end))
                {
                    MessageBox.Show("Bitmə tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // SQL-ə tarixləri düzgün formatda göndəririk
                string emr_tar_sql = emr_tar.ToString("dd-MM-yyyy");
                string mez_start_sql = mez_start.ToString("dd-MM-yyyy");
                string mez_end_sql = mez_end.ToString("dd-MM-yyyy");


                db.InsertMezuniyyetler(regnum, mez_type, emr_no, saa, emr_tar_sql, mez_gun, mez_start_sql, mez_end_sql, evezedici, cl.icraci_ad);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void InsertMezuniyyetHesablama()
        {
            try
            {
                string regnum = txb_QeydNo.Text;
                string mez_type = cmb_MTip.Text;
                string emr_no = txb_EmrNo.Text;
                string saa = txb_SAA.Text;
                string mez_gun = txb_MSay.Text;

                string odenis_sorgu = "";
                decimal mez_meb = 0;

                DateTime emr_tar, mez_start, mez_end;

                string[] dateFormats = { "dd-MM-yyyy", "d-MM-yyyy", "dd-M-yyyy", "d-M-yyyy", "yyyy-MM-dd", "MM/dd/yyyy" };

                if (!DateTime.TryParseExact(txb_EmrTar.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out emr_tar))
                {
                    MessageBox.Show("Əmr tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!DateTime.TryParseExact(txb_MDovrStart.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out mez_start))
                {
                    MessageBox.Show("Başlanğıc tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!DateTime.TryParseExact(txb_MDovrEnd.Text, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out mez_end))
                {
                    MessageBox.Show("Bitmə tarixini düzgün daxil edin! (dd-MM-yyyy)", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // SQL-ə tarixləri düzgün formatda göndəririk
                string emr_tar_sql = emr_tar.ToString("dd-MM-yyyy");
                string mez_start_sql = mez_start.ToString("dd-MM-yyyy");
                string mez_end_sql = mez_end.ToString("dd-MM-yyyy");


                db.InsertMezuniyyetHesablama(regnum, emr_no, mez_start_sql, mez_end_sql, mez_meb, odenis_sorgu, cl.icraci_ad);

                MessageBox.Show("Məlumat uğurla əlavə edildi!", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private int HesablaIkiTarixArasindakiGunleri()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(txb_MDovrStart.Text) && !string.IsNullOrWhiteSpace(txb_MDovrEnd.Text))
                {
                    DateTime StartTarix = DateTime.ParseExact(txb_MDovrStart.Text, "dd-MM-yyyy", null);
                    DateTime EndTarix = DateTime.ParseExact(txb_MDovrEnd.Text, "dd-MM-yyyy", null);

                    int gunFerqi = (EndTarix - StartTarix).Days + 1;

                    return gunFerqi;
                }
            }
            catch (FormatException)
            {
            }
            return 0;
        }
        private int HesablamIkiTarixArasindaBazadaIsGunleri(string startDate, string endDate)
        {
            if (!DateTime.TryParseExact(startDate, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedStartDate) ||
                !DateTime.TryParseExact(endDate, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedEndDate))
            {
                return 0;
            }

            // Tarixlərin saat hissəsini sıfırlayırıq (həmişə tam gün hesablayırıq)
            parsedStartDate = parsedStartDate.Date;
            parsedEndDate = parsedEndDate.Date.AddDays(1).AddSeconds(-1); // Son günü tam daxil etmək üçün
            decimal son12aymaas = Son12AyUzreTopMaas(parsedStartDate.ToString(),txb_QeydNo.Text);
            using (OracleConnection con = new OracleConnection(cly.con_odb))
            {
                con.Open();

                string query = @"
                                SELECT 
                        NVL(COUNT(date_oper), 0) - NVL(COUNT(CASE WHEN space_or_star = '*' 
                            AND TRIM(TO_CHAR(c1.date_oper, 'Day', 'NLS_DATE_LANGUAGE=ENGLISH')) IN ('Saturday', 'Sunday')
                        THEN 1 END), 0) AS bayramsiz_MG
                    FROM calendar c1
                    WHERE TRUNC(date_oper) BETWEEN :startDate AND :endDate";

                using (OracleCommand cmd = new OracleCommand(query, con))
                {
                    cmd.Parameters.Add(":startDate", OracleDbType.Date).Value = parsedStartDate;
                    cmd.Parameters.Add(":endDate", OracleDbType.Date).Value = parsedEndDate;

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            //decimal mezmaas=son12aymaas/
                            return reader.GetInt32(0);
                        }
                    }
                }
            }
            return 0;
        }
        private decimal Son12AyUzreTopMaas(string startDate, string qeydNo)
        {
            if (!DateTime.TryParseExact(startDate, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedStartDate))
                
            {
                return 0;
            }

            // Tarixlərin saat hissəsini sıfırlayırıq (həmişə tam gün hesablayırıq)
            parsedStartDate = parsedStartDate.Date;
            using (OracleConnection con = new OracleConnection(cly.con_odb))
            {
                con.Open();

                string query = @"
                                SELECT 
                          SUM(t.cemi_hao) AS total_last_12_months
                      FROM bmi_emek_haqqi_shtat t
                      WHERE TO_DATE(t.ay_il, 'MM-YYYY') 
                            BETWEEN ADD_MONTHS(TO_DATE(:startDate,'DD-MM-YYYY'), -13) 
                                AND ADD_MONTHS(:startDate, -1) and substr(t.cari_hesab,10,6)=:qeydNo;";

                using (OracleCommand cmd = new OracleCommand(query, con))
                {
                    cmd.Parameters.Add(":startDate", OracleDbType.Date).Value = parsedStartDate;
                    cmd.Parameters.Add(":startDate", OracleDbType.Date).Value = qeydNo;

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return reader.GetInt32(0);
                        }
                    }
                }
            }
            return 0;
        }
        private bool CheckFields()
        {
            if (string.IsNullOrWhiteSpace(txb_EmrTar.Text))
            {
                MessageBox.Show("Əmr tarixi doldurun!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txb_EmrNo.Text))
            {
                MessageBox.Show("Əmr nömrəsini doldurun!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (cmb_MTip.SelectedIndex == -1)
            {
                MessageBox.Show("Məzuniyyət növünü seçin!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txb_MDovrStart.Text))
            {
                MessageBox.Show("Başlanğıc tarixi doldurun!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txb_MDovrEnd.Text))
            {
                MessageBox.Show("Bitmə tarixi doldurun!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txb_MSay.Text))
            {
                MessageBox.Show("Məzuniyyət günü say sahəsini doldurun!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true; // Bütün sahələr doludursa, true qaytarırıq
        }
        private void LoadDataCombobox()
        {
            List<Isciler> iscilerList = db.GetIsciler(); // Oracle-dan məlumatları alırıq

            cmb_isci_siyahisi.Items.Clear(); // Köhnə məlumatları silirik

            foreach (var isciler in iscilerList)
            {
                cmb_isci_siyahisi.Items.Add(isciler.SAA);
            }
        }
        private void GetTarixVeHefteGunleri(string startDate, string endDate, ListBox listBox)
        {
            // Tarix formatının düzgünlüyünü yoxlayırıq
            if (!DateTime.TryParseExact(startDate, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedStartDate) ||
                !DateTime.TryParseExact(endDate, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedEndDate))
            {
                MessageBox.Show("Tarix formatı düzgün deyil!", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            listBox.Items.Clear(); // ListBox-u təmizləyirik

            using (OracleConnection con = new OracleConnection(cly.con_odb))
            {
                con.Open();

                string query = @"
            SELECT 
                date_oper, 
                TO_CHAR(date_oper, 'Day', 'NLS_DATE_LANGUAGE=ENGLISH') AS hefte_gunu
            FROM calendar
            WHERE TRUNC(date_oper) BETWEEN :startDate AND :endDate
            ORDER BY date_oper";

                using (OracleCommand cmd = new OracleCommand(query, con))
                {
                    cmd.Parameters.Add(":startDate", OracleDbType.Date).Value = parsedStartDate;
                    cmd.Parameters.Add(":endDate", OracleDbType.Date).Value = parsedEndDate;

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            DateTime tarix = reader.GetDateTime(0);
                            string hefteGunu = reader.GetString(1).Trim(); // Boşluqları silirik

                            string listItem = $"{tarix:dd-MM-yyyy} → {hefteGunu}";
                            listBox.Items.Add(listItem);
                        }
                    }
                }
            }
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btn_Elave_Click(object sender, EventArgs e)
        {
            //if (!CheckFields()) // Əgər sahələr boşdursa
            //{
            //    return; // Əməliyyat davam etmir
            //}
            //else
            //{
            //    InsertMezuniyyetler();
            //}
            GetTarixVeHefteGunleri("13-01-2025", "22-01-2025", listBox1);
        }
        private void txb_MDovrStart_TextChanged(object sender, EventArgs e)
        {
            txb_MSay.Text = HesablaIkiTarixArasindakiGunleri().ToString();
            HesablamIkiTarixArasindaBazadaIsGunleri(txb_MDovrStart.Text, txb_MDovrEnd.Text).ToString();


        }
        private void txb_MDovrEnd_TextChanged(object sender, EventArgs e)
        {
            txb_MSay.Text = HesablaIkiTarixArasindakiGunleri().ToString();
            txt_mebleg.Text = HesablamIkiTarixArasindaBazadaIsGunleri(txb_MDovrStart.Text, txb_MDovrEnd.Text).ToString();
        }
        private void frm_MXO_Load(object sender, EventArgs e)
        {
            LoadDataCombobox();
        }
    }
}
