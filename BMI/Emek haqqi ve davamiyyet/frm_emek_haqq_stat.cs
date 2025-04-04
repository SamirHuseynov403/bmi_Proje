using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMI.Muhasibat;
using Oracle.ManagedDataAccess.Client;

namespace BMI.Emek_haqqi_ve_davamiyyet
{
    public partial class frm_emek_haqq_stat : Form
    {
        public frm_emek_haqq_stat()
        {
            InitializeComponent();

        }
        cl_yanasmalar cl = new cl_yanasmalar();
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void frm_emek_haqq_stat_Load(object sender, EventArgs e)
        {
            dtg_Emek_haqqi.Columns[0].Frozen = true;
            dtg_Emek_haqqi.Columns[1].Frozen = true;
        }

        private void dtg_Emek_haqqi_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V) // Ctrl+V basıldıqda
            {
                PasteMultiLine();
                e.Handled = true;
            }
        }
        public static string RusToAzeTranslit(string text)
        {
            Dictionary<string, string> translitMap = new Dictionary<string, string>
    {
        {"А", "A"}, {"Б", "B"}, {"В", "V"}, {"Г", "Q"}, {"Д", "D"},
        {"Е", "E"}, {"Ё", "Yo"}, {"Ж", "J"}, {"З", "Z"}, {"И", "İ"},
        {"Й", "Y"}, {"К", "K"}, {"Л", "L"}, {"М", "M"}, {"Н", "N"},
        {"О", "O"}, {"П", "P"}, {"Р", "R"}, {"С", "S"}, {"Т", "T"},
        {"У", "U"}, {"Ф", "F"}, {"Х", "X"}, {"Ц", "Ü"}, {"Ч", "C"},
        {"Ш", "Ş"}, {"Щ", "H"}, {"Ъ", ""}, {"Ы", "I"}, {"Ь", "Ğ"},
        {"Э", "G"}, {"Ю", "Y"}, {"Я", "Ə"},

        {"а", "a"}, {"б", "b"}, {"в", "v"}, {"г", "q"}, {"д", "d"},
        {"е", "e"}, {"ё", "yo"}, {"ж", "j"}, {"з", "z"}, {"и", "i"},
        {"й", "y"}, {"к", "k"}, {"л", "l"}, {"м", "m"}, {"н", "n"},
        {"о", "o"}, {"п", "p"}, {"р", "r"}, {"с", "s"}, {"т", "t"},
        {"у", "u"}, {"ф", "f"}, {"х", "x"}, {"ц", "ü"}, {"ч", "c"},
        {"ш", "ş"}, {"щ", "h"}, {"ъ", ""}, {"ы", "ı"}, {"ь", "ğ"},
        {"э", "g"}, {"ю", "yu"}, {"я", "ə"}
    };

            foreach (var pair in translitMap)
            {
                text = text.Replace(pair.Key, pair.Value);
            }
            return text;
        }
        private void ConvertDataGridViewToAze()
        {
            foreach (DataGridViewRow row in dtg_Emek_haqqi.Rows)
            {
                if (row.Cells[0].Value != null) // 0 sütunun indeksidir
                {
                    row.Cells[0].Value = RusToAzeTranslit(row.Cells[0].Value.ToString());
                }
            }
        }
        private void PasteMultiLine()
        {
            try
            {
                string clipboardText = Clipboard.GetText();
                string[] lines = clipboardText.Split('\n');

                int rowIndex = dtg_Emek_haqqi.CurrentCell?.RowIndex ?? 0;
                int colIndex = dtg_Emek_haqqi.CurrentCell?.ColumnIndex ?? 0;

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] cells = line.Trim().Split('\t');
                    int tempColIndex = colIndex;

                    // Əgər yeni sətir lazımdırsa, əlavə et
                    if (rowIndex >= dtg_Emek_haqqi.Rows.Count - 1)
                    {
                        dtg_Emek_haqqi.Rows.Add();
                    }

                    foreach (string cell in cells)
                    {
                        if (tempColIndex < dtg_Emek_haqqi.ColumnCount)
                        {
                            dtg_Emek_haqqi[tempColIndex, rowIndex].Value = cell;
                            tempColIndex++;
                        }
                    }

                    rowIndex++; // Yeni sətrə keç
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Paste zamanı xəta baş verdi: " + ex.Message);
            }
        }
        public void InsertDataFromDataGridView(DataGridView dgv, string selectedAyIl)
        {
            using (OracleConnection con = new OracleConnection(cl.con_odb))
            {
                con.Open();
                using (OracleTransaction transaction = con.BeginTransaction()) // Tranzaksiya ilə işləyirik
                {
                    try
                    {
                        string query = @"INSERT INTO bmi_emek_haqqi_shtat 
                        (SAA, VEZIFE, EMEK_HAQQI, EMEK_H_HAQQI, EMR_07, IEMKO, ORTA_EH, MH, MUKAFAT, BAY_HEDIY, MQVH, XESTELIK_VER, 
                         MDSS_DG, M_98_2_1, M_98_2_3, ELAVE_EH, HYS, CEMI_HAO, OHYSH, I_OHYSH, MDSS, TISH, GV, ITS, AVANS, TUTULMUSDUR, 
                         ODENILMISDIR, ITO_MDSS, ITO_ITS_H, CARI_HESAB, AY_IL, CREATED_AT)
                        VALUES 
                        (:SAA, :VEZIFE, :EMRK_HAQQI, :H_EMEK_HAQQI, :EMR_07, :IEMKO, :ORTA_EH, :MH, :MUKAFAT, :BAY_HEDIY, :MQVH, 
                         :XESTELIK_VER, :MDSS_DC, :M_98_2_1, :M_98_2_3, :ELAVE_EH, :MJS, :CMT_HAO, :OHYSH, :I_OHYSH, :MDSS, 
                         :TISH, :GV, :ITS, :AVANS, :TUTULMUSDUR, :ODENILMISDIR, :ITO_MDSS, :ITO_ITS_H, :CARI_HESAB, :AY_IL, SYSDATE)";

                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            if (!row.IsNewRow) // Yalnız dolu olan sətirləri götür
                            {
                                using (OracleCommand cmd = new OracleCommand(query, con))
                                {
                                    cmd.Parameters.Add(":SAA", OracleDbType.Varchar2).Value = row.Cells[0].Value?.ToString() ?? "";
                                    cmd.Parameters.Add(":VEZIFE", OracleDbType.Varchar2).Value = row.Cells[1].Value?.ToString() ?? "";

                                    cmd.Parameters.Add(":EMRK_HAQQI", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[2].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal emekHaqqi) ? (object)emekHaqqi : DBNull.Value;
                                    cmd.Parameters.Add(":H_EMEK_HAQQI", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[3].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal hEmekHaqqi) ? (object)hEmekHaqqi : DBNull.Value;
                                    cmd.Parameters.Add(":EMR_07", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[4].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal emr07) ? (object)emr07 : DBNull.Value;
                                    cmd.Parameters.Add(":IEMKO", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[5].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal iemko) ? (object)iemko : DBNull.Value;
                                    cmd.Parameters.Add(":ORTA_EH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[6].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal ortaEh) ? (object)ortaEh : DBNull.Value;
                                    cmd.Parameters.Add(":MH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[7].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mh) ? (object)mh : DBNull.Value;
                                    cmd.Parameters.Add(":MUKAFAT", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[8].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mukafat) ? (object)mukafat : DBNull.Value;
                                    cmd.Parameters.Add(":BAY_HEDIY", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[9].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal bayHediy) ? (object)bayHediy : DBNull.Value;
                                    cmd.Parameters.Add(":MQVH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[10].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mqvh) ? (object)mqvh : DBNull.Value;
                                    cmd.Parameters.Add(":XESTELIK_VER", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[11].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal xestelikVer) ? (object)xestelikVer : DBNull.Value;
                                    cmd.Parameters.Add(":MDSS_DC", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[12].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mdssDc) ? (object)mdssDc : DBNull.Value;
                                    cmd.Parameters.Add(":M_98_2_1", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[13].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal m98_2_1) ? (object)m98_2_1 : DBNull.Value;
                                    cmd.Parameters.Add(":M_98_2_3", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[14].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal m98_2_3) ? (object)m98_2_3 : DBNull.Value;
                                    cmd.Parameters.Add(":ELAVE_EH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[15].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal elaveEh) ? (object)elaveEh : DBNull.Value;
                                    cmd.Parameters.Add(":MJS", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[16].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mjs) ? (object)mjs : DBNull.Value;
                                    cmd.Parameters.Add(":CMT_HAO", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[17].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal cmtHao) ? (object)cmtHao : DBNull.Value;
                                    cmd.Parameters.Add(":OHYSH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[18].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal ohysh) ? (object)ohysh : DBNull.Value;
                                    cmd.Parameters.Add(":I_OHYSH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[19].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal lOhysh) ? (object)lOhysh : DBNull.Value;
                                    cmd.Parameters.Add(":MDSS", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[20].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal mdss) ? (object)mdss : DBNull.Value;
                                    cmd.Parameters.Add(":TISH", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[21].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal tish) ? (object)tish : DBNull.Value;
                                    cmd.Parameters.Add(":GV", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[22].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal gv) ? (object)gv : DBNull.Value;
                                    cmd.Parameters.Add(":ITS", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[23].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal ints) ? (object)ints : DBNull.Value;
                                    cmd.Parameters.Add(":AVANS", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[24].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal avans) ? (object)avans : DBNull.Value;
                                    cmd.Parameters.Add(":TUTULMUSDUR", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[25].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal tutulmusdur) ? (object)tutulmusdur : DBNull.Value;
                                    cmd.Parameters.Add(":ODENILMISDIR", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[26].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal odenilmisdir) ? (object)odenilmisdir : DBNull.Value;
                                    cmd.Parameters.Add(":ITO_MDSS", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[27].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal itoMdss) ? (object)itoMdss : DBNull.Value;
                                    cmd.Parameters.Add(":ITO_ITS_H", OracleDbType.Decimal).Value =
                                        decimal.TryParse(row.Cells[28].Value?.ToString().Replace("'", "").Replace(",", "."), out decimal itoItsH) ? (object)itoItsH : DBNull.Value;

                                    cmd.Parameters.Add(":CARI_HESAB", OracleDbType.Varchar2).Value = row.Cells[29].Value?.ToString() ?? "";
                                    cmd.Parameters.Add(":AY_IL", OracleDbType.Varchar2).Value = selectedAyIl; // ComboBox-dan gələn dəyər


                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        transaction.Commit();
                        MessageBox.Show("Məlumat uğurla əlavə olundu!", "Uğur", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void btn_Elave_Click(object sender, EventArgs e)
        {
            ConvertDataGridViewToAze();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InsertDataFromDataGridView(dtg_Emek_haqqi,textBox1.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dtg_Emek_haqqi.Rows.Clear();
        }
    }
}
