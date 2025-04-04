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

namespace BMI.Esas
{
    public partial class frm_setup_icaze : Form
    {
        public frm_setup_icaze()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        public int icracino;
        public int formid;
        private void button1_Click(object sender, EventArgs e)
        {
            UpdateOrInsert(icracino, formid, cl.con_odb);
            this.Close();
        }

        private void frm_setup_icaze_Load(object sender, EventArgs e)
        {

        }
        public void UpdateOrInsert(int icracino, int formid, string connectionString)
        {
            try
            {
                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    connection.Open();

                    // Şərti yoxlamaq üçün SQL sorğusu
                    string checkQuery = @"
                SELECT COUNT(*) 
                FROM bmi_xususi_icazeler 
                WHERE ISTIFADECI_ID = :icracino AND form_id = :formid";

                    using (OracleCommand checkCommand = new OracleCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.Add(new OracleParameter("icracino", icracino));
                        checkCommand.Parameters.Add(new OracleParameter("formid", formid));

                        // Sətir mövcuddursa, `COUNT(*)` > 0 olacaq
                        int count = Convert.ToInt32(checkCommand.ExecuteScalar());

                        if (count > 0)
                        {
                            if (comboBox1.Text=="True")
                            {
                                // `UPDATE` əməliyyatı
                                string updateQuery = @"
                        UPDATE bmi_xususi_icazeler 
                        SET icaze = 1 
                        WHERE istifadeci_id = :icracino AND form_id = :formid";

                                using (OracleCommand updateCommand = new OracleCommand(updateQuery, connection))
                                {
                                    updateCommand.Parameters.Add(new OracleParameter("icracino", icracino));
                                    updateCommand.Parameters.Add(new OracleParameter("formid", formid));
                                    int rowsUpdated = updateCommand.ExecuteNonQuery();
                                    MessageBox.Show($"{rowsUpdated} sətir yeniləndi.", "Uğur", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }

                           else if (comboBox1.Text == "False")
                            {
                                // `UPDATE` əməliyyatı
                                string updateQuery = @"
                        UPDATE bmi_xususi_icazeler 
                        SET icaze = 0 
                        WHERE istifadeci_id = :icracino AND form_id = :formid";

                                using (OracleCommand updateCommand = new OracleCommand(updateQuery, connection))
                                {
                                    updateCommand.Parameters.Add(new OracleParameter("icracino", icracino));
                                    updateCommand.Parameters.Add(new OracleParameter("formid", formid));
                                    int rowsUpdated = updateCommand.ExecuteNonQuery();
                                    MessageBox.Show($"{rowsUpdated} sətir yeniləndi.", "Uğur", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                        else
                        {
                            // Yeni `ID` əldə etmək üçün `MAX(ID)` + 1
                            string getMaxIdQuery = @"
                        SELECT NVL(MAX(id), 0) + 1 
                        FROM bmi_xususi_icazeler";

                            using (OracleCommand getMaxIdCommand = new OracleCommand(getMaxIdQuery, connection))
                            {
                                int newId = Convert.ToInt32(getMaxIdCommand.ExecuteScalar());

                                // `INSERT INTO` əməliyyatı
                                string insertQuery = @"
                            INSERT INTO bmi_xususi_icazeler (id, istifadeci_id, form_id, icaze) 
                            VALUES (:newId, :icracino, :formid, 1)";

                                using (OracleCommand insertCommand = new OracleCommand(insertQuery, connection))
                                {
                                    insertCommand.Parameters.Add(new OracleParameter("newId", newId));
                                    insertCommand.Parameters.Add(new OracleParameter("icracino", icracino));
                                    insertCommand.Parameters.Add(new OracleParameter("formid", formid));
                                    int rowsInserted = insertCommand.ExecuteNonQuery();
                                    MessageBox.Show($"{rowsInserted} sətir əlavə edildi.", "Uğur", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}\n{ex.StackTrace}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
