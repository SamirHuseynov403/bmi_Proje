using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Data.OracleClient;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BMI.Muhasibat;
using Oracle.ManagedDataAccess.Client;

namespace BMI.Esas
{
    public partial class frm_setup : Form
    {
        private Form1 mainForm;
        public frm_setup(Form1 form1)
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
            mainForm = form1; // Form1 referansını saxla
                              // Paneli yarat və formaya əlavə et
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        int icraci;
        public Logiin frmana;
        string form;
        public void SaveMenuStripToDatabase(OracleConnection conn, MenuStrip menuStrip)
        {
            foreach (ToolStripItem item in menuStrip.Items)
            {
                if (item is ToolStripMenuItem topLevelMenuItem)
                {
                    // Top-Level menyu başlıqlarını və elementlərini saxla
                    SaveElementToDatabase(conn, topLevelMenuItem, topLevelMenuItem.Text);
                }
            }
        }

        private void SaveElementToDatabase(OracleConnection conn, ToolStripMenuItem menuItem, string topLevelMenuName)
        {
            // Hazırkı elementin adı
            string elementName = menuItem.Text;
            string formName = topLevelMenuName;

            // Element və forma birlikdə bazada yoxdursa, əlavə et
            if (!IsElementAndFormAlreadyExists(conn, elementName, formName))
            {
                AddElementToDatabase(conn, elementName, formName);
            }

            // Alt menyuları yoxla və daxil et
            foreach (ToolStripItem subItem in menuItem.DropDownItems)
            {
                if (subItem is ToolStripMenuItem subMenuItem)
                {
                    // Alt menyular üçün eyni Top-Level Menu adı ötürülür
                    SaveElementToDatabase(conn, subMenuItem, topLevelMenuName);
                }
            }
        }

        private void AddElementToDatabase(OracleConnection conn, string elementName, string formName)
        {
            // Son ID-ni tapmaq üçün sorğu
            string getLastIdQuery = "SELECT NVL(MAX(ID), 0) FROM odb.bmi_elementler";

            int newId = 0;
            using (OracleCommand getLastIdCmd = new OracleCommand(getLastIdQuery, conn))
            {
                newId = Convert.ToInt32(getLastIdCmd.ExecuteScalar()) + 1; // Son ID-yə 1 əlavə et
            }

            // Yeni element əlavə etmək üçün sorğu
            string insertQuery = "INSERT INTO odb.bmi_elementler (ID, ELEMENT_ADI, FORM_ADI) VALUES (:id, :elementName, :formName)";

            using (OracleCommand insertCmd = new OracleCommand(insertQuery, conn))
            {
                insertCmd.Parameters.Add(new OracleParameter("id", newId)); // Manual olaraq artırılan ID
                insertCmd.Parameters.Add(new OracleParameter("elementName", elementName));
                insertCmd.Parameters.Add(new OracleParameter("formName", formName));

                insertCmd.ExecuteNonQuery();
            }
        }

        private bool IsElementAndFormAlreadyExists(OracleConnection conn, string elementName, string formName)
        {
            string query = @"
        SELECT COUNT(*)
        FROM odb.bmi_elementler
        WHERE ELEMENT_ADI = :elementName AND FORM_ADI = :formName";

            using (OracleCommand cmd = new OracleCommand(query, conn))
            {
                cmd.Parameters.Add(new OracleParameter("elementName", elementName));
                cmd.Parameters.Add(new OracleParameter("formName", formName));

                int count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0; // Əgər say > 0 olarsa, mövcuddur
            }
        }


       
        private void button1_Click(object sender, EventArgs e)
        {
            // Bağlantı sətirini təyin edin
            string connectionString = "Data Source=BMI;User ID=odb;Password=Lah$2021!";

            using (OracleConnection conn = new OracleConnection(connectionString))
            {
                try
                {
                    conn.Open();

                    // MenuStrip-dəki bütün menyuları Top-Level başlıqlara görə bazaya əlavə edir
                    SaveMenuStripToDatabase(conn, mainForm.menuStrip1);

                    MessageBox.Show("Elementlər bazaya uğurla əlavə edildi!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xəta: " + ex.Message);
                }
            }
        }
        
        #region MyRegion
        private void CreateButtonsFromDatabase()
        {
            // SQL sorğusu
            string query = @"select i.i_kod icraci, i.istifadeci_adi adi from bmi_istifadeciler i";

            try
            {
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataReader reader = command.ExecuteReader())
                        {
                            int buttonWidth = 170; // Düymənin eni
                            int buttonHeight = 25; // Düymənin hündürlüyü
                            int margin = 4;       // Aralıq (düymələr arasında boşluq)
                            int currentX = margin; // Başlanğıc X mövqeyi
                            int currentY = margin; // Başlanğıc Y mövqeyi

                            while (reader.Read())
                            {
                                // Məlumatları oxu
                                string adi = reader["adi"]?.ToString() ?? string.Empty;

                                // icraci sütunu boşdursa, sətri keç
                                if (reader["icraci"] == DBNull.Value || string.IsNullOrEmpty(reader["icraci"].ToString()))
                                {
                                    continue;
                                }

                                icraci = Convert.ToInt32(reader["icraci"]);

                                // Yeni düymə yarat
                                Button button = new Button
                                {
                                    Text = adi,
                                    Size = new Size(buttonWidth, buttonHeight),
                                    Location = new Point(currentX, currentY),
                                    Font = new Font("Arial", 8), // Font stili
                                    BackColor = Color.PowderBlue, // Arxa fon rəngi
                                    FlatStyle = FlatStyle.Flat, // Səliqəli düymə üslubu
                                    Tag = icraci // i.i_kod dəyərini düymənin Tag sahəsinə yerləşdiririk
                                };

                                button.FlatAppearance.BorderSize = 1; // Sərhəd qalınlığı
                                button.FlatAppearance.MouseOverBackColor = Color.LightBlue; // Hover rəngi
                                button.Cursor = Cursors.Hand; // Hand göstəricisi

                                // Düyməyə klik hadisəsi əlavə et
                                button.Click += (sender, e) =>
                                {
                                    lbl_ad.Text = button.Text;
                                    int clickedIcraci = Convert.ToInt32(button.Tag);
                                    LoadDataToGrid(dtg_cedvel, clickedIcraci);
                                    //MessageBox.Show($"Seçilmiş İcracı: {clickedIcraci}");
                                };

                                // Düyməni panelə əlavə et
                                pnl_istifadeciler.Controls.Add(button);

                                // Növbəti düymə üçün X mövqeyini dəyişdir
                                currentX += buttonWidth + margin;

                                // Əgər X mövqeyi panelin eni ilə üst-üstə düşərsə, növbəti xəttə keç
                                if (currentX + buttonWidth > pnl_istifadeciler.Width)
                                {
                                    currentX = margin; // X-i sıfırla
                                    currentY += buttonHeight + margin; // Y mövqeyini artır
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}\n{ex.StackTrace}");
            }
        }
        #endregion


        private void CreateButtonsFromDatabaseformlar()
        {
            // SQL sorğusu
            string query = @"
                select distinct e.form_adi adi from bmi_elementler e";

            try
            {
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataReader reader = command.ExecuteReader())
                        {
                            int buttonWidth = 170; // Düymənin eni
                            int buttonHeight = 25; // Düymənin hündürlüyü
                            int margin = 4;       // Aralıq (düymələr arasında boşluq)
                            int currentX = margin; // Başlanğıc X mövqeyi
                            int currentY = margin; // Başlanğıc Y mövqeyi

                            while (reader.Read())
                            {
                                string adi = reader["adi"].ToString();

                                // Yeni düymə yarat
                                Button button = new Button
                                {
                                    Text = adi,
                                    Size = new Size(buttonWidth, buttonHeight),
                                    Location = new Point(currentX, currentY),
                                    Font = new Font("Arial", 8), // Font stili
                                    BackColor = Color.Gainsboro,                // Arxa fon rəngi
                                    //FlatStyle = FlatStyle.Flat,                 // Səliqəli düymə üslubu
                                };

                                button.FlatAppearance.BorderSize = 1;           // Sərhəd qalınlığı
                                button.FlatAppearance.MouseOverBackColor = Color.LightBlue; // Hover rəngi
                                button.Cursor = Cursors.Hand;                   // Hand göstəricisi

                                // Düyməyə klik hadisəsi əlavə et
                                button.Click += (sender, e) =>
                                {
                                    pnl_aletler.Controls.Clear();
                                    form =button.Text;
                                    lbl_form.Text = button.Text;
                                    CreateButtonsFromDatabasealetler();
                                };

                                // Düyməni panelə əlavə et
                                pnl_formlar.Controls.Add(button);

                                // X mövqeyini növbəti düymə üçün dəyişdir
                                currentX += buttonWidth + margin;

                                // Əgər X mövqeyi panelin eni ilə üst-üstə düşərsə, növbəti xəttə keç
                                if (currentX + buttonWidth > pnl_formlar.Width)
                                {
                                    currentX = margin;             // X-i sıfırla
                                    currentY += buttonHeight + margin; // Y mövqeyini artır
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}");
            }
        }
        private void CreateButtonsFromDatabasealetler()
        {
            // SQL sorğusu
            string query = @"
                select e.element_adi adi from bmi_elementler e where e.element_adi <>'Setup' and e.form_adi='"+form+"'";

            try
            {
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataReader reader = command.ExecuteReader())
                        {
                            int buttonWidth = 170; // Düymənin eni
                            int buttonHeight = 25; // Düymənin hündürlüyü
                            int margin = 4;       // Aralıq (düymələr arasında boşluq)
                            int currentX = margin; // Başlanğıc X mövqeyi
                            int currentY = margin; // Başlanğıc Y mövqeyi

                            while (reader.Read())
                            {
                                string adi = reader["adi"].ToString();

                                // Yeni düymə yarat
                                Button button = new Button
                                {
                                    Text = adi,
                                    Size = new Size(buttonWidth, buttonHeight),
                                    Location = new Point(currentX, currentY),
                                    Font = new Font("Arial", 8), // Font stili
                                    BackColor = Color.Gainsboro,                // Arxa fon rəngi
                                    //FlatStyle = FlatStyle.Flat,                 // Səliqəli düymə üslubu
                                };

                                button.FlatAppearance.BorderSize = 1;           // Sərhəd qalınlığı
                                button.FlatAppearance.MouseOverBackColor = Color.LightBlue; // Hover rəngi
                                button.Cursor = Cursors.Hand;                   // Hand göstəricisi

                                // Düyməyə klik hadisəsi əlavə et
                                button.Click += (sender, e) =>
                                {
                                    lbl_alet.Text = button.Text;
                                };

                                // Düyməni panelə əlavə et
                                pnl_aletler.Controls.Add(button);

                                // X mövqeyini növbəti düymə üçün dəyişdir
                                currentX += buttonWidth + margin;

                                // Əgər X mövqeyi panelin eni ilə üst-üstə düşərsə, növbəti xəttə keç
                                if (currentX + buttonWidth > pnl_aletler.Width)
                                {
                                    currentX = margin;             // X-i sıfırla
                                    currentY += buttonHeight + margin; // Y mövqeyini artır
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}");
            }
        }
        //public void LoadDataToGrid(DataGridView grid, int userCode)
        //{
        //    try
        //    {
        //        // SQL sorğusu
        //        string query = @"
        //SELECT e.id AS ID, 
        //       e.element_adi AS Element_Adı, 
        //       e.form_adi AS Form_Adı, 
        //       CASE 
        //           WHEN i.derece >= e.derece THEN '1'
        //           ELSE '0'
        //       END AS İcazə 
        //FROM bmi_elementler e, bmi_istifadeciler i 
        //WHERE i.i_kod = :userCode";

        //        // Oracle bağlantısı
        //        using (OracleConnection connection = new OracleConnection(cl.con_odb))
        //        {
        //            connection.Open();

        //            // SQL əmri
        //            using (OracleCommand command = new OracleCommand(query, connection))
        //            {
        //                command.Parameters.Add(new OracleParameter("userCode", userCode));

        //                // Məlumatları oxumaq üçün adapter
        //                using (OracleDataAdapter adapter = new OracleDataAdapter(command))
        //                {
        //                    DataTable dataTable = new DataTable();
        //                    adapter.Fill(dataTable);

        //                    // DataGridView-ə məlumat yükləyin
        //                    grid.DataSource = dataTable;

        //                    // DataGridView sütun başlıqlarını avtomatik genişləndirin
        //                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        //                    // Checkbox sütunu əlavə et və idarə et
        //                    AddCheckboxColumn(grid, dataTable);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Məlumat yüklənərkən xəta baş verdi: " + ex.Message);
        //    }
        //}
        //public void LoadDataToGrid(DataGridView grid, int userCode)
        //{
        //    try
        //    {
        //        // SQL sorğusu
        //        string query = @"
        //SELECT e.id AS ID, 
        //       e.element_adi AS Element_Adı, 
        //       e.form_adi AS Form_Adı, 
        //       CASE 
        //       WHEN i.derece >= e.derece THEN 'true'
        //       ELSE 'false'
        //       END AS İcazə 
        //FROM bmi_elementler e, bmi_istifadeciler i 
        //WHERE i.i_kod = :userCode";

        //        // Oracle bağlantısı
        //        using (OracleConnection connection = new OracleConnection(cl.con_odb))
        //        {
        //            connection.Open();

        //            // SQL əmri
        //            using (OracleCommand command = new OracleCommand(query, connection))
        //            {
        //                command.Parameters.Add(new OracleParameter("userCode", userCode));

        //                // Məlumatları oxumaq üçün adapter
        //                using (OracleDataAdapter adapter = new OracleDataAdapter(command))
        //                {
        //                    DataTable dataTable = new DataTable();
        //                    adapter.Fill(dataTable);

        //                    // DataGridView-ə məlumat yükləyin
        //                    grid.DataSource = dataTable;

        //                    // DataGridView sütun başlıqlarını avtomatik genişləndirin
        //                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        //                    // Customize DataGridView görünüşü
        //                    CustomizeDataGridView(grid);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Məlumat yüklənərkən xəta baş verdi: " + ex.Message);
        //    }
        //}
        public void LoadDataToGrid(DataGridView grid, int userCode)
        {
            try
            {
                // SQL sorğusu
                string query = @"
                   SELECT distinct t.id AS ID,t.e_adi AS Element_Adı, t.f_adi AS Form_Adı, 
                   CASE 
                   WHEN i.derece >= t.derece THEN 'true'
                   ELSE 'false'
                   END AS İcazə 
                    FROM (SELECT e.element_adi e_adi,e.form_adi f_adi,e.id id,
                    case when e.id=x.form_id and x.icaze=1 then 1 else e.derece end derece
                    FROM bmi_elementler e,bmi_xususi_icazeler x where e.id=x.form_id(+))t, bmi_istifadeciler i 
                    WHERE i.i_kod  =  :userCode";

                // Oracle bağlantısı
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();

                    // SQL əmri
                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        command.Parameters.Add(new OracleParameter("userCode", userCode));

                        // Məlumatları oxumaq üçün adapter
                        using (OracleDataAdapter adapter = new OracleDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            // DataGridView-ə məlumat yükləyin
                            grid.DataSource = dataTable;

                            // Görünüş və "Edit" düyməsi əlavə et
                            CustomizeDataGridView(grid);
                            AddEditButtonToDataGridView(grid);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Məlumat yüklənərkən xəta baş verdi: " + ex.Message);
            }
        }

        private void CreateButtonsFromDatabase1()
        {
            // SQL sorğusu
            string query = @"SELECT i.i_kod AS icraci, i.istifadeci_adi AS adi FROM bmi_istifadeciler i";

            try
            {
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open(); // Bağlantını açırıq

                    using (OracleCommand command = new OracleCommand(query, connection))
                    {
                        using (OracleDataReader reader = command.ExecuteReader())
                        {
                            int buttonWidth = 170; // Düymənin eni
                            int buttonHeight = 25; // Düymənin hündürlüyü
                            int margin = 4;       // Aralıq (düymələr arasında boşluq)
                            int currentX = margin; // Başlanğıc X mövqeyi
                            int currentY = margin; // Başlanğıc Y mövqeyi

                            while (reader.Read())
                            {
                                // Məlumatları oxu
                                string adi = reader["adi"]?.ToString() ?? string.Empty;

                                

                                // Yeni düymə yarat
                                Button button = new Button
                                {
                                    Text = adi,
                                    Size = new Size(buttonWidth, buttonHeight),
                                    Location = new Point(currentX, currentY),
                                    Font = new Font("Arial", 8), // Font stili
                                    BackColor = Color.PowderBlue, // Arxa fon rəngi
                                    FlatStyle = FlatStyle.Flat, // Səliqəli düymə üslubu
                                    Tag = icraci // i.i_kod dəyərini düymənin Tag sahəsinə yerləşdiririk
                                };

                                button.FlatAppearance.BorderSize = 1; // Sərhəd qalınlığı
                                button.FlatAppearance.MouseOverBackColor = Color.LightBlue; // Hover rəngi
                                button.Cursor = Cursors.Hand; // Hand göstəricisi

                                // Düyməyə klik hadisəsi əlavə et
                                button.Click += (sender, e) =>
                                {
                                    Button clickedButton = sender as Button;
                                    if (clickedButton != null)
                                    {
                                        string clickedAdi = clickedButton.Text;

                                        // SQL sorğusu ilə klik edilən istifadəçinin icraci dəyərini al
                                        string getIcraciQuery = @"
                                    SELECT i.i_kod 
                                    FROM bmi_istifadeciler i 
                                    WHERE i.istifadeci_adi = :istifadeciAdi";

                                        using (OracleConnection innerConnection = new OracleConnection(cl.con_odb))
                                        {
                                            innerConnection.Open(); // Bağlantını açırıq

                                            using (OracleCommand getIcraciCommand = new OracleCommand(getIcraciQuery, innerConnection))
                                            {
                                                getIcraciCommand.Parameters.Add(new OracleParameter("istifadeciAdi", clickedAdi));

                                                object result = getIcraciCommand.ExecuteScalar();

                                                if (result != null)
                                                {
                                                    int clickedIcraci = Convert.ToInt32(result);
                                                    lbl_ad.Text = clickedButton.Text;

                                                    // Verilənləri yüklə
                                                    LoadDataToGrid(dtg_cedvel, clickedIcraci);
                                                    icraci = clickedIcraci;
                                                }
                                                else
                                                {
                                                    MessageBox.Show("İstifadəçi tapılmadı.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                }
                                            }
                                        }
                                    }
                                };

                                // Düyməni panelə əlavə et
                                pnl_istifadeciler.Controls.Add(button);

                                // Növbəti düymə üçün X mövqeyini dəyişdir
                                currentX += buttonWidth + margin;

                                // Əgər X mövqeyi panelin eni ilə üst-üstə düşərsə, növbəti xəttə keç
                                if (currentX + buttonWidth > pnl_istifadeciler.Width)
                                {
                                    currentX = margin; // X-i sıfırla
                                    currentY += buttonHeight + margin; // Y mövqeyini artır
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


        private void CustomizeDataGridView(DataGridView grid)
        {
            // Bütün sətiri seçmək üçün seçim rejimi
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false; // Yalnız bir sətir seçilsin

            // Başlıq sətirinin dizaynı
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkSlateBlue;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            grid.EnableHeadersVisualStyles = false;

            // Alternativ sətir rəngləri
            grid.RowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

            // Xətlərin görünüşü
            grid.GridColor = Color.Silver;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;

            // Sətir seçildikdə rəng dəyişikliyi
            grid.DefaultCellStyle.SelectionBackColor = Color.MediumPurple;
            grid.DefaultCellStyle.SelectionForeColor = Color.White;

            // Avtomatik sütun genişliyi
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AutoGenerateColumns = false; // Sütunlar əl ilə təyin ediləcək
        }
        private void AddEditButtonToDataGridView(DataGridView grid)
        {
            // "Edit" sütununu əlavə et
            if (!grid.Columns.Contains("Edit"))
            {
                DataGridViewButtonColumn editButton = new DataGridViewButtonColumn
                {
                    Name = "Edit",
                    HeaderText = "Edit",
                    Text = "Edit",
                    UseColumnTextForButtonValue = true,
                    FlatStyle = FlatStyle.Popup,
                    DefaultCellStyle = { BackColor = Color.DarkSlateBlue, ForeColor = Color.Black }
                };

                grid.Columns.Add(editButton);
            }
        }


        private void AddCheckboxColumn(DataGridView grid, DataTable dataTable)
        {
            // Əgər checkbox sütunu mövcud deyilsə, əlavə edin
            if (!grid.Columns.Contains("Checkbox"))
            {
                DataGridViewCheckBoxColumn checkboxColumn = new DataGridViewCheckBoxColumn
                {
                    Name = "Checkbox",
                    HeaderText = "İcazə",
                    TrueValue = "1", // SQL-də "1" varsa seçilir
                    FalseValue = "0", // SQL-də "0" varsa seçilmir
                    DataPropertyName = "İcazə" // DataTable-dan gələn sütun adı
                };
                grid.Columns.Add(checkboxColumn);
            }

            // Hər bir satır üçün checkbox-u təyin edin
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.Cells["İcazə"].Value != null && row.Cells["Checkbox"] != null)
                {
                    string izin = row.Cells["İcazə"].Value.ToString();
                    row.Cells["Checkbox"].Value = izin == "1";
                }
            }
        }
        private void frm_setup_Load(object sender, EventArgs e)
        {
            CreateButtonsFromDatabase1();
            CreateButtonsFromDatabaseformlar();
        }

        private void dtg_cedvel_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dtg_cedvel.Columns["Edit"].Index && e.RowIndex >= 0)
            {
                DataGridViewRow selectedRow = dtg_cedvel.Rows[e.RowIndex];
                int id = Convert.ToInt32(selectedRow.Cells["ID"].Value); // ID dəyərini al

                Esas.frm_setup_icaze fr=new frm_setup_icaze();
                fr.lbl_adi.Text = lbl_ad.Text;
                fr.lbl_bolme.Text = selectedRow.Cells["FORM_ADI"].Value.ToString();
                fr.lbl_form.Text = selectedRow.Cells["ELEMENT_ADI"].Value.ToString();
                fr.formid = id; 
                fr.icracino=icraci;
                fr.lbl_status.Text= selectedRow.Cells["İCAZƏ"].Value.ToString();
                fr.ShowDialog();
                //MessageBox.Show($"Edit düyməsinə basıldı. ID: {id}");

                // Burada istənilən redaktə əməliyyatını edə bilərsiniz
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Esas.frm_isci_derece fr=new Esas.frm_isci_derece();
            fr.ShowDialog();

        }
    }
}
