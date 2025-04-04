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
    public partial class frm_isci_derece : Form
    {
        public frm_isci_derece()
        {
            InitializeComponent();
            this.Icon = Aletler.DefaultIcon;
        }
        cl_yanasmalar cl = new cl_yanasmalar();
        private void LoadDataToGrid(DataGridView grid)
        {
            try
            {
                // SQL sorğusu
                string query = @"
            SELECT 
                i.istifadeci_adi AS Adi, 
                i.derece AS Derece
            FROM bmi_istifadeciler i";

                // Oracle bağlantısı
                using (OracleConnection connection = new OracleConnection(cl.con_odb))
                {
                    connection.Open();

                    using (OracleDataAdapter adapter = new OracleDataAdapter(query, connection))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);

                        // DataGridView-i doldurun
                        grid.DataSource = dataTable;

                        // Sütunların genişliyini tənzimləyin
                        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                        // "Adi" sütunun genişliyini dəyişmək
                        if (grid.Columns["Adi"] != null)
                        {
                            grid.Columns["Adi"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }
                        if (grid.Columns["Derece"] != null)
                        {
                            grid.Columns["Derece"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        // Checkbox sütunları əlavə et
                        if (!grid.Columns.Contains("Check1"))
                        {
                            DataGridViewCheckBoxColumn check1 = new DataGridViewCheckBoxColumn
                            {
                                Name = "Check1",
                                HeaderText = "Dərəcə",
                                Width = 50
                            };
                            grid.Columns.Add(check1);
                        }

                        if (!grid.Columns.Contains("Check2"))
                        {
                            DataGridViewCheckBoxColumn check2 = new DataGridViewCheckBoxColumn
                            {
                                Name = "Check2",
                                HeaderText = "Dərəcə",
                                Width = 50
                            };
                            grid.Columns.Add(check2);
                        }

                        if (!grid.Columns.Contains("Check3"))
                        {
                            DataGridViewCheckBoxColumn check3 = new DataGridViewCheckBoxColumn
                            {
                                Name = "Check3",
                                HeaderText = "Dərəcə",
                                Width = 50
                            };
                            grid.Columns.Add(check3);
                        }

                        // Edit düyməsi əlavə et
                        if (!grid.Columns.Contains("Edit"))
                        {
                            DataGridViewButtonColumn btnEdit = new DataGridViewButtonColumn
                            {
                                Name = "Edit",
                                HeaderText = "Edit",
                                Text = "Edit",
                                UseColumnTextForButtonValue = true

                            };
                            grid.Columns.Add(btnEdit);
                        }
                        

                        // Sətirləri iterasiya et və checkboxları doldur
                        foreach (DataGridViewRow row in grid.Rows)
                        {
                            if (row.Cells["Derece"]?.Value != null &&
                                int.TryParse(row.Cells["Derece"].Value.ToString(), out int derece))
                            {
                                // Derece dəyərinə əsasən checkboxları işarələyirik
                                row.Cells["Check1"].Value = (derece >= 1);
                                row.Cells["Check2"].Value = (derece >= 2);
                                row.Cells["Check3"].Value = (derece >= 3);
                            }
                            else
                            {
                                // Derece dəyəri null və ya boşdursa, bütün checkboxları false et
                                row.Cells["Check1"].Value = false;
                                row.Cells["Check2"].Value = false;
                                row.Cells["Check3"].Value = false;
                            }
                        }
                        // Başlıq sətirinin dizaynı
                        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 229, 235);
                        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                        grid.EnableHeadersVisualStyles = false;

                        // Alternativ sətir rəngləri
                        grid.RowsDefaultCellStyle.BackColor = Color.WhiteSmoke;
                        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                        // Xətlərin görünüşü
                        grid.GridColor = Color.Silver;
                        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
                        // Bütün sətiri seçmək üçün seçim rejimi
                        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        grid.MultiSelect = false; // Yalnız bir sətir seçilsin
                        // Sətir seçildikdə rəng dəyişikliyi
                        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(207, 219, 239);
                        grid.DefaultCellStyle.SelectionForeColor = Color.White;
                        // Görünüşü yaxşılaşdırır
                        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        // Avtomatik sütun genişliyi
                        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        grid.AutoGenerateColumns = false; // Sütunlar əl ilə təyin ediləcək
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xəta baş verdi: {ex.Message}", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CustomizeDataGridView(DataGridView grid)
        {
            // Başlıq Dizaynı
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 128, 185); // Mavi rəng
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; // Ağ mətn
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold); // Font
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            grid.ColumnHeadersHeight = 40; // Başlıq hündürlüyü

            // Sətir Dizaynı
            grid.RowsDefaultCellStyle.BackColor = Color.White; // Ağ əsas rəng
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240); // Alternativ rəng
            grid.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219); // Seçilmiş sətir rəngi
            grid.RowsDefaultCellStyle.SelectionForeColor = Color.White; // Seçilmiş mətn rəngi
            grid.RowsDefaultCellStyle.Font = new Font("Segoe UI", 9); // Font stili

            // Hüceyrə Sərhəd Xətləri
            grid.GridColor = Color.FromArgb(224, 224, 224); // Sərhəd xətləri
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal; // Xətər üslubu

            // Sütun Genişliyi və Tam Sətir Seçimi
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;

            // Düzəlişlər üçün Buton Sütunu
            DataGridViewButtonColumn btnEdit = new DataGridViewButtonColumn
            {
                Name = "Edit",
                HeaderText = "Edit",
                Text = "Edit",
                UseColumnTextForButtonValue = true, // Buton mətnini istifadə et
                FlatStyle = FlatStyle.Flat // Yastı düymə
            };
            btnEdit.DefaultCellStyle.BackColor = Color.FromArgb(46, 204, 113); // Yaşıl düymə rəngi
            btnEdit.DefaultCellStyle.ForeColor = Color.White; // Ağ mətn
            grid.Columns.Add(btnEdit);
        }

        // Form yükləndikdə DatagridView-i özelleştir
        private void Form1_Load(object sender, EventArgs e)
        {
            //CustomizeDataGridView(dtg_dereceler);
        }


        private void frm_isci_derece_Load(object sender, EventArgs e)
        {
            LoadDataToGrid(dtg_dereceler);
            //CustomizeDataGridView(dtg_dereceler);
        }

        private void dtg_dereceler_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dtg_dereceler.Columns["Edit"].Index && e.RowIndex >= 0)
            {
                string adi = dtg_dereceler.Rows[e.RowIndex].Cells["Adi"].Value.ToString();
                MessageBox.Show($"Edit düyməsinə klik olundu: {adi}", "Edit", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dtg_dereceler_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }
    }
}
