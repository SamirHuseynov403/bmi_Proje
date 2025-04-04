using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using System.Windows.Forms;
using System.Globalization;

namespace BMI
{
    public class mektubdaxil:Xmlbaglanti
    {

        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;

        public string il = string.Empty;
        int kecencavab = 0;
        public BindingSource mybing;

        public string tamnom = string.Empty;
        public string siranom = string.Empty;
        public string Daxiltarix = string.Empty;
        public string Teshkilatadi = string.Empty;
        public string GonTarix = string.Empty;
        public string Mektubnom=string.Empty;

        public mektubdaxil()
        {
            mybing = new BindingSource();
        }

        public void Daxilmektubgos(DataGridView dgw)
        {
            try
            {
                //
                //select *from odb.daxil_mektub where il=" + il + "
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("Select d.nom1, d.dax_tarix, d.idare_adi, d.gon_tarix, d.dax_nom, d.mek_unvan, n.f_i_o, d.il  from odb.daxil_mektub d left outer join odb.nameoi n on n.code = d.mek_unvan where d.il=" + il + " order by d.nom desc", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Qeydiyyat №";
                dgw.Columns[0].Width = 125;

                dgw.Columns[1].HeaderText = "Daxil olma tarixi";
                dgw.Columns[1].Width = 150;

                dgw.Columns[2].HeaderText = "Göndərən adı";
                dgw.Columns[2].Width = 400;

                dgw.Columns[3].HeaderText = "Göndərilmə tarixi";
                dgw.Columns[3].Width = 150;

                dgw.Columns[4].HeaderText = "Məktub №";
                dgw.Columns[4].Width = 160;

                dgw.Columns[5].HeaderText = "İcraçı №";
                dgw.Columns[5].Width = 110;

                dgw.Columns[6].HeaderText = "İcraçı";
                dgw.Columns[6].Width = 180;

                dgw.Columns[7].HeaderText = "İl";
                dgw.Columns[7].Width = 80;
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void DaxilmektubAxtarish(DataGridView dgw, string axtar)
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("Select d.nom1, d.dax_tarix, d.idare_adi, d.gon_tarix, d.dax_nom, d.mek_unvan, n.f_i_o, d.il  from odb.daxil_mektub d left outer join odb.nameoi n on n.code = d.mek_unvan where " + axtar + " order by d.nom desc", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Qeydiyyat №";
                dgw.Columns[0].Width = 125;

                dgw.Columns[1].HeaderText = "Daxil olma tarixi";
                dgw.Columns[1].Width = 150;

                dgw.Columns[2].HeaderText = "Göndərən adı";
                dgw.Columns[2].Width = 400;

                dgw.Columns[3].HeaderText = "Göndərilmə tarixi";
                dgw.Columns[3].Width = 150;

                dgw.Columns[4].HeaderText = "Məktub №";
                dgw.Columns[4].Width = 160;

                dgw.Columns[5].HeaderText = "İcraçı №";
                dgw.Columns[5].Width = 110;

                dgw.Columns[6].HeaderText = "İcraçı";
                dgw.Columns[6].Width = 180;

                dgw.Columns[7].HeaderText = "İl";
                dgw.Columns[7].Width = 80;
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Daxilmktbinsert()
        {
            try
            {
        Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
        Orcon.Open();
        Orcom = new OracleCommand("insert into odb.daxil_mektub t (t.dax_tarix, t.idare_adi, t.gon_tarix, t.dax_nom, t.mek_unvan, t.il) values (TO_DATE('" + Daxiltarix + "', 'dd-MM-yyyy'), '" + Teshkilatadi + "', TO_DATE('" + GonTarix + "', 'dd-MM-yyyy'), '" + Mektubnom + "', " + icracikodu + ", " + il + ")", Orcon);
        kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktub qeydiyyata alındı...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub qeydiyyata alınmadı... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Daxilmktbupdate()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("Update odb.daxil_mektub t set t.dax_tarix = TO_DATE('" + Daxiltarix + "', 'dd-MM-yyyy'), t.idare_adi ='" + Teshkilatadi + "', t.gon_tarix=TO_DATE('" + GonTarix + "', 'dd-MM-yyyy'), t.dax_nom='" + Mektubnom + "' where t.nom="+tamnom+" and t.nom1="+siranom+"", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktuba Düzəliş olundu...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktuba Düzəliş olunmadı... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Daxilmktboxu()
        {
            try
            {
      Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
      Orcon.Open();
      Orcom = new OracleCommand("Select d.nom, d.dax_tarix, d.idare_adi, d.gon_tarix, d.dax_nom from odb.daxil_mektub d where d.nom1=" + siranom + " and d.mek_unvan=" + icracikodu + " and d.il=" + il + " order by d.nom desc", Orcon);
         Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    tamnom = Ordr["nom"].ToString();
                    Daxiltarix = Ordr["dax_tarix"].ToString();
                    Teshkilatadi = Ordr["idare_adi"].ToString();
                    GonTarix = Ordr["gon_tarix"].ToString();
                    Mektubnom = Ordr["dax_nom"].ToString();
                }
                Ordr.Close();
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Daxilmktbdelete()
        {
            try
            {
     Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
     Orcon.Open();
     Orcom = new OracleCommand("delete from odb.daxil_mektub d where d.nom1=" + siranom + " and d.mek_unvan=" + icracikodu + " and d.il="+il+"", Orcon);

         kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktub məlumat bazasından silindi...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub silinmədi... !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
