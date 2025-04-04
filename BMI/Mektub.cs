using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Data;
using System.Globalization;

namespace BMI
{
    public class Mektub : Xmlbaglanti
    {
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;

        public string il = string.Empty;
        int kecencavab = 0;
        public BindingSource mybing;

        public string Xmgeydiyyatnom = string.Empty;
        public string Xmgonderilenyer = string.Empty;
        public string Xaricmektbtarix = string.Empty;
        public string Xmgisamezmun = string.Empty;
        public string Xmmektubmetn = string.Empty;

        public string gh_hevnom = string.Empty;
        public string gh_hesnom = string.Empty;
        public string gh_adi = string.Empty;
        public string gh_mebleg = string.Empty;
        public string gh_valtip = string.Empty;
        public string gh_tarix = string.Empty;
        public string gh_hevtip = string.Empty;
        public string gh_menolke = string.Empty;
        public string gh_olke = string.Empty;
        public string gh_gontip = string.Empty;
        public string gh_albank = string.Empty;

        public Mektub()
        {
            mybing = new BindingSource();
        }

        public void Xaricmektub(DataGridView dgw)
        {
            try
            {

                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                //Orcom = new OracleCommand("Select d.nom1, d.dax_tarix, d.idare_adi, d.gon_tarix, d.dax_nom, d.mek_unvan, n.f_i_o, d.il  from odb.daxil_mektub d left outer join odb.nameoi n on n.code = d.mek_unvan where d.il=" + il + " order by d.nom desc", Orcon);

                Orcom = new OracleCommand("select m.qey_nom, m.tarix, m.gon_yer, m.qisa_mez, m.icraci, n.f_i_o, m.il from odb.xaric_mektub m left outer join odb.nameoi n on n.code = m.icraci Where m.il=" + il + " order by m.kod desc", Orcon); Orda = new OracleDataAdapter(Orcom); 
                 DataTable dt111 = new DataTable();
                Orda.Fill(dt111);
                //mybing.DataSource = dt111;
                dgw.DataSource = dt111;
                Orcon.Close();
                //dgw.DefaultCellStyle.Encoding = Encoding.UTF8; // DataGridView'deki hücreler için karakter kodlaması ayarı

                dgw.Columns[0].HeaderText = "Məktubun №";
                dgw.Columns[0].Width = 125; 

                dgw.Columns[1].HeaderText = "Tarix";
                dgw.Columns[1].Width = 100;

                dgw.Columns[2].HeaderText = "Göndərilən yer";
                dgw.Columns[2].Width = 350;

                dgw.Columns[3].HeaderText = "Qısa məzmun";
                dgw.Columns[3].Width = 420;

                dgw.Columns[4].HeaderText = "İcraçı kodu";
                dgw.Columns[4].Width = 110;

                dgw.Columns[5].HeaderText = "İcraçı";
                dgw.Columns[5].Width = 190;

                dgw.Columns[6].HeaderText = "İl";
                dgw.Columns[6].Width = 80;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
        }
        public void listelegedenhevale(DataGridView dgw)
        {
            //where to_char( tarix,'yyyy')='"+tarixIl+"'
            string tarixIl = DateTime.Now.Date.Year.ToString();
            OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
            Ocon.Open();
            OracleCommand Ocom = new OracleCommand("Select hev_nom,hes_nom,saa,tip_res,mebleg,val_tip,tarix,men_olke,olke,hev_tip,gon_tip,al_bank From odb.geden_hevale where to_char( tarix,'yyyy')='" + tarixIl + "' order by to_number(substr(hev_nom,6,4)) desc ", Ocon);

            OracleDataAdapter Oda = new OracleDataAdapter(Ocom);
            DataTable Odt = new DataTable();
            Oda.Fill(Odt);
            mybing.DataSource = Odt;
            dgw.DataSource = mybing;

            Ocon.Close();
            dgw.Columns[0].HeaderText = "Həvalə №";
            dgw.Columns[0].Width = 125;
            dgw.Columns[1].HeaderText = "Hesab №";
            dgw.Columns[1].Width = 150;
            dgw.Columns[2].HeaderText = "Adı";
            dgw.Columns[2].Width = 200;
            dgw.Columns[3].HeaderText = "Rezident tipi ";
            dgw.Columns[3].Width = 125;
            dgw.Columns[4].HeaderText = "Məbləğ №";
            dgw.Columns[4].Width = 125;
            dgw.Columns[5].HeaderText = "Valyuta növü ";
            dgw.Columns[5].Width = 125;
            dgw.Columns[6].HeaderText = "Tarix ";
            dgw.Columns[6].Width = 100;
            dgw.Columns[7].HeaderText = "Ölkə mənşəyi ";
            dgw.Columns[7].Width = 150;
            dgw.Columns[8].HeaderText = "Ölkə ";
            dgw.Columns[8].Width = 150;
            dgw.Columns[9].HeaderText = "Həvalə tipi ";
            dgw.Columns[9].Width = 125;
            dgw.Columns[10].HeaderText = "Göndərən tipi";
            dgw.Columns[10].Width = 125;
            dgw.Columns[11].HeaderText = "Alan bank";
            dgw.Columns[11].Width = 200;
            //dgwgedenhevale.Columns[12].HeaderText = "İcraçı";
            //dgwgedenhevale.Columns[12].Width = 200;

        }

        public void XaricmektubAxtarish(DataGridView dgw, string axtar)
        {
            try
            {
                // dgw.Refresh();
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("select m.qey_nom, m.tarix, m.gon_yer, m.qisa_mez, m.icraci, n.f_i_o, m.il from odb.xaric_mektub m left outer join odb.nameoi n on n.code = m.icraci Where " + axtar + " order by m.kod desc", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Məktubun №";
                dgw.Columns[0].Width = 125;

                dgw.Columns[1].HeaderText = "Tarix";
                dgw.Columns[1].Width = 100;

                dgw.Columns[2].HeaderText = "Göndərilən yer";
                dgw.Columns[2].Width = 350;

                dgw.Columns[3].HeaderText = "Qısa məzmun";
                dgw.Columns[3].Width = 420;

                dgw.Columns[4].HeaderText = "İcraçı kodu";
                dgw.Columns[4].Width = 110;

                dgw.Columns[5].HeaderText = "İcraçı";
                dgw.Columns[5].Width = 190;

                dgw.Columns[6].HeaderText = "İl";
                dgw.Columns[6].Width = 80;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void gelenhevAxtarish(DataGridView dgw, string axtar)
        {
            try
            {
                // dgw.Refresh();
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                //Orcom = new OracleCommand("select m.qey_nom, m.tarix, m.gon_yer, m.qisa_mez, m.icraci, n.f_i_o, m.il from odb.xaric_mektub m left outer join odb.nameoi n on n.code = m.icraci Where " + axtar + " order by m.kod desc", Orcon);
                Orcom = new OracleCommand("Select m.hev_nom,m.hes_nom,m.saa,m.tip_res,m.mebleg,m.val_tip,m.tarix,m.men_olke,m.hev_tip,m.gel_olke,m.gon_tip,m.al_bank,m.icra,n.f_i_o From odb.gelen_hevale m left outer join odb.nameoi n on n.code = icra Where " + axtar + " order by m.kod desc", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Həvalə №";
                dgw.Columns[0].Width = 125;
                dgw.Columns[1].HeaderText = "Hesab №";
                dgw.Columns[1].Width = 150;
                dgw.Columns[2].HeaderText = "Adı";
                dgw.Columns[2].Width = 200;
                dgw.Columns[3].HeaderText = "Rezident tipi ";
                dgw.Columns[3].Width = 125;
                dgw.Columns[4].HeaderText = "Məbləğ №";
                dgw.Columns[4].Width = 125;
                dgw.Columns[5].HeaderText = "Valyuta növü ";
                dgw.Columns[5].Width = 125;
                dgw.Columns[6].HeaderText = "Tarix ";
                dgw.Columns[6].Width = 100;
                dgw.Columns[7].HeaderText = "Ölkə mənşəyi ";
                dgw.Columns[7].Width = 150;

                dgw.Columns[9].HeaderText = "Həvalə tipi ";
                dgw.Columns[9].Width = 125;
                dgw.Columns[8].HeaderText = "Ölkə ";
                dgw.Columns[8].Width = 150;

                dgw.Columns[10].HeaderText = "Göndərən tipi";
                dgw.Columns[10].Width = 125;
                dgw.Columns[11].HeaderText = "Alan bank";
                dgw.Columns[11].Width = 200;
                dgw.Columns[12].HeaderText = "İcraçı";
                dgw.Columns[12].Width = 200;
                dgw.Columns[13].HeaderText = "İcraçı adı";
                dgw.Columns[13].Width = 200;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Xaricmektuboxu()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("select x.tarix, x.gon_yer, x.qisa_mez, x.mektub_metn from odb.xaric_mektub x where x.qey_nom='" + Xmgeydiyyatnom + "'", Orcon);
                Orcom.InitialLONGFetchSize = 10000;
                Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    Xaricmektbtarix = Ordr["tarix"].ToString();
                    Xmgonderilenyer = Ordr["gon_yer"].ToString();
                    Xmgisamezmun = Ordr["qisa_mez"].ToString();
                    Xmmektubmetn = Ordr.IsDBNull(3) ? null : Ordr.GetString(3);
                    //Xmmektubmetn = Ordr.GetString(3);
                }
                Ordr.Close();
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Xaricmektinsert()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("insert into odb.xaric_mektub x (x.gon_yer, x.tarix, x.qisa_mez, x.icraci, x.mektub_metn, x.il) values ('" + Xmgonderilenyer + "', TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), '" + Xmgisamezmun + "', " + icracikodu + ", '" + Xmmektubmetn + "', " + il + ")", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktub qeydiyyata alındı...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub qeydiyyata alınmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
                Orcon.Close();
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    Orcon.Close();
            //}
        }

        public void m_geden_hevale_insert()
        {
            try
            {
        

                OracleConnection Ocon = new OracleConnection("DATA SOURCE=BMI;USER ID=FOXPRO;Password=pass");
                Ocon.Open();

                //Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                //Orcon.Open();
                Orcom = new OracleCommand("insert into odb.geden_hevale (HEV_NOM,HES_NOM,SAA,Mebleg,Val_tip,tarix,hev_tip,men_olke,olke,gon_tip,al_bank) values ('" + gh_hevnom + "','" + gh_hesnom + "','" + gh_adi + "','" + gh_mebleg + "','" + gh_valtip + "', TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), '" + gh_hevtip + "', " + gh_menolke + ", '" + gh_menolke + "', " + gh_olke + ", " + gh_gontip + ", " + gh_albank + ")", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktub qeydiyyata alındı...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub qeydiyyata alınmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Orcon.Close();
            }
        }

        //(select '"+il+"-'||(substr(m.qey_nom,6, length(m.qey_nom)-5)+1) from odb.xaric_mektub m where m.kod =(select max(t.kod) from odb.xaric_mektub t)),  --- x.qey_nom,

        public void Xaricmektupdate()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("Update odb.xaric_mektub x set x.tarix = TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), x.gon_yer='" + Xmgonderilenyer + "', x.qisa_mez='" + Xmgisamezmun + "', x.mektub_metn='" + Xmmektubmetn + "' where x.qey_nom='" + Xmgeydiyyatnom + "'", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktuba Düzəliş olundu...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktuba Düzəliş olunmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void gedhevaleupdate()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("Update odb.gelen_hevale x set x.tarix = TO_DATE('" + Xaricmektbtarix + "', 'dd-MM-yyyy'), x.gon_yer='" + Xmgonderilenyer + "', x.qisa_mez='" + Xmgisamezmun + "', x.mektub_metn='" + Xmmektubmetn + "' where x.qey_nom='" + Xmgeydiyyatnom + "'", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktuba Düzəliş olundu...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktuba Düzəliş olunmadı... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Orcon.Close();
            }
            catch (Exception ex)
            {
                Orcon.Close();
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Xaricmektdelete()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("delete from odb.xaric_mektub x where x.qey_nom='" + Xmgeydiyyatnom + "'", Orcon);

                kecencavab = Orcom.ExecuteNonQuery();

                if (kecencavab > 0)
                {
                    MessageBox.Show("Məktub  silindi...", "Məlumat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Məktub silinmədi... !", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
