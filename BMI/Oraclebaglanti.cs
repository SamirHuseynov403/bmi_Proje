using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Oracle.DataAccess.Client;
using System.Windows.Forms;
using System.Data;
using System.Xml;
using System.Globalization;
using System.Threading;

namespace BMI
{
    public class Oraclebaglanti:Xmlbaglanti
    {
        public OracleCommand Orcom;
        public OracleDataAdapter Orda;
        public OracleDataReader Ordr;
        public DataTable dt;

        public ComboBox cmbName;
        public CheckBox ckbSaxla;

        DateTime yoxlaBash = DateTime.Parse("01-01-2013");
        DateTime yoxlaOrta = DateTime.Parse("22-02-2016");
        DateTime yoxla;
        public string tarix;
        public string[] gunayil = new string[3];

        public Label sglkod;

        public BindingSource mybing;

        public Oraclebaglanti()
        {
            tesdig = false;
            tarix = string.Empty;
            sglkod = new Label();
            sglkod.Text = string.Empty;
            mybing = new BindingSource();
        }

        public void Userlogin()
        {
            try
            {
              
                    cmbName.Items.Clear();
                    Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                    Orcon.Open();
                    Orcom = new OracleCommand("select n.name from odb.nameoi n where n.code not in (0, 2, 3, 5, 6, 9, 10, 11, 13, 14, 17, 19, 23, 24, 26, 27, 29, 31, 41) order by n.name asc", Orcon);
                    Ordr = Orcom.ExecuteReader();
                    while (Ordr.Read())
                    {
                        cmbName.Items.Add(Ordr["name"]).ToString();
                    }
                    Ordr.Close();
                    Orcon.Close();

                    if (Currentuser == "Noname")
                    {
                        cmbName.SelectedIndex = 0;
                    }
                    else
                    {
                        cmbName.SelectedItem = Currentuser;
                        ckbSaxla.Checked = true;
                    }
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public override void Tesdigal()
        {
            try
            {
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                if (Orcon.State == ConnectionState.Open)
                {
                    tesdig = true;
                }
                Orcon.Close();
            }
            catch (Exception)
            {
                Orcon.Close();
                tesdig = false;
                MessageBox.Show("Şifrə doğru deyil", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public override void Xmlyenile()
        {
            try
            {
                if (ckbSaxla.Checked == true)
                {
                    xmlfilepatch = @"C:\BMI_\BMI\Servername.xml";

                    if (File.Exists(xmlfilepatch) == false)
                    {
                        MessageBox.Show("XML file tapılmadı", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        xmldoc.Load(xmlfilepatch);
                        xmlnodes = xmldoc.SelectNodes("/Server/OracleServer");

                        foreach (XmlElement element in xmlnodes)
                        {
                            if (element.Attributes["ID"].Value == "1")
                            {
                                element.SelectSingleNode("Currentuser").InnerText = Currentuser;
                                xmldoc.Save("Servername.xml");
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Adtap()
        {
            try
            {
                Orcon = new OracleConnection("DATA SOURCE=" + DataSource + "; USER ID=" + UserID + "; Password=" + Password + "");
                Orcon.Open();
                Orcom = new OracleCommand("select n.f_i_o, n.code from odb.nameoi n where n.name='" + cmbName.Text + "'", Orcon);
                Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    Currentuser = Ordr["f_i_o"].ToString();
                    icracikodu = Ordr["code"].ToString();
                }
                Ordr.Close();
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Olkeler()
        {
            try
            {
                cmbName.Items.Clear();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select c.name from odb.countrycode c where c.code is not null and c.name is not null order by c.name", Orcon);
                Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    cmbName.Items.Add(Ordr["name"]).ToString();
                }
                Ordr.Close();
                Orcon.Close();

                cmbName.SelectedItem = "Азярбайъан Республикасы";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Rekvizit(TextBox Rehberadi, TextBox Rehbervez)
        {
            try
            {
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select t.f_i_o_rukovoditela, t.status_rukovoditela from odb.attrib t", Orcon);
                Ordr = Orcom.ExecuteReader();
                while (Ordr.Read())
                {
                    Rehberadi.Text = Ordr["f_i_o_rukovoditela"].ToString();
                    Rehbervez.Text = Ordr["status_rukovoditela"].ToString();
                }
                Ordr.Close();
                Orcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Rekvizit(DataGridView dgw)
        {
            try
            {
                dgw.Columns.Clear();
                dgw.Refresh();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select r.name_regnom  from odb.regnom r where r.insider=1 and r.svazanniy=1", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Ады";
                dgw.Columns[0].Width = 478;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Rekvizit(string axtar, DataGridView dgw)
        {
            try
            {
                dgw.Columns.Clear();
                dgw.Refresh();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select r.name_regnom  from odb.regnom r where r.insider=1 and r.svazanniy=1 and r.name_regnom like '" + axtar + "%'", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Ады";
                dgw.Columns[0].Width = 478;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Hesabacmamug(DataGridView dgw)
        {
            try
            {
                yoxla = DateTime.Parse(tarix);

                if (yoxla >= yoxlaOrta)
                {
                    dgw.Columns.Clear();
                    dgw.Refresh();
                    Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                    Orcon.Open();
                    Orcom = new OracleCommand("Select l.az||l.nr||l.bank||l.licsch, k.ekv, r.name_regnom, r.mesto_rojdeniya, r.data_rojdeniya, r.registrac, r.adress, r.grajdanstvo, r.senedi_veren_orqaninin_adi, r.senedin_verilme_tarixi, r.passport, f.fealiyyet_ve_peshesi, f.vezifesi, f.gelir_menbeyi, r.telefon, r.faks  from odb.licsch l inner join odb.regnom r on r.regnom = l.registrac_nomer inner join odb.ikv_ekv k on k.ikv = SUBSTR(l.licsch,7,1) inner join odb.fiziki_shexs f on f.regnom = l.registrac_nomer where l.date_open_licsch=to_date('" + tarix + "', 'dd-MM-yyyy') and SUBSTR( l.licsch,0,5) in ('41010','41020','41015','41025') and SUBSTR(l.licsch,16,1) != 9" + sglkod.Text + "", Orcon);
                    Orda = new OracleDataAdapter(Orcom);
                    dt = new DataTable();
                    Orda.Fill(dt);
                    mybing.DataSource = dt;
                    dgw.DataSource = mybing;
                    Orcon.Close();

                    dgw.Columns[0].HeaderText = "Щесаб нюмряси";
                    dgw.Columns[0].Width = 260;

                    dgw.Columns[1].HeaderText = "Валйута";
                    dgw.Columns[1].Width = 120;

                    dgw.Columns[2].HeaderText = "Ады";
                    dgw.Columns[2].Width = 340;

                    dgw.Columns[3].HeaderText = "Доьулдуьу йер";
                    dgw.Columns[3].Width = 270;

                    dgw.Columns[4].HeaderText = "Доьум тарихи";
                    dgw.Columns[4].Width = 130;

                    dgw.Columns[5].HeaderText = "Гейдиййат цнваны";
                    dgw.Columns[5].Width = 350;

                    dgw.Columns[6].HeaderText = "Йашадыь цнван";
                    dgw.Columns[6].Width = 350;

                    dgw.Columns[7].HeaderText = "Вятяндашлыг";
                    dgw.Columns[7].Width = 120;

                    dgw.Columns[8].HeaderText = "Сяняди верен орган";
                    dgw.Columns[8].Width = 170;

                    dgw.Columns[9].HeaderText = "Сянядин верилмя тарихи";
                    dgw.Columns[9].Width = 180;

                    dgw.Columns[10].HeaderText = "Ш/В Серийа вя нюмряси";
                    dgw.Columns[10].Width = 180;

                    dgw.Columns[11].HeaderText = "Иш йери";
                    dgw.Columns[11].Width = 420;

                    dgw.Columns[12].HeaderText = "Вязифяси";
                    dgw.Columns[12].Width = 250;

                    dgw.Columns[13].HeaderText = "Эялир мянбяйи";
                    dgw.Columns[13].Width = 150;

                    dgw.Columns[14].HeaderText = "Телефон";
                    dgw.Columns[14].Width = 480;

                    dgw.Columns[15].HeaderText = "Факс";
                    dgw.Columns[15].Width = 200;
                }

                else if (yoxla < yoxlaOrta && yoxla >= yoxlaBash)
                {
                    dgw.Columns.Clear();
                    dgw.Refresh();
                    Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                    Orcon.Open();
                    Orcom = new OracleCommand("Select l.az||l.nr||l.bank||l.licsch, k.ekv, r.name_regnom, r.mesto_rojdeniya, r.data_rojdeniya, r.registrac, r.adress, r.grajdanstvo, r.kem_i_kogda_vidan, r.passport, f.fealiyyet_ve_peshesi, f.vezifesi, f.gelir_menbeyi, r.telefon, r.faks from odb.licsch l inner join odb.regnom r on r.regnom = l.registrac_nomer inner join odb.ikv_ekv k on k.ikv = SUBSTR(l.licsch,7,1) inner join odb.fiziki_shexs f on f.regnom = l.registrac_nomer where l.date_open_licsch =to_date('" + tarix + "', 'dd-MM-yyyy') and SUBSTR(l.licsch,0,5) in ('41010','41020','41015','41025')" + sglkod.Text + "", Orcon);
                    Orda = new OracleDataAdapter(Orcom);
                    dt = new DataTable();
                    Orda.Fill(dt);
                    mybing.DataSource = dt;
                    dgw.DataSource = mybing;
                    Orcon.Close();

                    dgw.Columns[0].HeaderText = "Щесаб нюмряси";
                    dgw.Columns[0].Width = 260;

                    dgw.Columns[1].HeaderText = "Валйута";
                    dgw.Columns[1].Width = 120;

                    dgw.Columns[2].HeaderText = "Ады";
                    dgw.Columns[2].Width = 340;

                    dgw.Columns[3].HeaderText = "Доьулдуьу йер";
                    dgw.Columns[3].Width = 270;

                    dgw.Columns[4].HeaderText = "Доьум тарихи";
                    dgw.Columns[4].Width = 130;

                    dgw.Columns[5].HeaderText = "Гейдиййат цнваны";
                    dgw.Columns[5].Width = 350;

                    dgw.Columns[6].HeaderText = "Йашадыь цнваны";
                    dgw.Columns[6].Width = 350;

                    dgw.Columns[7].HeaderText = "Вятяндашлыг";
                    dgw.Columns[7].Width = 120;

                    dgw.Columns[8].HeaderText = "Сянядин верилмя тарихи вя орган";
                    dgw.Columns[8].Width = 250;

                    dgw.Columns[9].HeaderText = "Ш/В Серийа вя нюмряси";
                    dgw.Columns[9].Width = 180;

                    dgw.Columns[10].HeaderText = "Иш йери";
                    dgw.Columns[10].Width = 420;

                    dgw.Columns[11].HeaderText = "Вязифяси";
                    dgw.Columns[11].Width = 250;

                    dgw.Columns[12].HeaderText = "Эялир мянбяйи";
                    dgw.Columns[12].Width = 150;

                    dgw.Columns[13].HeaderText = "Телефон";
                    dgw.Columns[13].Width = 480;

                    dgw.Columns[14].HeaderText = "Факс";
                    dgw.Columns[14].Width = 200;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Sahibkar(DataGridView dgw)
        {
            try
            {
                dgw.Columns.Clear();
                dgw.Refresh();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("Select l.az||l.nr||l.bank||l.licsch, k.ekv, r.name_regnom, r.inn_regnom, r.grajdanstvo, r.registrac, r.adress, r.telefon, r.faks from odb.licsch l inner join odb.regnom r on r.regnom = l.registrac_nomer inner join odb.ikv_ekv k on k.ikv = SUBSTR(l.licsch,7,1) where l.date_open_licsch=to_date('" + tarix + "', 'dd-MM-yyyy') and SUBSTR(l.licsch,0,5) in ('41040','41050','41045','41055')" + sglkod.Text + "", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Щесаб нюмряси";
                dgw.Columns[0].Width = 260;

                dgw.Columns[1].HeaderText = "Валйута";
                dgw.Columns[1].Width = 90;

                dgw.Columns[2].HeaderText = "Ады";
                dgw.Columns[2].Width = 335;

                dgw.Columns[3].HeaderText = "ВЮЕН";
                dgw.Columns[3].Width = 120;

                dgw.Columns[4].HeaderText = "Вятяндашлыг";
                dgw.Columns[4].Width = 120;

                dgw.Columns[5].HeaderText = "Гейдиййат цнваны";
                dgw.Columns[5].Width = 350;

                dgw.Columns[6].HeaderText = "Йашадыь цнван";
                dgw.Columns[6].Width = 480;

                dgw.Columns[7].HeaderText = "Телефон";
                dgw.Columns[7].Width = 480;

                dgw.Columns[8].HeaderText = "Факс";
                dgw.Columns[8].Width = 200;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Hugugishexs(DataGridView dgw)
        {
            try
            {
                dgw.Columns.Clear();
                dgw.Refresh();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("Select l.az||l.nr||l.bank||l.licsch, k.ekv, r.name_regnom, r.inn_regnom, r.registrac, r.telefon, r.faks  from odb.licsch l inner join odb.regnom r on r.regnom = l.registrac_nomer inner join odb.ikv_ekv k on k.ikv = SUBSTR(l.licsch,7,1) where l.date_open_licsch=to_date('" + tarix + "', 'dd-MM-yyyy') and (l.licsch like '3%' or l.licsch like '40%')" + sglkod.Text + "", Orcon);
                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Щесаб нюмряси";
                dgw.Columns[0].Width = 260;

                dgw.Columns[1].HeaderText = "Валйута";
                dgw.Columns[1].Width = 90;

                dgw.Columns[2].HeaderText = "Ады";
                dgw.Columns[2].Width = 335;

                dgw.Columns[3].HeaderText = "ВЮЕН";
                dgw.Columns[3].Width = 120;

                dgw.Columns[4].HeaderText = "Цнваны";
                dgw.Columns[4].Width = 350;

                dgw.Columns[5].HeaderText = "Телефон";
                dgw.Columns[5].Width = 480;

                dgw.Columns[6].HeaderText = "Факс";
                dgw.Columns[6].Width = 200;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Plastickart(DataGridView dgw)
        {
            try
            {
                dgw.Columns.Clear();
                dgw.Refresh();
                Orcon = new OracleConnection("Data Source=BMI;User ID=FOXPRO;Password=pass");
                Orcon.Open();
                Orcom = new OracleCommand("select l.az||l.nr||l.bank||l.licsch as Щесаб_нюмряси, k.ekv, r.name_regnom, r.mesto_rojdeniya, r.data_rojdeniya, (case when r.m=1 then 'Киши' else 'Гадын' end) as Ъинси, r.semeynoye_polojenie, r.registrac, r.adress, r.grajdanstvo, r.senedi_veren_orqaninin_adi, r.senedin_verilme_tarixi, r.passport, f.fealiyyet_ve_peshesi, f.vezifesi, f.gelir_menbeyi, r.telefon, (case SUBSTR(l.licsch,16, 2) when '91' then 'Локал (Ямяк Щаггы)'  when '92' then 'Локал Кредит' when '93' then 'Локал Дебет' else 'Мювъуд дейил' end) as Картын_нювц from odb.licsch l left outer join odb.regnom r on r.regnom = l.registrac_nomer left outer join odb.ikv_ekv k on k.ikv = SUBSTR(l.licsch,7,1) left outer join odb.fiziki_shexs f on f.regnom = l.registrac_nomer where l.date_open_licsch=to_date('" + tarix + "', 'dd-MM-yyyy') and SUBSTR( l.licsch,0,5) in ('41010','41020','41015','41025') and SUBSTR(l.licsch,16,1)=9" + sglkod.Text + "", Orcon);

                Orda = new OracleDataAdapter(Orcom);
                dt = new DataTable();
                Orda.Fill(dt);
                mybing.DataSource = dt;
                dgw.DataSource = mybing;
                Orcon.Close();

                dgw.Columns[0].HeaderText = "Щесаб нюмряси";
                dgw.Columns[0].Width = 260;

                dgw.Columns[1].HeaderText = "Валйута";
                dgw.Columns[1].Width = 120;

                dgw.Columns[2].HeaderText = "Ады";
                dgw.Columns[2].Width = 340;

                dgw.Columns[3].HeaderText = "Доьулдуьу йер";
                dgw.Columns[3].Width = 270;

                dgw.Columns[4].HeaderText = "Доьум тарихи";
                dgw.Columns[4].Width = 130;

                dgw.Columns[5].HeaderText = "Ъинси";
                dgw.Columns[5].Width = 130;

                dgw.Columns[6].HeaderText = "Аиля вязиййяти";
                dgw.Columns[6].Width = 130;

                dgw.Columns[7].HeaderText = "Гейдиййат цнваны";
                dgw.Columns[7].Width = 360;

                dgw.Columns[8].HeaderText = "Йашадыь цнван";
                dgw.Columns[8].Width = 360;

                dgw.Columns[9].HeaderText = "Вятяндашлыг";
                dgw.Columns[9].Width = 120;

                dgw.Columns[10].HeaderText = "Сяняди верен орган";
                dgw.Columns[10].Width = 180;

                dgw.Columns[11].HeaderText = "Сянядин верилмя тарихи";
                dgw.Columns[11].Width = 180;

                dgw.Columns[12].HeaderText = "Ш/В Серийа вя нюмряси";
                dgw.Columns[12].Width = 180;

                dgw.Columns[13].HeaderText = "Иш йери";
                dgw.Columns[13].Width = 420;

                dgw.Columns[14].HeaderText = "Вязифяси";
                dgw.Columns[14].Width = 250;

                dgw.Columns[15].HeaderText = "Эялир мянбяйи";
                dgw.Columns[15].Width = 150;

                dgw.Columns[16].HeaderText = "Телефон";
                dgw.Columns[16].Width = 480;

                dgw.Columns[17].HeaderText = "Картын нювц";
                dgw.Columns[17].Width = 150;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string Hesabacilmatarix()
        {
            gunayil = tarix.Split('-');
            //gunayil[0] = "“" + gunayil[0] + "”";    //"«" + gunayil[0] + "»";

            switch (gunayil[1])
            {
                case "01": gunayil[1] = "yanvar";
                    break;
                case "02": gunayil[1] = "fevral";
                    break;
                case "03": gunayil[1] = "mart";
                    break;
                case "04": gunayil[1] = "aprel";
                    break;
                case "05": gunayil[1] = "may";
                    break;
                case "06": gunayil[1] = "iyun";
                    break;
                case "07": gunayil[1] = "iyul";
                    break;
                case "08": gunayil[1] = "avqust";
                    break;
                case "09": gunayil[1] = "sentyabr";
                    break;
                case "10": gunayil[1] = "oktyabr";
                    break;
                case "11": gunayil[1] = "noyabr";
                    break;
                case "12": gunayil[1] = "dekabr";
                    break;
                default: gunayil[1] = "Hansı ay olduğu məlum deyil";
                    break;
            }
            switch (gunayil[2].Substring(2, 2))
            {
                case "13": gunayil[2] = gunayil[2] + "-cü il";
                    break;
                case "14": gunayil[2] = gunayil[2] + "-cü il";
                    break;
                case "16": gunayil[2] = gunayil[2] + "-cı il";
                    break;
                case "17": gunayil[2] = gunayil[2] + "-ci il";
                    break;
                case "18": gunayil[2] = gunayil[2] + "-ci il";
                    break;
                case "19": gunayil[2] = gunayil[2] + "-cu il";
                    break;
                case "20": gunayil[2] = gunayil[2] + "-ci il";
                    break;
                case "21": gunayil[2] = gunayil[2] + "-ci il";
                    break;
                case "22": gunayil[2] = gunayil[2] + "-ci il";
                    break;
                case "23": gunayil[2] = gunayil[2] + "-cü il";
                    break;
                case "24": gunayil[2] = gunayil[2] + "-cü il";
                    break;
                default: gunayil[2] = gunayil[2] + "-ci il";
                    break;
            }
            return gunayil[0] + " " + gunayil[1] + " " + gunayil[2];
        }

        public void ChangeKeyboardENG()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage);
            InputLanguage.CurrentInputLanguage = l;
        }

        public void ChangeKeyboardRUS()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("ru-RUS");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage); InputLanguage.CurrentInputLanguage = l;
        }

        public void ChangeKeyboardAZE()
        {
            CultureInfo TypeOfLanguage = CultureInfo.CreateSpecificCulture("az-AZE");
            Thread.CurrentThread.CurrentCulture = TypeOfLanguage;
            InputLanguage l = InputLanguage.FromCulture(TypeOfLanguage); InputLanguage.CurrentInputLanguage = l;
        }
    }
}
