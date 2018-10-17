using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace Units_import_system
{
    public partial class Form1 : Form
    {
        string strConn = "server=tcp:WIN-3AKK5VRL0P3\\SQLEXPRESS,49172;database=test;User ID=test;Password=Oct0b1vju1v38;";//登入資料庫帳密跟port和Database    
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        DataTable a = new DataTable();
        Form2 fm2;

        public Form1()
        {
            InitializeComponent();
            fm2 = new Form2(this);
            SqlConnection myConn = new SqlConnection(strConn);//建立連接   
            cmd.Connection = myConn;//讓cmd連資料庫
            try
            {
                myConn.Open();
            }
            catch
            {
                MessageBox.Show("Cannot connect to database.");
            }

            if (myConn.State != ConnectionState.Open)
            {
                MessageBox.Show("Database state:" + myConn.State.ToString());//資料庫連結失敗，show狀態
            }
            //以下year1的下拉選項
            year1.Items.Add("");
            year1.Items.Add("2008");
            year1.Items.Add("2009");
            year1.Items.Add("2010");
            year1.Items.Add("2011");
            year1.Items.Add("2012");
            year1.Items.Add("2013");
            year1.Items.Add("2014");
            year1.Items.Add("2015");
            year1.Items.Add("2016");
            year1.Items.Add("2011");
            year1.Items.Add("2017");
            year1.Items.Add("2018");
            year1.Items.Add("2019");
            year1.Items.Add("2020");
            year1.Items.Add("2021");
            year1.Items.Add("2022");
            year1.Items.Add("2023");
            year1.Items.Add("2024");

            //以下year2的下拉選項
            year2.Enabled = false;//先不能選除非type==CNB
            year2.Items.Add("");
            year2.Items.Add("C1");
            year2.Items.Add("C2");
            year2.Items.Add("C3");

            //以下type的下拉選項
            type.Items.Add("");
            type.Items.Add("CNB");
            type.Items.Add("BNB");
            type.Items.Add("CDT");
            type.Items.Add("BDT");
            type.Items.Add("CAIO");
            type.Items.Add("BAIO");

            //以下phase的下拉選項
            phase1.Items.Add("");
            phase1.Items.Add("DB");
            phase1.Items.Add("SI");
            phase1.Items.Add("PV");
            phase1.Items.Add("MV");

            //phase2.Items.Add("-1");
            phase2.Items.Add("");
            phase2.Items.Add("-2");
            phase2.Items.Add("-3");
            phase2.Items.Add("-4");

            a.Columns.Add("No.");
            a.Columns.Add("SN");//0
            a.Columns.Add("SKU");//1

            phase2.DropDownStyle = ComboBoxStyle.DropDownList;
            phase1.DropDownStyle = ComboBoxStyle.DropDownList;
            year1.DropDownStyle = ComboBoxStyle.DropDownList;
            year2.DropDownStyle = ComboBoxStyle.DropDownList;
            type.DropDownStyle = ComboBoxStyle.DropDownList;


        }

        private void type_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((type.Text == ("CNB")) || (type.Text == ("CDT")) || (type.Text == ("CAIO")))
            {
                year2.Enabled = true;
            }
            else
            {
                year2.Enabled = false;
                year2.Text = "";
            }
        }

        private void sn1_TextChanged(object sender, EventArgs e)
        {
            if (sn1.Text.Length >= 10)
            {
                sn2.Focus();
            }
        }

        private void sn2_TextChanged(object sender, EventArgs e)
        {
            if (sn2.Text.Length >= 10)
            {
                sn3.Focus();
            }
        }

        private void sn3_TextChanged(object sender, EventArgs e)
        {
            if (sn3.Text.Length >= 10)
            {
                sn4.Focus();
            }
        }

        private void sn4_TextChanged(object sender, EventArgs e)
        {
            if (sn4.Text.Length >= 10)
            {
                sn5.Focus();
            }
        }

        private void sn5_TextChanged(object sender, EventArgs e)
        {
            if (sn5.Text.Length >= 10)
            {
                sn6.Focus();
            }
        }

        private void sn6_TextChanged(object sender, EventArgs e)
        {
            if (sn6.Text.Length >= 10)
            {
                sn7.Focus();
            }
        }

        private void sn7_TextChanged(object sender, EventArgs e)
        {
            if (sn7.Text.Length >= 10)
            {
                sn8.Focus();
            }
        }

        private void sn8_TextChanged(object sender, EventArgs e)
        {
            if (sn8.Text.Length >= 10)
            {
                sn9.Focus();
            }
        }

        private void sn9_TextChanged(object sender, EventArgs e)
        {
            if (sn9.Text.Length >= 10)
            {
                sn10.Focus();
            }
        }

        private void sn10_TextChanged(object sender, EventArgs e)
        {
            if (sn10.Text.Length >= 10)
            {
                sn11.Focus();
            }
        }

        private void sn11_TextChanged(object sender, EventArgs e)
        {
            if (sn11.Text.Length >= 10)
            {
                sn12.Focus();
            }
        }

        private void sn12_TextChanged(object sender, EventArgs e)
        {
            if (sn12.Text.Length >= 10)
            {
                sn13.Focus();
            }
        }

        private void sn13_TextChanged(object sender, EventArgs e)
        {
            if (sn13.Text.Length >= 10)
            {
                sn14.Focus();
            }
        }

        private void sn14_TextChanged(object sender, EventArgs e)
        {
            if (sn14.Text.Length >= 10)
            {
                sn15.Focus();
            }
        }

        private void sn15_TextChanged(object sender, EventArgs e)
        {
            if (sn15.Text.Length >= 10)
            {
                sn16.Focus();
            }
        }

        private void sn16_TextChanged(object sender, EventArgs e)
        {
            if (sn16.Text.Length >= 10)
            {
                sn17.Focus();
            }
        }

        private void sn17_TextChanged(object sender, EventArgs e)
        {
            if (sn17.Text.Length >= 10)
            {
                sn18.Focus();
            }
        }

        private void sn18_TextChanged(object sender, EventArgs e)
        {
            if (sn18.Text.Length >= 10)
            {
                sn19.Focus();
            }
        }

        private void sn19_TextChanged(object sender, EventArgs e)
        {
            if (sn19.Text.Length >= 10)
            {
                sn20.Focus();
            }
        }

        private void sn20_TextChanged(object sender, EventArgs e)
        {

        }

        public string cmdString1 = "";
        public string cmdString2 = "";
        public string cmdString3 = "";

        private void submit_Click(object sender, EventArgs e)
        {
            a.Clear();
            String erroMsg = "";
            cmdString1 = "";
            cmdString2 = "";
            cmdString3 = "";
            bool empty = false;

            string[] name = new string[20];
            if (checkBox1.Checked)
            {
                if (sn1.Text == "")
                {
                    empty = true;
                }
                name[0] = sn1.Text;
            }
            if (checkBox2.Checked)
            {
                if (sn2.Text == "")
                {
                    empty = true;
                }
                name[1] = sn2.Text;
            }
            if (checkBox3.Checked)
            {
                if (sn3.Text == "")
                {
                    empty = true;
                }
                name[2] = sn3.Text;
            }
            if (checkBox4.Checked)
            {
                if (sn4.Text == "")
                {
                    empty = true;
                }
                name[3] = sn4.Text;
            }
            if (checkBox5.Checked)
            {
                if (sn5.Text == "")
                {
                    empty = true;
                }
                name[4] = sn5.Text;
            }
            if (checkBox6.Checked)
            {
                if (sn6.Text == "")
                {
                    empty = true;
                }
                name[5] = sn6.Text;
            }
            if (checkBox7.Checked)
            {
                if (sn7.Text == "")
                {
                    empty = true;
                }
                name[6] = sn7.Text;
            }
            if (checkBox8.Checked)
            {
                if (sn8.Text == "")
                {
                    empty = true;
                }
                name[7] = sn8.Text;
            }
            if (checkBox9.Checked)
            {
                if (sn9.Text == "")
                {
                    empty = true;
                }
                name[8] = sn9.Text;
            }
            if (checkBox10.Checked)
            {
                if (sn10.Text == "")
                {
                    empty = true;
                }
                name[9] = sn10.Text;
            }
            if (checkBox11.Checked)
            {
                if (sn11.Text == "")
                {
                    empty = true;
                }
                name[10] = sn11.Text;
            }
            if (checkBox12.Checked)
            {
                if (sn12.Text == "")
                {
                    empty = true;
                }
                name[11] = sn12.Text;
            }
            if (checkBox13.Checked)
            {
                if (sn13.Text == "")
                {
                    empty = true;
                }
                name[12] = sn13.Text;
            }
            if (checkBox14.Checked)
            {
                if (sn14.Text == "")
                {
                    empty = true;
                }
                name[13] = sn14.Text;
            }
            if (checkBox15.Checked)
            {
                if (sn15.Text == "")
                {
                    empty = true;
                }
                name[14] = sn15.Text;
            }
            if (checkBox16.Checked)
            {
                if (sn16.Text == "")
                {
                    empty = true;
                }
                name[15] = sn16.Text;
            }
            if (checkBox17.Checked)
            {
                if (sn17.Text == "")
                {
                    empty = true;
                }
                name[16] = sn17.Text;
            }
            if (checkBox18.Checked)
            {
                if (sn18.Text == "")
                {
                    empty = true;
                }
                name[17] = sn18.Text;
            }
            if (checkBox19.Checked)
            {
                if (sn19.Text == "")
                {
                    empty = true;
                }
                name[18] = sn19.Text;
            }
            if (checkBox20.Checked)
            {
                if (sn20.Text == "")
                {
                    empty = true;
                }
                name[19] = sn20.Text;
            }

            if(empty==false)
            {
                for (int i = 0; i < 20; i++)
                {
                    if (name[i] != "" && name[i] != null)
                    {
                        for (int j = i + 1; j < 20; j++)
                        {
                            if (name[i] == name[j] && name[j] != "")
                            {
                                erroMsg += "Unit " + (i + 1) + ": SN: " + name[i] + " & " + "Unit " + (j + 1) + ": SN: " + name[j] + ": Data duplicate! \n";
                            }
                        }
                    }
                }

                fm2.dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(51, 153, 102);
                fm2.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial Rounded MT", 13, FontStyle.Bold);
                fm2.dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                fm2.dataGridView1.DefaultCellStyle.Font = new Font("Arial Rounded MT", 12);
                fm2.dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(242, 255, 204);
                fm2.dataGridView1.EnableHeadersVisualStyles = false;
                fm2.dataGridView1.RowHeadersWidth = 20;

                fm2.platformName2.Text = platformName.Text;
                fm2.phaseName.Text = phase1.Text + phase2.Text;
                fm2.yearName.Text = year1.Text + year2.Text;
                fm2.typeName.Text = type.Text;


                SqlConnection myConn = new SqlConnection(strConn);//建立連接     
                cmd.Connection = myConn;//讓cmd連資料庫
                myConn.Open();
                int rowNum = 0;

                if ((checkBox1.Checked == true) && (sn1.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn1.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn1.Text, SKU1.Text);
                        addRow(rowNum, sn1.Text, SKU1.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 1: SN: " + sn1.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox2.Checked == true) && (sn2.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn2.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn2.Text, SKU2.Text);
                        addRow(rowNum, sn2.Text, SKU2.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 2: SN: " + sn2.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox3.Checked == true) && (sn3.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn3.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn3.Text, SKU3.Text);
                        addRow(rowNum, sn3.Text, SKU3.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 3: SN: " + sn3.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }

                if ((checkBox4.Checked == true) && (sn4.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn4.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn4.Text, SKU4.Text);
                        addRow(rowNum, sn4.Text, SKU4.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 4: SN: " + sn4.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }

                if ((checkBox5.Checked == true) && (sn5.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn5.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn5.Text, SKU5.Text);
                        addRow(rowNum, sn5.Text, SKU5.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 5: SN: " + sn5.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox6.Checked == true) && (sn6.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn6.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn6.Text, SKU6.Text);
                        addRow(rowNum, sn6.Text, SKU6.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 6: SN: " + sn6.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox7.Checked == true) && (sn7.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn7.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString1 += submit1(sn7.Text, SKU7.Text);
                        addRow(rowNum, sn7.Text, SKU7.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 7: SN: " + sn7.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox8.Checked == true) && (sn8.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn8.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn8.Text, SKU8.Text);
                        addRow(rowNum, sn8.Text, SKU8.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 8: SN: " + sn8.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox9.Checked == true) && (sn9.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn9.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn9.Text, SKU9.Text);
                        addRow(rowNum, sn9.Text, SKU9.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 9: SN: " + sn9.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox10.Checked == true) && (sn10.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn10.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn10.Text, SKU10.Text);
                        addRow(rowNum, sn10.Text, SKU10.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 10: SN: " + sn10.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox11.Checked == true) && (sn11.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn11.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn11.Text, SKU11.Text);
                        addRow(rowNum, sn11.Text, SKU11.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 11: SN: " + sn11.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox12.Checked == true) && (sn12.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn12.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn12.Text, SKU12.Text);
                        addRow(rowNum, sn12.Text, SKU12.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 12: SN: " + sn12.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox13.Checked == true) && (sn13.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn13.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn13.Text, SKU13.Text);
                        addRow(rowNum, sn13.Text, SKU13.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 13: SN: " + sn13.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox14.Checked == true) && (sn14.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn14.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString2 += submit1(sn14.Text, SKU14.Text);
                        addRow(rowNum, sn14.Text, SKU14.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 14: SN: " + sn14.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox15.Checked == true) && (sn15.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn15.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn15.Text, SKU15.Text);
                        addRow(rowNum, sn15.Text, SKU15.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 15: SN: " + sn15.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox16.Checked == true) && (sn16.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn16.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn16.Text, SKU16.Text);
                        addRow(rowNum, sn16.Text, SKU16.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 16: SN: " + sn16.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox17.Checked == true) && (sn17.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn17.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn17.Text, SKU17.Text);
                        addRow(rowNum, sn17.Text, SKU17.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 17: SN: " + sn17.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox18.Checked == true) && (sn18.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn18.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn18.Text, SKU18.Text);
                        addRow(rowNum, sn18.Text, SKU18.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 18: SN: " + sn18.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox19.Checked == true) && (sn19.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn19.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn19.Text, SKU19.Text);
                        addRow(rowNum, sn19.Text, SKU19.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 19: SN: " + sn19.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if ((checkBox20.Checked == true) && (sn20.Text != ""))
                {
                    cmd.CommandText = "Select * from UnitTable where SN='" + sn20.Text + "'";
                    dr = cmd.ExecuteReader();
                    rowNum++;

                    if (!dr.HasRows)
                    {
                        cmdString3 += submit1(sn20.Text, SKU20.Text);
                        addRow(rowNum, sn20.Text, SKU20.Text);
                    }
                    else
                    {
                        erroMsg += "Unit 20: SN: " + sn20.Text + ": Data duplicate with database! \n";
                    }
                    dr.Close();
                }
                if (erroMsg == "")
                {
                    //cmd.CommandText = cmdString;
                    //cmd.ExecuteNonQuery();
                    if (cmdString1 != "" || cmdString2 != "" || cmdString3 != "")
                    {
                        fm2.dataGridView1.DataSource = a;
                        fm2.Show();/* lForm.Show();*/
                    }
                    else
                    {
                        MessageBox.Show("Please tick which you want to record.");
                    }
                }
                else
                {
                    MessageBox.Show(erroMsg);
                }
            }
            else
            {
                MessageBox.Show("SN cannot be empty!");
            }
            
        }
        private string submit1(String SN_text, String SKU_text)
        {
            String cmd_text = "";
            cmd_text += "INSERT INTO UnitTable VALUES('" + SN_text + "','"//1,SN
                        + platformName.Text + "','"//2,pf
                        + phase1.Text + phase2.Text + "','"//3,ph
                        + SKU_text + "'," //4,sku
                        + "NULL," //5,cat
                        + "NULL,"//6,borrower
                        + "NULL,'"//7,status
                        + year1.Text + year2.Text + "','"//8,year+cycle
                        + type.Text + "',"//9,catagegry(type)
                        + "NULL,"//10,note
                        + "NULL,"//11,borrowdate1
                        + "NULL,'"//12,positon
                        + DateTime.Now.ToString("yyyy - MM - dd HH: mm: ss") + "',"//13,keyintime
                        + "NULL,"//14,mail-1
                        + "NULL,"//15,mail-2
                        + "NULL,"//16,mail-3
                        + "NULL,"//17,borrowdate2
                        + "NULL,"//18,borrowdate3
                        + "NULL,"//19,CPU
                        + "NULL,"//20,WLAN
                        + "NULL,"//21,note1
                        + "NULL,"//22,note2
                        + "NULL); ";//23,note3
            return cmd_text;
        }
        private void addRow(int num, String SN_text, String SKU_text)
        {
            DataRow b = a.NewRow();
            b[0] = num.ToString();
            b[1] = SN_text;
            b[2] = SKU_text;
            a.Rows.Add(b);
        }

        private void checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAll.Checked == true)
            {
                if ((sn1.Text != "") || (SKU1.Text != ""))
                {
                    checkBox1.Checked = true;
                }
                if ((sn2.Text != "") || (SKU2.Text != ""))
                {
                    checkBox2.Checked = true;
                }
                if ((sn3.Text != "") || (SKU3.Text != ""))
                {
                    checkBox3.Checked = true;
                }
                if ((sn4.Text != "") || (SKU4.Text != ""))
                {
                    checkBox4.Checked = true;
                }
                if ((sn5.Text != "") || (SKU5.Text != ""))
                {
                    checkBox5.Checked = true;
                }
                if ((sn6.Text != "") || (SKU6.Text != ""))
                {
                    checkBox6.Checked = true;
                }
                if ((sn7.Text != "") || (SKU7.Text != ""))
                {
                    checkBox7.Checked = true;
                }
                if ((sn8.Text != "") || (SKU8.Text != ""))
                {
                    checkBox8.Checked = true;
                }
                if ((sn9.Text != "") || (SKU9.Text != ""))
                {
                    checkBox9.Checked = true;
                }
                if ((sn10.Text != "") || (SKU10.Text != ""))
                {
                    checkBox10.Checked = true;
                }
                if ((sn11.Text != "") || (SKU11.Text != ""))
                {
                    checkBox11.Checked = true;
                }
                if ((sn12.Text != "") || (SKU12.Text != ""))
                {
                    checkBox12.Checked = true;
                }
                if ((sn13.Text != "") || (SKU13.Text != ""))
                {
                    checkBox13.Checked = true;
                }
                if ((sn14.Text != "") || (SKU14.Text != ""))
                {
                    checkBox14.Checked = true;
                }
                if ((sn15.Text != "") || (SKU15.Text != ""))
                {
                    checkBox15.Checked = true;
                }
                if ((sn16.Text != "") || (SKU16.Text != ""))
                {
                    checkBox16.Checked = true;
                }
                if ((sn17.Text != "") || (SKU17.Text != ""))
                {
                    checkBox17.Checked = true;
                }
                if ((sn18.Text != "") || (SKU18.Text != ""))
                {
                    checkBox18.Checked = true;
                }
                if ((sn19.Text != "") || (SKU19.Text != ""))
                {
                    checkBox19.Checked = true;
                }
                if ((sn20.Text != "") || (SKU20.Text != ""))
                {
                    checkBox20.Checked = true;
                }
            }
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox19.Checked = false;
                checkBox20.Checked = false;
            }

        }

        public string cmd1()
        {
            return cmdString1;
        }

        public string cmd2()
        {
            return cmdString2;
        }

        public string cmd3()
        {
            return cmdString3;
        }

        public void clearAll()
        {
            sn1.Text = "";
            SKU1.Text = "";
            sn2.Text = "";
            SKU2.Text = "";
            sn3.Text = "";
            SKU3.Text = "";
            sn4.Text = "";
            SKU4.Text = "";
            sn5.Text = "";
            SKU5.Text = "";
            sn6.Text = "";
            SKU6.Text = "";
            sn7.Text = "";
            SKU7.Text = "";
            sn8.Text = "";
            SKU8.Text = "";
            sn9.Text = "";
            SKU9.Text = "";
            sn10.Text = "";
            SKU10.Text = "";
            sn11.Text = "";
            SKU11.Text = "";
            sn12.Text = "";
            SKU12.Text = "";
            sn13.Text = "";
            SKU13.Text = "";
            sn14.Text = "";
            SKU14.Text = "";
            sn15.Text = "";
            SKU15.Text = "";
            sn16.Text = "";
            SKU16.Text = "";
            sn17.Text = "";
            SKU17.Text = "";
            sn18.Text = "";
            SKU18.Text = "";
            sn19.Text = "";
            SKU19.Text = "";
            sn20.Text = "";
            SKU20.Text = "";

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox19.Checked = false;
            checkBox20.Checked = false;
            checkBoxAll.Checked = false;

            platformName.Text = "";
            phase1.Text = "";
            phase2.Text = "";
            year1.Text = "";
            year2.Text = "";
            type.Text = "";
        }
        private void clear_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                sn1.Text = "";
                SKU1.Text = "";
            }
            if (checkBox2.Checked == true)
            {
                sn2.Text = "";
                SKU2.Text = "";
            }
            if (checkBox3.Checked == true)
            {
                sn3.Text = "";
                SKU3.Text = "";
            }
            if (checkBox4.Checked == true)
            {
                sn4.Text = "";
                SKU4.Text = "";
            }
            if (checkBox5.Checked == true)
            {
                sn5.Text = "";
                SKU5.Text = "";
            }
            if (checkBox6.Checked == true)
            {
                sn6.Text = "";
                SKU6.Text = "";
            }
            if (checkBox7.Checked == true)
            {
                sn7.Text = "";
                SKU7.Text = "";
            }
            if (checkBox8.Checked == true)
            {
                sn8.Text = "";
                SKU8.Text = "";
            }
            if (checkBox9.Checked == true)
            {
                sn9.Text = "";
                SKU9.Text = "";
            }
            if (checkBox10.Checked == true)
            {
                sn10.Text = "";
                SKU10.Text = "";
            }
            if (checkBox11.Checked == true)
            {
                sn11.Text = "";
                SKU11.Text = "";
            }
            if (checkBox12.Checked == true)
            {
                sn12.Text = "";
                SKU12.Text = "";
            }
            if (checkBox13.Checked == true)
            {
                sn13.Text = "";
                SKU13.Text = "";
            }
            if (checkBox14.Checked == true)
            {
                sn14.Text = "";
                SKU14.Text = "";
            }
            if (checkBox15.Checked == true)
            {
                sn15.Text = "";
                SKU15.Text = "";
            }
            if (checkBox16.Checked == true)
            {
                sn16.Text = "";
                SKU16.Text = "";
            }
            if (checkBox17.Checked == true)
            {
                sn17.Text = "";
                SKU17.Text = "";
            }
            if (checkBox18.Checked == true)
            {
                sn18.Text = "";
                SKU18.Text = "";
            }
            if (checkBox19.Checked == true)
            {
                sn19.Text = "";
                SKU19.Text = "";
            }
            if (checkBox20.Checked == true)
            {
                sn20.Text = "";
                SKU20.Text = "";
            }
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            checkBox19.Checked = false;
            checkBox20.Checked = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
