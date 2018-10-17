using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Win32Mapi;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace Units_display
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        string strConn = "server=tcp:WIN-3AKK5VRL0P3\\SQLEXPRESS,49172;database=test;User ID=test;Password=Oct0b1vju1v38;";//登入資料庫帳密跟port和Database    
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        DataTable a = new DataTable();
        string[] SN_original;
        private Mapi ma = new Mapi();
        private bool first_activated = false;
        //string outputText = "";

        public Form1()
        {
            InitializeComponent();
            table.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(51, 153, 102);
            table.ColumnHeadersDefaultCellStyle.Font = new Font("Arial Rounded MT", 13, FontStyle.Bold);
            table.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            table.DefaultCellStyle.Font = new Font("Arial Rounded MT", 12);
            table.DefaultCellStyle.BackColor = Color.FromArgb(242, 255, 204);
            table.EnableHeadersVisualStyles = false;
            //連資料庫
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

            //表單預設
            table.RowHeadersWidth = 20;
            a.Columns.Add(" ", typeof(Boolean));
            a.Columns.Add("Platform");//0
            a.Columns.Add("Phase");//1
            a.Columns.Add("SKU");//2
            a.Columns.Add("S/N");//3
            a.Columns.Add("Borrower");//4
            a.Columns.Add("Status");//5
            a.Columns.Add("Year");//6
            a.Columns.Add("Category");//7
            a.Columns.Add("Note");//8
            a.Columns.Add("Date");//9
            a.Columns.Add("Position");//10
            a.Columns.Add("Mail");//11
            a.Columns.Add("CPU");//12
            a.Columns.Add("WLAN");//13
            setUpDisplay();

            //table.Columns[3].SortMode = 

            //下拉選單設定
            checkComboboxSetup();

            //button預設
            back.Visible = false;
            submit.Visible = false;
        }

        private void checkComboboxSetup()
        {
            SN.Text = "";
            platform.Text = "";
            phase.Text = "";
            SKU.Text = "";
            borrower.Text = "";
            status.Text = "";
            year.Text = "";
            category.Text = "";
            note.Text = "";
            borrowDate.Text = "";
            position.Text = "";
            CPU.Text = "";
            WLAN.Text = "";

            SN.Items.Clear();
            platform.Items.Clear();
            phase.Items.Clear();
            SKU.Items.Clear();
            borrower.Items.Clear();
            status.Items.Clear();
            year.Items.Clear();
            category.Items.Clear();
            note.Items.Clear();
            borrowDate.Items.Clear();
            position.Items.Clear();
            CPU.Items.Clear();
            WLAN.Items.Clear();

            SN.Items.Add("Select All");
            platform.Items.Add("Select All");
            phase.Items.Add("Select All");
            SKU.Items.Add("Select All");
            borrower.Items.Add("Select All");
            status.Items.Add("Select All");
            year.Items.Add("Select All");
            category.Items.Add("Select All");
            note.Items.Add("Select All");
            borrowDate.Items.Add("Select All");
            position.Items.Add("Select All");
            CPU.Items.Add("Select All");
            WLAN.Items.Add("Select All");

            SN.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.SN_ItemCheck);
            platform.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.platform_ItemCheck);
            phase.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.phase_ItemCheck);
            SKU.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.SKU_ItemCheck);
            borrower.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.borrower_ItemCheck);
            status.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.status_ItemCheck);
            year.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.year_ItemCheck);
            category.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.category_ItemCheck);
            note.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.note_ItemCheck);
            borrowDate.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.borrowDate_ItemCheck);
            position.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.position_ItemCheck);
            CPU.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.SN_ItemCheck);
            WLAN.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.SN_ItemCheck);

            //for (int i = 0; i < table.Rows.Count - 1; i++)
            //{
            //    listItems(platform, table.Rows[i].Cells[1].Value.ToString());//1
            //    listItems(phase, table.Rows[i].Cells[2].Value.ToString());//2
            //    listItems(SKU, table.Rows[i].Cells[3].Value.ToString());//3
            //    SN.Items.Add(table.Rows[i].Cells[4].Value.ToString());//4
            //    listItems(borrower, table.Rows[i].Cells[5].Value.ToString());//5
            //    listItems(status, table.Rows[i].Cells[6].Value.ToString());//6
            //    listItems(year, table.Rows[i].Cells[7].Value.ToString());//7
            //    listItems(category, table.Rows[i].Cells[8].Value.ToString());//8
            //    listItems(note, table.Rows[i].Cells[9].Value.ToString());//9
            //    listItems(borrowDate, table.Rows[i].Cells[10].Value.ToString());//10
            //    listItems(position, table.Rows[i].Cells[11].Value.ToString());//11
            //}

            cmd.CommandText = "Select * From UnitTable";
            dr = cmd.ExecuteReader();


            if (dr.HasRows)
            {
                //添加表格内容
                while (dr.Read())
                {
                    //SN=0, platformName=1, phase=2, SKU=3, CAT=4, borrower=5, unitStatus=6,
                    //yearCycle=7, category=8, note=9, borrowingDate=10, position=11, keyInTime=12,
                    //mailOne=13, mailTwo=14, mailThree=15
                    SN.Items.Add(dr[0].ToString());
                    listItems(platform, dr[1].ToString());
                    listItems(phase, dr[2].ToString());
                    listItems(SKU, dr[3].ToString());
                    listItems(borrower, dr[5].ToString());
                    listItems(status, dr[6].ToString());
                    listItems(year, dr[7].ToString());
                    listItems(category, dr[8].ToString());
                    listItems(note, dr[9].ToString());
                    listItems(borrowDate, dr[12].ToString().Substring(0, 10));
                    listItems(position, dr[11].ToString());
                    listItems(CPU, dr[18].ToString());
                    listItems(CPU, dr[19].ToString());
                }
            }
            dr.Close();
        }

        private void SN_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = SN.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(SN);
            }
        }
        private void platform_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = platform.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(platform);
            }
        }
        private void phase_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = phase.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(phase);
            }
        }
        private void SKU_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = SKU.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(SKU);
            }
        }
        private void borrower_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = borrower.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(borrower);
            }
        }
        private void status_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = status.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(status);
            }
        }
        private void year_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = year.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(year);
            }
        }
        private void category_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = category.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(category);
            }
        }
        private void note_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = note.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(note);
            }
        }
        private void borrowDate_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = borrowDate.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(borrowDate);
            }
        }
        private void position_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //CCBoxItem item = position.Items[e.Index] as CCBoxItem;
            if (e.Index == 0)
            {
                selectAll(position);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //SendForm sf = new SendForm(ref ma);
            //sf.Show();
        }

        private void setUpDisplay()
        {
            cmd.CommandText = "Select * From UnitTable";
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    DataRow b = a.NewRow();
                    //Button mail = new Button();
                    //mail.Click += new EventHandler(this.mail_click);
                    
                    b[0] = false;
                    b[1] = dr[1].ToString();
                    b[2] = dr[2].ToString();
                    b[3] = dr[3].ToString();
                    b[4] = dr[0].ToString();
                    b[5] = dr[5].ToString();
                    b[6] = dr[6].ToString();
                    b[7] = dr[7].ToString();
                    b[8] = dr[8].ToString();
                    b[9] = dr[9].ToString();
                    b[10] = dr[12].ToString().Substring(0, 10);
                    b[11] = dr[11].ToString();
                    b[13] = dr[18].ToString();
                    b[14] = dr[19].ToString();
                    a.Rows.Add(b);
                 }
            }
            dr.Close();
            table.DataSource = a;
            DataGridViewColumn dataGridViewColumn = table.Columns[4];
            dataGridViewColumn.HeaderCell.Style.BackColor = Color.Orange;
            dataGridViewColumn.DefaultCellStyle.BackColor = Color.FromArgb(255, 239, 204);
            //DataGridViewImageColumn mailColumn = new DataGridViewImageColumn();

            //mailColumn.Name = "Mail";
            //DataGridViewCell cell = new DataGridViewTextBoxCell();
            //mailColumn.UseColumnTextForButtonValue = true;
            //table.Columns.Add(mailColumn);

            table.Columns[1].ReadOnly = true;
            table.Columns[2].ReadOnly = true;
            table.Columns[3].ReadOnly = true;
            table.Columns[4].ReadOnly = true;
            table.Columns[5].ReadOnly = true;
            table.Columns[6].ReadOnly = true;
            table.Columns[7].ReadOnly = true;
            table.Columns[8].ReadOnly = true;
            table.Columns[9].ReadOnly = true;
            table.Columns[10].ReadOnly = true;
            table.Columns[11].ReadOnly = true;
            table.Columns[13].ReadOnly = true;
            table.Columns[14].ReadOnly = true;
            table.Columns[12].ReadOnly = true;

            table.Columns[0].Width = 50;
            //table.Columns[10].Width = 175;
            table.Columns[12].Width = 50;
            //table.SortCompare += customSortCompare;
        }

        //private void mail_click(Object sender,  EventArgs e)
        //{

        //}
        private void listItems(CheckComboBoxTest.CheckedComboBox s, string drString)
        {
            

            int i;
            for (i = 1; i < s.Items.Count; i++)
            {
                if (s.Items[i].ToString() == drString)
                {
                    break;
                }
            }
            if (i == s.Items.Count)
            {

                s.Items.Add(drString);
            }
        }

        private void selectAll(CheckComboBoxTest.CheckedComboBox s)
        {
            if (s.GetItemCheckState(0) == CheckState.Unchecked)
            {
                for (int i = 1; i < s.Items.Count; i++)
                {
                    s.SetItemChecked(i, true);
                }
            }
            else if (s.GetItemCheckState(0) == CheckState.Checked)
            {
                for (int i = 1; i < s.Items.Count; i++)
                {
                    s.SetItemChecked(i, false);
                }
            }
        }

        private void clearChecked(CheckComboBoxTest.CheckedComboBox s)
        {
            for (int i = 0; i < s.Items.Count; i++)
            {
                s.SetItemChecked(i, false);
            }
        }

        private void decreseTableItem(CheckComboBoxTest.CheckedComboBox ccb, int num)
        {
            bool exist = false;
            for (int i = 0; i < a.Rows.Count; i++)
            {
                foreach (string s in ccb.CheckedItems)
                {
                    if ((string)a.Rows[i][num] == s)
                    {
                        exist = true;
                    }
                }
                if (exist == false)
                {
                    a.Rows.Remove(a.Rows[i]);
                    i--;
                }
                else
                {
                    exist = false;
                }
            }

            table.DataSource = a;
        }

        private void increaseTableItem()
        {
            a.Clear();
            setUpDisplay();

            if (SN.CheckedItems.Count != 0)
            {
                decreseTableItem(SN, 4);
            }
            if (platform.CheckedItems.Count != 0)
            {
                decreseTableItem(platform, 1);
            }
            if (phase.CheckedItems.Count != 0)
            {
                decreseTableItem(phase, 2);
            }
            if (SKU.CheckedItems.Count != 0)
            {
                decreseTableItem(SKU, 3);
            }
            if (borrower.CheckedItems.Count != 0)
            {
                decreseTableItem(borrower, 5);
            }
            if (status.CheckedItems.Count != 0)
            {
                decreseTableItem(status, 6);
            }
            if (year.CheckedItems.Count != 0)
            {
                decreseTableItem(year, 7);
            }
            if (category.CheckedItems.Count != 0)
            {
                decreseTableItem(category, 8);
            }
            if (note.CheckedItems.Count != 0)
            {
                decreseTableItem(note, 9);
            }
            if (borrowDate.CheckedItems.Count != 0)
            {
                decreseTableItem(borrowDate, 10);
            }
            if (position.CheckedItems.Count != 0)
            {
                decreseTableItem(position, 11);
            }
            if (CPU.CheckedItems.Count != 0)
            {
                decreseTableItem(SN, 4);
            }
            if (WLAN.CheckedItems.Count != 0)
            {
                decreseTableItem(SN, 4);
            }
        }

        private void clear_Click(object sender, EventArgs e)
        {
            a.Clear();
            setUpDisplay();
            //checkComboboxSetup();
            clearChecked(platform);
            clearChecked(SN);
            clearChecked(phase);
            clearChecked(SKU);
            clearChecked(borrower);
            clearChecked(status);
            clearChecked(year);
            clearChecked(category);
            clearChecked(note);
            clearChecked(borrowDate);
            clearChecked(position);
            clearChecked(CPU);
            clearChecked(WLAN);
        }

        public struct unitTB
        {
            public string type;
            public string platform;
            public string phase;
            public int qty;
        }

        public struct unitInfo
        {
            public string Platform;
            public string Phase;
            public string SKU;
            public string SN;
            public string Borrower;
            public string Status;
            public string Year;
            public string Category;
            public string Note;
            public string Date;
        }

        public unitTB[] unitTBs = { };
        public unitInfo[] unitInfos = { };

        private void sendMail_Click(object sender, EventArgs e)
        {
            Array.Clear(unitTBs, 0, unitTBs.Length);
            Array.Resize(ref unitTBs,0);
            Array.Clear(unitInfos, 0, unitInfos.Length);
            Array.Resize(ref unitInfos, 0);

            bool hasCheck = false;
            
            for(int i=0; i<table.Rows.Count-1; i++)
            {
                if (table.Rows[i].Cells[0].Value.ToString()=="True")
                {
                    hasCheck = true;
                    bool hasRecord = false;
                    //MessageBox.Show(table.Rows[i].Cells[4].Value.ToString());
                    for (int j=0; j<unitTBs.Length; j++)
                    {
                        if((unitTBs[j].type == table.Rows[i].Cells[8].Value.ToString())&&(unitTBs[j].platform == table.Rows[i].Cells[1].Value.ToString()) &&(unitTBs[j].phase == table.Rows[i].Cells[2].Value.ToString()))
                        {
                            unitTBs[j].qty += 1;
                            hasRecord = true;
                        }
                    }
                    if (!hasRecord)
                    {
                        unitTB newUnit = new unitTB();
                        newUnit.type = table.Rows[i].Cells[8].Value.ToString();
                        newUnit.platform = table.Rows[i].Cells[1].Value.ToString();
                        newUnit.phase = table.Rows[i].Cells[2].Value.ToString();
                        newUnit.qty = 1;
                        Array.Resize(ref unitTBs, unitTBs.Length + 1);
                        unitTBs[unitTBs.Length - 1] = newUnit;
                    }

                    unitInfo newUnitInfo = new unitInfo();
                    newUnitInfo.Platform = table.Rows[i].Cells[1].Value.ToString();
                    newUnitInfo.Phase = table.Rows[i].Cells[2].Value.ToString();
                    newUnitInfo.SKU = table.Rows[i].Cells[3].Value.ToString();
                    newUnitInfo.SN = table.Rows[i].Cells[4].Value.ToString();
                    newUnitInfo.Borrower = table.Rows[i].Cells[5].Value.ToString();
                    newUnitInfo.Status = table.Rows[i].Cells[6].Value.ToString();
                    newUnitInfo.Year = table.Rows[i].Cells[7].Value.ToString();
                    newUnitInfo.Category = table.Rows[i].Cells[8].Value.ToString();
                    newUnitInfo.Note = table.Rows[i].Cells[9].Value.ToString();
                    newUnitInfo.Date = table.Rows[i].Cells[10].Value.ToString();
                    Array.Resize(ref unitInfos, unitInfos.Length + 1);
                    unitInfos[unitInfos.Length - 1] = newUnitInfo;
                }
            }
            if (hasCheck == false)
            {
                MessageBox.Show("Please tick which you want to sent the mail.");
            }
            else
            {
                Form3 fm3 = new Form3(this);
                fm3.Show();
            }

        }

        private void edit_Click(object sender, EventArgs e)
        {
            SN_original = new string[table.Rows.Count - 1];
            for(int i = 0; i < table.Rows.Count - 1; i++){
                SN_original[i] = table.Rows[i].Cells[4].Value.ToString();
            }
            table.Columns[1].ReadOnly = false;
            table.Columns[2].ReadOnly = false;
            table.Columns[3].ReadOnly = false;
            table.Columns[4].ReadOnly = false;//
            table.Columns[5].ReadOnly = false;
            table.Columns[6].ReadOnly = false;
            table.Columns[7].ReadOnly = false;
            table.Columns[8].ReadOnly = false;
            table.Columns[9].ReadOnly = false;
            table.Columns[10].ReadOnly = true;//
            table.Columns[11].ReadOnly = false;
            table.Columns[12].ReadOnly = true;//
            table.Columns[13].ReadOnly = false;//
            table.Columns[14].ReadOnly = false;//

            table.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[10].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[11].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[12].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[13].SortMode = DataGridViewColumnSortMode.NotSortable;
            table.Columns[14].SortMode = DataGridViewColumnSortMode.NotSortable;

            platform.Enabled = false;
            SN.Enabled = false;
            phase.Enabled = false;
            SKU.Enabled = false;
            borrower.Enabled = false;
            status.Enabled = false;
            year.Enabled = false;
            category.Enabled = false;
            note.Enabled = false;
            borrowDate.Enabled = false;
            position.Enabled = false;
            CPU.Enabled = false;
            WLAN.Enabled = false;

            sendMail.Visible = false;
            clear.Visible = false;
            edit.Visible = false;
            back.Visible = true;
            submit.Visible = true;
            output.Visible = false;
            selectAllItems.Visible = false;

            table.Columns[0].Visible = false;

        }

        private void submit_Click(object sender, EventArgs e)
        {
            bool changed = false;
            bool SN_OK = true;
            bool noError = true;

            dr.Close();

            DataTable newA = new DataTable();
            //newA.Clear();
            newA.Columns.Add("Select", typeof(Boolean));
            newA.Columns.Add("Platform");//0
            newA.Columns.Add("Phase");//1
            newA.Columns.Add("SKU");//2
            newA.Columns.Add("S/N");//3
            newA.Columns.Add("Borrower");//4
            newA.Columns.Add("Status");//5
            newA.Columns.Add("Year");//6
            newA.Columns.Add("Category");//7
            newA.Columns.Add("Note");//8
            newA.Columns.Add("Date");//9
            newA.Columns.Add("Position");//10
            newA.Columns.Add("Mail", typeof(Image));//11
            newA.Columns.Add("CPU");//12
            newA.Columns.Add("WLAN");//10

            List<string> Duplicate_SN = new List<string>();
            List<string> cml = new List<string>();

            for (int i = 0; i < table.Rows.Count - 1; i++)
            {
                if (table.Rows[i].Cells[4].Value.ToString() != SN_original[i])
                {
                    cmd.CommandText = "select * from UnitTable where SN='" + table.Rows[i].Cells[4].Value.ToString() + "'";
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows)
                    {
                        if (!Duplicate_SN.Contains(table.Rows[i].Cells[4].Value.ToString(), StringComparer.OrdinalIgnoreCase))
                        {
                            MessageBox.Show("SN: \"" + table.Rows[i].Cells[4].Value.ToString() + "\" Duplicate!");
                            Duplicate_SN.Add(table.Rows[i].Cells[4].Value.ToString());
                            noError = false;
                        }
                        SN_OK = false;
                    }
                    else if (table.Rows[i].Cells[4].Value.ToString() == "")
                    {
                        MessageBox.Show("SN cannot be empty!");
                        noError = false;
                        SN_OK = false;
                    }
                    else
                    {
                        for (int j = 0; j < table.Rows.Count - 1; j++)
                        {
                            if ((table.Rows[i].Cells[4].Value.ToString() == table.Rows[j].Cells[4].Value.ToString()) && (i != j))
                            {
                                if (!Duplicate_SN.Contains(table.Rows[i].Cells[4].Value.ToString(), StringComparer.OrdinalIgnoreCase))
                                {
                                    MessageBox.Show("SN: \"" + table.Rows[i].Cells[4].Value.ToString() + "\" Duplicate!");
                                    Duplicate_SN.Add(table.Rows[i].Cells[4].Value.ToString());
                                    noError = false;
                                }
                                SN_OK = false;
                                break;
                            }
                        }
                    }
                    dr.Close();
                    if (SN_OK)
                    {
                        changed = true;
                        cmd.CommandText = "Update UnitTable Set SN = '" + table.Rows[i].Cells[4].Value.ToString() + "' where SN='" + SN_original[i] + "'";
                        cmd.ExecuteNonQuery();
                    }
                }
                
                if(SN_OK == true)
                {
                    cmd.CommandText = "select * from UnitTable where SN='" + table.Rows[i].Cells[4].Value.ToString() + "'";
                    dr = cmd.ExecuteReader();
                    cmd.CommandText = "Update UnitTable Set platformName = '" + table.Rows[i].Cells[1].Value.ToString() + "'";
                    dr.Read();
                    if (dr[1].ToString() != table.Rows[i].Cells[1].Value.ToString())
                    {
                        changed = true;
                    }
                    else if (dr[2].ToString() != table.Rows[i].Cells[2].Value.ToString())
                    {
                        cmd.CommandText += ", phase='" + table.Rows[i].Cells[2].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[3].ToString() != table.Rows[i].Cells[3].Value.ToString())
                    {
                        cmd.CommandText += ",SKU = '" + table.Rows[i].Cells[3].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[5].ToString() != table.Rows[i].Cells[5].Value.ToString())
                    {
                       
                        if (dr[15].ToString() != "")//如果三個都滿，全部往前移，加新的
                        {
                            cmd.CommandText += ",mailOne = '" + dr[14].ToString() + "'";
                            cmd.CommandText += ",noteOne = '" + dr[21].ToString() + "'";
                            cmd.CommandText += ",borrowingDate1 = '" + dr[16].ToString() + "'";
                            cmd.CommandText += ",mailTwo = '" + dr[15].ToString() + "'";
                            cmd.CommandText += ",noteTwo = '" + dr[22].ToString() + "'";
                            cmd.CommandText += ",borrowingDate2 = '" + dr[17].ToString() + "'";
                            cmd.CommandText += ",mailThree = '" + table.Rows[i].Cells[5].Value.ToString() + "'";
                            cmd.CommandText += ",noteThree = '" + table.Rows[i].Cells[9].Value.ToString() + "'";
                            cmd.CommandText += ",borrowingDate3 = '" + DateTime.Now.ToString("yyyy/MM/dd") + "'";

                        }
                        else if (dr[14].ToString() != "")//如果兩個滿，新的加在第三個
                        {
                            cmd.CommandText += ",mailThree = '" + table.Rows[i].Cells[5].Value.ToString() + "'";
                            cmd.CommandText += ",noteThree = '" + table.Rows[i].Cells[9].Value.ToString() + "'";
                            cmd.CommandText += ",borrowingDate3 = '" + DateTime.Now.ToString("yyyy/MM/dd") + "'";
                        }
                        else if (dr[13].ToString() != "")//如果只有一個，新的加在第二個
                        {
                            cmd.CommandText += ",mailTwo = '" + table.Rows[i].Cells[5].Value.ToString() + "'";
                            cmd.CommandText += ",noteTwo = '" + table.Rows[i].Cells[9].Value.ToString() + "'";
                            cmd.CommandText += ",borrowingDate2 = '" + DateTime.Now.ToString("yyyy/MM/dd") + "'";
                        }
                        else if (dr[13].ToString() == "")//如果三個全空，新的加在第一個
                        {
                            cmd.CommandText += ",mailOne = '" + table.Rows[i].Cells[5].Value.ToString() + "'";
                            cmd.CommandText += ",noteOne = '" + table.Rows[i].Cells[9].Value.ToString() + "'";
                            cmd.CommandText += ",borrowingDate1 = '" + DateTime.Now.ToString("yyyy/MM/dd") + "'";
                        }
                        cmd.CommandText += ",borrower='" + table.Rows[i].Cells[5].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[6].ToString() != table.Rows[i].Cells[6].Value.ToString())
                    {
                        cmd.CommandText += ",unitStatus='" + table.Rows[i].Cells[6].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[7].ToString() != table.Rows[i].Cells[7].Value.ToString())
                    {
                        cmd.CommandText += ",yearCycle='" + table.Rows[i].Cells[7].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[8].ToString() != table.Rows[i].Cells[8].Value.ToString())
                    {
                        cmd.CommandText += ",category='" + table.Rows[i].Cells[8].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[9].ToString() != table.Rows[i].Cells[9].Value.ToString())
                    {
                        cmd.CommandText += ",note='" + table.Rows[i].Cells[9].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[11].ToString() != table.Rows[i].Cells[11].Value.ToString())
                    {
                        cmd.CommandText += ",position='" + table.Rows[i].Cells[11].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[18].ToString() != table.Rows[i].Cells[13].Value.ToString())
                    {
                        cmd.CommandText += ",CPU='" + table.Rows[i].Cells[13].Value.ToString() + "'";
                        changed = true;
                    }
                    else if (dr[19].ToString() != table.Rows[i].Cells[14].Value.ToString())
                    {
                        cmd.CommandText += ",WLAN='" + table.Rows[i].Cells[14].Value.ToString() + "'";
                        changed = true;
                    }
                    dr.Close();

                    if (changed == true)
                    {
                        cmd.CommandText += ",keyInTime = '" + DateTime.Now.ToString("yyyy/MM/dd") + "' where SN='" + table.Rows[i].Cells[4].Value.ToString() + "'";
                        cml.Add(cmd.CommandText);
                        DataRow b = newA.NewRow();
                        b[0] = false;
                        b[1] = table.Rows[i].Cells[1].Value.ToString();
                        b[2] = table.Rows[i].Cells[2].Value.ToString();
                        b[3] = table.Rows[i].Cells[3].Value.ToString();
                        b[4] = table.Rows[i].Cells[4].Value.ToString();
                        b[5] = table.Rows[i].Cells[5].Value.ToString();
                        b[6] = table.Rows[i].Cells[6].Value.ToString();
                        b[7] = table.Rows[i].Cells[7].Value.ToString();
                        b[8] = table.Rows[i].Cells[8].Value.ToString();
                        b[9] = table.Rows[i].Cells[9].Value.ToString();
                        b[10] = table.Rows[i].Cells[10].Value.ToString();
                        b[11] = table.Rows[i].Cells[11].Value.ToString();
                        b[12] = table.Rows[i].Cells[12].Value;
                        b[13] = table.Rows[i].Cells[13].Value.ToString();
                        b[14] = table.Rows[i].Cells[14].Value.ToString();
                        newA.Rows.Add(b);}
                    changed = false;
                }
                SN_OK = true;
            }

            if (noError == true)
            {
                for(int i=0; i< cml.Count; i++)
                {
                    cmd.CommandText = cml[i];
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException excep)
                    {
                        MessageBox.Show(excep.Message);
                        noError = false;
                    }
                    dr.Close();
                }
                MessageBox.Show("Done!");

                table.DataSource = newA;

                table.Columns[1].ReadOnly = true;
                table.Columns[2].ReadOnly = true;
                table.Columns[3].ReadOnly = true;
                table.Columns[4].ReadOnly = true;
                table.Columns[5].ReadOnly = true;
                table.Columns[6].ReadOnly = true;
                table.Columns[7].ReadOnly = true;
                table.Columns[8].ReadOnly = true;
                table.Columns[9].ReadOnly = true;
                table.Columns[10].ReadOnly = true;
                table.Columns[11].ReadOnly = true;
                table.Columns[12].ReadOnly = true;
                table.Columns[13].ReadOnly = true;
                table.Columns[14].ReadOnly = true;

                checkComboboxSetup();

                platform.Enabled = true;
                SN.Enabled = true;
                phase.Enabled = true;
                SKU.Enabled = true;
                borrower.Enabled = true;
                status.Enabled = true;
                year.Enabled = true;
                category.Enabled = true;
                note.Enabled = true;
                borrowDate.Enabled = true;
                position.Enabled = true;
                CPU.Enabled = true;
                WLAN.Enabled = true;

                sendMail.Visible = true;
                clear.Visible = true;
                edit.Visible = true;
                back.Visible = false;
                submit.Visible = false;
                output.Visible = true;
                selectAllItems.Visible = true;
                table.Columns[0].Visible = true;
            }

            table.Columns[1].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[2].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[3].SortMode = DataGridViewColumnSortMode.Programmatic;
            table.Columns[4].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[5].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[6].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[7].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[8].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[9].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[10].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[11].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[12].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[13].SortMode = DataGridViewColumnSortMode.Automatic;
            table.Columns[14].SortMode = DataGridViewColumnSortMode.Automatic;
        }

        private void selectAllItems_CheckedChanged(object sender, EventArgs e)
        {
            if (selectAllItems.Checked == true)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    table.Rows[i].Cells[0].Value = true;
                }
            }
            else
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    table.Rows[i].Cells[0].Value = false;
                }
            }
        }

        private void back_Click(object sender, EventArgs e)
        {
            a.Clear();
            setUpDisplay();
            checkComboboxSetup();

            platform.Enabled = true;
            SN.Enabled = true;
            phase.Enabled = true;
            SKU.Enabled = true;
            borrower.Enabled = true;
            status.Enabled = true;
            year.Enabled = true;
            category.Enabled = true;
            note.Enabled = true;
            borrowDate.Enabled = true;
            position.Enabled = true;
            CPU.Enabled = true;
            WLAN.Enabled = true;

            sendMail.Visible = true;
            clear.Visible = true;
            edit.Visible = true;
            back.Visible = false;
            submit.Visible = false;
            selectAllItems.Visible = true;
            output.Visible = true;
            table.Columns[0].Visible = true;
        }

        private void SN_DropDownClosed(object sender, EventArgs e)
        {
            if (SN.CheckedItems.Count != 0) {
                increaseTableItem();
            }
        }
        private void platform_DropDownClosed(object sender, EventArgs e)
        {
            if (platform.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void phase_DropDownClosed(object sender, EventArgs e)
        {
            if (phase.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void SKU_DropDownClosed(object sender, EventArgs e)
        {
            if (SKU.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void borrower_DropDownClosed(object sender, EventArgs e)
        {
            if (borrower.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void status_DropDownClosed(object sender, EventArgs e)
        {
            if (status.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void year_DropDownClosed(object sender, EventArgs e)
        {
            if (year.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void category_DropDownClosed(object sender, EventArgs e)
        {
            if (category.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void note_DropDownClosed(object sender, EventArgs e)
        {
            if (note.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void borrowDate_DropDownClosed(object sender, EventArgs e)
        {
            if (borrowDate.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }
        private void position_DropDownClosed(object sender, EventArgs e)
        {
            if (position.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }

        private void CPU_DropDownClosed(object sender, EventArgs e)
        {
            if (CPU.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }

        private void WLAN_DropDownClosed(object sender, EventArgs e)
        {
            if (WLAN.CheckedItems.Count != 0)
            {
                increaseTableItem();
            }
        }

        private void table_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 12 && e.RowIndex>=0 && e.RowIndex<table.Rows.Count-1)
            {
                dr.Close();
                cmd.CommandText = "Select * From UnitTable where SN='" + table.Rows[e.RowIndex].Cells[4].Value.ToString() + "'";
                dr = cmd.ExecuteReader();
                dr.Read();
                e.Paint(e.CellBounds, DataGridViewPaintParts.All);
                if (dr[14].ToString() != "")
                {
                   e.Graphics.DrawImage(Units_display.Properties.Resources._02, e.CellBounds.Left, e.CellBounds.Top, 25, 15);
                }
                else if (dr[13].ToString() != "")
                {
                    e.Graphics.DrawImage(Units_display.Properties.Resources._01, e.CellBounds.Left, e.CellBounds.Top, 25, 15);
                }
                else
                {
                    e.Graphics.DrawImage(Units_display.Properties.Resources._00, e.CellBounds.Left, e.CellBounds.Top,25, 15);
                }
                e.Handled = true;
                dr.Close();
            }
        }

        private void table_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dr.Close();
            if (e.ColumnIndex == 12)
            {
                Form2 fm2 = new Form2();

                cmd.CommandText = "Select * From UnitTable where SN='" + table.Rows[e.RowIndex].Cells[4].Value.ToString() + "'";
                dr = cmd.ExecuteReader();
                dr.Read();

                fm2.b1_Text = dr[13].ToString();
                fm2.d1_Text = dr[10].ToString();
                fm2.b2_Text = dr[14].ToString();
                fm2.d2_Text = dr[16].ToString();
                fm2.b3_Text = dr[15].ToString();
                fm2.d3_Text = dr[17].ToString();
                fm2.n1_Text = dr[20].ToString();
                fm2.n2_Text = dr[21].ToString();
                fm2.n3_Text = dr[22].ToString();

                dr.Close();
                fm2.Show();
            }
        }

        private void table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void copyAlltoClipboard()
        {
            table.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DataObject dataObj = table.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        public void CopyToClipboardWithHeaders()
        {
            //Copy to clipboard
            table.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            table.SelectAll();
            DataObject dataObj = table.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void output_Click(object sender, EventArgs e)
        {
            table.Columns[0].Visible = false;
            table.RowHeadersVisible = false;
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            try
            {
                CopyToClipboardWithHeaders();
                xlexcel.Visible = false;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    xlWorkBook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                xlexcel.Quit();
                xlWorkBook = null;
                xlexcel = null;
            }
            table.Columns[0].Visible = true;
            table.RowHeadersVisible = true;
            table.ClearSelection();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            if (!first_activated)
            {
                first_activated = true;
                ma.Logon(IntPtr.Zero);
            }
        }

        private void table_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void table_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index==3)
            {
                e.SortResult = int.Parse(e.CellValue1.ToString()).CompareTo(int.Parse(e.CellValue2.ToString()));
                e.Handled = true;//pass by the default sorting

        //    //    int a = int.Parse(e.CellValue1.ToString()), b = int.Parse(e.CellValue2.ToString());

        //    //    // If the cell value is already an integer, just cast it instead of parsing

        //    //    e.SortResult = a.CompareTo(b);

        //    //    e.Handled = true;
            }
        }

        private void table_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //DataGridView dgv = sender as DataGridView;
            //if (dgv.Columns[e.ColumnIndex].SortMode == DataGridViewColumnSortMode.Programmatic)
            //{
            //     string columnBindingName = dgv.Columns[e.ColumnIndex].DataPropertyName;
            //     switch (dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection)
            //     {
            //        case System.Windows.Forms.SortOrder.None:
            //        case System.Windows.Forms.SortOrder.Ascending:
            //            CustomSort(columnBindingName, "desc");
            //            dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Descending;
            //            break;
            //       case System.Windows.Forms.SortOrder.Descending:
            //        CustomSort(columnBindingName, "asc");
            //        dgv.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = System.Windows.Forms.SortOrder.Ascending;
            //              break;
            //        }
            //}
        }

        private void CustomSort(string columnBindingName, string sortMode)
        {
            //DataTable dt = this.table.DataSource as DataTable;
            //DataView dv = dt.DefaultView;
            //dv.Sort = columnBindingName + " " + sortMode;
            //this.table.DataSource = dv.ToTable();
            //this.table.Refresh();
        }

        public unitTB[] return_UnitTB()
        {
            return unitTBs;
        }

        public unitInfo[] return_unitInfo()
        {
            return unitInfos;
        }
        private void SN_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //private void customSortCompare(object sender, DataGridViewSortCompareEventArgs e)
        //{
        //    if (e.Column == table.Columns[3])
        //    {
        //        int a = int.Parse(e.CellValue1.ToString()), b = int.Parse(e.CellValue2.ToString());

        //        // If the cell value is already an integer, just cast it instead of parsing

        //        e.SortResult = a.CompareTo(b);

        //        e.Handled = true;
        //    }

        //}
    }
}
