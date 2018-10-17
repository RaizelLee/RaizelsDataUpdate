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
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace real_import_IUR_to_sql_system
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;

        string strConn = "server=tcp:WIN-3AKK5VRL0P3\\SQLEXPRESS,49172;database=test;User ID=test;Password=Oct0b1vju1v38;";//登入資料庫帳密跟port和Database    
        SqlCommand cmd = new SqlCommand();
        //SqlDataReader dr;
        //int[] errorNum = { 215, 216, 217, 404, 405, 490, 491, 492, 493, 494, 495, 496, 497, 498, 499, 500,
        //    501, 733, 1533, 1534, 1535, 1536, 1537, 1538, 1539, 1540, 1541, 1542, 1543, 1544, 1545, 1546, 1547,
        //    1548, 1549, 1550, 1551, 1552, 1553, 1554, 1555, 1556, 1557, 1558, 1559, 1560, 1561, 1562, 1563, 1564,
        //    1565, 1566, 1567, 1568, 1569, 1570, 1571, 1572, 1573, 1574, 1575, 1576, 1577, 1578, 1579, 1580, 1581,
        //    1582, 1583, 1584, 1585, 1586, 1587, 1588, 1589, 1590, 1591, 1592, 1593, 1594, 1595, 1596, 1597, 1598,
        //    1599, 1600, 1601, 1602, 1603, 1604, 1605, 1606, 1607, 1608, 1609, 1610, 1611, 1612, 1613, 1614, 1615,
        //    1616, 1617, 1618, 1619, 1620, 1621, 1622, 1623, 1624, 1625, 1626, 1627, 1636, 1876, 1964, 1965, 1966,
        //    1967, 1968, 1969, 1970, 1971, 1972, 1973, 1974, 1975, 1976, 2038, 2039, 2040, 2041, 2210, 2212, 2448,
        //    2457, 2489, 2867, 2868, 2869, 2952, 3279, 3289, 3524, 3525, 3577, 3605, 3606, 3607, 3619, 3620, 3625,
        //    3634, 3635, 3645, 3646, 3690, 3751, 3752, 3754 };

        public Form1()
        {
            InitializeComponent();
            
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
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\HP\Desktop\all IUR list Raziel.xlsx");
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;

            
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string sn, pf, ph, sku, borrower, status, year, type, note, borrowDate, position, keyintime, mail1, note1;
            //string name, mail;
            for (int i = 2; i <= rowCount; i++)
            {
                label1.Text = i.ToString();

                
                if(xlRange.Cells[i, 4].Value2!= null)
                {
                    sn = xlRange.Cells[i, 4].Value2.ToString();
                }
                else
                {
                    sn = null;
                }
                if (xlRange.Cells[i, 2].Value2!= null)
                {
                    ph = xlRange.Cells[i, 2].Value2.ToString();
                }
                else
                {
                    ph = null;
                }

                if (xlRange.Cells[i, 1].Value2 != null)
                {
                    pf = xlRange.Cells[i, 1].Value2.ToString();
                }
                else
                {
                    pf = null;
                }
                if (xlRange.Cells[i, 3].Value2 != null)
                {
                    sku = xlRange.Cells[i, 3].Value2.ToString();
                }
                else
                {
                    sku = null;
                }
                if (xlRange.Cells[i, 5].Value2 != null)
                {
                    borrower = xlRange.Cells[i, 5].Value2.ToString();
                }
                else
                {
                    borrower = null;
                }
                if (xlRange.Cells[i, 6].Value2 != null)
                {
                    status = xlRange.Cells[i, 6].Value2.ToString();
                }
                else
                {
                    status = null;
                }
                if (xlRange.Cells[i, 7].Value2 != null)
                {
                    year = xlRange.Cells[i, 7].Value2.ToString();
                }
                else
                {
                    year = null;
                }
                if (xlRange.Cells[i, 8].Value2 != null)
                {
                    type = xlRange.Cells[i, 8].Value2.ToString();
                }
                else
                {
                    type = null;
                }
                if (xlRange.Cells[i, 9].Value2 != null)
                {
                    note = xlRange.Cells[i, 9].Value2.ToString();
                }
                else
                {
                    note = null;
                }
                if (xlRange.Cells[i, 10].Value2 != null)
                {
                    borrowDate = xlRange.Cells[i, 10].Text;
                }
                else
                {
                    borrowDate = null;
                }
                if (xlRange.Cells[i, 11].Value2!= null)
                {
                    position = xlRange.Cells[i, 11].Value2.ToString();
                }
                else
                {
                    position = null;
                }

                if((borrower!="N/A")&& (borrower != "Broken")&& (borrower != "storage")&& 
                    (borrower != "scrap")&& (borrower != "Scrap")&&(borrower != "Storage")&&
                    (borrower != "broken")&&(borrower != "4444444"))
                {
                    mail1 = borrower;
                }
                else
                {
                    mail1 = "";
                }

                if (note != "")
                {
                    note1 = note;
                }
                else
                {
                    note1 = "";
                }
                keyintime = DateTime.Now.ToString("yyyy/M/d");

                //name = xlRange.Cells[i, 1].Value2.ToString();
                //mail = xlRange.Cells[i, 2].Value2.ToString();
                string cmdText = "";
                cmdText += "INSERT INTO UnitTable VALUES('" + sn + "','"//1,SN //4
                            + pf + "','"//2,pf //1
                            + ph + "','"//3,ph //2
                            + sku + "'," //4,sku //3
                            + "NULL,'" //5,cat
                            + borrower + "','"//6,borrower //5
                            + status + "','"//7,status //6
                            + year + "','"//8,year+cycle //7
                            + type + "','"//9,catagegry(type) //8
                            + note + "','"//10,note //9
                            + borrowDate + "','"//11,borrowdate //10
                            + position + "','"//12,positon //11
                            + keyintime + "','"//13,keyintime
                            + mail1 + "',"//14,mail-1
                            + "NULL,"//15,mail-2
                            + "NULL,"//16,mail-3
                            + "NULL,"//17,borrowdate2
                            + "NULL,"//18,borrowdate3
                            + "NULL,"//19,CPU
                            + "NULL,'"//20,WLAN
                            + note1 + "'," //noteOne
                            + "NULL,"//noteTwo
                            + "NULL);";//noteThree
                //cmdText += "INSERT INTO IUR_borrower_name_list VALUES('" + name + "','" + mail + "');";
                cmd.CommandText = cmdText;
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (SqlException sqlEX)
                {
                    string lines = i.ToString() + " : " + sqlEX.Message.ToString() + "\r\n";
                    System.IO.File.AppendAllText(@"C:\Users\HP\Desktop\IUR_IMPORT_SQL_ERROR.txt", lines);
                }
            }

            //label1.Text = i.ToString() + " , " + j.ToString();
            //if ((xlRange.Cells[i, 4].Value2.ToString() == xlRange.Cells[j, 4].Value2.ToString()) && (i != j))
            //{
            //    string[] lines = { "" };
            //    lines += xlRange.Cells[i, 4].Value2.ToString() + "\r\n";
            //    System.IO.File.WriteAllLines(@"C:\Users\HP\Desktop\duplicate.txt", lines);
            //    //MessageBox.Show(xlRange.Cells[i, 4].Value2.ToString());
            xlWorkbook.Close();
            MessageBox.Show("Done!");
        }
    }
}
