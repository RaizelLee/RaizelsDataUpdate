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

namespace Units_import_system
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public Form1 mainForm;
        public Form2(Form callingForm)
        {
            mainForm = callingForm as Form1;
            InitializeComponent();
        }
        private void record_Click(object sender, EventArgs e)
        {
            String strConn = "server=tcp:WIN-3AKK5VRL0P3\\SQLEXPRESS,49172;database=test;User ID=test;Password=Oct0b1vju1v38;";//登入資料庫帳密跟port和Database
            SqlConnection myConn = new SqlConnection(strConn);//建立連接     
            SqlCommand cmd = new SqlCommand();//下sql指令用
            //SqlDataReader dr;
            cmd.Connection = myConn;//讓cmd連資料庫
            myConn.Open();

            //Form1 fm1 = new Form1();
            string cmd1, cmd2, cmd3;
            cmd1 = this.mainForm.cmd1();
            cmd2 = this.mainForm.cmd2();
            cmd3 = this.mainForm.cmd3();

            if (cmd1 != "")
            {
                try
                { 
                    cmd.CommandText = cmd1;
                    cmd.ExecuteNonQuery();
                }
                catch(SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            if (cmd2 != "")
            {
                try
                { 
                    cmd.CommandText = cmd2;
                    cmd.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }
            if (cmd3 != "")
            {
                try
                { 
                    cmd.CommandText = cmd3;
                    cmd.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            this.mainForm.clearAll();
            this.Hide();

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            //e.Cancel = true;
            this.Hide();
        }
    }
}
