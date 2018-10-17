using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Win32Mapi;

namespace Units_display
{
    public partial class Form3 : Form
    {
        private Mapi ma = new Mapi();
        //private bool first_activated = false;
        //private Font boldFont;
        //MailEnvelop currentMail;
        public int chooseNum = 1;
        

        public Form3()
        {
            InitializeComponent();
            chooseNum = 1;
            //ma.Logon(this.Handle);
        }

        public Form1 mainForm;
        public Form1.unitTB[] unitTBs;
        public Form1.unitInfo[] unitInfos;
        public Form3(Form callingForm)
        {
            mainForm = callingForm as Form1;
            InitializeComponent();
            chooseNum = 1;
            unitTBs = mainForm.return_UnitTB();
            unitInfos = mainForm.return_unitInfo();
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        

        private void confirm_Click(object sender, EventArgs e)
        {
            this.Hide();
            SendForm sfm = new SendForm(this,ref ma);
            sfm.ShowDialog(this);
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void arrivedMail_Click(object sender, EventArgs e)
        {
            chooseNum = 1;
        }

        private void borrowMail_Click(object sender, EventArgs e)
        {
            chooseNum = 2;
        }
        public int returnChoose()
        {
            return chooseNum;
        }

        public Form1.unitTB[] returnUnitTB()
        {
            return unitTBs;
        }

        public Form1.unitInfo[] returnUnitInfo()
        {
            return unitInfos;
        }
    }
}
