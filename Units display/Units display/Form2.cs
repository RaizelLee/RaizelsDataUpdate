using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Units_display
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public string b1_Text
        {
            set
            {
                this.label_b1.Text = value;
            }
        }
        public string b2_Text
        {
            set
            {
                this.label_b2.Text = value;
            }
        }
        public string b3_Text
        {
            set
            {
                this.label_b3.Text = value;
            }
        }
        public string d1_Text
        {
            set
            {
                this.label_d1.Text = value;
            }
        }
        public string d2_Text
        {
            set
            {
                this.label_d2.Text = value;
            }
        }
        public string d3_Text
        {
            set
            {
                this.label_d3.Text = value;
            }
        }
        public string n1_Text
        {
            set
            {
                this.label_n1.Text = value;
            }
        }
        public string n2_Text
        {
            set
            {
                this.label_n2.Text = value;
            }
        }
        public string n3_Text
        {
            set
            {
                this.label_n3.Text = value;
            }
        }

        private void OK_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
