namespace Units_import_system
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.record = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.platformName2 = new System.Windows.Forms.Label();
            this.phaseName = new System.Windows.Forms.Label();
            this.yearName = new System.Windows.Forms.Label();
            this.typeName = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 69);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(291, 540);
            this.dataGridView1.TabIndex = 0;
            // 
            // record
            // 
            this.record.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.record.Location = new System.Drawing.Point(118, 626);
            this.record.Name = "record";
            this.record.Size = new System.Drawing.Size(75, 23);
            this.record.TabIndex = 1;
            this.record.Text = "Record";
            this.record.UseVisualStyleBackColor = true;
            this.record.Click += new System.EventHandler(this.record_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Platform: ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Phase: ";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(197, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Year: ";
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(197, 42);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Type: ";
            // 
            // platformName2
            // 
            this.platformName2.AutoSize = true;
            this.platformName2.Location = new System.Drawing.Point(90, 19);
            this.platformName2.Name = "platformName2";
            this.platformName2.Size = new System.Drawing.Size(35, 13);
            this.platformName2.TabIndex = 6;
            this.platformName2.Text = "label5";
            // 
            // phaseName
            // 
            this.phaseName.AutoSize = true;
            this.phaseName.Location = new System.Drawing.Point(83, 42);
            this.phaseName.Name = "phaseName";
            this.phaseName.Size = new System.Drawing.Size(35, 13);
            this.phaseName.TabIndex = 7;
            this.phaseName.Text = "label6";
            // 
            // yearName
            // 
            this.yearName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.yearName.AutoSize = true;
            this.yearName.Location = new System.Drawing.Point(239, 19);
            this.yearName.Name = "yearName";
            this.yearName.Size = new System.Drawing.Size(35, 13);
            this.yearName.TabIndex = 8;
            this.yearName.Text = "label7";
            // 
            // typeName
            // 
            this.typeName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.typeName.AutoSize = true;
            this.typeName.Location = new System.Drawing.Point(239, 42);
            this.typeName.Name = "typeName";
            this.typeName.Size = new System.Drawing.Size(35, 13);
            this.typeName.TabIndex = 9;
            this.typeName.Text = "label8";
            // 
            // Form2
            // 
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(315, 661);
            this.Controls.Add(this.typeName);
            this.Controls.Add(this.yearName);
            this.Controls.Add(this.phaseName);
            this.Controls.Add(this.platformName2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.record);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form2";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form2_FormClosed);
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.DataGridView dataGridView1;
        public System.Windows.Forms.Button record;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.Label platformName2;
        public System.Windows.Forms.Label phaseName;
        public System.Windows.Forms.Label yearName;
        public System.Windows.Forms.Label typeName;
    }
}