namespace Units_display
{
    partial class Form3
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
            this.arrivedMail = new System.Windows.Forms.Button();
            this.borrowMail = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.confirm = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // arrivedMail
            // 
            this.arrivedMail.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.arrivedMail.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.arrivedMail.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.arrivedMail.Location = new System.Drawing.Point(57, 46);
            this.arrivedMail.Name = "arrivedMail";
            this.arrivedMail.Size = new System.Drawing.Size(385, 77);
            this.arrivedMail.TabIndex = 0;
            this.arrivedMail.Text = "Sent \"Machines Arrived Mail\"";
            this.arrivedMail.UseVisualStyleBackColor = false;
            this.arrivedMail.Click += new System.EventHandler(this.arrivedMail_Click);
            // 
            // borrowMail
            // 
            this.borrowMail.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.borrowMail.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.borrowMail.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.borrowMail.Location = new System.Drawing.Point(57, 129);
            this.borrowMail.Name = "borrowMail";
            this.borrowMail.Size = new System.Drawing.Size(385, 77);
            this.borrowMail.TabIndex = 1;
            this.borrowMail.Text = "Sent \"Borrow Record Mail\"";
            this.borrowMail.UseVisualStyleBackColor = false;
            this.borrowMail.Click += new System.EventHandler(this.borrowMail_Click);
            // 
            // cancel
            // 
            this.cancel.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.cancel.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.cancel.Location = new System.Drawing.Point(111, 236);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(84, 35);
            this.cancel.TabIndex = 2;
            this.cancel.Text = "Cancel";
            this.cancel.UseVisualStyleBackColor = false;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // confirm
            // 
            this.confirm.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.confirm.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.confirm.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.confirm.Location = new System.Drawing.Point(298, 236);
            this.confirm.Name = "confirm";
            this.confirm.Size = new System.Drawing.Size(84, 35);
            this.confirm.TabIndex = 3;
            this.confirm.Text = "Confirm";
            this.confirm.UseVisualStyleBackColor = false;
            this.confirm.Click += new System.EventHandler(this.confirm_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(501, 298);
            this.Controls.Add(this.confirm);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.borrowMail);
            this.Controls.Add(this.arrivedMail);
            this.Name = "Form3";
            this.Text = "Form3";
            this.Load += new System.EventHandler(this.Form3_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button arrivedMail;
        private System.Windows.Forms.Button borrowMail;
        private System.Windows.Forms.Button cancel;
        private System.Windows.Forms.Button confirm;
    }
}