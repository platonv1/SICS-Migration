namespace Bordereaux_SICS_Mapping.Forms
{
    partial class frmPolicyYear
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
        public void InitializeComponent()
        {
            this.txt_inputPolicyYear = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txt_inputPolicyYear
            // 
            this.txt_inputPolicyYear.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_inputPolicyYear.Location = new System.Drawing.Point(104, 50);
            this.txt_inputPolicyYear.Multiline = true;
            this.txt_inputPolicyYear.Name = "txt_inputPolicyYear";
            this.txt_inputPolicyYear.Size = new System.Drawing.Size(195, 33);
            this.txt_inputPolicyYear.TabIndex = 0;
            this.txt_inputPolicyYear.TextChanged += new System.EventHandler(this.txt_inputPolicyYear_TextChanged);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(158, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(79, 37);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmPolicyYear
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 162);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txt_inputPolicyYear);
            this.Name = "frmPolicyYear";
            this.Text = "Input PolicyYear";
            this.Load += new System.EventHandler(this.frmPolicyYear_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox txt_inputPolicyYear;
        private System.Windows.Forms.Button button1;
    }
}