using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;
using Bordereaux_SICS_Mapping.BAL;

namespace Bordereaux_SICS_Mapping.Forms
{
    public partial class frmPolicyYear : Form
    {
        public frmPolicyYear()
        {
            InitializeComponent();
        }

        private void txt_inputPolicyYear_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void frmPolicyYear_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txt_inputPolicyYear.Text == "")
            {
                MessageBox.Show("Policy Year cannot be blank","Enter  a Policy year");
            }
            else
            {
                Variables.strBmYear = txt_inputPolicyYear.Text;
            
                this.Close();
            }
        }
    }
}
