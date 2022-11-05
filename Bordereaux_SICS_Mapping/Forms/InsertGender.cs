using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Bordereaux_SICS_Mapping.Forms
{
    public partial class frmGender : Form
    {
        public frmGender()
        {
            InitializeComponent();
        }

        private void btnInsertGender_Click(object sender, EventArgs e)
        {
            BAL.HelperV21 objHlpr = new BAL.HelperV21();


            string strFirstName = txtFirstName.Text.ToUpper();
            string strGender;
            string strMultipleNames = txtGender.Text;
            string strUserID = Environment.UserName;

            objHlpr.fn_getuserid(strUserID, out string strUser);
            
            if (string.IsNullOrEmpty(strFirstName))
            {
                if (!string.IsNullOrEmpty(strMultipleNames))
                {
                   
                    var dialog = MessageBox.Show("Would you like to import all the names from this file?", "Import Successful!", MessageBoxButtons.YesNo);
                        if (dialog == DialogResult.Yes)
                        {
                            objHlpr.fn_importmultiplenamesdb(strMultipleNames, strUser);
                            var result = MessageBox.Show("Upload Completed!", "Completed", MessageBoxButtons.OK);
                        }
                        else if (dialog == DialogResult.No)
                        {
                           //do nothing
                        }

                }
                else if (strFirstName == "")
                {
                    var result = MessageBox.Show("Add name to First name field", "Import Denied", MessageBoxButtons.OK);
                    
                }
            }
            else if (rbFemale.Checked == false && rbMale.Checked == false)
            {
                var result = MessageBox.Show("Select gender for " + strFirstName, "Import Denied", MessageBoxButtons.OK);
            }
            else
            {
                if (rbMale.Checked == true)
                {
                    rbFemale.Checked = false;
                    strGender = "M";
                    objHlpr.fn_importnamegenderdb(strFirstName, strGender, strUser);

                    var result = MessageBox.Show("Upload Completed!", "Completed", MessageBoxButtons.OK);

                }
                else if (rbFemale.Checked == true)
                {
                    rbMale.Checked = false;
                    strGender = "F";
                    objHlpr.fn_importnamegenderdb(strFirstName, strGender, strUser);


                    var result = MessageBox.Show("Upload Completed!", "Completed", MessageBoxButtons.OK);


                }


            }
        }

        private void rbMale_CheckedChanged(object sender, EventArgs e)
        {
            if (rbMale.Checked)
            {
                rbFemale.Checked = false;
            }
        }

        private void rbFemale_CheckedChanged(object sender, EventArgs e)
        {
            if (rbFemale.Checked)
            {
                rbMale.Checked = false;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                txtGender.Enabled = true;
                txtFirstName.Enabled = false;
                rbMale.Enabled = false;
                rbFemale.Enabled = false;
                rbMale.Checked = false;
                rbFemale.Checked = false;

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "All Files|*.*";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtGender.Text = ofd.FileName;
                }
                ofd.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //fn_reset();
            }
        }

        private void frmGender_Load(object sender, EventArgs e)
        {
            BAL.HelperV21 objHlpr = new BAL.HelperV21();
            txtGender.Enabled = false;
            txtFirstName.Enabled = true;
            rbFemale.Enabled = true;
            rbMale.Enabled = true;
            string strUserid = Environment.UserName;
            objHlpr.fn_getuserid(strUserid, out string strUserName);
            lblName.Text = strUserName + "!";
        }





        private void button1_Click(object sender, EventArgs e)
        {
            frmGender.ActiveForm.Close();

        }

        private void btnCheckName_Click(object sender, EventArgs e)
        {
             BAL.HelperV21 objHlpr = new BAL.HelperV21();
             string strFirstname = txtName.Text;
             strFirstname = strFirstname.Trim().ToUpper().Replace(" ", string.Empty);

            objHlpr.fn_searchnamesdb(strFirstname,out string strGender,out string strAuthor);

            if (strGender == "M")
            {
                lblGender.ForeColor = System.Drawing.Color.MediumBlue;
                lblGender.Text = "MALE";
                lblAuthor.Text = strAuthor;
               

            }
            else if (strGender == "F")
            {
                lblGender.ForeColor = System.Drawing.Color.HotPink;
                lblGender.Text = "FEMALE";
                lblAuthor.Text = strAuthor;

            }
            else if (string.IsNullOrEmpty(strFirstname))
            {
                var result = MessageBox.Show("Input a name to search", "Search Name", MessageBoxButtons.OK);
                lblGender.Text = "";
                lblAuthor.Text = "";


            }
            else
            {
                var result = MessageBox.Show("No record was found in the Gender Database", "Search Name", MessageBoxButtons.OK);
                lblGender.Text = "";
                lblAuthor.Text = "";


            }
        }
    }
}
