using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.BAL;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping
{
    public partial class B2SM : Form
    {
        private void B2SM_Load(object sender, EventArgs e)
        {
            _Global _var = new _Global();
            
            lbl_5.Text = string.Format(lbl_5.Text, _var.str_ver);

#if DEBUG
            //txt_dir.Text = "D:\\SICS\\Output";
            txt_dir.Text = "G:\\Shared drives\\Information Technology\\Employee Folders\\Platon.vm\\Application Development\\SICS Migration\\Output File";
            //txt_file.Text = DateTime.Now.ToString("dd-mm-yy_h-mm-ss");
            //txt_gender.Text = "C:\\Users\\basobas.ps\\Desktop\\SICS Migration Testing files\\Gender Database.xlsx";
#endif
        }
        private void B2SM_FormClosing(Object sender, FormClosingEventArgs e)
        {
            fn_killexcel();
        }
        private void B2SM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                fn_extract();
            }
        }

        public B2SM()
        {
            InitializeComponent();
            cmb_bm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmb_sheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }
        private void fn_reset() {
            txt_raw.Text = string.Empty;
            txt_dir.Text = string.Empty;
            txt_file.Text = string.Empty;


            cmb_sheet.Enabled = false;
            cmb_sheet.DataSource = null;

            cmb_bm.SelectedIndex = -1;

            chk_gender.Checked = false;
            chk_macro.Checked = false;
            fn_edgender();
            fn_edmacro();
            
            lbl_10.Text = "Ready...       ";
            chk_clean.Checked = false;

            Refresh();
        }
        private void btn_reset_Click(object sender, EventArgs e)
        {
            fn_reset();
        }
        private void btn_browse1_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbraw;
                Microsoft.Office.Interop.Excel.Worksheet wsraw;

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "All Files|*.*";
                List<string> wrkshtname = new List<string>();
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txt_raw.Text = ofd.FileName;
                    wbraw = eapp.Workbooks.Open(txt_raw.Text);
                    for (int i = 1; i <= wbraw.Sheets.Count; i++)
                    {
                        wsraw = wbraw.Worksheets[i];
                        wrkshtname.Add(wsraw.Name);
                    }
                    cmb_sheet.DataSource = wrkshtname;
                    
                    wsraw = null;
                    wbraw.Close();
                    wbraw = null;
                    
                    eapp = null;

                    cmb_sheet.Enabled = true;

                    foreach (string i in cmb_bm.Items)
                    {
                        string ii = "BM" + i;
                        Console.WriteLine(i);
                        if (txt_raw.Text.Contains(ii) && i != "000") 
                        {
                            cmb_bm.SelectedItem = i;
                            break;
                        }
                    }
                }
                ofd.Dispose();
                //fn_killexcel();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fn_reset();
            }
        }
        
        private void btn_browse3_Click(object sender, EventArgs e)
        {
            try
            {
            
                FolderBrowserDialog fbd = new FolderBrowserDialog();
            
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txt_dir.Text = fbd.SelectedPath;
                }
                fbd.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fn_reset();
            }
        }
        private void btn_browse4_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "All Files|*.*";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txt_gender.Text = ofd.FileName;
                    ofd.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fn_reset();
            }
        }
        private void btn_browse5_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "All Files|*.*";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txt_macro.Text = ofd.FileName;
                }
                ofd.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fn_reset();
            }
        }
        private void fn_edmacro() {
            if (chk_macro.Checked == false)
            {
                btn_browse5.Enabled = false;
                txt_macro.Enabled = false;
                txt_macro.Text = string.Empty;
            }
            else
            {
                btn_browse5.Enabled = true;
                txt_macro.Enabled = true;
            }
        }
      
        private void chk_macro_CheckedChanged(object sender, EventArgs e)
        {
            fn_edmacro();
        }
        private void fn_edgender()
        {
            if (chk_gender.Checked == false)
            {
                btn_browse4.Enabled = false;
                txt_gender.Enabled = false;
                txt_gender.Text = string.Empty;
            }
            else
            {
                btn_browse4.Enabled = true;
                txt_gender.Enabled = true;
            }
        }
        private void chk_gender_CheckedChanged(object sender, EventArgs e)
        {
            fn_edgender();
        }
       
      
        private string fn_validation()
        {
            string msg = string.Empty;

            if (string.IsNullOrEmpty(txt_dir.Text))
            {
                msg += "- Save directory is empty" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(txt_file.Text))
            {
                msg += "- Save filename is empty" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(txt_raw.Text))
            {
                msg += "- Raw file for processing is empty" + Environment.NewLine;
            }

            //if (string.IsNullOrEmpty(txt_sicstemp.Text))
            //{
            //    msg += "- SICS template is empty" + Environment.NewLine;
            //}

            if (cmb_bm.SelectedIndex == -1)
            {
                msg += "- Select a BM program to process" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(txt_gender.Text) && chk_gender.Checked)
            {
                msg += "- Gender database is empty" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(txt_macro.Text) && chk_macro.Checked)
            {
                msg += "- Macro database is empty" + Environment.NewLine;
            }
            return msg;
        }
        private void btn_extract_Click(object sender, EventArgs e)
        {
            fn_extract();
            
        }
        private void fn_extract()
        {
#if DEBUG
            //txt_file.Text = DateTime.Now.ToString("dd-mm-yy_h-mm-ss");
#endif
           
            string msg = fn_validation();
            if (msg != string.Empty)
            {
                MessageBox.Show("Complete details below before processing:" + Environment.NewLine + msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                DialogResult res = MessageBox.Show("Start the Extraction?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    lbl_10.Text = "Processing...";
                    this.Enabled = false;

                    
                    Variables.strBmYear = txt_policyYear.Text;
                    string str_error = string.Empty;
                    var watch = System.Diagnostics.Stopwatch.StartNew();

                    switch (cmb_bm.SelectedItem.ToString())
                    {
                        case "001":
                            BM001 objBM001 = new BM001();
                            str_error = objBM001.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM001 = null;
                            break;

                        case "001A":
                            BM001A objBM001A = new BM001A();
                            str_error = objBM001A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM001 = null;
                            break;

                        case "003":
                            BM003 objBM003 = new BM003();
                            str_error = objBM003.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM003 = null;
                            break;

                        case "004":
                            BM004 objBM004 = new BM004();
                            str_error = objBM004.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM004 = null;
                            break;

                        case "005":
                            BM005 objBM005 = new BM005();
                            str_error = objBM005.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM005 = null;
                            break;
                        case "007":
                            BM007 objBM007 = new BM007();
                            str_error = objBM007.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM007 = null;
                            break;


                        case "009":
                            BM009 objBM009 = new BM009();
                            str_error = objBM009.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM009 = null;
                            break;

                        case "010":
                            BM010 objBM010 = new BM010();
                            str_error = objBM010.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM010 = null;
                            break;

                        case "011":
                        BM011 objBM011 = new BM011();
                        str_error = objBM011.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM011 = null;
                        break;

                        case "013":
                            BM013 objBM013 = new BM013();
                            str_error = objBM013.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM013 = null;
                            break;

                        case "013 - A":
                            BM013_A objBM013_A = new BM013_A();
                            str_error = objBM013_A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM013_A = null;
                            break;

                        case "013 - CGL22":
                        BM013_CGL22 objBM013_CGL22 = new BM013_CGL22();
                        str_error = objBM013_CGL22.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM013_CGL22 = null;
                        break;

                        case "014":
                            BM014 objBM014 = new BM014();
                            str_error = objBM014.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM014 = null;
                            break;

                        case "014 - A":
                            BM014_A objBM014_A = new BM014_A();
                            str_error = objBM014_A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM014_A = null;
                            break;

                        case "014 - CBL22":
                        BM014_CBL22 objBM014_CBL22 = new BM014_CBL22();
                        str_error = objBM014_CBL22.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM014_CBL22 = null;
                        break;

                        case "015":
                            BM015 objBM015 = new BM015();
                            str_error = objBM015.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM015 = null;
                            break;

                        case "015 - A":
                            BM015_A objBM015_A = new BM015_A();
                            str_error = objBM015_A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM015_A = null;
                            break;

                        case "015 - B":
                            BM015_B objBM015_B = new BM015_B();
                            str_error = objBM015_B.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM015_B = null;
                            break;

                        case "016":
                            BM016 objBM016 = new BM016();
                            str_error = objBM016.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM016 = null;
                            break;

                        case "016 - A":
                            BM016_A objBM016_A = new BM016_A();
                            str_error = objBM016_A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM016_A = null;
                            break;

                        case "016 - B":
                            BM016_B objBM016_B = new BM016_B();
                            str_error = objBM016_B.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM016_B = null;
                            break;

                        case "016 - C":
                            BM016_C objBM016_C = new BM016_C();
                            str_error = objBM016_C.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM016_C = null;
                            break;

                        case "016 - Facul22":
                            BM016_Facul22 objBM016_Facul22 = new BM016_Facul22();
                            str_error = objBM016_Facul22.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM016_Facul22 = null;
                            break;

                        case "016 - Ind22":
                        BM016_IND22 objBM016_IND22 = new BM016_IND22();
                        str_error = objBM016_IND22.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM016_Facul22 = null;
                        break;

                        case "019":
                            BM019 objBM019 = new BM019();
                            str_error = objBM019.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM019 = null;
                            break;

                        case "019 - A":
                            BM019_A objBM019_A = new BM019_A();
                            str_error = objBM019_A.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM019_A = null;
                            break;

                        case "020":
                            BM020 objBM020 = new BM020();
                            str_error = objBM020.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM020 = null;
                            break;

                        case "021":
                            BM021 objBM021 = new BM021();
                            str_error = objBM021.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM021 = null;
                            break;

                        case "021 - A":
                        BM021_A objBM021_A = new BM021_A();
                        str_error = objBM021_A.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM021_A = null;
                        break;

                        case "022":

                            BM022 objBM022 = new BM022();
                            str_error = objBM022.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                            objBM022 = null;
                            break;

                        case "023":
                            BM023 objBM023 = new BM023();
                            str_error = objBM023.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM023 = null;
                            break;

                        case "024":
                            BM024 objBM024 = new BM024();
                            str_error = objBM024.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM024 = null;
                            break;

                        case "025":
                            BM025 objBM025 = new BM025();
                            str_error = objBM025.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM025 = null;
                            break;


                        case "026":
                            BM026 objBM026 = new BM026();
                            str_error = objBM026.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM026 = null;
                            break;

                        case "030":
                            BM030 objBM030 = new BM030();
                            str_error = objBM030.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM030 = null;
                            break;

                        case "031":

                            BM031 objBM031 = new BM031();
                            str_error = objBM031.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                            objBM031 = null;
                            break;

                        case "032":

                            BM032 objBM032 = new BM032();
                            str_error = objBM032.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                            objBM032 = null;
                            break;

                        case "033":

                            BM033 objBM033 = new BM033();
                            str_error = objBM033.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM033 = null;
                            break;

                        //case "036":
                        //    BM036 objBM036 = new BM036();
                        //    str_error = objBM036.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM036 = null;
                        //    break;

                        //case "037":
                        //    BM037 objBM037 = new BM037();
                        //    str_error = objBM037.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM037 = null;
                        //    break;

                        //case "038":
                        //    BM038 objBM038 = new BM038();
                        //    str_error = objBM038.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM038 = null;
                        //    break;

                        //case "039":
                        //    BM039 objBM039 = new BM039();
                        //    str_error = objBM039.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM039 = null;
                        //    break;

                        //////case "031":
                        //////    BM031 objBM031 = new BM031();
                        //////    str_error = objBM031.fn_process(
                        //////        txt_raw.Text,
                        //////        txt_sicstemp.Text,
                        //////        cmb_sheet.SelectedItem.ToString(),
                        //////        txt_dir.Text,
                        //////        txt_file.Text, txt_gender.Text,
                        //////        chk_open.Checked);

                        //////    objBM031 = null;
                        //////    break;

                        case "041-NB":
                            BM041_NB objBM041_NB = new BM041_NB();
                            str_error = objBM041_NB.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM041_NB = null;
                            break;

                        case "041-PTF":
                            BM041_PTF objBM041_PTF = new BM041_PTF();
                            str_error = objBM041_PTF.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM041_PTF = null;
                            break;

                        case "041-new":
                            BM041_NB_new objBM041_NB_new = new BM041_NB_new();
                            str_error = objBM041_NB_new.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM041_NB = null;
                            break;

                        case "041-PRIOR":
                            BM041_PRIOR objBM041_PRIOR = new BM041_PRIOR();
                            str_error = objBM041_PRIOR.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM041_PRIOR = null;
                            break;

                        case "041_21":
                        BM041_21 objBM041_21 = new BM041_21();
                        str_error = objBM041_21.fn_process(
                            txt_raw.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked);

                        objBM041_21 = null;
                        break;

                        case "042":
                            BM042 objBM042 = new BM042();
                            str_error = objBM042.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM042 = null;
                            break;

                        case "043":
                            BM043 objBM043 = new BM043();
                            str_error = objBM043.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM043 = null;
                            break;

                        case "044":
                            BM044 objBM044 = new BM044();
                            str_error = objBM044.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM044 = null;
                            break;


                        case "048":
                            BM048 objBM048 = new BM048();
                            str_error = objBM048.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM048 = null;
                            break;

                        case "049":
                            BM049 objBM049 = new BM049();
                            str_error = objBM049.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM049 = null;
                            break;

                        //case "050":
                        //    BM050 objBM050 = new BM050();
                        //    str_error = objBM050.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM050 = null;
                        //    break;


                        case "051":
                            BM051 objBM051 = new BM051();
                            str_error = objBM051.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM051 = null;
                            break;

                        case "052":
                            BM052 objBM052 = new BM052();
                            str_error = objBM052.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM052 = null;
                            break;


                        case "053":
                            BM053 objBM053 = new BM053();
                            str_error = objBM053.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM053 = null;
                            break;

                        //case "058":
                        //    BM058 objBM058 = new BM058();
                        //    str_error = objBM058.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM058 = null;
                        //    break;

                        case "059":
                            BM059 objBM059 = new BM059();
                            str_error = objBM059.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM059 = null;
                            break;


                        case "060":
                            BM060 objBM060 = new BM060();
                            str_error = 
                            str_error = objBM060.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                            objBM060 = null;
                            break;


                        case "061":
                            BM061 objBM061 = new BM061();
                            str_error = objBM061.fn_process(
                                txt_raw.Text,
                                //txt_sicstemp.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                            objBM061 = null;
                            break;

                        case "061-RA":
                        BM061_RA objBM061_RA = new BM061_RA();
                        str_error = objBM061_RA.fn_process(
                            txt_raw.Text,
                            //txt_sicstemp.Text,
                            cmb_sheet.SelectedItem.ToString(),
                            txt_dir.Text,
                            txt_file.Text, txt_gender.Text,
                            chk_open.Checked, chk_clean.Checked, txt_macro.Text);

                        objBM061_RA = null;
                        break;

                        case "062":
                            BM062 objBM062 = new BM062();
                            str_error = objBM062.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM062 = null;
                            break;

                        case "063":
                            BM063 objBM063 = new BM063();
                            str_error = objBM063.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM063 = null;
                            break;


                        case "064":
                            BM064 objBM064 = new BM064();
                            str_error = objBM064.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM064 = null;
                            break;

                        case "065":
                            BM065 objBM065 = new BM065();
                            str_error = objBM065.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked);
                            objBM065 = null;
                            break;

                        case "066":
                            BM066 objBM066 = new BM066();
                            str_error = objBM066.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM066 = null;
                            break;

                        case "067":
                            BM067 objBM067 = new BM067();
                            str_error = objBM067.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM067 = null;
                            break;

                        case "068":
                            BM068 objBM068 = new BM068();
                            str_error = objBM068.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM068 = null;
                            break;

                        case "069":
                            BM069 objBM069 = new BM069();
                            str_error = objBM069.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM069 = null;
                            break;

                        case "070":
                            BM070 objBM070 = new BM070();
                            str_error = objBM070.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM070 = null;
                            break;

                        case "072":
                            BM072 objBM072 = new BM072();
                            str_error = objBM072.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM072 = null;
                            break;

                        case "073":
                            BM073 objBM073 = new BM073();
                            str_error = objBM073.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM073 = null;
                            break;

                        case "074":
                            BM074 objBM074 = new BM074();
                            str_error = objBM074.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM074 = null;
                            break;

                        //case "087":
                        //    BM087 objBM087 = new BM087();
                        //    str_error = objBM087.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM087 = null;
                        //    break;

                        //case "089":
                        //    BM089 objBM089 = new BM089();
                        //    str_error = objBM089.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM089 = null;
                        //    break;

                        case "090":
                            BM090 objBM090 = new BM090();
                            str_error = objBM090.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM090 = null;
                            break;

                        //case "094":
                        //    BM094 objBM094 = new BM094();
                        //    str_error = objBM094.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM094 = null;
                        //    break;

                        case "098":
                            BM098 objBM098 = new BM098();
                            str_error = objBM098.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM098 = null;
                            break;

                        case "097":
                            BM097 objBM097 = new BM097();
                            str_error = objBM097.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM097 = null;
                            break;

                        case "096":
                            BM096 objBM096 = new BM096();
                            str_error = objBM096.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM096 = null;
                            break;

                        case "099":
                            BM099 objBM099 = new BM099();
                            str_error = objBM099.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM099 = null;
                            break;

                        //case "100":
                        //    BM100 objBM100 = new BM100();
                        //    str_error = objBM100.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM100 = null;
                        //    break;

                        //case "101":
                        //    BM101 objBM101 = new BM101();
                        //    str_error = objBM101.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM101 = null;
                        //    break;

                        case "106":
                            BM106 objBM106 = new BM106();
                            str_error = objBM106.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM106 = null;
                            break;

                        case "117":
                            BM117 objBM117 = new BM117();
                            str_error = objBM117.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM117 = null;
                            break;

                        //case "120":
                        //    BM120 objBM120 = new BM120();
                        //    str_error = objBM120.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM120 = null;
                        //    break;

                        //case "121":
                        //    BM121 objBM121 = new BM121();
                        //    str_error = objBM121.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM121 = null;
                        //    break;

                        //case "122":
                        //    BM122 objBM122 = new BM122();
                        //    str_error = objBM122.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM122 = null;
                        //    break;

                        case "123":
                            BM123 objBM123 = new BM123();
                            str_error = objBM123.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked, txt_macro.Text);
                            objBM123 = null;
                            break;


                        //case "132":
                        //    BM132 objBM132 = new BM132();
                        //    str_error = objBM132.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM132 = null;
                        //    break;

                        //case "134":
                        //    BM134 objBM134 = new BM134();
                        //    str_error = objBM134.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM134 = null;
                        //    break;


                        //case "140":
                        //    BM140 objBM140 = new BM140();
                        //    str_error = objBM140.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM140 = null;
                        //    break;

                        case "141":
                            BM141 objBM141 = new BM141();
                            str_error = objBM141.fn_process(
                                txt_raw.Text,
                                cmb_sheet.SelectedItem.ToString(),
                                txt_dir.Text,
                                txt_file.Text, txt_gender.Text,
                                chk_open.Checked, chk_clean.Checked);

                            objBM141 = null;
                            break;

                        //case "142":
                        //    BM142 objBM142 = new BM142();
                        //    str_error = objBM142.fn_process(
                        //        txt_raw.Text,
                        //        txt_sicstemp.Text,
                        //        cmb_sheet.SelectedItem.ToString(),
                        //        txt_dir.Text,
                        //        txt_file.Text, txt_gender.Text,
                        //        chk_open.Checked, chk_clean.Checked);

                        //    objBM142 = null;
                        //    break;

                        default:
                            break;
                    }

                    
                    watch.Stop();
                    lbl_10.Text = watch.ElapsedMilliseconds.ToString();
                    watch = null;

                    if (string.IsNullOrEmpty(str_error))
                    {
                        MessageBox.Show("Processing complete!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else
                    
                    {
                        MessageBox.Show("Error during processing" + Environment.NewLine + "Please capture this message for tracking " + Environment.NewLine + "Contact your system administrator " + Environment.NewLine + str_error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        fn_killexcel();
                        
                    }

                    this.Enabled = true;

                    lbl_10.Text = "Ready...       "
                        + "Duration: "
                        + (Int64.Parse(lbl_10.Text) / 1000).ToString()
                        + "(s)";
                }
            }
        }
        private void fn_killexcel() 
        {
            Helper objHlpr = new Helper();
            objHlpr.fn_killexcel();
            objHlpr = null;
        }

        private void cmb_bm_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void grpb_pc_Enter(object sender, EventArgs e)
        {

        }

        private void grpb_optn_Enter(object sender, EventArgs e)
        {

        }

        private void txt_raw_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
                frmGender frmGender = new frmGender();
                frmGender.ShowDialog();
        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {
            MacroDB macroDB = new MacroDB();
            macroDB.ShowDialog();

        }
    }
}
