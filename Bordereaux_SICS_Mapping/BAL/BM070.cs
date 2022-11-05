using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM070
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets [str_sheet];

            int intLastRow = wsraw.Cells [wsraw.Rows.Count, 1].End [XlDirection.xlUp].row;

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;
            string valueTransEffectiveDate = string.Empty;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {

                if(string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            if(str_sheet.ToUpper().Contains("TRAD") || str_sheet.ToUpper().Contains("ACC"))
            {
                for(int i = 1; i <= intLastRow; i++)
                {
                    if(wsraw.Range ["B" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range ["B" + i].Value);
                        if(Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            dtDataRow [0] = wsraw.Range ["B" + i].Value; // Policy Number
                            dtDataRow [36] = objHlpr2.fn_MaleOrFemale(wsraw.Range ["E" + i].Value); // Gender
                            objHlpr.fn_separatefullnamev2(wsraw.Range ["F" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow [32] = strLastName; // Last Name
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow [34] = objHlpr2.fn_removeCharacters(strMiddleInitial); // Middle Initials
                            string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow [37] = strBirthday; // Birthday
                            dtDataRow [29] = "NATREID"; // Life ID Type
                            dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow [79] = wsraw.Range ["H" + i].Value; // Life Issue Age
                            //dtDataRow[25] = wsraw.Range["K" + i].Value; // Original Sum Assured
                            //dtDataRow[27] = wsraw.Range["N" + i].Value; // Initial Sum at Risk
                            dtDataRow [28] = wsraw.Range ["M" + i].Value; // Cedant Retention
                            dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow [9] = "PAFM"; // Type of Business
                            dtDataRow [10] = "S"; // Reinsurance Methods
                            dtDataRow [13] = "IND"; // Class of Business
                            dtDataRow [14] = "T"; // Business Type
                            dtDataRow [24] = "YLY"; // Premium Frequency
                            dtDataRow [38] = "NONE"; // Smoker Status
                            dtDataRow [23] = "PHP"; // Cession Currency
                            dtDataRow [5] = wsraw.Range ["C" + i].Value; //Branded Product
                            dtDataRow [41] = Variables.strBmYear; // Policy Year
                            dtDataRow [39] = "STANDARD";


                            dtDataRow [20] = wsraw.Range ["D" + i].Value.ToString("MM/dd/yyyy"); // Policy Start Date
                            dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                            dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                            objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range ["K" + i].Value), Convert.ToString(wsraw.Range ["N" + i].Value), null, out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk

                            string strTcode = "TRENEW";
                            dtDataRow [21] = strTcode; // Transcode
                            dtDataRow [58] = "4001"; // Entry Code
                            dtDataRow [59] = wsraw.Range ["O" + i].Value; // Premium

                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");
                            dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                            dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                            dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range ["O" + i].Value);
                            dblTotalSumAtRisk = dblTotalSumAtRisk + Convert.ToDecimal(strSumAtRisk);
                        }
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 070", "Information");
                return "";
            }


            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM070-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
