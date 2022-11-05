using System;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    public class BM009
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
            Worksheet wsraw = wbraw.Worksheets[str_sheet];

            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 1].End[XlDirection.xlUp].row;

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            #region CONNECTION TO DATABASE
            string szConnect = "DSN=SICS_Postgres_DB;" +
                                   "UID=sics;" +
                                   "PWD=sics_1";

            OdbcConnection cnDB = new OdbcConnection(szConnect);

            //try
            //{
            cnDB.Open();
            string query = "SELECT * FROM dbo_gender";
            OdbcCommand command = new OdbcCommand(query, cnDB);

            OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            while (reader.Read() == true)
            {
                Console.WriteLine("New Row:");
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    Console.WriteLine(reader.GetString(i));
                }
            }
            reader.Close();
            cnDB.Close();
            #endregion

            if (Regex.IsMatch(str_sheet, @"^\d+$"))
            {
                for (int i = 3; i <= intLastRow; i++)
                {
                    //if (wsraw.Range["C" + i].Value != null)
                    //{
                    //    string strCessionNo = Convert.ToString(wsraw.Range["C" + i].Value);
                    //    if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                    //    {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    var strPolno = wsraw.Range["D" + i].Value;
                    string strPolicyNumber;
                    if (strPolno.GetType() != typeof(string))
                    {
                        strPolicyNumber = strPolno.ToString("0");
                    }
                    else
                    {
                        strPolicyNumber = strPolno;
                    }
                    objHlpr.fn_separatefullnamev8(wsraw.Range["E" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                    dtDataRow[32] = strLastName; // Last Name
                    dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName); // First Name
                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strBirthday; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    strBirthday = strBirthday.Replace("/", "");
                    if (strPolicyNumber.Length > 7)
                    {
                        strPolicyNumber = strPolicyNumber.Substring(strPolicyNumber.Length - 7);
                    }
                    dtDataRow[0] = strPolicyNumber + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                    dtDataRow[1] = wsraw.Range["D" + i].Value; // Cedent Cession Number
                    dtDataRow[7] = wsraw.Range["D" + i].Value; // Group Scheme ID
                    dtDataRow[78] = wsraw.Range["H" + i].Value; // Attained Age
                    dtDataRow[82] = wsraw.Range["I" + i].Value; // Group Policy Holder
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[23] = "PHP"; // Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status
                    dtDataRow[39] = "STANDARD"; // Preferred Classific

                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                    dtDataRow[28] = wsraw.Range["O" + i].Value; // Cedent Retention
                    if (wsraw.Range["T" + i].Value == "TPD")
                    {
                        dtDataRow[5] = "TPD"; // Branded Product Cedent Code
                    }
                    else
                    {
                        dtDataRow[5] = "GEB"; // Branded Product Cedent Code
                    }
                    string strPolicyStartDate = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[20] = strPolicyStartDate; // Policy Start Date
                    dtDataRow[19] = strPolicyStartDate.Substring(0, 6) + Variables.strBmYear; // REINSURANCE_START_DATE
                    dtDataRow[22] = strPolicyStartDate.Substring(0, 6) + Variables.strBmYear; // TRANS_EFFECTIVE_DATE

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["N" + i].Value), Convert.ToString(wsraw.Range["Q" + i].Value), Convert.ToString(wsraw.Range["Q" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk

                    if (wsraw.Range["C" + i].Value == "NB")
                    {
                        dtDataRow[56] = "4000"; // Entry Code
                        dtDataRow[57] = wsraw.Range["S" + i].Value; // Premium
                        dtDataRow[21] = "TNEWBUS"; // Transcode
                    }
                    else
                    {
                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["S" + i].Value; // Premium
                        dtDataRow[21] = "TRENEW"; // Transcode
                    }
                    dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["S" + i].Value);
                    dblTotalSumAtRisk = dblTotalSumAtRisk + Convert.ToDecimal(wsraw.Range["Q" + i].Value);
                    //    }
                    //}
                }
            }
            //Adjustment
            else if (str_sheet.ToUpper().Contains("ADJUSTMENT"))
            {
                for (int i = 3; i <= intLastRow; i++)
                {

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    var strPolno = wsraw.Range["C" + i].Value;
                    string strPolicyNumber;
                    if (strPolno.GetType() != typeof(string))
                    {
                        strPolicyNumber = strPolno.ToString("0");
                    }
                    else
                    {
                        strPolicyNumber = strPolno;
                    }
                    objHlpr.fn_separatefullnamev8(wsraw.Range["D" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                    dtDataRow[32] = strLastName; // Last Name
                    dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName); // First Name
                    dtDataRow [34] = strMiddleInitial; // Middle Initials
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; //Gender
                    string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strBirthday; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    strBirthday = strBirthday.Replace("/", "");
                    if (strPolicyNumber.Length > 7)
                    {
                        strPolicyNumber = strPolicyNumber.Substring(strPolicyNumber.Length - 7);
                    }
                    dtDataRow[0] = strPolicyNumber + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                    dtDataRow[1] = wsraw.Range["C" + i].Value; // Cedent Cession Number
                    dtDataRow[7] = wsraw.Range["C" + i].Value; // Group Scheme ID
                    //dtDataRow[79] = wsraw.Range["G" + i].Value; // Life Issue Age
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[23] = "PHP"; // Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status
                    dtDataRow[39] = "STANDARD"; // Preferred Classific
                    dtDataRow[41] = Variables.strBmYear; // Policy Year

                    dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["J" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["U" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                    dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["U" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["M" + i].Value), Convert.ToString(wsraw.Range["P" + i].Value), Convert.ToString(wsraw.Range["P" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                    dtDataRow[25] = (objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) < 0) ? "1" : strOriginalSum; // Original Sum Assured
                    dtDataRow[27] = (objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) < 0) ? "1" : strInitialSum; // Initial Sum at Risk
                    dtDataRow[77] = (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) < 0) ? "1" : strSumAtRisk; // Sum at Risk

                    dtDataRow[78] = wsraw.Range["G" + i].Value; // Attained Age
                    if (wsraw.Range["S" + i].Value == "TPD")
                    {
                        dtDataRow[5] = "TPD"; // Branded Product Cedent Code
                    }
                    else
                    {
                        dtDataRow[5] = "GEB"; // Branded Product Cedent Code
                    }
                    string strtranscode = objHlpr.fn_CheckTransCodeV3(wsraw.Range["T" + i].Value);
                    if (strtranscode == "ADJUST")
                    {
                        dtDataRow[21] = strtranscode; // Transcode
                        dtDataRow[62] = "4004"; // Entry Code
                        dtDataRow[63] = wsraw.Range["R" + i].Value; // Premium
                        dtDataRow[76] = wsraw.Range["T" + i].Value; // Remarks

                    }
                    else
                    {
                        dtDataRow[21] = strtranscode; // Transcode
                        dtDataRow[60] = "4004"; // Entry Code
                        dtDataRow[61] = wsraw.Range["R" + i].Value; // Premium
                    }
                    //dtDataRow[21] = "ADJUST"; // Transcode
                    //dtDataRow[60] = "4002"; // Entry Code
                    //dtDataRow[61] = wsraw.Range["R" + i].Value; // Premium


                    dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["R" + i].Value);


                }
            }

            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 009", "Information");
                return "";
            }
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM009-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";


        }
    }
}