using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM011
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets [str_sheet];

            int intLastRow = wsraw.Cells [wsraw.Rows.Count, 2].End [XlDirection.xlUp].row;

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;
            string strTranscode = "";

            while ( string.IsNullOrEmpty(Variables.strBmYear) )
            {

                if ( string.IsNullOrEmpty(Variables.strBmYear) )
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            for ( int i = 1; i <= intLastRow; i++ )
            {
                if ( wsraw.Range ["A" + i].Value != null )
                {
                    var transcode = wsraw.Range ["A" + i].Value;
                    if ( transcode.GetType() == typeof(string) )
                    {
                        if ( transcode == "FIRST YEAR BUSINESS" )
                        {
                            strTranscode = "TNEWBUS";
                        }
                        else if ( transcode == "RENEWALS" )
                        {
                            strTranscode = "TRENEW";
                        }
                        else if ( transcode == "REFUNDS & ADJUSTMENTS**" )
                        {
                            strTranscode = "ADJUST";
                        }
                    }
                    string strCessionNo = Convert.ToString(wsraw.Range ["A" + i].Value);
                    if ( strCessionNo.Contains("-") )
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        string strPolicyNumber = wsraw.Range ["A" + i].Value;
                        dtDataRow [36] = wsraw.Range ["F" + i].Value; // Gender
                        objHlpr.fn_separatefullnamev9(wsraw.Range ["C" + i].Value, out string strFirstName, out string strLastName);
                        dtDataRow [31] = strLastName + ", " + strFirstName;
                        dtDataRow [32] = strLastName; // Last Name
                        dtDataRow [33] = strFirstName; // First Name
                        string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["D" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow [37] = strBirthday; // Birthday
                        dtDataRow [29] = "NATREID"; // Life ID Type
                        dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        strBirthday = strBirthday.Replace("/", "");
                        dtDataRow [0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                        dtDataRow [1] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // Cedent Cession Number
                        dtDataRow [79] = wsraw.Range ["E" + i].Value; // Life Issue Age
                        dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow [9] = "PAFM"; // Type of Business
                        dtDataRow [10] = "S"; // Reinsurance Methods
                        dtDataRow [13] = "IND"; // Class of Business
                        string strBusinessType = wsraw.Range ["B" + i].Value; ; // Business Type
                        if ( strBusinessType.ToUpper().Contains("YES") )
                        {
                            dtDataRow [14] = "F";
                        }
                        else
                        {
                            dtDataRow [14] = "T";
                        }
                        dtDataRow [24] = "ORIGINAL"; // Premium Frequency
                        dtDataRow [38] = "NONE"; // Smoker Status
                        dtDataRow [23] = "PHP"; // Cession Currency
                        dtDataRow [41] = Variables.strBmYear; // Policy Year
                        dtDataRow [05] = wsraw.Range ["G" + i].Value; // Branded Product
                        dtDataRow [20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["H" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range ["O" + i].Value), Convert.ToString( wsraw.Range ["Q" + i].Value), null, out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                        dtDataRow [76] = strRemarksCode; // Remarks
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                        dtDataRow [21] = strTranscode; // Transcode
                        dtDataRow[39] = objHlpr.fn_getmortality(Convert.ToString(wsraw.Range["J" + i].Value)); // Preffered Classific
                        if ( strTranscode == "TNEWBUS" )
                        {
                            dtDataRow [56] = "4000"; // Entry Code
                            dtDataRow [57] = wsraw.Range ["S" + i].Value; // Premium
                        }
                        else if ( strTranscode == "TRENEW" )
                        {
                            dtDataRow [58] = "4001"; // Entry Code
                            dtDataRow [59] = wsraw.Range ["S" + i].Value; // Premium
                        }
                        else
                        {
                            dtDataRow [60] = "4002"; // Entry Code
                            dtDataRow [61] = wsraw.Range ["S" + i].Value; // Premium
                        }

                        dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range ["V" + i].Value);
                        dblTotalSumAtRisk = dblTotalSumAtRisk + objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);

                        if (wsraw.Range["P" + i].Value != 0)
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            strPolicyNumber = wsraw.Range["A" + i].Value;
                            dtDataRow[36] = wsraw.Range["F" + i].Value; // Gender
                            //objHlpr.fn_separatefullnamev9(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName);
                            dtDataRow[31] = strLastName + ", " + strFirstName;
                            dtDataRow[32] = strLastName; // Last Name
                            dtDataRow[33] = strFirstName; // First Name
                            strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            strBirthday = strBirthday.Replace("/", "");
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                            dtDataRow[1] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // Cedent Cession Number
                            dtDataRow[79] = wsraw.Range["E" + i].Value; // Life Issue Age
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            strBusinessType = wsraw.Range["B" + i].Value; ; // Business Type
                            if (strBusinessType.ToUpper().Contains("YES"))
                            {
                                dtDataRow[14] = "F";
                            }
                            else
                            {
                                dtDataRow[14] = "T";
                            }
                            dtDataRow[24] = "ORIGINAL"; // Premium Frequency
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow[23] = "PHP"; // Cession Currency
                            dtDataRow[41] = Variables.strBmYear; // Policy Year
                            dtDataRow[05] = wsraw.Range["G" + i].Value; // Branded Product
                            dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                            //objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["P" + i].Value), Convert.ToString(wsraw.Range["R" + i].Value), null, out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                            dtDataRow[76] = strRemarksCode; // Remarks
                            dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["P" + i].Value)); // Original Sum Assured
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["R" + i].Value)); // Initial Sum at Risk
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["R" + i].Value)); // Sum at Risk
                            dtDataRow[21] = strTranscode; // Transcode
                            dtDataRow[39] = objHlpr.fn_getmortality(Convert.ToString(wsraw.Range["J" + i].Value)); // Preffered Classific
                            if (strTranscode == "TNEWBUS")
                            {
                                dtDataRow[56] = "4000"; // Entry Code
                                dtDataRow[57] = wsraw.Range["U" + i].Value; // Premium
                            }
                            else if (strTranscode == "TRENEW")
                            {
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["U" + i].Value; // Premium
                            }
                            else
                            {
                                dtDataRow[60] = "4002"; // Entry Code
                                dtDataRow[61] = wsraw.Range["U" + i].Value; // Premium
                            }

                            dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["V" + i].Value);
                            dblTotalSumAtRisk = dblTotalSumAtRisk + objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);
                        }
                    }
                }
            }
            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);
            #endregion

            string despath = str_saved + @"\BM011-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
