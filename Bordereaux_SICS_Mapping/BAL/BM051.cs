using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
namespace Bordereaux_SICS_Mapping.BAL
{
    class BM051
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            //try
            //{
                _Global _var = new _Global();
                Helper objHlpr = new Helper();
                System.Data.DataTable objdt_template = new System.Data.DataTable();

                objdt_template = objHlpr.dt_formtemplate(str_sheet);
                Application eapp = new Application();
                Workbook wbraw = eapp.Workbooks.Open(str_raw);
                Worksheet wsraw = wbraw.Worksheets[str_sheet];
                int intLastRow = wsraw.Cells [wsraw.Rows.Count, 1].End [XlDirection.xlUp].row;
                //Range rawrange = wsraw.Columns["A:A"];
                //var result = rawrange.Find("FOOTER", LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlWhole);
                //int intLastRow = result.Row -1;
                decimal dblTotalPremiumPHP = 0, dblTotalPremiumUSD = 0, dblTotalSumAtRiskPHP = 0, dblTotalSumAtRiskUSD = 0;

                DataRow dtDataRow;

                // BM051
                if (str_sheet.ToUpper().Contains("INF_NEC_UNIRE_PH") || str_sheet.ToUpper().Contains("INF_NEC_RGA_PH"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {

                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AT" + i].Value), Convert.ToString(wsraw.Range["AV" + i].Value), Convert.ToString(wsraw.Range["AW" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                            dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        Console.WriteLine(strLastName + ", " + strFirstName + " " + strMiddleInitial);
                        dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initials
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        dtDataRow[23] = wsraw.Range["AZ" + i].Value; //  Currency
                                                                     //string strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                     //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        var polyear = wsraw.Range["BE" + i].Value;
                        string strPolicyYear = polyear.ToString("0");
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                        dtDataRow[5] = wsraw.Range["Y" + i].Value; // Branded Product Cedent Code
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency 
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        //dtDataRow[19] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                        //dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        //dtDataRow[22] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date

                        var issdate = wsraw.Range["T" + i].Value;
                        var polstartdate = wsraw.Range["BE" + i].Value;
                        string strIssueDate = issdate.ToString("0.00");
                        string strPolicyStartDate = polstartdate.ToString("0.00");

                        if (wsraw.Range["AZ" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["BF" + i].Value != 0)
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["BF" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BF" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BF" + i].Value);
                            }

                            if (wsraw.Range["BG" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initials
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                dtDataRow[23] = wsraw.Range["AZ" + i].Value; //  Currency
                                                                             //strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                             //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet;
                                polyear = wsraw.Range["BE" + i].Value;
                                strPolicyYear = polyear.ToString("0.00");
                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                                dtDataRow[5] = wsraw.Range["Y" + i].Value; // Branded Product Cedent Code
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                                dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[21] = "TRENEW"; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["BG" + i].Value; // Premium

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                                if (wsraw.Range["AZ" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                            }
                        }
                        else if (wsraw.Range["BG" + i].Value != 0)
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["BG" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                            }
                        }
                        else if (strIssueDate.Substring(0, 4) == strPolicyStartDate.Substring(0, 4))
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["BF" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BF" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BF" + i].Value);
                            }
                        }
                        else
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["BG" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BG" + i].Value);
                            }
                        }

                    }
                }
                // BM051B
                else if (str_sheet.ToUpper().Contains("INF_UNR_RGA_PH") || str_sheet.ToUpper().Contains("INF_UNR_UNIR_PH"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AT" + i].Value), Convert.ToString(wsraw.Range["AV" + i].Value), Convert.ToString(wsraw.Range["AW" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initial
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        dtDataRow[23] = wsraw.Range["AZ" + i].Value; //  Currency
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                   //string strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                   //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        var polyear = wsraw.Range["BG" + i].Value;
                        string strPolicyYear = polyear.ToString("0.00");
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                        dtDataRow[5] = wsraw.Range["Y" + i].Value; // Branded Product Cedent Code
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        var issdate = wsraw.Range["T" + i].Value;
                        var polstartdate = wsraw.Range["BG" + i].Value;
                        string strIssueDate = issdate.ToString("0.00");
                        string strPolicyStartDate = polstartdate.ToString("0.00");

                        if (wsraw.Range["AZ" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["BH" + i].Value != 0)
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["BH" + i].Value; // Premium
                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BH" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BH" + i].Value);
                            }

                            if (wsraw.Range["BI" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initial
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                dtDataRow[23] = wsraw.Range["AZ" + i].Value; //  Currency
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                           //strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                           //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet;
                                polyear = wsraw.Range["BG" + i].Value;
                                strPolicyYear = polyear.ToString("0.00");
                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                                dtDataRow[5] = wsraw.Range["Y" + i].Value; // Branded Product Cedent Code
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                                dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[21] = "TRENEW"; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["BI" + i].Value; // Premium

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                                if (wsraw.Range["AZ" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                            }
                        }
                        else if (wsraw.Range["BI" + i].Value != 0)
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["BI" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                            }
                        }
                        else if (strIssueDate.Substring(0, 4) == strPolicyStartDate.Substring(0, 4))
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["BH" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BH" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BH" + i].Value);
                            }
                        }
                        else
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["BI" + i].Value; // Premium

                            if (wsraw.Range["AZ" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["BI" + i].Value);
                            }
                        }

                    }
                }
                // BM051C
                else if (str_sheet.ToUpper().Contains("ADB_NEC_RGA_PH") || str_sheet.ToUpper().Contains("ADB_NEC_UNIRE_PH"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AH" + i].Value), Convert.ToString(wsraw.Range["AI" + i].Value), Convert.ToString(wsraw.Range["AJ" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initial
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        dtDataRow[23] = wsraw.Range["AN" + i].Value; //  Currency
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                   //string strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                   //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        var polyear = wsraw.Range["AT" + i].Value;
                        string strPolicyYear = polyear.ToString("0.00");
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                        dtDataRow[5] = wsraw.Range["AO" + i].Value; // Branded Product Cedent Code
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        var issdate = wsraw.Range["T" + i].Value;
                        var polstartdate = wsraw.Range["AT" + i].Value;
                        string strIssueDate = issdate.ToString("0.00");
                        string strPolicyStartDate = polstartdate.ToString("0.00");

                        if (wsraw.Range["AN" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["AU" + i].Value != 0)
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AU" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }

                            if (wsraw.Range["AV" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initial
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                dtDataRow[23] = wsraw.Range["AN" + i].Value; //  Currency
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                           //strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                           //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet;
                                polyear = wsraw.Range["AT" + i].Value;
                                strPolicyYear = polyear.ToString("0.00");
                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                                dtDataRow[5] = wsraw.Range["AO" + i].Value; // Branded Product Cedent Code
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                                dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[21] = "TRENEW"; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                                if (wsraw.Range["AN" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }

                            }
                        }
                        else if (wsraw.Range["AV" + i].Value != 0)
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                        }
                        else if (strIssueDate.Substring(0, 4) == strPolicyStartDate.Substring(0, 4))
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AU" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }
                        }
                        else
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                        }
                    }
                }
                // BM051C
                else if (str_sheet.ToUpper().Contains("RDR_UNR_RGA_PH") || str_sheet.ToUpper().Contains("RDR_UNR_UNIR_PH"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AH" + i].Value), Convert.ToString(wsraw.Range["AI" + i].Value), Convert.ToString(wsraw.Range["AJ" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullnamev2(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initial
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        dtDataRow[23] = wsraw.Range["AN" + i].Value; //  Currency
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                   //string strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                   //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet + " " + strRemarksCode;
                        var polyear = wsraw.Range["AT" + i].Value;
                        string strPolicyYear = polyear.ToString("0.00");
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                        dtDataRow[5] = wsraw.Range["AO" + i].Value; // Branded Product Cedent Code
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        var issdate = wsraw.Range["T" + i].Value;
                        var polstartdate = wsraw.Range["AT" + i].Value;
                        string strIssueDate = issdate.ToString("0.00");
                        string strPolicyStartDate = polstartdate.ToString("0.00");

                        if (wsraw.Range["AN" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["AU" + i].Value != 0)
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AU" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }

                            if (wsraw.Range["AV" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initial
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                dtDataRow[23] = wsraw.Range["AN" + i].Value; //  Currency
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                           //strRemarks = wsraw.Range["A" + i].Value + "_" + wsraw.Range["B" + i].Value;
                                                                                                           //dtDataRow[76] = strRemarks.Replace(" ", "") + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet;
                                polyear = wsraw.Range["AT" + i].Value;
                                strPolicyYear = polyear.ToString("0.00");
                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                                dtDataRow[5] = wsraw.Range["AO" + i].Value; // Branded Product Cedent Code
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                                dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[21] = "TRENEW"; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                                if (wsraw.Range["AN" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }

                            }
                        }
                        else if (wsraw.Range["AV" + i].Value != 0)
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }

                        }
                        else if (strIssueDate.Substring(0, 4) == strPolicyStartDate.Substring(0, 4))
                        {
                            dtDataRow[21] = "TNEWBUS"; // Transcode
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AU" + i].Value; // Premium

                            if (wsraw.Range["AN" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AU" + i].Value);
                            }
                        }
                        else
                        {
                            dtDataRow[21] = "TRENEW"; // Transcode
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AV" + i].Value; // Premium

                            if (wsraw.Range["AV" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range["AV" + i].Value);
                            }
                        }
                    }
                }
                // BM051 LAPSED
                else if (str_sheet.ToUpper().Contains("LAPEXT_PH_UNR_UNI") || str_sheet.ToUpper().Contains("LAPEXT_PH_NEC_UNIRE"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["R" + i].Value), Convert.ToString(wsraw.Range["Q" + i].Value), Convert.ToString(wsraw.Range["M" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initials
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        dtDataRow[5] = wsraw.Range["S" + i].Value; // Branded Product Cedent Code
                                                                   //string strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                                                                   //dtDataRow[76] = strRemarks.Trim() + " " + wbraw.Name + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet + " " + strRemarksCode;

                        dtDataRow[21] = "TLAPSE"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        string strFilePath = Path.GetDirectoryName(str_raw);
                        string strPolicyYear = strFilePath.Substring(strFilePath.Length - 6);
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year


                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        if (wsraw.Range["K" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["N" + i].Value != 0)
                        {
                            dtDataRow[60] = "4002"; // Entry Code
                            dtDataRow[61] = (wsraw.Range["N" + i].Value / 100) * -1; ; // FY R&A

                            if (wsraw.Range["K" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / -100);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / -100);
                            }


                            if (wsraw.Range["O" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initials
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                dtDataRow[5] = wsraw.Range["S" + i].Value; // Branded Product Cedent Code
                                                                           //strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                                                                           //dtDataRow[76] = strRemarks.Trim() + " " + wbraw.Name + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet + " " + strRemarksCode;
                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year
                                dtDataRow[21] = "TLAPSE"; // Transcode
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[38] = "NONE"; // Smoker Status
                                dtDataRow[39] = "STANDARD"; // Preferred Classific

                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date


                                dtDataRow[62] = "4004"; // Entry Code   
                                dtDataRow[63] = (wsraw.Range["O" + i].Value / 100) * -1; ; // Premium

                                if (wsraw.Range["K" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["O" + i].Value) / -100);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["O" + i].Value) / -100);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                            }
                        }
                        else if (wsraw.Range["O" + i].Value != 0)
                        {
                            dtDataRow[62] = "4004"; // Entry Code
                            dtDataRow[63] = (wsraw.Range["O" + i].Value / 100) * -1;

                            if (wsraw.Range["K" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["O" + i].Value) / -100);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["O" + i].Value) / -100);
                            }
                        }
                    }
                }
                // BM051 REINSTATEMENT
                else if (str_sheet.ToUpper().Contains("RSTEXT_PH_UNR_UNIRE") || str_sheet.ToUpper().Contains("RSTEXT_PH_NEC_UNIRE"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["Q" + i].Value), Convert.ToString(wsraw.Range["P" + i].Value), Convert.ToString(wsraw.Range["L" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initials
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                   //string strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                                                                                                   //dtDataRow[76] = strRemarks.Trim() + " " + wbraw.Name + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        //dtDataRow[41] = strPolicyYear.Substring(2, 4); // Policy Year
                        dtDataRow[5] = wsraw.Range["R" + i].Value; // Branded Product Cedent Code
                        dtDataRow[21] = "TREINS"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        string strFilePath = Path.GetDirectoryName(str_raw);
                        string strPolicyYear = strFilePath.Substring(strFilePath.Length - 6);
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        if (wsraw.Range["K" + i].Value == "PHP")
                        {
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }

                        if (wsraw.Range["M" + i].Value != 0)
                        {
                            dtDataRow[60] = "4002"; // Entry Code
                            dtDataRow[61] = (wsraw.Range["M" + i].Value / 100); // FY R&A

                            if (wsraw.Range["K" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["M" + i].Value) / 100);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["M" + i].Value) / 100);
                            }


                            if (wsraw.Range["N" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                                dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initials
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) / 100; // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                                dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                                strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[5] = wsraw.Range["R" + i].Value; // Branded Product Cedent Code
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                                                                           //strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                                                                                                           //dtDataRow[76] = strRemarks.Trim() + " " + wbraw.Name + " " + strRemarksCode; // Remarks
                                dtDataRow[76] = str_sheet;
                                //dtDataRow[41] = strPolicyYear.Substring(2, 4); // Policy Year
                                dtDataRow[21] = "TREINS"; // Transcode
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PAFM"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[39] = "STANDARD"; // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                //dtDataRow[19] = strPolicyYear.Substring(0, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(2, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                                                                                                                                     //dtDataRow[22] = strPolicyYear.Substring(0, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(2, 4); // Trans Effective Date

                                dtDataRow[62] = "4004"; // Entry Code
                                dtDataRow[63] = (wsraw.Range["N" + i].Value / 100);

                                if (wsraw.Range["K" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                            }
                        }
                        else if (wsraw.Range["N" + i].Value != 0)
                        {
                            dtDataRow[62] = "4004"; // Entry Code
                            dtDataRow[63] = (wsraw.Range["N" + i].Value / 100);

                            if (wsraw.Range["K" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100);
                            }
                        }


                    }
                }
                // Premium Refund
                else if (str_sheet.ToUpper().Contains("RBT_UNI_PH_NEC") || str_sheet.ToUpper().Contains("RBT_UNI_PH_UNR"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AA" + i].Value), Convert.ToString(wsraw.Range["Z" + i].Value), Convert.ToString(wsraw.Range["V" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initials
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) / 100; // Initial Sum at Risk
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100; // Sum at Risk
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        dtDataRow[23] = wsraw.Range["AG" + i].Value; //  Currency
                        dtDataRow[76] = wbraw.Name + strRemarksCode;
                        string strPolicyYear = Path.GetFileNameWithoutExtension(wbraw.Name);
                        dtDataRow[41] = strPolicyYear.Substring(strPolicyYear.Length - 6, 4); // Policy Year
                        dtDataRow[21] = "ADJUST"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[5] = (wsraw.Range["O" + i].Value == "DB") ? "LIFE" : ""; // Branded Product Cedent Code
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[21] = "ADJUST"; // Transcode
                        dtDataRow[62] = "4004"; // Entry Code
                        dtDataRow[63] = (wsraw.Range["X" + i].Value / 100) * -1; // Premium

                        dtDataRow[19] = strPolicyYear.Substring(strPolicyYear.Length - 2, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(strPolicyYear.Length - 6, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(strPolicyYear.Length - 2, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(strPolicyYear.Length - 6, 4); // Trans Effective Date

                        if (wsraw.Range["AG" + i].Value == "PHP")
                        {
                            dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["X" + i].Value) / -100);
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["X" + i].Value) / -100 );
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                    }
                }

                else
                {
                    System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 051", "Information");
                    return "";
                }

                #region Computing Hash 
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Premium PHP:";
                dtDataRow[1] = dblTotalPremiumPHP;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Premium USD:";
                dtDataRow[1] = dblTotalPremiumUSD;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk PHP:";
                dtDataRow[1] = dblTotalSumAtRiskPHP;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk USD:";
                dtDataRow[1] = dblTotalSumAtRiskUSD;
                objdt_template.Rows.Add(dtDataRow);
                #endregion



                string despath = str_saved + @"\BM051-" + str_sheet + str_savef + ".xlsx";
                objHlpr.fn_savefile(objdt_template, despath);

                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";


            //}
            //catch (Exception RowData)
            //{
            //    return "Check the data for row" + RowData;
            //}


        }

    }
}