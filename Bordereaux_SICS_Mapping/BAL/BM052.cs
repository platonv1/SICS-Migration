using System;
using System.Data;
using System.Linq;
using System.Globalization;
using System.IO;
using Microsoft.Office.Interop.Excel;
namespace Bordereaux_SICS_Mapping.BAL
{
    class BM052
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);

            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 5].End[XlDirection.xlUp].row;
            string strFilePath = Path.GetDirectoryName(str_raw);
            string strPolicyYear = strFilePath.Substring(strFilePath.Length - 6);
            //strPolicyYear = strPolicyYear.Insert(2, "/");
            //DateTime PolicyYear = DateTime.ParseExact(strPolicyYear, "MM/yyyy", CultureInfo.InvariantCulture);
            //strPolicyYear = PolicyYear.ToString("MM/yyyy");

            DataRow dtDataRow;
            decimal dblTotalPremiumPHP = 0, dblTotalPremiumUSD = 0, dblTotalSumAtRiskPHP = 0, dblTotalSumAtRiskUSD = 0;

            //try
            //{
                // NB
                if (str_sheet.ToUpper().Contains("NBEXT_PH"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["O" + i].Value), Convert.ToString(wsraw.Range["N" + i].Value), Convert.ToString(wsraw.Range["L" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

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
                        string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                        //string strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                        //dtDataRow[76] = strRemarks.Trim() + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet + strRemarksCode + "Dummy Column AN";
                        dtDataRow[5] = wsraw.Range["Q" + i].Value; // Branded Product Cedent Code
                        dtDataRow[21] = "TNEWBUS"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[56] = "4000"; // Entry Code
                        dtDataRow[57] = wsraw.Range["M" + i].Value / 100; // Premium

                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        if (wsraw.Range["K" + i].Value == "PHP")
                        {
                            dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["M" + i].Value) / 100);
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["M" + i].Value) / 100);
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                    }
                }

                // LIFE
                else if (str_sheet.ToUpper().Contains("INF_MIL"))
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AT" + i].Value), Convert.ToString(wsraw.Range["AV" + i].Value), Convert.ToString(wsraw.Range["AW" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + "," + strFirstName + " " + strMiddleInitial + ".";
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
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        dtDataRow[23] = wsraw.Range["AZ" + i].Value; //  Currency
                        dtDataRow[5] = wsraw.Range["Y" + i].Value; // Branded Product Cedent Code
                        //string strRemarks = wsraw.Range["B" + i].Value + "_" + wsraw.Range["A" + i].Value;
                        //dtDataRow[76] = strRemarks.Trim() + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        dtDataRow[21] = "TRENEW"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["AS" + i].Value / 100; // Premium
                        
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        if (wsraw.Range["AZ" + i].Value == "PHP")
                        {
                            dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["AS" + i].Value) / 100);
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["AS" + i].Value) / 100);
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                    }
                }

                // ADB
                else if (str_sheet.ToUpper().Contains("ADB_MIL")) 
                {
                    for (int i = 2; i <= intLastRow; i++)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["AH" + i].Value), Convert.ToString(wsraw.Range["AI" + i].Value), Convert.ToString(wsraw.Range["AJ" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                        dtDataRow[0] = wsraw.Range["F" + i].Value; // Policy Number
                        dtDataRow[36] = wsraw.Range["I" + i].Value; // Gender
                        objHlpr.fn_separatefullname(wsraw.Range["H" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + "," + strFirstName + " " + strMiddleInitial + ".";
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
                        dtDataRow[79] = wsraw.Range["U" + i].Value; // Life Issue Age
                        //string strRemarks = wsraw.Range["B" + i].Value + "_" + wsraw.Range["A" + i].Value;
                        //dtDataRow[76] = strRemarks.Trim() + " " + strRemarksCode; // Remarks
                        dtDataRow[76] = str_sheet;
                        dtDataRow[23] = wsraw.Range["AN" + i].Value; //  Currency
                        dtDataRow[5] = wsraw.Range["AO" + i].Value; // Branded Product Cedent Code
                        dtDataRow[21] = "TRENEW"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["W" + i].Value); // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["AG" + i].Value / 100; // Premium

                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date


                        if (wsraw.Range["AN" + i].Value == "PHP")
                        {
                            dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["AG" + i].Value) / 100);
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["AG" + i].Value) / 100);
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                    }
                }

                // LAPSED
                else if (str_sheet.ToUpper().Contains("LAPEXT_PH"))
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
                        dtDataRow[76] = str_sheet;
                        dtDataRow[21] = "TLAPSE"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[38] = "NONE"; // Smoker Status
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

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
                            dtDataRow[61] = (wsraw.Range["N" + i].Value / 100) * -1; // FY R&A

                            if (wsraw.Range["K" + i].Value == "PHP")
                            {
                                dblTotalPremiumPHP = dblTotalPremiumPHP + ((Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100) * -1);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + ((Convert.ToDecimal(wsraw.Range["N" + i].Value) / 100) * -1);
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
                                dtDataRow[76] = str_sheet;
                                dtDataRow[21] = "TLAPSE"; // Transcode
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PA"; // Type of Business
                                dtDataRow[10] = "PT"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[39] = "STANDARD"; // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date


                                dtDataRow[62] = "4004"; // Entry Code   
                                dtDataRow[63] = (wsraw.Range["O" + i].Value / 100) * -1;

                                if (wsraw.Range["K" + i].Value == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + ((Convert.ToDecimal(wsraw.Range["O" + i].Value) / 100) * -1);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + ((Convert.ToDecimal(wsraw.Range["O" + i].Value) / 100) * -1);
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
                                dblTotalPremiumPHP = dblTotalPremiumPHP + ((Convert.ToDecimal(wsraw.Range["O" + i].Value) / 100) * -1);
                            }
                            else
                            {
                                dblTotalPremiumUSD = dblTotalPremiumUSD + ((Convert.ToDecimal(wsraw.Range["O" + i].Value) / 100) * -1);
                            }
                        }
                    }
                }

                // REINSTATEMENT
                else if (str_sheet.ToUpper().Contains("RSTEXT_PH"))
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
                        dtDataRow[5] = wsraw.Range["R" + i].Value; // Branded Product Cedent Code
                        dtDataRow[21] = "TREINS"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

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
                                dtDataRow[21] = "TREINS"; // Transcode
                                dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "PA"; // Type of Business
                                dtDataRow[10] = "PT"; // Reinsurance Methods
                                dtDataRow[13] = "IND"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[24] = "MLY"; // Premium Frequency
                                dtDataRow[39] = "STANDARD"; // Preferred Classific
                                dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                                dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                                dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                                dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

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
                else if (str_sheet.ToUpper().Contains("RBT_UMRE"))
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
                        dtDataRow[23] = wsraw.Range["K" + i].Value; //  Currency
                        dtDataRow[76] = Path.GetFileNameWithoutExtension(wbraw.Name) + strRemarksCode;
                        dtDataRow[21] = "ADJUST"; // Transcode
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PA"; // Type of Business
                        dtDataRow[10] = "PT"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[24] = "MLY"; // Premium Frequency
                        dtDataRow[39] = "STANDARD"; // Preferred Classific
                        dtDataRow[5] = (wsraw.Range["O" + i].Value == "DB") ? "LIFE" : ""; // Branded Product Cedent Code
                        dtDataRow[38] = (wsraw.Range["J" + i].Value == "B") ? "BLENDED" : "NONE"; // Smoker Status

                        dtDataRow[21] = "ADJUST"; // Transcode
                        dtDataRow[60] = "4002"; // Entry Code
                        dtDataRow[61] = wsraw.Range["X" + i].Value / -100; // Premium
                        
                        dtDataRow[41] = strPolicyYear.Substring(0, 4); // Policy Year

                        dtDataRow[19] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = strPolicyYear.Substring(4, 2) + "/" + objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("dd") + "/" + strPolicyYear.Substring(0, 4); // Trans Effective Date

                        if (wsraw.Range["K" + i].Value == "PHP")
                        {
                            dblTotalPremiumPHP = dblTotalPremiumPHP + (Convert.ToDecimal(wsraw.Range["X" + i].Value) / -100);
                            dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                        else
                        {
                            dblTotalPremiumUSD = dblTotalPremiumUSD + (Convert.ToDecimal(wsraw.Range["X" + i].Value) / -100);
                            dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + (objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) / 100);
                        }
                    }
                }

                else
                {
                    System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 052", "Information");
                    return "";
                }
            //}
            //catch (Exception ex)
            //{
            //    return ex.Message;
            //}

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

            string despath = str_saved + @"\BM052-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }
}