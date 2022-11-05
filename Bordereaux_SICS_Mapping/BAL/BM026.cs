using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM026
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
            double dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            if (str_sheet.ToUpper().Contains("DETA"))
            {
                for (int i = 3; i <= intLastRow; i++)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    var strPolno = wsraw.Range["A" + i].Value;
                    Console.WriteLine(strPolno);
                    string strPolicyNumber = strPolno.ToString();
                    string strGender = objHlpr.fn_getgenderv3( wsraw.Range["E" + i].Value); // Gender
                    string strBirthday = Convert.ToDateTime(wsraw.Range["F" + i].Value).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strBirthday; // Birthday
                    if (strGender.ToUpper().Contains("FEMALE"))
                    {
                        dtDataRow[36] = "F";
                    }
                    else
                    {
                        dtDataRow[36] = "M";
                    }

                    string strfullname = Convert.ToString(wsraw.Range["C" + i].Value);
                    Console.WriteLine(strfullname);
                    if (Regex.IsMatch(strfullname, @"^\d+$") || Regex.IsMatch(strfullname, @"PHP\d|\d") || strfullname.ToUpper().Contains(";"))
                    {
                        dtDataRow[31] = "DummyLastName" + ", " + "DummyFirstName" + " " + "DummyMiddleName" + "."; // Full name
                        dtDataRow[32] = "DummyLastName"; // Last Name
                        dtDataRow[33] = "DummyFirstName"; // First Name
                        dtDataRow[34] = "DummyMiddleName"; // Middle Initials
                        dtDataRow[30] = "DUMMYDU07011900"; // Life ID
                        dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + "DD07011900"; // Policy Number
                    }
                    else
                    {
                        objHlpr.fn_separatefullnamev3(strfullname, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        strFirstName = objHlpr2.fn_checkFirstname(strFirstName);
                        dtDataRow [33] = strFirstName; // First Name
                        dtDataRow [34] = strMiddleInitial; // Middle Initials
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        string strBirthday1 = strBirthday.Replace("/", "");
                        dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday1; // PolicyNumber
                    }

                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[78] = wsraw.Range["G" + i].Value; // Attained Age
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product    
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[23] = "PHP"; // Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status 
                    dtDataRow[5] = wsraw.Range["V" + i].Value; // Plan Code
                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                    dtDataRow[82] = wsraw.Range["B" + i].Value; // Group Policy Holder
                    dtDataRow[7] = wsraw.Range["A" + i].Value; // Group Scheme ID
                    dtDataRow[76] = wsraw.Range["K" + i].Value; // Remarks
                    string strPreferredClassific = Convert.ToString(wsraw.Range["H" + i].Value);
                    if (strPreferredClassific == "1")
                    {
                        strPreferredClassific = "CLASS1";
                    }
                    else if (strPreferredClassific == "2")
                    {
                        strPreferredClassific = "CLASS2";
                    }
                    else if (strPreferredClassific == "3")
                    {
                        strPreferredClassific = "CLASS3";
                    }
                    else if (strPreferredClassific == "4")
                    {
                        strPreferredClassific = "CLASS4";
                    }       
                    else
                    {
                        strPreferredClassific = "CLASS5";
                    }
                    dtDataRow[39] = strPreferredClassific; // Preferred Classific)
                    dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                    dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["Q" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date

                    if (wsraw.Range["T" + i].Value == null)
                    {
                        dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["S" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                    }
                    else
                    {
                        dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                    }    
                    objHlpr.fn_CheckTransCode(wsraw.Range["J" + i].Value, out string transcode);
                    dtDataRow[21] = transcode; // Transcode

                    string strBusinessType = wsraw.Range["D" + i].Value; // Business Type
                    if (strBusinessType.ToUpper().Contains("STANDARD"))
                    {
                        dtDataRow[14] = "T";
                    }
                    else
                    {
                        dtDataRow[14] = "F";
                    }

                    if (transcode == "TNEWBUS")
                    {
                        dtDataRow[56] = "4000"; // Entry Code
                        dtDataRow[57] = wsraw.Range["AK" + i].Value; // Premium
                    }
                    else
                    {
                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["AK" + i].Value; // Premium
                    }
                    dtDataRow[28] = (wsraw.Range["AE" + i].Value < 0 )? "1": wsraw.Range["AE" + i].Value; // Cedent Retention
                    dtDataRow[25] = (wsraw.Range["AB" + i].Value < 0) ? "1" : wsraw.Range["AB" + i].Value; // Original Sum Assured
                    dtDataRow[27] = (wsraw.Range["AI" + i].Value < 0 )? "1": wsraw.Range["AI" + i].Value; // Initial Sum at Risk
                    dtDataRow[77] = (wsraw.Range["AI" + i].Value < 0 )? "1": wsraw.Range["AI" + i].Value; // Sum at Risk

                    dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDouble(wsraw.Range["AI" + i].Value));
                    dblTotalPremium = dblTotalPremium + (Convert.ToDouble(wsraw.Range["AK" + i].Value));


                    //ADDB RIDER
                    if (wsraw.Range["AC" + i].Value != 0)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        strPolno = wsraw.Range["A" + i].Value;
                        strPolicyNumber = strPolno.ToString();
                        strGender = objHlpr.fn_getgenderv3(wsraw.Range ["E" + i].Value); // Gender
                        if (strGender.ToUpper().Contains("FEMALE"))
                        {
                            dtDataRow[36] = "F";
                        }
                        else
                        {
                            dtDataRow[36] = "M";
                        }

                        strfullname = Convert.ToString(wsraw.Range["C" + i].Value);
                        if(Regex.IsMatch(strfullname, @"^\d+$") || Regex.IsMatch(strfullname, @"PHP\d|\d") || strfullname.ToUpper().Contains(";"))
                        {
                            dtDataRow[31] = "DummyLastName" + ", " + "DummyFirstName" + " " + "DummyMiddleName" + "."; // Full name
                            dtDataRow[32] = "DummyLastName"; // Last Name
                            dtDataRow[33] = "DummyFirstName"; // First Name
                            dtDataRow[34] = "DummyMiddleName"; // Middle Initials
                            dtDataRow[30] = "DUMMYDU07011900"; // Life ID
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + "DD07011900"; // Policy Number
                        }
                        else
                        {
                            objHlpr.fn_separatefullnamev3(strfullname, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow[32] = strLastName; // Last Name
                            strFirstName = objHlpr2.fn_checkFirstname(strFirstName);
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow[34] = strMiddleInitial; // Middle Initials
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            string strBirthday1 = strBirthday.Replace("/", "");
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday1; // PolicyNumber
                        }
                        //objHlpr.fn_separatefullnamev3(wsraw.Range["E" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        //dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        //dtDataRow[32] = strLastName; // Last Name
                        //dtDataRow[33] = strFirstName; // First Name
                        //dtDataRow[34] = strMiddleInitial; // Middle Initials
                        //strBirthday = Convert.ToString(wsraw.Range["F" + i].Value).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        //dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        //strBirthday = strBirthday.Replace("/", "");
                        //dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                        dtDataRow[78] = wsraw.Range["G" + i].Value; // Life Issue Age
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "GRP"; // Class of Business
                        dtDataRow[23] = "PHP"; // Cession Currency
                        dtDataRow[24] = "YLY"; // Premium Frequency
                        dtDataRow[38] = "NONE"; // Smoker Status
                        dtDataRow[5] = "ADDB"; // Plan Code
                        dtDataRow[41] = Variables.strBmYear; // Policy Year
                        dtDataRow[82] = wsraw.Range["B" + i].Value; // Group Policy Holder
                        dtDataRow[7] = wsraw.Range["A" + i].Value; // Group Scheme ID
                        dtDataRow[76] = wsraw.Range["K" + i].Value; // Remarks
                        dtDataRow[39] = strPreferredClassific; // Preferred Classific)
                        dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["Q" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date

                        if (wsraw.Range["T" + i].Value == null)
                        {
                            dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["S" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                        }
                        else
                        {
                            dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                        }

                        //objHlpr.fn_CheckTransCode(wsraw.Range["J" + i].Value, out string transcode);
                        dtDataRow[21] = transcode; // Transcode

                        strBusinessType = wsraw.Range["D" + i].Value; // Business Type
                        if (strBusinessType.ToUpper().Contains("STANDARD"))
                        {
                            dtDataRow[14] = "T";
                        }
                        else
                        {
                            dtDataRow[14] = "F";
                        }

                        if (transcode == "TNEWBUS")
                        {
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AN" + i].Value; // Premium
                        }
                        else
                        {
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AN" + i].Value; // Premium
                        }
                        
                        dtDataRow[28] = (wsraw.Range["AF" + i].Value < 0) ? "1" : wsraw.Range["AF" + i].Value; // Cedent Retention
                        dtDataRow[25] = (wsraw.Range["AC" + i].Value < 0) ? "1" : wsraw.Range["AC" + i].Value; // Original Sum Assured
                        dtDataRow[27] = (wsraw.Range["AL" + i].Value < 0) ? "1" : wsraw.Range["AL" + i].Value; // Initial Sum at Risk
                        dtDataRow[77] = (wsraw.Range["AL" + i].Value < 0) ? "1" : wsraw.Range["AL" + i].Value; // Sum at Risk

                        dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDouble(wsraw.Range["AL" + i].Value));
                        dblTotalPremium = dblTotalPremium + (Convert.ToDouble(wsraw.Range["AN" + i].Value));

                        
                    }
                    // TPD RIDER
                    if (wsraw.Range["AD" + i].Value != 0)
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        strPolno = wsraw.Range["A" + i].Value;
                        strPolicyNumber = strPolno.ToString();
                        strGender = objHlpr.fn_getgenderv3(wsraw.Range ["E" + i].Value); // Gender
                        if (strGender.ToUpper().Contains("FEMALE"))
                        {
                            dtDataRow[36] = "F";
                        }
                        else
                        {
                            dtDataRow[36] = "M";
                        }
                        strfullname = Convert.ToString(wsraw.Range["C" + i].Value);
                        Console.WriteLine(strfullname);
                        if(Regex.IsMatch(strfullname, @"^\d+$") || Regex.IsMatch(strfullname, @"PHP\d|\d") || strfullname.ToUpper().Contains(";"))
                        {
                            dtDataRow[31] = "DummyLastName" + ", " + "DummyFirstName" + " " + "DummyMiddleName" + "."; // Full name
                            dtDataRow[32] = "DummyLastName"; // Last Name
                            dtDataRow[33] = "DummyFirstName"; // First Name
                            dtDataRow[34] = "DummyMiddleName"; // Middle Initials
                            dtDataRow[30] = "DUMMYDU07011900"; // Life ID
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + "DD07011900"; // Policy Number
                        }
                        else
                        {
                            objHlpr.fn_separatefullnamev3(strfullname, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";   
                            dtDataRow[32] = strLastName; // Last Name
                            strFirstName = objHlpr2.fn_checkFirstname(strFirstName);
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow [34] = strMiddleInitial; // Middle Initials
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            string strBirthday1 = strBirthday.Replace("/", "");
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday1; // PolicyNumber
                        }
                        //objHlpr.fn_separatefullnamev3(wsraw.Range["E" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        //dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        //dtDataRow[32] = strLastName; // Last Name
                        //dtDataRow[33] = strFirstName; // First Name
                        //dtDataRow[34] = strMiddleInitial; // Middle Initials
                        //strBirthday = Convert.ToString(wsraw.Range["F" + i].Value).ToString("MM/dd/yyyy");
                        dtDataRow[37] = strBirthday; // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                                                   //dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                                   //strBirthday = strBirthday.Replace("/", "");
                                                   //dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                        dtDataRow[78] = wsraw.Range["G" + i].Value; // Attained Age
                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "GRP"; // Class of Business
                        dtDataRow[23] = "PHP"; // Cession Currency
                        dtDataRow[24] = "YLY"; // Premium Frequency
                        dtDataRow[38] = "NONE"; // Smoker Status
                        dtDataRow[5] = "TPD"; // Plan Code
                        dtDataRow[41] = Variables.strBmYear; // Policy Year
                        dtDataRow[82] = wsraw.Range["B" + i].Value; // Group Policy Holder
                        dtDataRow[7] = wsraw.Range["A" + i].Value; // Group Scheme ID
                        dtDataRow[76] = wsraw.Range["K" + i].Value; // Remarks
                        dtDataRow[39] = strPreferredClassific; // Preferred Classific)
                        dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                        dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["Q" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                        dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["R" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date

                        if (wsraw.Range["T" + i].Value == null)
                        {
                            dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["S" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                        }
                        else
                        {
                            dtDataRow[40] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["T" + i].Value)).ToString("MM/dd/yyyy"); ; // Risk Expiry Date
                        }

                        //objHlpr.fn_CheckTransCode(wsraw.Range["J" + i].Value, out string transcode);
                        dtDataRow[21] = transcode; // Transcode

                        strBusinessType = wsraw.Range["D" + i].Value; // Business Type
                        if (strBusinessType.ToUpper().Contains("STANDARD"))
                        {
                            dtDataRow[14] = "T";
                        }
                        else
                        {
                            dtDataRow[14] = "F";
                        }

                        if (transcode == "TNEWBUS")
                        {
                            dtDataRow[56] = "4000"; // Entry Code
                            dtDataRow[57] = wsraw.Range["AQ" + i].Value; // Premium
                        }
                        else
                        {
                            dtDataRow[58] = "4001"; // Entry Code
                            dtDataRow[59] = wsraw.Range["AQ" + i].Value; // Premium
                        }

                        dtDataRow[28] = (wsraw.Range["AG" + i].Value < 0) ? "1" : wsraw.Range["AG" + i].Value; // Cedent Retention
                        dtDataRow[25] = (wsraw.Range["AD" + i].Value < 0) ? "1" : wsraw.Range["AD" + i].Value; // Original Sum Assured
                        dtDataRow[27] = (wsraw.Range["AO" + i].Value < 0) ? "1" : wsraw.Range["AO" + i].Value; // Initial Sum at Risk
                        dtDataRow[77] = (wsraw.Range["AO" + i].Value < 0) ? "1" : wsraw.Range["AO" + i].Value; // Sum at Risk

                        dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDouble(wsraw.Range["AO" + i].Value));
                        dblTotalPremium = dblTotalPremium + (Convert.ToDouble(wsraw.Range["AQ" + i].Value));
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 026", "Information");
                return "";
            }

            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);
            #endregion



            string despath = str_saved + @"\BM026-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
