using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM066
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);


            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int intLastRow = wsraw.Range["A1"].End[XlDirection.xlDown].Row;
            string strFilePath = wbraw.Path;

            /*string strPolicyYear = strFilePath.Substring(strFilePath.Length - 6);
            strPolicyYear = strPolicyYear.Insert(2, "/");
            DateTime PolicyYear = DateTime.ParseExact(strPolicyYear, "MM/yyyy", CultureInfo.InvariantCulture);
            strPolicyYear = PolicyYear.ToString("MM/yyyy");*/
            
            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;
            string valueTransEffectiveDate = string.Empty; string strTcode = string.Empty;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }
            //LIFE - COMPLETED
            if (str_sheet == "Life")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }
                   
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["Z" + i].Value), Convert.ToString(wsraw.Range["AG" + i].Value), Convert.ToString(wsraw.Range["AG" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    
                    string strGender = wsraw.Range["K" + i].Value;
                    dtDataRow[36] = strGender; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + "  " + strFirstName;
                    dtDataRow[31] = strFullName;
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product

                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["J" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthday
                   
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow[23] = "PHP"; //  Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strSmoker = wsraw.Range["L" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCodeV2(strSmoker.ToUpper()); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow [79] = wsraw.Range ["M" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");
                    decimal.TryParse(Convert.ToString(wsraw.Range["AG" + i].Value), out decimal sumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AL" + i].Value), out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AN" + i].Value), out decimal premiumFY);

                    dblTotalPremium += premiumRY + premiumFY;

                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strGender, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|"+ strRemarksCode; // Remarks
                    
                }
               
            }

            //LIFE  - GIO COMPLETED
            else if (str_sheet == "Life - GIO")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                  
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["X" + i].Value), Convert.ToString(wsraw.Range["AH" + i].Value), Convert.ToString(wsraw.Range["AH" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    string strSex = wsraw.Range["K" + i].Value;
                    dtDataRow[36] = strSex; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["J" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; //Reinsurance Product
                    dtDataRow[9] = "PA"; //Type of Business
                    dtDataRow[10] = "Q"; //Reinsurance Methods
                    dtDataRow[13] = "IND"; //Class of Business
                    dtDataRow[14] = "T"; //Business Type
                    dtDataRow[24] = "YLY"; //Premium Frequency
                    string strSmoker = wsraw.Range["L" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCodeV2(strSmoker.ToUpper()); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow [79] = wsraw.Range ["M" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");

                    decimal.TryParse(Convert.ToString(wsraw.Range ["AH" + i].Value),out decimal sumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AM" + i].Value), out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AO" + i].Value), out decimal premiumFY);

                    dblTotalPremium += premiumRY + premiumFY;

                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dblTotalSumAtRisk += sumAtRisk;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dblTotalSumAtRisk += sumAtRisk;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumFY; // FY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                    }
                 
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                   

                
                   
                }
            }

            //HEALTH - COMPLETED
            else if (str_sheet == "Health")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range ["X" + i].Value), Convert.ToString(wsraw.Range ["AC" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    string strGender = Convert.ToString(wsraw.Range["L" + i].Value);
                    dtDataRow[36] = strGender; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthdays
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow[23] = wsraw.Range["T" + i].Value; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strSmoker = wsraw.Range["M" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCodeV2(strSmoker); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow [79] = wsraw.Range ["N" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");

                    decimal.TryParse(Convert.ToString(wsraw.Range ["X" + i].Value), out decimal origSum);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AC" + i].Value), out decimal sumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AH" + i].Value), out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AJ" + i].Value), out decimal premiumFY);
                    dblTotalPremium += premiumRY + premiumFY;
                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dblTotalSumAtRisk += sumAtRisk;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;//ceded retention
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dblTotalSumAtRisk += sumAtRisk;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumFY; //FY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;//ceded retention
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;//ceded retention
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = valueTransEffectiveDate;//Policy Start Date

                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strGender, out string strRemarksCode);
                    dtDataRow [76] = strRemarksAABBZ + " |" + strRemarksCode; // Remarks


                }
            }

            //ACCIDENT - for update
            else if (str_sheet == "Accident")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range ["Y" + i].Value), Convert.ToString(wsraw.Range ["AH" + i].Value), Convert.ToString(wsraw.Range ["AH" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strSex = wsraw.Range["L" + i].Value;
                    dtDataRow[36] = strSex; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + ", " + strFirstName;
                    dtDataRow[31] = strFullName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthday
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strSmoker = wsraw.Range["M" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCodeV2(strSmoker); //Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow [79] = wsraw.Range ["N" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AH" + i].Value), out decimal sumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AM" + i].Value), out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AO" + i].Value), out decimal premiumFY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["Y" + i].Value), out decimal origSum);
                    dblTotalPremium += premiumRY + premiumFY;
                  
                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dblTotalSumAtRisk += sumAtRisk;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumFY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = valueTransEffectiveDate;//Policy Start Date
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + " |" + strRemarksCode; // Remarks
                }
            }

            //CI - COMPLETED
            else if (str_sheet == "CI")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["Y" + i].Value), Convert.ToString(wsraw.Range["AH" + i].Value), Convert.ToString(wsraw.Range["AH" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strSex = wsraw.Range["L" + i].Value;
                    dtDataRow[36] = strSex; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; //Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;//First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + "  " + strFirstName;
                    dtDataRow[31] = strFullName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; //Birthday
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP";//Currency
                    dtDataRow[41] = Variables.strBmYear; //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strSmoker = wsraw.Range["M" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCodeV2(strSmoker); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow [79] = wsraw.Range ["N" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");

                    decimal.TryParse(strSumAtRisk, out decimal sumAtRisk); //sumAtRisk
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AM" + i].Value),out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AO" + i].Value),out decimal premiumFY);
                    dblTotalPremium += premiumRY + premiumFY;
                    
                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumFY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 0;
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;//REINSURANCE START DATE

                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + strRemarksCode;// Remarks


                }
            }
            //WP 
            else if(str_sheet == "WP")
            {
                for(int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range ["A" + i].Value; // Policy Number
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range ["B" + i].Value, wsraw.Range ["C" + i].Value, wsraw.Range ["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range ["AH" + i].Value), Convert.ToString(wsraw.Range ["AQ" + i].Value), Convert.ToString(wsraw.Range ["AQ" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    string strSex = wsraw.Range ["L" + i].Value;
                    dtDataRow [36] = strSex; // Gender
                    dtDataRow [5] = wsraw.Range ["E" + i].Value; //branded product
                    string strLastName = wsraw.Range ["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range ["C" + i].Value;// First Name                      
                    dtDataRow [33] = strFirstName;
                    dtDataRow [32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow [31] = strFullName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow [37] = strDOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [23] = "PHP"; //Currency
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [8] = "QA"; // Reinsurance Product
                    dtDataRow [9] = "PA"; // Type of Business
                    dtDataRow [10] = "Q"; // Reinsurance Methods
                    dtDataRow [13] = "IND"; // Class of Business
                    dtDataRow [14] = "T"; // Business Type
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    string strSmoker = wsraw.Range ["M" + i].Value;
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(strSmoker); // Smoker Status
                    dtDataRow [79] = wsraw.Range ["N" + i].Value; // Issue Age
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");

                    decimal.TryParse(Convert.ToString(wsraw.Range ["AL" + i].Value), out decimal sumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AV" + i].Value), out decimal premiumRY);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["AX" + i].Value), out decimal premiumFY);

                    dblTotalPremium += premiumRY + premiumFY;
                 
                    if(premiumRY > 0 && premiumFY == 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else if(premiumRY == 0 && premiumFY > 0)
                    {
                        dblTotalSumAtRisk += sumAtRisk;
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = "4000"; // Entry code
                        dtDataRow [57] = premiumRY; // RY  
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = "4001"; // Entry code
                        dtDataRow [59] = premiumRY; // RY
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                        dtDataRow [26] = 1;
                    }
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow [19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date

                }
            }



            //"Life Past Accts" 
            else if (str_sheet == "Life Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, null, Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID 
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear;//Policy Year
                    dtDataRow[21] = strTcode; // Transcode
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal.TryParse(Convert.ToString(wsraw.Range["M" + i].Value), out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range["G" + i].Value), out decimal dclSAR);
                    dclPremium = dclPremium * -1;

                    dtDataRow [62] = "4004"; // Entry code
                    dtDataRow [63] = dclPremium; // RY Adjustments
                    dblTotalPremium += dclPremium;
                    if (dclPremium != 0)
                    {
                        dblTotalSumAtRisk += dclSAR;
                    }
                  
                }
            }

            else if (str_sheet == "Life GIO Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range["G" + i].Value), Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + "  " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex;// Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; //Policy Year
                    dtDataRow[21] = "TNEWBUS"; // Transcode
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow[62] = "4004"; // Entry code
                    dtDataRow[63] = Convert.ToString(wsraw.Range["M" + i].Value); // RY Adjustments
                    decimal.TryParse(Convert.ToString(wsraw.Range["M" + i].Value), out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range["G" + i].Value), out decimal dclSAR);
                    dclPremium = dclPremium * -1;
                    dblTotalPremium += dclPremium;
                    if(dclPremium != 0)
                    {
                        dblTotalSumAtRisk += dclSAR;
                    }
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                }
            }

            //HEALTH PAST ACCTS 
            else if (str_sheet == "Health Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range["G" + i].Value), Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal.TryParse(Convert.ToString(wsraw.Range["J" + i].Value), out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range["G" + i].Value), out decimal dclSAR);
                    dclPremium = dclPremium * -1;
                    dblTotalPremium += dclPremium;
                    if(dclPremium != 0)
                    {
                        dblTotalSumAtRisk += dclSAR;
                    }
                    dtDataRow [62] = "4004";
                    dtDataRow[63] = dclPremium;
                }
            }

            //Accident Past Accts
            else if (str_sheet == "Accident Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range["G" + i].Value), Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); //Life ID
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal.TryParse(Convert.ToString(wsraw.Range["J" + i].Value),out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range["G" + i].Value), out decimal dclSAR);
                    dclPremium = dclPremium * -1;
                    dtDataRow [62] = "4004";
                    dtDataRow[63] = dclPremium;
                    dblTotalPremium += dclPremium;
                    dblTotalSumAtRisk += dclSAR;

                }
            }

            //WP PAST ACCTS
            else if (str_sheet == "WP Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range["G" + i].Value), Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks
                    decimal.TryParse(Convert.ToString(wsraw.Range["M" + i].Value), out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out decimal dclSAR);
                    dclPremium = dclPremium * -1;
                    dtDataRow [62] = "4004";
                    dtDataRow[63] = dclPremium;
                    dblTotalPremium += dclPremium;
                    if(dclPremium != 0)
                    {
                        dblTotalSumAtRisk += dclSAR;
                    }
                }
            }

            //CI Past Accts
            else if (str_sheet == "CI Past Accts")
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range["G" + i].Value), Convert.ToString(wsraw.Range["G" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    //objHlpr.fn_osabreakdown(Convert.ToDouble(strSumAtRisk), Convert.ToDouble(strInitialSum), Convert.ToDouble(strOriginalSum), out string strOrignalSum);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // ceded sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[5] = wsraw.Range["E" + i].Value; //branded product
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strFullName = strLastName + " " + strFirstName;
                    dtDataRow[31] = strFullName;
                    objHlpr.fn_Getfirstname(strFirstName, out strFirstName);
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB;
                    strTcode = "ADJUST"; //Transcode
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); //Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = objHlpr.fn_SmokerCodeV2(null); // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal.TryParse(Convert.ToString(wsraw.Range["M" + i].Value), out decimal dclPremium);
                    decimal.TryParse(Convert.ToString(wsraw.Range["G" + i].Value), out decimal dclSAR);
                    dtDataRow[62] = "4004";
                    dtDataRow[63] = dclPremium;
                    dclPremium = dclPremium * -1;
                    dblTotalPremium += dclPremium;
                    if(dclPremium != 0)
                    {
                        dblTotalSumAtRisk += dclSAR;
                    }

                }
            }

            #region
            /*else if (str_sheet.ToUpper().Contains("LIFE PAST ACCTS"))
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    string strPolicyNo = wsraw.Range["A" + i].Value; // Policy Number
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range["B" + i].Value, wsraw.Range["C" + i].Value, wsraw.Range["D" + i].Value))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(null, null, Convert.ToString(wsraw.Range["J" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk);
                    dtDataRow[36] = wsraw.Range["L" + i].Value; // Gender
                    string strLastName = wsraw.Range["B" + i].Value; // Last Name
                    string strFirstName = wsraw.Range["C" + i].Value;// First Name                      
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[31] = strLastName + ", " + strFirstName;
                    dtDataRow[34] = ""; // Middle Initials
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    string strBirthday = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["K" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strBirthday; // Birthday
                    string strPolicyStartDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[20] = strPolicyStartDate;
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    dtDataRow[23] = wsraw.Range["T" + i].Value; //  Currency
                    string strRemarks = wsraw.Range["D" + i].Value + "_" + wsraw.Range["E" + i].Value;
                    dtDataRow[76] = strRemarks.Trim(); // Remarks
                    //dtDataRow[5] = wsraw.Range["Q" + i].Value; // Branded Product Cedent Code
                    dtDataRow[41] = str_policyYear;/*PolicyYear.ToString("MM/yyyy"); //Policy Year
                    dtDataRow[21] = "TNEWBUS"; // Transcode
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strSmoker = wsraw.Range["L" + i].Value;
                    dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                    dtDataRow[56] = "4000"; // Entry Code
                    dtDataRow[57] = wsraw.Range["AN" + i].Value; // FY
                    dtDataRow[58] = "4004"; // Entry code
                    dtDataRow[59] = wsraw.Range["AL" + i].Value; // RY
                    dtDataRow[19] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                    dtDataRow[22] = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date
                                                                                                                                         //dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["M" + i].Value);
                    dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["AJ" + i].Value);
                    dblTotalSumAtRisk = dblTotalSumAtRisk + objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);

                }
            }*/
            #endregion

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


            string despath = str_saved + @"\BM066-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dblTotalPremium = 0;
            dblTotalSumAtRisk = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}
