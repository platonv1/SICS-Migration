using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;
using Bordereaux_SICS_Mapping.BAL;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM064
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

            //int intLastRow = wsraw.Range["B12"].End[XlDirection.xlDown].Row;
            int erawrow = rawrange.Rows.Count;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();


            }

            DataRow dtDataRow;
            string valueTransEffectiveDate = string.Empty;
            decimal dclTotalPremium = 0, dclTotalSAR = 0;

            if (str_sheet == "Life" || str_sheet == "Accident")
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 1].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    string strBrandedProduct = wsraw.Cells[intLoop, 5].Value;
                    dtDataRow[5] = strBrandedProduct.ToUpper(); //Branded Product
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[29] = "NATREID"; //Life ID Type
                    string strTcode = "TNEWBUS";
                    dtDataRow[21] = strTcode; //Transcode
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 9].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[38] = objHlpr.fn_SmokerCode(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SMOKER
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 10].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[31] = strFullName;//Full Name
                    dtDataRow[37] = strDOB; //Birthday
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);//life ID 
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 11].Value); //Gender
                    dtDataRow[36] = strSex;
                    dtDataRow[79] = Convert.ToString(wsraw.Cells[intLoop, 13].Value);//Issue Age

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 26].Value), Convert.ToString(wsraw.Cells[intLoop, 33].Value), Convert.ToString(wsraw.Cells[intLoop, 33].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                    decimal dclSAR = Convert.ToDecimal(wsraw.Cells[intLoop, 33].Value);//Sum at Risk
                    decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 35].Value);//Premium
                    decimal dclRewewaYear = Convert.ToDecimal(wsraw.Cells[intLoop, 38].Value);//Renewal Year
                    decimal dclFirstYear = Convert.ToDecimal(wsraw.Cells[intLoop, 40].Value);//First Year

                    dtDataRow[56] = "4000";
                    dtDataRow[57] = dclFirstYear; 
                    dtDataRow[59] = dclRewewaYear;
                    dtDataRow[58] = "4001";
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum);//osa
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum);//isr
                    dtDataRow[26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 28].Value));//cededretention
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));

                    dclTotalPremium += dclFirstYear + dclRewewaYear;
                    dclTotalSAR += dclSAR;

                }
            }

            else if (str_sheet == "CI")
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 1].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    string strBrandedProduct = wsraw.Cells[intLoop, 5].Value;
                    dtDataRow[5] = strBrandedProduct.ToUpper(); //Branded Product
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[29] = "NATREID"; //Life ID Type
                    string strTcode = "TNEWBUS";
                    dtDataRow[21] = strTcode; //Transcode
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 9].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[38] = objHlpr.fn_SmokerCode(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SMOKER
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 10].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[31] = strFullName;//Full Name
                    dtDataRow[37] = strDOB; //Birthday
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);//life ID 
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 11].Value); //Gender
                    dtDataRow[36] = strSex;
                    dtDataRow[79] = Convert.ToString(wsraw.Cells[intLoop, 13].Value);//Issue Age

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 26].Value), Convert.ToString(wsraw.Cells[intLoop, 33].Value), Convert.ToString(wsraw.Cells[intLoop, 33].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                    decimal dclSAR = Convert.ToDecimal(wsraw.Cells[intLoop, 33].Value);//Sum at Risk
                    decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 35].Value);//Premium
                    decimal dclRewewaYear = Convert.ToDecimal(wsraw.Cells[intLoop, 38].Value);//Renewal Year
                    decimal dclFirstYear = Convert.ToDecimal(wsraw.Cells[intLoop, 40].Value);//First Year

                    dtDataRow[56] = "4000";
                    dtDataRow[57] = dclFirstYear;
                    dtDataRow[59] = dclRewewaYear;
                    dtDataRow[58] = "4001";
                    dtDataRow[26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 28].Value));//ceded
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum);//osa
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum);//isr
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));

                    dclTotalPremium += dclFirstYear + dclRewewaYear;
                    dclTotalSAR += dclSAR;

                }
            }


            else if (str_sheet == "WP")
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 1].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 11].Value); //Gender
                    dtDataRow[36] = strSex; // Gender
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 10].Value)).ToString("MM/dd/yyyy"); 
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[31] = strFullName;
                    dtDataRow[37] = strDOB;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    string strTcode = "TNEWBUS";
                    dtDataRow[21] = strTcode; //Transcode
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 9].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                    dtDataRow[38] = objHlpr.fn_SmokerCode(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SMOKER
                    dtDataRow[29] = "NATREID"; // Life ID TypedtDataRow[23] = "USD"; //  Cession Currency
                    string strBrandedProduct = wsraw.Cells[intLoop, 5].Value;
                    dtDataRow[5] = strBrandedProduct.ToUpper(); //Branded Product
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[29] = "NATREID"; //Life ID Type
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 35].Value), Convert.ToString(wsraw.Cells[intLoop, 42].Value), Convert.ToString(wsraw.Cells[intLoop, 42].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 37].Value));//ceded

                    decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 44].Value);//Premium
                    decimal dclRewewaYear = Convert.ToDecimal(wsraw.Cells[intLoop, 47].Value);//Renewal Year
                    decimal dclFirstYear = Convert.ToDecimal(wsraw.Cells[intLoop, 49].Value);//First Year
                   
                    dtDataRow[56] = "4000";
                    dtDataRow[57] = dclFirstYear;
                    dtDataRow[59] = dclRewewaYear;
                    dtDataRow[58] = "4001";

                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal dclSAR = Convert.ToDecimal(strSumAtRisk);


                    dclTotalPremium += dclFirstYear + dclRewewaYear;
                    dclTotalSAR += dclSAR;
                    
                }
            }

            else if (str_sheet == "Lives Prev Qtr" || str_sheet == "Accident Prev Qtr" || str_sheet == "CI Prev Qtr" || str_sheet == "WP Prev Qtr")
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 1].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;

                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[31] = strFullName;
                    dtDataRow[37] = strDOB;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB); // Life ID
                    objHlpr2.fn_getgenderv2(strFirstName, out string strSex);
                    dtDataRow[36] = strSex;// Gender
                    string strTcode = "TRENEW";
                    dtDataRow[21] = strTcode; //Transcode
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells[intLoop, 6].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode,Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;//REINSURANCE START DATE
                    dtDataRow [38] = objHlpr.fn_SmokerCode(null);//SMOKER
                    dtDataRow[29] = "NATREID"; // Life ID TypedtDataRow[23] = "USD"; //  Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    string strBrandedProduct = wsraw.Cells[intLoop, 5].Value;
                    dtDataRow[5] = strBrandedProduct.ToUpper(); //Branded Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[23] = "PHP"; //Currency
                    dtDataRow[29] = "NATREID"; //Life ID Type
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 7].Value), Convert.ToString(wsraw.Cells[intLoop, 7].Value), Convert.ToString(wsraw.Cells[intLoop, 7].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                    decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 10].Value);//Premium
                    dtDataRow[62] = "4004";
                    dtDataRow[63] = dclPremium;
                    decimal dclSAR = Convert.ToDecimal(strSumAtRisk);
                    dclTotalPremium += dclPremium;
                    dclTotalSAR += dclSAR;

                }
            }


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dclTotalSAR;
            objdt_template.Rows.Add(dtDataRow);
            #endregion

             dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if (Variables.boogenderfail)
            {
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Please check for blank genders";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }

            string despath = str_saved + @"\BM064-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSAR = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}