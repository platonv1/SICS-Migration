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
    class BM063
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
            string valueTransEffectiveDate = string.Empty;
            decimal dclTotalPremium = 0;
            decimal dclTotalSumAtRisk = 0;
            decimal premium = 0;
            decimal sumAtRisk = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }
            string strTcode = string.Empty;
            string strPremiumYear = string.Empty;
            DataRow dtDataRow;

            if (str_sheet == "Premiums")
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
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 14].Value), Convert.ToString(wsraw.Cells[intLoop, 15].Value), Convert.ToString(wsraw.Cells[intLoop, 15].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //OSR
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //ISR
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);//SAR
                    dtDataRow[8] = "QA"; //REINSURANCE PRODUCT
                    dtDataRow[9] = "PA"; //TYPE OF BUSINESS
                    dtDataRow[10] = "Q"; //REINSURANCE_METHODS
                    dtDataRow[13] = "IND"; //CLASS OF BUSINESS
                    dtDataRow[14] = "T"; //BUSINESS TYPE
                    dtDataRow[23] = "PHP"; //CESSION CURRENCY
                    dtDataRow[24] = "YLY"; //PREMIUM FREQUENCY
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[41] = Variables.strBmYear;//Policy Year
                    dtDataRow [6] = Convert.ToString(wsraw.Cells [intLoop, 23].Value);//Branded Product
                    string strFullName = Convert.ToString(wsraw.Cells[intLoop, 3].Value);
                    objHlpr2.fn_separateLastNameFirstNameV7(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                    //objHlpr.fn_separatefullname(strFullName, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    //objHlpr.fn_separatefullnamev8(strFullName, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    dtDataRow [31] = strFullName; //Full Name
                    dtDataRow[32] = strLastName;
                    dtDataRow[33] = strFirstName;
                    dtDataRow[34] = strMiddleInitial;
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB; //DOB
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, "07/01/1900"); //life ID 
                    objHlpr.fn_getbusinessTypeRefundingCode(Convert.ToString(wsraw.Cells[intLoop, 13].Value), out string strBusinessType, out string strRefundingCode);
                    dtDataRow[14] = strBusinessType;
                    //dtDataRow[83] = strRefundingCode;
                    dtDataRow[38] = objHlpr.fn_SmokerCode(null);
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null);
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    strPremiumYear = wsraw.Cells [intLoop, 10].Value;
                    //strTcode = "TNEWBUS";
                    //dtDataRow [21] = strTcode; // Transcode
                    //dtDataRow [56] = "4000";
                    //dtDataRow [57] = wsraw.Cells [intLoop, 16].Value; //PREMIUMS
                    if(strPremiumYear.ToUpper() == "FY")
                    {
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode; // Transcode
                        dtDataRow [56] = "4000";
                        dtDataRow [57] = wsraw.Cells [intLoop, 17].Value; //PREMIUMS    
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode; // Transcode
                        dtDataRow [58] = "4001";
                        dtDataRow [59] = wsraw.Cells [intLoop, 17].Value; //PREMIUMS
                    }


                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    dtDataRow [19] = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");  // Reinsurance Start Date

                    //dtDataRow [79] = objHlpr.fn_getIssueAge(strDOB, strIssueDate);//issue age
                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, str_gender, out string remarksCode);
                    dtDataRow [76] = remarksCode;
                    premium = Convert.ToDecimal(wsraw.Cells [intLoop, 16].Value);
                    sumAtRisk = Convert.ToDecimal(wsraw.Cells [intLoop, 15].Value);

                    #region HashTotal
                    dclTotalPremium += premium;
                    dclTotalSumAtRisk += sumAtRisk;
                    #endregion

                }
            }

            else if (str_sheet == "Refunds")
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
                    //objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 12].Value), Convert.ToString(wsraw.Cells[intLoop, 15].Value), Convert.ToString(wsraw.Cells[intLoop, 15].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //OSR
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ISR
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SAR
                    dtDataRow[8] = "QA"; //REINSURANCE PRODUCT
                    dtDataRow[9] = "PA"; //TYPE OF BUSINESS
                    dtDataRow[10] = "Q"; //REINSURANCE_METHODS
                    dtDataRow[13] = "IND"; //CLASS OF BUSINESS
                    dtDataRow[14] = "T"; //BUSINESS TYPE
                    dtDataRow[23] = "PHP"; //CESSION CURRENCY
                    dtDataRow[24] = "YLY"; //PREMIUM FREQUENCY
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow [6] = Convert.ToString(wsraw.Cells [intLoop, 24].Value);//Branded Product
                    dtDataRow [21] = strTcode; // Transcode
                    dtDataRow[41] = Variables.strBmYear;//Policy Year
                    string strFullName = Convert.ToString(wsraw.Cells[intLoop, 3].Value);
                    objHlpr2.fn_separateLastNameFirstNameV7(strFullName, out string strLastName, out string strFirstName ,out string strMiddleInitial);
                    dtDataRow[31] = strFullName; //Full Name
                    dtDataRow[32] = strLastName;
                    dtDataRow[33] = strFirstName;
                    dtDataRow[34] = strMiddleInitial;
                    string strDOB = objHlpr.fn_getDOB(null);
                    dtDataRow[37] = strDOB; //DOB
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow[36] = strSex; // Gender
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, "07/01/1900"); //life ID 

                    objHlpr.fn_getbusinessTypeRefundingCode(Convert.ToString(wsraw.Cells[intLoop, 13].Value), out string strBusinessType, out string strRefundingCode);
                    dtDataRow[14] = strBusinessType;
                    //dtDataRow[83] = strRefundingCode;
                    dtDataRow[38] = objHlpr.fn_SmokerCode(null);
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null);


                    objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                    strPremiumYear = wsraw.Cells [intLoop, 11].Value;
                    strTcode = "ADJUST";

                    if(strPremiumYear.ToUpper() == "FY")
                    {
                        dtDataRow [21] = strTcode; // Transcode
                        dtDataRow [60] = "4002";
                        dtDataRow [61] = wsraw.Cells [intLoop, 17].Value; //PREMIUMS    

                    }
                    else
                    {
                        dtDataRow [21] = strTcode; // Transcode
                        dtDataRow [62] = "4004";
                        dtDataRow [63] = wsraw.Cells [intLoop, 17].Value; //PREMIUMS
                    }

                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 10].Value).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = Convert.ToDateTime(wsraw.Cells [intLoop, 8].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    dtDataRow [19] = Convert.ToDateTime(wsraw.Cells [intLoop, 8].Value).ToString("MM/dd/yyyy");  // Reinsurance Start Date

                    premium = Convert.ToDecimal(wsraw.Cells [intLoop, 17].Value);
                    sumAtRisk = Convert.ToDecimal(wsraw.Cells [intLoop, 15].Value);
                    #region HashTotal
                    dclTotalPremium += premium;
                    dclTotalSumAtRisk += sumAtRisk;
                    #endregion

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
            dtDataRow[1] = dclTotalSumAtRisk;
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

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM063-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSumAtRisk = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}