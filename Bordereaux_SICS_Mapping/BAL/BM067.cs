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
    class BM067
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

            string strFilePath = wbraw.Path;
            string strRemarksAABBZ = string.Empty;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();


            }

            DataRow dtDataRow;
            string valueTransEffectiveDate = string.Empty; string strTcode = string.Empty;
            decimal dclTotalPremium = 0; decimal dclTotalSumAtRisk = 0;

            if (str_sheet.ToUpper() == "BASIC" || str_sheet.ToUpper() == "MAJOR CC" || str_sheet.ToUpper() == "MINOR" || str_sheet.ToUpper() == "CI")
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 1].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text.ToString(), wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [23] = wsraw.Cells [intLoop, 19].Value; //  Cession Currency
                    dtDataRow [24] = "MLY"; // Premium Frequency
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow [5] = wsraw.Cells [intLoop, 5].Value; //branded product
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PA"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [13] = "IND"; // Class of Business
                    dtDataRow [14] = "T"; // Business Type
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    string strSmoker = wsraw.Cells [intLoop, 12].Value;
                    dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker);//SMOKER
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells [intLoop, 3].Value));
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells [intLoop, 2].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells [intLoop, 10].Value)).ToString("MM/dd/yyyy");
                    string strSex = Convert.ToString(wsraw.Cells [intLoop, 11].Value); // Gender
                    dtDataRow [33] = strFirstName;
                    dtDataRow [32] = strLastName;
                    dtDataRow [31] = strFullName;//Full Name
                    dtDataRow [37] = strDOB; // Birthday
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow [36] = strSex;
                    dtDataRow [79] = Convert.ToString(wsraw.Cells [intLoop, 13].Value); //Issue Age
                    //objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells [intLoop, 33].Value), Convert.ToString(wsraw.Cells [intLoop, 33].Value), Convert.ToString(wsraw.Cells [intLoop, 33].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                   
                    //decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 34].Value);//Premium
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 38].Value), out decimal dclRewewaYear);//Renewal Year
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 40].Value),out decimal dclFirstYear);//First Year
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 42].Value), out decimal dclComission);//Comission
                    //dtDataRow [21] = strTcode; // Transcode
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells [intLoop, 9].Value)).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = valueTransEffectiveDate;//Policy Start Date

                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 26].Text), out decimal dclOrigSum);//
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 33].Text), out decimal dclSumAtRisk);//Sum at Risk and initial
                    dclTotalPremium += dclRewewaYear + dclFirstYear;
                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;

                    if(dclFirstYear != 0)
                    {
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = "4000";
                        dtDataRow [57] = dclFirstYear;
                        dtDataRow [64] = "5004";
                        dtDataRow [65] = dclComission;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                        dtDataRow [26] = dclOrigSum;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                        dclTotalSumAtRisk += dclSumAtRisk;

                    }
                    if(dclRewewaYear != 0)
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [59] = dclRewewaYear;
                        dtDataRow [58] = "4001";
                        dtDataRow [66] = "5005";
                        dtDataRow [67] = dclComission;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                        dtDataRow [26] = dclOrigSum;
                        dclTotalSumAtRisk += dclSumAtRisk;
                    }
                    if (dclFirstYear == 0 && dclRewewaYear == 0)
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [66] = "5005";
                        dtDataRow [67] = dclComission;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                        dtDataRow [26] = dclOrigSum;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        dtDataRow [59] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk)); ;
                        dtDataRow [58] = "4001";
                    }
                }
               
            }
            else
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 1].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text.ToString(), wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [23] = "PHP"; //  Cession Currency
                    dtDataRow [24] = "MLY"; // Premium Frequency
                    dtDataRow [9] = "PA"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [13] = "IND"; // Class of Business
                    dtDataRow [14] = "T"; // Business Type
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow [5] = wsraw.Cells [intLoop, 5].Value; //branded product
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [38] = objHlpr.fn_SmokerCode("");//SMOKER
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(null); // mortality
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells [intLoop, 2].Value));
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells [intLoop, 3].Value));
                    string strFullName = strLastName + " " + strFirstName;
                    string strDOB = objHlpr2.fn_checkDOB("");
                    dtDataRow [37] = strDOB;// Birthday
                    dtDataRow [33] = strFirstName;
                    dtDataRow [32] = strLastName;
                    dtDataRow [31] = strFullName;//Full Name
                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                    dtDataRow [36] = strSex;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    string strIssueDate = objHlpr.fn_convertStringtoDate(Convert.ToString(wsraw.Cells [intLoop, 6].Value)).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = valueTransEffectiveDate;//Policy Start Date

                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value), out decimal dclPremium);//Renewal Year
                    //decimal dclFirstYear = Convert.ToDecimal(wsraw.Cells [intLoop, 40].Value);//First Year   
                    //dtDataRow [21] = strTcode; // Transcode

                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Text), out decimal dclOrigSum);//
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 33].Text), out decimal dclSumAtRisk);//Sum at Risk and initial
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 17].Value), out decimal dclComission);//Comission
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                    dtDataRow [26] = 1;
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                    dtDataRow [62] = "4004";
                    dtDataRow [63] = dclPremium;
                    dtDataRow [66] = "5005";
                    dtDataRow [67] = dclComission;
                    dclTotalPremium += dclPremium;
                    dclTotalSumAtRisk += dclSumAtRisk;

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


            string despath = str_saved + @"\BM067-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalSumAtRisk = 0;
            dclTotalPremium = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}