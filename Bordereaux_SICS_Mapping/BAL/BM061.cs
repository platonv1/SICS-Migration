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
    class BM061
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false, string str_policyYear = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);


            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets [str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            string strFilePath = wbraw.Path;
            int erawrow = rawrange.Rows.Count;

            decimal dclTotalCommission = 0;
            decimal dclTotalPremium = 0;
            decimal dclTotalSumAtRisk = 0;
            string valueTransEffectiveDate = string.Empty;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }

            DataRow dtDataRow;
            if(str_sheet.ToUpper().Contains("SHEET"))
            {
                for(int intLoop = 2; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    string strCurrency = objHlpr2.fn_getcurrency(Convert.ToString(wsraw.Cells [intLoop, 4].Value));
                    string strCOB = objHlpr2.fn_getcob(Convert.ToString(wsraw.Cells [intLoop, 3].Value));
                    string strRisk = (Convert.ToString(wsraw.Cells [intLoop, 13].Value));
                    dtDataRow [23] = strCurrency; //  Cession Currency
                    dtDataRow [13] = strCOB; // Class of Business
                    dtDataRow [4] = objHlpr2.fn_gettransactionproduct(Convert.ToString(wsraw.Cells [intLoop, 3].Value), Convert.ToString(wsraw.Cells [intLoop, 4].Value), strRisk);
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells [intLoop, 18].Value), Convert.ToString(wsraw.Cells [intLoop, 19].Value), Convert.ToString(wsraw.Cells [intLoop, 19].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    string strTcode = "TRENEW";
                    dtDataRow [21] = strTcode; // Transcode
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [14] = wsraw.Cells [intLoop, 8].Value;//Business Type
                    string strSmoker = wsraw.Cells [intLoop, 11].Value;// Smoker Status
                    dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker);
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells [intLoop, 16].Value));
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    string strGender = wsraw.Cells [intLoop, 10].Value;
                    dtDataRow [36] = strGender; //Gender
                    string strFullName = wsraw.Cells [intLoop, 5].Value;
                    dtDataRow [31] = strFullName;
                    objHlpr2.fn_seperateforeignamesV2(strFullName, out string strFirstName, out string strLastName, out string strMI);
                    strLastName = objHlpr2.fn_checkLastname(strLastName);
                    dtDataRow [32] = strLastName;
                    strFirstName = objHlpr2.fn_checkFirstname(strFirstName);
                    dtDataRow [33] = strFirstName;
                    dtDataRow [34] = strMI;
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); //orignal sum risk
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); //RIS
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                    string strIssueDate = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Cells [intLoop, 6].Value));
                    dtDataRow [22] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = strIssueDate;//Policy Start Date
                    dtDataRow [79] = Convert.ToString(wsraw.Cells [intLoop, 7].Value);//LIFE ISSUE AGE

                    string strBirthday = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Cells [intLoop, 9].Value));
                    dtDataRow [37] = strBirthday; // Birthday
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); //Life ID
                    objHlpr.fn_GetRemarksCode(strBirthday, strFullName, strGender, out string strRemarksCode);
                    dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                    decimal dclComissionX = Convert.ToDecimal(wsraw.Cells [intLoop, 24].Value);
                    decimal dclComissionZ = Convert.ToDecimal(wsraw.Cells [intLoop, 26].Value);
                    decimal dclComissionAB = Convert.ToDecimal(wsraw.Cells [intLoop, 28].Value);
                    decimal dclPremiumT = Convert.ToDecimal(wsraw.Cells [intLoop, 20].Value);
                    decimal dclPremiumV = Convert.ToDecimal(wsraw.Cells [intLoop, 22].Value);
                    decimal dclSumAtRisk = Convert.ToDecimal(strSumAtRisk);

                    dtDataRow [58] = "4001"; // Entry code
                    dtDataRow [59] = dclPremiumT + dclPremiumV;
                    dtDataRow [66] = "5005";
                    dtDataRow [67] = dclComissionX + dclComissionZ;

                    dclTotalCommission += dclComissionX + dclComissionZ + dclComissionAB;
                    dclTotalPremium += dclPremiumT + dclPremiumV;
                    dclTotalSumAtRisk += dclSumAtRisk;


                }
            }


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dclTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Commission:";
            dtDataRow [1] = dclTotalCommission;
            objdt_template.Rows.Add(dtDataRow);
            #endregion


            string despath = str_saved + @"\BM061-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSumAtRisk = 0;
            dclTotalCommission = 0;

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";




        }
    }

}