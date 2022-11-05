using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;
using Bordereaux_SICS_Mapping.BAL;
using System.Text.RegularExpressions;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM041_21
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
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets [str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            //int intLastRow = wsraw.Range["B12"].End[XlDirection.xlDown].Row;
            int erawrow = rawrange.Rows.Count;
            DataRow dtDataRow;

            string strFilePath = wbraw.Path;
            decimal dclTotalSumAtRisk = 0, dclTotalPremium = 0;
            decimal dclSumAtRisk = 0, dclPremium = 0;
            string strTcode = "";
            string strFyRy = "";
        

            if(str_sheet.ToUpper().Contains("NB") || str_sheet.ToUpper().Contains("REN"))
            {

                for(int intLoop = 3; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 1].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text.ToString(), wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = "'" + strPolicyNo.ToString().Trim(new char [0]);
                    string strFullName = wsraw.Cells [intLoop, 2].Value;
                    dtDataRow [31] = strFullName;
                    objHlpr2.fn_separateLastNameFirstNameV9(strFullName, out string strLastname, out string stFirstname, out string strMI);
                    dtDataRow [34] = strMI;
                    dtDataRow [33] = objHlpr2.fn_checkFirstname(stFirstname);
                    dtDataRow [32] = objHlpr2.fn_removeCharacters(strLastname);
                    string strDOB = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");
                    dtDataRow [37] = strDOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(stFirstname, strLastname, strDOB);// life ID 
                    dtDataRow [3] = objHlpr2.fn_benefitcover(Convert.ToString(wsraw.Cells [intLoop, 4].Value));//benefit cover
                    objHlpr2.fn_RemarksCode(strPolicyNo, Convert.ToString(wsraw.Cells [intLoop, 4].Value), Convert.ToString(wsraw.Cells [intLoop, 24].Value), out string insuredProd, out string remarksCode);//Insurance Prod
                    dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; //Branded Product
                    dtDataRow [4] = insuredProd;
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance method
                    dtDataRow [13] = "IND"; // Class of Business
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // life id type
                    dtDataRow [38] = objHlpr.fn_SmokerCode(wsraw.Cells [intLoop, 10].Value);// smoker status
                    dtDataRow [36] = wsraw.Cells [intLoop, 9].Value; //gender
                    dtDataRow [14] = objHlpr2.fn_businessTypeV2(Convert.ToString(wsraw.Cells [intLoop, 12].Value)); // Business Type
                    dtDataRow [23] = objHlpr2.fn_getcurrencyV2(Convert.ToString(wsraw.Cells [intLoop, 25].Value)); //  Cession Currency
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells [intLoop, 15].Text));// preffered classific
                    string bmYear = Convert.ToString(wsraw.Cells [intLoop, 19].Value);
                    dtDataRow [41] = bmYear; //Policy Year
                    dtDataRow [79] = wsraw.Cells [intLoop, 11].Value; //Life Issue age
                    dtDataRow [83] = objHlpr2.fn_businessType(Convert.ToString(wsraw.Cells [intLoop, 12].Value));//refunding code
               
                    dtDataRow [25] = 1; //original sum
                    dtDataRow [27] = wsraw.Cells [intLoop, 8].Value; //initial sum
                    dtDataRow [76] = remarksCode;//Remarks Code
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out dclSumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 23].Value), out dclPremium);
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                    dclTotalSumAtRisk += dclSumAtRisk;
                    dclTotalPremium += dclPremium;
                    strFyRy = wsraw.Cells [intLoop, 6].Value; //FYRY

                    if (strFyRy == "FY")
                    {
                        strTcode = "TNEWBUS";
                        dtDataRow [21] = strTcode;
                        dtDataRow [56] = 4000;
                        dtDataRow [57] = dclPremium;
                    }
                    else
                    {
                        strTcode = "TRENEW";
                        dtDataRow [21] = strTcode;
                        dtDataRow [58] = 4002;
                        dtDataRow [59] = dclPremium;
                    }

                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 14].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    objHlpr2.fn_getTransReinsuranceDateV2(strIssueDate, strTcode, bmYear, out string transEffectiveDate);
                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                    dtDataRow [20] = strIssueDate;//Policy Start Date
                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date

                    
                }
            }
            else if(str_sheet.ToUpper().Contains("ADJ"))
            {
                for(int intLoop = 3; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 1].Text.ToString();
                    Regex checkPolicy = new Regex(@"\d");
                    if(!checkPolicy.IsMatch(strPolicyNo))
                    {
                        continue;
                    }
                    //if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text.ToString(), wsraw.Cells [intLoop, 3].Text.ToString(), wsraw.Cells [intLoop, 4].Text.ToString()))
                    //{
                    //    continue;
                    //}
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = "'" + strPolicyNo.ToString().Trim(new char [0]);
                    string strFullName = wsraw.Cells [intLoop, 2].Value;
                    dtDataRow [31] = strFullName;
                    objHlpr2.fn_separateLastNameFirstNameV9(strFullName, out string strLastname, out string stFirstname, out string strMI);
                    dtDataRow [34] = strMI;
                    dtDataRow [33] = objHlpr2.fn_checkFirstname(stFirstname);
                    dtDataRow [32] = objHlpr2.fn_removeCharacters(strLastname);
                    string strDOB = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");
                    dtDataRow [37] = strDOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(stFirstname, strLastname, strDOB);// life ID 
                    objHlpr2.fn_RemarksCode(strPolicyNo, Convert.ToString(wsraw.Cells [intLoop, 4].Value), Convert.ToString(wsraw.Cells [intLoop, 24].Value), out string insuredProd, out string remarksCode);//Insurance Prod
                    dtDataRow [3] = objHlpr2.fn_benefitcover(Convert.ToString(wsraw.Cells [intLoop, 4].Value));//benefit cover
                    dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; //Branded Product
                    dtDataRow [4] = insuredProd;//insured prod
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance method
                    dtDataRow [13] = "IND"; // Class of Business
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // life id type
                    dtDataRow [38] = objHlpr.fn_SmokerCode(wsraw.Cells [intLoop, 10].Value);// smoker status
                    dtDataRow [36] = wsraw.Cells [intLoop, 9].Value; //gender
                    dtDataRow [14] = objHlpr2.fn_businessTypeV2(Convert.ToString(wsraw.Cells [intLoop, 12].Value)); // Business Type
                    dtDataRow [23] = objHlpr2.fn_getcurrencyV2(Convert.ToString(wsraw.Cells [intLoop, 25].Value)); //  Cession Currency
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells [intLoop, 15].Text));// preffered classific
                    dtDataRow [41] = Convert.ToString(wsraw.Cells [intLoop, 19].Value); //Policy Year
                    dtDataRow [79] = wsraw.Cells [intLoop, 11].Value; //Life Issue age
                    dtDataRow [83] = objHlpr2.fn_businessType(Convert.ToString(wsraw.Cells [intLoop, 12].Value));//refunding code
                    dtDataRow [76] = remarksCode;//Remarks Code

                    dtDataRow [22] = Convert.ToDateTime(wsraw.Cells [intLoop, 28].Value).ToString("MM/dd/yyyy"); // Trans Effective Date
                    dtDataRow [20] = Convert.ToDateTime(wsraw.Cells [intLoop, 14].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    dtDataRow [19] = Convert.ToDateTime(wsraw.Cells [intLoop, 28].Value).ToString("MM/dd/yyyy");  // Reinsurance Start Date
                    dtDataRow [25] = 1; //original sum
                    dtDataRow [27] = wsraw.Cells [intLoop, 8].Value; //initial sum
                    dtDataRow [21] = "ADJUST"; //Trans Code
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out dclSumAtRisk);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 31].Value), out dclPremium);
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));
                    dclTotalSumAtRisk += dclSumAtRisk;
                    dclTotalPremium += dclPremium;

                    string adjYear = Convert.ToString(wsraw.Cells [intLoop, 6].Value);
                    if(adjYear.ToUpper() == "FY")
                    {
                        dtDataRow [60] = "4002";
                        dtDataRow [61] = dclPremium;
                    }
                    else {
                        dtDataRow [62] = "4004";
                        dtDataRow [63] = dclPremium;
                    }

                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM041_21", "Information");
                return "";
            }
            #region Hash Total 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium";
            dtDataRow [1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk";
            dtDataRow [1] = dclTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

          
            #endregion

            if(Variables.boogenderfail)
            {
                //objdt_template.Rows.Add(dtDataRow);
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01 [0] = "Please check for blank genders";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }



            string despath = str_saved + @"\BM041_21" + str_sheet + str_savef + ".xlsx";
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