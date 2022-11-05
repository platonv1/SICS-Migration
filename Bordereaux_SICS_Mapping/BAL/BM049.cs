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
    class BM049
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
            DataRow dtDataRow;

            string strFilePath = wbraw.Path;
            string strRemarksAABBZ = string.Empty;
            string strIssueDate = string.Empty;
            string valueTransEffectiveDate = string.Empty;
            decimal dclTotalPremiumQuota = 0;
            decimal dclTotalPremiumSurplus = 0;
            decimal dclTotalQuotaSAR = 0;
            decimal dclTotalSurplusSAR = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
            }


            if(str_sheet.ToUpper().Contains("AGRICOM") || str_sheet.ToUpper().Contains("ARGEM") || str_sheet.ToUpper().Contains("AYALA") || str_sheet.ToUpper().Contains("BATANGAS")
            || str_sheet.ToUpper().Contains("CAMAVEMCO") || str_sheet.ToUpper().Contains("INFANTA") || str_sheet.ToUpper().Contains("JH")
            || str_sheet.ToUpper().Contains("MASAG") || str_sheet.ToUpper().Contains("MASANTO") || str_sheet.ToUpper().Contains("MULTISAVE")
            || str_sheet.ToUpper().Contains("CAVINTI") || str_sheet.ToUpper().Contains("POLA") || str_sheet.ToUpper().Contains("POZORRUBIO") 
            || str_sheet.ToUpper().Contains("SUMMIT") || (str_sheet.ToUpper().Contains("AIR FORCE")  || str_sheet.ToUpper().Contains("CARE MBA") || str_sheet.ToUpper().Contains("CHRISTIAN") || str_sheet.ToUpper().Contains("CITYSTATE")
            || str_sheet.ToUpper().Contains("REALTY") || str_sheet.ToUpper().Contains("DAYLIGHT") || str_sheet.ToUpper().Contains("EAST") || str_sheet.ToUpper().Contains("ENTERPRISE") || str_sheet.ToUpper().Contains("ENTREPRENEUR")
            || str_sheet.ToUpper().Contains("FABSLAI") || str_sheet.ToUpper().Contains("FIRST COMMUNITY") || str_sheet.ToUpper().Contains("LIMCOMA") || str_sheet.ToUpper().Contains("LWUA") || str_sheet.ToUpper().Contains("MAGCOOP") || str_sheet.ToUpper().Contains("MAKILING")
            || str_sheet.ToUpper().Contains("METRO") || str_sheet.ToUpper().Contains("MWSS") || str_sheet.ToUpper().Contains("EMPCO") || str_sheet.ToUpper().Contains("PCGSLAI") || str_sheet.ToUpper().Contains("PENCOOP") || str_sheet.ToUpper().Contains("PESO LINE") || str_sheet.ToUpper().Contains("PHIL STAR")
            || str_sheet.ToUpper().Contains("PNP LAKAS") || str_sheet.ToUpper().Contains("PROGRESSIVE") || str_sheet.ToUpper().Contains("PSSLAI") || str_sheet.ToUpper().Contains("JAEN") || str_sheet.ToUpper().Contains("SAN LUIS")
            || str_sheet.ToUpper().Contains("ROSARIAN") || str_sheet.ToUpper().Contains("BAREMCOOP") || str_sheet.ToUpper().Contains("SHANGRI-LA MACTAN COOP") || str_sheet.ToUpper().Contains("ST JOSEPH") || str_sheet.ToUpper().Contains("SUNRISE") 
            || str_sheet.ToUpper().Contains("NGCP") || str_sheet.ToUpper().Contains("SHANGRI") || str_sheet.ToUpper().Contains("ASIANMINES") || str_sheet.ToUpper().Contains("BATAAN DEV") || str_sheet.ToUpper().Contains("MUTUAL") || str_sheet.ToUpper().Contains("COOP") ||
            str_sheet.ToUpper().Contains("LIPA") || str_sheet.ToUpper().Contains("RAC FUND ") || str_sheet.ToUpper().Contains("ANGELES") || str_sheet.ToUpper().Contains("SEVEN")  || str_sheet.ToUpper().Contains("CALAMBA") || str_sheet.ToUpper().Contains("PROVIDENT") 
            || str_sheet.ToUpper().Contains("FUND") || str_sheet.ToUpper().Contains("JOSEPH") || str_sheet.ToUpper().Contains("PREDOMINANT")))
            {
                for (int intLoop = 8; intLoop <= erawrow + 1; intLoop++)
                {
                    string strIssueAge = Convert.ToString(wsraw.Cells[intLoop, 6].Value); //Issue Age
                    if (string.IsNullOrEmpty(strIssueAge))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strFullName = strLastName.TrimEnd() + " " + strFirstName.TrimEnd();
                    objHlpr2.fn_separateLastNameFirstNameV2(strFullName,out strFullName, out strLastName, out strFirstName, out string strMiddleInitial);
                    //objHlpr.fn_checksheetnameV2(str_sheet,strFullName, out strLastName, out strFirstName, out string strMiddleInitial);
                    dtDataRow[34] =strMiddleInitial;
                    dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName);
                    strLastName = objHlpr2.fn_checkLastname(strLastName);
                    dtDataRow[32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow[31] = strFullName;
                    string strDOB = Convert.ToDateTime(wsraw.Cells[intLoop, 5].Value).ToString("MM/dd/yyyy");
                    //strDOB = objHlpr2.fn_reformatDatev1(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");//Birthday
                    dtDataRow[37] = objHlpr2.fn_checkDOB(strDOB);
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 4].Value);// Gender
                    if (string.IsNullOrEmpty(strSex))
                    {
                        dtDataRow[36] = objHlpr.fn_getgenderv2(strSex);
                    }
                    else
                    {
                        dtDataRow[36] = strSex;
                    }
                    string strPolicyNo = Convert.ToString(wsraw.Cells[2, 3].Value);
                    dtDataRow[0] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleInitial, strLastName, strDOB);//PolicyNo
                    dtDataRow[7] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleInitial, strLastName, strDOB);//GroupSchemeID
                    dtDataRow[8] = "COMBINE"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[23] ="PHP"; //  Cession Currency
                    dtDataRow[14] = "T"; // Business Type
                    string strTcode = "";
                    dtDataRow[21] = strTcode; // Transcode
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[5] = "LIFE";
                    dtDataRow [58] = "4001";
                    dtDataRow [39] = "STANDARD";
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    
                    if (Variables.strBmYear == "2020" || Variables.strBmYear == "2021" && str_raw.ToUpper().Contains("2Q") || Variables.strBmYear == "2021" && str_raw.ToUpper().Contains("3Q")
                     || Variables.strBmYear == "2021" && str_raw.ToUpper().Contains("4Q") || Variables.strBmYear == "2022")
                    {
                        strIssueDate = objHlpr2.fn_checkIssueDate(wsraw.Cells [3, 3].Text);
                    }
                    else
                    {
                        strIssueDate = objHlpr2.fn_checkIssueDate(wsraw.Cells [2, 3].Text);
                    }
                   
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow[79] = strIssueAge;
                    dtDataRow[82] = Convert.ToString(wsraw.Cells[1, 3].Value); //Group Policy holder
                    objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksCode;

                    decimal dclPremQuota = Convert.ToDecimal(wsraw.Cells[intLoop, 14].Value); //Premium
                    decimal dclPremSurplus = Convert.ToDecimal(wsraw.Cells[intLoop, 15].Value);//Premium
                    decimal dclQuotaCeded = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);//SUM AT RISK
                    decimal dclSurplusCeded = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);//SUM AT RISK
                    decimal dclOriginalSum = Convert.ToDecimal(wsraw.Cells[intLoop, 10].Value);//Original Sum
                    decimal dclCededRetention = Convert.ToDecimal(wsraw.Cells[intLoop, 11].Value);//Ceded Retention

                    dclTotalPremiumQuota += dclPremQuota;
                    dclTotalPremiumSurplus += dclPremSurplus;
                    dclTotalQuotaSAR += dclQuotaCeded;
                    dclTotalSurplusSAR += dclSurplusCeded;

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    #region Quota & Surplus Premium
                    if (dclPremQuota != 0 && dclPremSurplus == 0)
                    {
                        dtDataRow[59] = dclPremQuota;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded));// sum at risk
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Method

                    }
                    else if (dclPremSurplus != 0 && dclPremQuota == 0)
                    {
                        dtDataRow[59] = dclPremSurplus;
                     
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded));//sum at risk
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Method

                    }
                    else if (dclPremQuota != 0 && dclPremSurplus != 0)
                    {
                        dtDataRow[59] = dclPremQuota;
                       
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded));// sum at risk
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Method
                        _var.dtworkRow02[59] = dclPremSurplus;
                        
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded));// sum at risk
                        _var.dtworkRow02 [28] = 0;
                        dtDataRow [10] = "S"; // Reinsurance Method
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    #endregion

                }
            }

           else if(str_sheet.ToUpper().Contains("TFSI"))
           
            {
                for(int intLoop = 8; intLoop <= erawrow + 1; intLoop++)
                {
                    string strLastName = ""; string strFirstName = ""; string strMI = "";
                    string strIssueAge = Convert.ToString(wsraw.Cells [intLoop, 6].Value); //Issue Age
                    if(string.IsNullOrEmpty(strIssueAge))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                  
                    string strFullName = Convert.ToString(wsraw.Cells [intLoop, 2].Value);
                    
                    if(strFullName.Contains(","))
                    {
                        objHlpr2.fn_separateLastNameFirstNameV8(strFullName, out strLastName, out strFirstName, out strMI);
                     
                    }
                    else
                    {
                        objHlpr.fn_separatefullnamev5(strFullName, out strFirstName, out strLastName, out strMI);
                    }
                    dtDataRow [34] = strMI;
                    dtDataRow [33] = strFirstName;
                    dtDataRow [32] = strLastName;
                    dtDataRow [31] = strFullName;

                    //objHlpr.fn_checksheetnameV2(str_sheet,strFullName, out strLastName, out strFirstName, out string strMiddleInitial);

                    string strDOB = Convert.ToDateTime(wsraw.Cells [intLoop, 5].Value).ToString("MM/dd/yyyy");
                    //strDOB = objHlpr2.fn_reformatDatev1(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");//Birthday
                    dtDataRow [37] = objHlpr2.fn_checkDOB(strDOB);
                    string strSex = Convert.ToString(wsraw.Cells [intLoop, 4].Value);// Gender
                    if(string.IsNullOrEmpty(strSex))
                    {
                        dtDataRow [36] = objHlpr.fn_getgenderv2(strSex);
                    }
                    else
                    {
                        dtDataRow [36] = strSex;
                    }
 
                    string strPolicyNo = Convert.ToString(wsraw.Cells [2, 3].Value);
                    dtDataRow [0] = objHlpr2.fn_generatePolicyno(strPolicyNo, strFirstName, strMI, strLastName, strDOB);//PolicyNo
                    dtDataRow [7] = objHlpr2.fn_generatePolicyno(strPolicyNo, strFirstName, strMI, strLastName, strDOB);//GroupSchemeID
                    dtDataRow [8] = "COMBINE"; // Reinsurance Product
                    dtDataRow [9] = "PA"; // Type of Business
                    dtDataRow [13] = "GRP"; // Class of Business
                    dtDataRow [23] = "PHP"; //  Cession Currency
                    dtDataRow [14] = "T"; // Business Type
                    string strTcode = "";
                    dtDataRow [21] = strTcode; // Transcode
                    dtDataRow [24] = "MLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [5] = "LIFE";
                    dtDataRow [58] = "4001";
                    dtDataRow [39] = "STANDARD";
                    dtDataRow [41] = Variables.strBmYear; //Policy Year

                    if(Variables.strBmYear == "2020" || Variables.strBmYear == "2021" && str_sheet.ToUpper().Contains("2Q") || Variables.strBmYear == "2021" && str_sheet.ToUpper().Contains("3Q")
                     || Variables.strBmYear == "2021" && str_sheet.ToUpper().Contains("4Q") || Variables.strBmYear == "2022")
                    {
                        strIssueDate = objHlpr2.fn_checkIssueDate(wsraw.Cells [3, 3].Text);
                    }
                    else
                    {
                        strIssueDate = objHlpr2.fn_checkIssueDate(wsraw.Cells [2, 3].Text);
                    }

                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow [79] = strIssueAge;
                    dtDataRow [82] = Convert.ToString(wsraw.Cells [1, 3].Value); //Group Policy holder
                    objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strSex, out string strRemarksCode);
                    dtDataRow [76] = strRemarksCode;

                    decimal dclPremQuota = Convert.ToDecimal(wsraw.Cells [intLoop, 14].Value); //Premium
                    decimal dclPremSurplus = Convert.ToDecimal(wsraw.Cells [intLoop, 15].Value);//Premium
                    decimal dclQuotaCeded = Convert.ToDecimal(wsraw.Cells [intLoop, 12].Value);//SUM AT RISK
                    decimal dclSurplusCeded = Convert.ToDecimal(wsraw.Cells [intLoop, 13].Value);//SUM AT RISK
                    decimal dclOriginalSum = Convert.ToDecimal(wsraw.Cells [intLoop, 10].Value);//Original Sum
                    decimal dclCededRetention = Convert.ToDecimal(wsraw.Cells [intLoop, 11].Value);//Ceded Retention

                    dclTotalPremiumQuota += dclPremQuota;
                    dclTotalPremiumSurplus += dclPremSurplus;
                    dclTotalQuotaSAR += dclQuotaCeded;
                    dclTotalSurplusSAR += dclSurplusCeded;

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    #region Quota & Surplus Premium
                    if(dclPremQuota != 0 && dclPremSurplus == 0)
                    {
                        dtDataRow [59] = dclPremQuota;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded)); //Initial Sum
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded));// sum at risk
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Method

                    }
                    else if(dclPremSurplus != 0 && dclPremQuota == 0)
                    {
                        dtDataRow [59] = dclPremSurplus;

                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded)); //Initial Sum
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded));//sum at risk
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Method

                    }
                    else if(dclPremQuota != 0 && dclPremSurplus != 0)
                    {
                        dtDataRow [59] = dclPremQuota;

                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded)); //Initial Sum
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclQuotaCeded));// sum at risk
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Method
                        _var.dtworkRow02 [59] = dclPremSurplus;

                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded)); //Initial Sum
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSurplusCeded));// sum at risk
                        _var.dtworkRow02 [28] = 1;
                        dtDataRow [10] = "S"; // Reinsurance Method
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    #endregion

                }
            }

            else if (str_sheet.ToUpper().Contains("TRANSNATIONAL") || str_sheet.ToUpper().Contains("VETERAN") || str_sheet.ToUpper().Contains("CWSLAI"))
            {
                for (int intLoop = 8; intLoop <= erawrow + 1; intLoop++)
                {
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 4].Value);// Gender
                    
                    if (string.IsNullOrEmpty(strSex))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strFullName = Convert.ToString(wsraw.Cells[intLoop, 2].Value);
                    objHlpr2.fn_separateLastNameFirstNameV3(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                    dtDataRow [34] = objHlpr2.fn_removeCharacters(strMiddleInitial);
                    dtDataRow [33] = objHlpr2.fn_checkFirstname(strFirstName);
                    strLastName = objHlpr2.fn_checkLastname(strLastName);
                    dtDataRow [32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow [31] = strFullName;
                    string strDOB = Convert.ToDateTime(wsraw.Cells[intLoop, 5].Value).ToString("MM/dd/yyyy");
                    dtDataRow[37] = objHlpr2.fn_checkDOB(strDOB);
                    
                    if (string.IsNullOrEmpty(strSex))
                    {
                        dtDataRow[36] = objHlpr.fn_getgenderv2(strSex);
                    }
                    else
                    {
                        dtDataRow[36] = strSex;
                    }
                    string strPolicyNo = Convert.ToString(wsraw.Cells[2, 3].Value);
                    dtDataRow[0] = objHlpr2.fn_generatePolicyno(strPolicyNo, strFirstName, strMiddleInitial, strLastName, strDOB);//PolicyNo
                    dtDataRow[7] = objHlpr2.fn_generatePolicyno(strPolicyNo, strFirstName, strMiddleInitial, strLastName, strDOB);//GroupSchemeID
                    dtDataRow[8] = "COMBINE"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow [13] = "GRP"; // Class of Business
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[14] = "T"; // Business Type
                    string strTcode = "TRENEW";
                    dtDataRow[21] = strTcode; // Transcode
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow [5] = "GCL"; //Branded Product

                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [39] = "STANDARD"; //preferred classific
                    strIssueDate = Convert.ToDateTime(wsraw.Cells [3, 3].Value).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow[79] = Convert.ToString(wsraw.Cells[intLoop, 6].Value); //Issue Age
                    dtDataRow[82] = Convert.ToString(wsraw.Cells[1, 3].Value); //Group Policy holder
                    objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksCode;
                    dtDataRow [58] = "4001";
                    decimal dclPremQuota = Convert.ToDecimal(wsraw.Cells[intLoop, 14].Value); //Premium
                    decimal dclPremSurplus = Convert.ToDecimal(wsraw.Cells[intLoop, 15].Value);//Premium
                    decimal dclCededQuota = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);//SUM AT RISK
                    decimal dclCededSurplus = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);//SUM AT RISK
                    decimal dclOriginalSum = Convert.ToDecimal(wsraw.Cells[intLoop, 10].Value);//Original Sum
                    decimal dclCededRetention = Convert.ToDecimal(wsraw.Cells[intLoop, 11].Value);//Ceded Retention

                    dclTotalPremiumQuota += dclPremQuota;
                    dclTotalPremiumSurplus += dclPremSurplus;
                    dclTotalQuotaSAR += dclCededQuota;
                    dclTotalSurplusSAR += dclCededSurplus;

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    #region Quota & Surplus Premium
                    if (dclPremQuota != 0 && dclPremSurplus == 0)
                    {
                        dtDataRow[59] = dclPremQuota;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Method
                    }
                    else if (dclPremSurplus != 0 && dclPremQuota == 0)
                    {
                        dtDataRow[59] = dclPremSurplus;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "S"; // Reinsurance Method

                    }
                    else if (dclPremQuota != 0 && dclPremSurplus != 0)
                    {
                        dtDataRow[59] = dclPremQuota;
                        dtDataRow [25] = dclOriginalSum;
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Method

                        _var.dtworkRow02[59] = dclPremSurplus;
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOriginalSum));
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02 [28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Method
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    #endregion

                }
            }

            else if (str_sheet.ToUpper().Contains("BDO") ||  str_sheet.ToUpper().Contains("FAIRCHILD"))
            {
                for (int intLoop = 9; intLoop <= erawrow + 1; intLoop++)
                {
                    string strIssueAge = Convert.ToString(wsraw.Cells[intLoop, 6].Value); //Issue Age
                 
                    if (string.IsNullOrEmpty(strIssueAge))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strFullName = strLastName.TrimEnd() + " " + strFirstName.TrimEnd();
                    objHlpr2.fn_separateLastNameFirstNameV2(strFullName, out strFullName, out strLastName, out strFirstName, out string strMiddleInitial);
                    dtDataRow [34] = strMiddleInitial;
                    dtDataRow [33] = objHlpr2.fn_checkFirstname(strFirstName);
                    strLastName = objHlpr2.fn_checkLastname(strLastName);
                    dtDataRow [32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow [31] = strLastName + " " + strFirstName;
                    string strDOB = Convert.ToDateTime(wsraw.Cells [intLoop, 5].Value).ToString("MM/dd/yyyy");
                    dtDataRow [37] = objHlpr2.fn_checkDOB(strDOB);
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 4].Value);// Gender
                    if (string.IsNullOrEmpty(strSex))
                    {
                        dtDataRow[36] = objHlpr.fn_getgenderv2(strSex);
                    }
                    else
                    {
                        dtDataRow[36] = strSex;
                    }
                    string strPolicyNo = Convert.ToString(wsraw.Cells[2, 3].Value);
                    dtDataRow[0] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleInitial, strLastName, strDOB);//PolicyNo
                    dtDataRow[7] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleInitial, strLastName, strDOB);//GroupSchemeID
                    dtDataRow[8] = "COMBINE"; // Reinsurance Product
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    string strTcode = "TRENEW";
                    dtDataRow[21] = strTcode; // Transcode
                    dtDataRow[14] = "T"; //Bussiness Type
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow [5] = "GCL"; //Branded Product
                    dtDataRow [13] = "GRP"; // Class of Business
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow [39] = "STANDARD"; //preferred classific
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    strIssueDate = Convert.ToDateTime(wsraw.Cells [3, 3].Value).ToString("MM/dd/yyyy");
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow[28] = Convert.ToString(wsraw.Cells[intLoop, 11].Value);//Cedent Retention
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow[79] = strIssueAge; //Issue Age
                    dtDataRow[82] = Convert.ToString(wsraw.Cells[1, 3].Value); //Group Policy holder
                    objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksCode;
                    dtDataRow [58] = "4001";

                    decimal dclLifeQuotaPrem = Convert.ToDecimal(wsraw.Cells[intLoop, 14].Value); //Premium
                    decimal dclWiddQuotaPrem = Convert.ToDecimal(wsraw.Cells[intLoop, 15].Value); //Premium
                    decimal dclLifeSurplusPrem = Convert.ToDecimal(wsraw.Cells[intLoop, 16].Value); //Premium
                    decimal dclWiddSurplusPrem = Convert.ToDecimal(wsraw.Cells[intLoop, 17].Value); //Premium
                    decimal dclCededQuota = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value); //Sum at risk
                    decimal dclCededSurplus = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);//Sum at risk
                    decimal dclCededRetention = Convert.ToDecimal(wsraw.Cells [intLoop, 11].Value);//Ceded Retention

                    dclTotalPremiumQuota += dclLifeQuotaPrem + dclWiddQuotaPrem;
                    dclTotalPremiumSurplus += dclLifeSurplusPrem + dclWiddSurplusPrem;
                    dclTotalQuotaSAR += dclCededQuota;
                    dclTotalSurplusSAR += dclCededSurplus;


                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                    #region Quota & Surplus Premium
                    if (dclLifeQuotaPrem != 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclLifeQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods
                    }
                    else if (dclLifeQuotaPrem == 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclWiddQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods
                    }
                    else if (dclLifeQuotaPrem == 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclLifeSurplusPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLifeSurplusPrem)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "S"; // Reinsurance Methods
                    }
                    else if (dclLifeQuotaPrem == 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow [59] = dclWiddSurplusPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLifeSurplusPrem)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "S"; // Reinsurance Methods
                    }
                    else if (dclLifeQuotaPrem != 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclLifeQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclWiddQuotaPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(null)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(null));// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclLifeQuotaPrem != 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclLifeQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclWiddQuotaPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "Q"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[59] = dclLifeSurplusPrem;
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow03[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if (dclLifeQuotaPrem == 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclWiddQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclLifeSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclLifeQuotaPrem == 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow[59] = dclWiddQuotaPrem;
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclWiddSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclLifeQuotaPrem == 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow[59] = dclWiddQuotaPrem;
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = dclCededRetention;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclLifeSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[59] = dclWiddSurplusPrem;
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial Sum
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow03[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if(dclLifeQuotaPrem == 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow[59] = dclLifeSurplusPrem;
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "S"; // Reinsurance Methods


                        _var.dtworkRow02[59] = dclWiddSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow03[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclLifeQuotaPrem != 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem == 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow[59] = dclLifeQuotaPrem;
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclWiddSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclLifeQuotaPrem != 0 && dclWiddQuotaPrem == 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem == 0)
                    {
                        dtDataRow[59] = dclLifeQuotaPrem;
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 10].Value)); //Original Sum
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclLifeSurplusPrem;
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02[28] = 1;
                        _var.dtworkRow02 [10] = "Q"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclLifeQuotaPrem != 0 && dclWiddQuotaPrem != 0 && dclLifeSurplusPrem != 0 && dclWiddSurplusPrem != 0)
                    {
                        dtDataRow [59] = dclLifeQuotaPrem;
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota)); //Initial Sum
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuota));// sum at risk
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededRetention));
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclWiddQuotaPrem;
                        _var.dtworkRow02 [58] = "4001";
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial Sum
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow02 [28] = 1;
                        _var.dtworkRow02 [10] = "Q"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [59] = dclLifeSurplusPrem;
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus)); //Initial Sum
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplus));// sum at risk
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow03 [28] = 1;
                        _var.dtworkRow03 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow03);


                        _var.dtworkRow04 [59] = dclWiddSurplusPrem;
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial Sum
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 10].Value)); //Original Sum
                        _var.dtworkRow04 [28] = 1;
                        _var.dtworkRow04 [10] = "S"; // Reinsurance Methods
                        objdt_template.Rows.Add(_var.dtworkRow04);

                    }
                    #endregion
                }
            }

            else if (str_sheet.ToUpper().Contains("AVIDA"))
            {
                for (int intLoop = 11; intLoop <= erawrow + 1; intLoop++)
                {
                    string strDOB = Convert.ToDateTime(wsraw.Cells[intLoop, 5].Value).ToString("MM/dd/yyyy");
                    if (string.IsNullOrEmpty(strDOB))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strLastName = objHlpr2.fn_checkLastname(Convert.ToString(wsraw.Cells[intLoop, 2].Value));
                    string strFirstName = objHlpr2.fn_checkFirstname(Convert.ToString(wsraw.Cells[intLoop, 3].Value));
                    string strFullName = strLastName.TrimEnd() + " " + strFirstName.TrimEnd();
                    objHlpr2.fn_separateLastNameFirstNameV2(strFullName, out strFullName, out strLastName, out strFirstName, out string strMiddleName);
                    dtDataRow [34] = strMiddleName;
                    dtDataRow [33] = strFirstName;
                    dtDataRow [32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow [31] = strLastName + " " + strFirstName;
                    dtDataRow [37] = objHlpr2.fn_checkDOB(strDOB);
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 4].Value);// Gender
                    if (string.IsNullOrEmpty(strSex))
                    {
                        dtDataRow[36] = objHlpr.fn_getgenderv2(strSex);
                    }
                    else
                    {
                        dtDataRow[36] = strSex;
                    }
                    string strPolicyNo = Convert.ToString(wsraw.Cells[2, 3].Value);
                    dtDataRow[0] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleName, strLastName, strDOB);//PolicyNo 
                    dtDataRow[7] = objHlpr2.fn_generatePolicyno(strPolicyNo,strFirstName, strMiddleName, strLastName, strDOB);//GroupSchemeID
                    dtDataRow[8] = "COMBINE"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[14] = "T"; // Business Type
                    string strTcode = "TRENEW";
                    dtDataRow[21] = strTcode; // Transcode
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow [5] = "GCL"; //Branded Product
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [39] = "STANDARD"; //preferred classific
                    strIssueDate = Convert.ToDateTime(wsraw.Cells[3, 3].Value).ToString("MM/dd/yyyy");
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow[19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    //dtDataRow[78] = objHlpr.fn_getAttainAge(Variables.strBmYear, strDOB);//Life Attain Age;
                    dtDataRow[79] = Convert.ToString(wsraw.Cells[intLoop, 6].Value); //Issue Age
                    dtDataRow[82] = Convert.ToString(wsraw.Cells[1, 3].Value); //Group Policy holder
                    objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksCode;
                    dtDataRow [58] = "4001";

                    decimal dclCededQuotaLife = Convert.ToDecimal(wsraw.Cells[intLoop, 16].Value);//SAR
                    decimal dclCededQuotaWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 17].Value);//SAR
                    decimal dclCededSurplusLife = Convert.ToDecimal(wsraw.Cells[intLoop, 18].Value);//SAR
                    decimal dclCededSurplusWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 19].Value);//SAR
                    decimal dclPremQuotaLife = Convert.ToDecimal(wsraw.Cells[intLoop, 20].Value);//Premium Quota
                    decimal dclPremQuotaWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 21].Value);//Premium Quota
                    decimal dclPremSurplusLife = Convert.ToDecimal(wsraw.Cells[intLoop, 22].Value);//Premium Surplus
                    decimal dclPremSurplusWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 23].Value);//Premium Surplus

                    decimal dclCoverageLife = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);//Original Sum
                    decimal dclCoverageWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);//Original Sum
                    decimal dclCrLife = Convert.ToDecimal(wsraw.Cells[intLoop, 14].Value);//Cedent Retention
                    decimal dclCrWidd = Convert.ToDecimal(wsraw.Cells[intLoop, 15].Value);//Cedent Retention

                    dclTotalQuotaSAR += dclCededQuotaLife + dclCededQuotaWidd;
                    dclTotalSurplusSAR += dclCededSurplusLife + dclCededSurplusWidd;
                    dclTotalPremiumQuota += dclPremQuotaLife + dclPremQuotaWidd;
                    dclTotalPremiumSurplus += dclPremSurplusLife + dclPremSurplusWidd;

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                    #region Sum at risk / Premium / Orig Sum / Cedent Retetion
                    if (dclCededQuotaLife != 0 && dclCededQuotaWidd == 0 && dclCededSurplusLife == 0 && dclCededSurplusWidd == 0)
                    {
                        dtDataRow[59] = dclPremQuotaLife;
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Sum at Risk
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife)); //Initial Sum
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife)); // Original Sum Assured
                        //dtDataRow[5] = "LIFE";                        
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrLife));
                        dtDataRow [10] = "Q"; // Reinsurance Methods
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife == 0 && dclCededSurplusWidd == 0)
                    {
                        dtDataRow[59] = dclPremQuotaWidd;
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Sum at Risk
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd)); //Initial Sum
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));//Original Sum
                        dtDataRow[5] = "WIDD";
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrWidd));
                        dtDataRow [10] = "Q"; // Reinsurance Methods
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd == 0 && dclCededSurplusLife == 0 && dclCededSurplusWidd != 0)
                    {
                        dtDataRow [59] = dclPremSurplusWidd;
                        dtDataRow [5] = "WIDD";
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrWidd));
                        dtDataRow [10] = "S"; // Reinsurance Methods
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd == 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd == 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclPremSurplusLife;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrLife));
                        dtDataRow [10] = "S"; // Reinsurance Methods
                    }
                    else if(dclCededQuotaLife != 0 && dclCededQuotaWidd == 0 && dclCededSurplusLife == 0 && dclCededSurplusWidd != 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclPremQuotaLife;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrLife));
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclPremSurplusWidd;
                        _var.dtworkRow02 [5] = "WIDD";
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        _var.dtworkRow02 [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCrWidd));
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd == 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd != 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclPremSurplusLife;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow [28] = dclCrLife;
                        dtDataRow [10] = "S"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclPremSurplusWidd;
                        _var.dtworkRow02 [5] = "WIDD";
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        _var.dtworkRow02 [28] = dclCrWidd;
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd == 0)
                    {
                        dtDataRow [5] = "WIDD";
                        dtDataRow [59] = dclPremQuotaWidd;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Initial Sum
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        dtDataRow [28] = dclCrWidd;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclPremSurplusLife;
                        //_var.dtworkRow02 [5] = "LIFE";
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));//Original Sum
                        _var.dtworkRow02 [28] = 0;
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclCededQuotaLife != 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd == 0)
                    {
                        dtDataRow[59] = dclPremQuotaLife;
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Sum at Risk
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));//Original Sum
                        //dtDataRow[5] = "LIFE";
                        dtDataRow[28] = dclCrLife;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[59] = dclPremQuotaWidd;
                        _var.dtworkRow02[5] = "WIDD";
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Sum at Risk
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Initial Sum
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));//Original Sum
                        _var.dtworkRow02[28] = dclCrWidd;
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[59] = dclPremSurplusLife;
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));//Original Sum
                        //_var.dtworkRow03[5] = "LIFE";
                        _var.dtworkRow03[28] = 0;
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if(dclCededQuotaLife == 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd != 0)
                    {
                        dtDataRow [5] = "WIDD";
                        dtDataRow [59] = dclPremQuotaWidd;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Initial Sum
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        dtDataRow [28] = dclCrWidd;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclPremSurplusLife;;
                        //_var.dtworkRow02 [5] = "LIFE";
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));//Original Sum
                        _var.dtworkRow02 [28] = 0;
                        _var.dtworkRow02 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [59] = dclPremSurplusWidd;
                        _var.dtworkRow03 [5] = "WIDD";
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(Convert.ToString(dclCoverageWidd)));
                        _var.dtworkRow03 [28] = 0;
                        _var.dtworkRow03 [10] = "S"; 
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if(dclCededQuotaLife != 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife == 0 && dclCededSurplusWidd == 0)
                    {
                        dtDataRow [59] = dclPremQuotaLife;
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Sum at Risk
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));//Original Sum
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [28] = dclCrLife;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02 [59] = dclPremQuotaWidd;
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Sum at Risk
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Initial Sum
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));//Original Sum
                        _var.dtworkRow02 [5] = "WIDD";
                        _var.dtworkRow02 [28] = dclCrWidd;
                        _var.dtworkRow02 [10] = "Q";
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclCededQuotaLife != 0 && dclCededQuotaWidd != 0 && dclCededSurplusLife != 0 && dclCededSurplusWidd != 0)
                    {
                        //dtDataRow[5] = "LIFE";
                        dtDataRow[59] = dclPremQuotaLife;
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageLife));
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Initial Sum
                        dtDataRow[28] = dclCrLife;
                        dtDataRow [10] = "Q"; // Reinsurance Methods

                        _var.dtworkRow02[5] = "WIDD";
                        _var.dtworkRow02[59] = dclPremQuotaWidd;
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaWidd));//Initial Sum
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCoverageWidd));
                        _var.dtworkRow02[28] = dclCrWidd;
                        _var.dtworkRow02 [10] = "Q";
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[59] = dclPremSurplusLife;
                        //_var.dtworkRow03[5] = "LIFE";
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusLife));
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededQuotaLife));//Original Sum
                        _var.dtworkRow03[28] = 0;
                        _var.dtworkRow03 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04[59] = dclPremSurplusWidd;
                        _var.dtworkRow04 [5] = "WIDD";
                        _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededSurplusWidd));
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(Convert.ToString(dclCoverageWidd)));
                        _var.dtworkRow04[28] = 0;
                        _var.dtworkRow04 [10] = "S";
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }
                    #endregion

                }
            }


            #region Hash Total 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium Quota:";
            dtDataRow[1] = dclTotalPremiumQuota;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk Quota:";
            dtDataRow[1] = dclTotalQuotaSAR;
            objdt_template.Rows.Add(dtDataRow);
         
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium Surplus:";
            dtDataRow[1] = dclTotalPremiumSurplus;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk Surplus:";
            dtDataRow[1] = dclTotalSurplusSAR;
            objdt_template.Rows.Add(dtDataRow);
            #endregion

            if (Variables.boogenderfail)
            {
                //objdt_template.Rows.Add(dtDataRow);
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Please check for blank genders";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }

            

            string despath = str_saved + @"\BM049-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);
            objHlpr.fn_openfile(despath);

            dclTotalPremiumQuota = 0;
            dclTotalQuotaSAR = 0;
            dclTotalSurplusSAR = 0;
            dclTotalPremiumSurplus = 0;

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}