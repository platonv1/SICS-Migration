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
    class BM141
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
            string strSmoker = string.Empty;
            string strCheckSheetName = string.Empty;
            string strOriginalSum = string.Empty;
            string strInitialSum = string.Empty;
            string strRemarksAABBZ = string.Empty;
            string strSumAtRisk = string.Empty;
            string valueTransEffectiveDate = string.Empty;
            string CessionCode = string.Empty;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
                newform.ShowDialog();

            }

            DataRow dtDataRow;

            #region 
            if (str_sheet.ToUpper().Contains("DOLLAR REN")) 
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    if (wsraw.Cells[intLoop, 5].Value != null)
                    {
                        var BusinessType = wsraw.Cells[intLoop, 5].Value;
                        if (BusinessType.GetType() == typeof(string))
                        {
                            BusinessType = objHlpr2.fn_removeCharacters(BusinessType);
                            if (BusinessType == "FACULTATIVE")
                            {
                                CessionCode = "F";
                            }
                            else if (BusinessType == "AUTOMATIC")
                            {
                                CessionCode = "T";
                            }
                        }
                        string strPolicyNo = wsraw.Cells[intLoop, 2].Text.ToString();
                        if (Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            dtDataRow[0] = strPolicyNo;
                            dtDataRow[5] = Convert.ToString(wsraw.Cells[intLoop, 3].Value);
                            string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                            dtDataRow[23] = strCurrency; //  Cession Currency
                            dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                            string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                            dtDataRow[21] = strTcode; // Transcode
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            objHlpr.fn_CheckingforA_AB_BZColumn(null, null, Convert.ToString(wsraw.Cells[intLoop, 21].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                            string strFullName = wsraw.Cells[intLoop, 5].Value;
                            objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                            out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                            out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                            out string strLife);
                            string strIssueDate = Convert.ToDateTime(wsraw.Cells[intLoop, 1].Value).ToString("MM/dd/yyyy");
                            dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                            dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date
                            dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                            dtDataRow[31] = strFullName; //Full Name
                            dtDataRow[33] = strFirstName;
                            dtDataRow[32] = strLastName;
                            dtDataRow[34] = strMiddleName;
                            strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strDOB; //Birthday
                            dtDataRow[36] = strSex; //Gender
                            dtDataRow[30] = strLifeID;//life ID 
                            dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                            dtDataRow[58] = "4001"; // Entry code
                            dtDataRow[79] = strIssueAge;//issue age
                            dtDataRow[39] = objHlpr2.fn_getmortalityrating(null);
                            objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                            dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks

                            decimal dclPremLife = Convert.ToDecimal(wsraw.Cells[intLoop, 6].Value);
                            decimal dclPremADB = Convert.ToDecimal(wsraw.Cells[intLoop, 7].Value);
                            decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 8].Value);
                            decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells[intLoop, 9].Value);

                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow04 = objdt_template.NewRow();
                            _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                            #region Premiums and SaR
                            if (CessionCode == "T")
                            {
                                if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;


                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                    _var.dtworkRow04[5] = "PDD";
                                    _var.dtworkRow04[14] = CessionCode;//Business Type;
                                    _var.dtworkRow04[59] = dclPremPDD;
                                    _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow04);
                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "PDD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremPDD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;
                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }

                            }
                            else if (CessionCode == "F")
                            {
                                if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;


                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                    _var.dtworkRow04[5] = "PDD";
                                    _var.dtworkRow04[14] = CessionCode;//Business Type;
                                    _var.dtworkRow04[59] = dclPremPDD;
                                    _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow04);
                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "PDD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremPDD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;
                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                            }
                            #endregion

                            objHlpr.fn_getTotalPremiumV2(strCurrency, CessionCode, dclPremLife, dclPremADB, dclPremWPD, dclPremPDD);
                            objHlpr.fn_getTotalSumAtRiskV3(CessionCode, strCurrency, Convert.ToDecimal(strSumAtRisk));
                        }
                    }
                }

            }
            else if (str_sheet.ToUpper().Contains("PESO REN"))
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    if (wsraw.Cells[intLoop, 5].Value != null)
                    {
                        var BusinessType = wsraw.Cells[intLoop, 5].Value;
                        if (BusinessType.GetType() == typeof(string))
                        {
                            BusinessType = objHlpr2.fn_removeCharacters(BusinessType);
                            if (BusinessType == "FACULTATIVE")
                            {
                                CessionCode = "F";
                            }
                            else if (BusinessType == "AUTOMATIC")
                            {
                                CessionCode = "T";
                            }
                        }

                        string strPolicyNo = wsraw.Cells[intLoop, 2].Text.ToString();
                        if (Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            dtDataRow[0] = strPolicyNo;
                            dtDataRow[5] = Convert.ToString(wsraw.Cells[intLoop, 3].Value);//Branded Product Cedent Code
                            string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                            dtDataRow[23] = strCurrency; //  Cession Currency
                            dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                            string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                            dtDataRow[21] = strTcode; // Transcode
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Cells[intLoop, 21].Value), Convert.ToString(wsraw.Cells[intLoop, 21].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                            string strFullName = wsraw.Cells[intLoop, 5].Value;
                            objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                            out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                            out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                            out string strLife);
                            dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                            string strIssueDate = Convert.ToDateTime(wsraw.Cells[intLoop, 1].Value).ToString("MM/dd/yyyy");
                            dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                            dtDataRow[20] = valueTransEffectiveDate;//Policy Start Date

                            dtDataRow[31] = strFullName; //Full Name
                            dtDataRow[33] = strFirstName;
                            dtDataRow[32] = strLastName;
                            dtDataRow[34] = strMiddleName;
                            strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strDOB; // Birthday
                            dtDataRow[36] = strSex; // Gender
                            dtDataRow[30] = strLifeID;// life ID 
                            dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                            dtDataRow[58] = "4001"; // Entry code
                            dtDataRow[79] = strIssueAge;//issue age
                            dtDataRow[39] = objHlpr2.fn_getmortalityrating(null);
                            objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                            dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                            decimal dclPremLife = Convert.ToDecimal(wsraw.Cells[intLoop, 6].Value);
                            decimal dclPremADB = Convert.ToDecimal(wsraw.Cells[intLoop, 7].Value);
                            decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 8].Value);
                            decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells[intLoop, 9].Value);

                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow04 = objdt_template.NewRow();
                            _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                            #region Premiums and SaR
                            if (CessionCode == "T")
                            {
                                if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;


                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                    _var.dtworkRow04[5] = "PDD";
                                    _var.dtworkRow04[14] = CessionCode;//Business Type;
                                    _var.dtworkRow04[59] = dclPremPDD;
                                    _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow04);
                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "PDD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremPDD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;
                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }

                            }
                            else if (CessionCode == "F")
                            {
                                if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;


                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                }
                                else if (dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "ADB";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremADB;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "WPD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremWPD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);

                                    _var.dtworkRow04[5] = "PDD";
                                    _var.dtworkRow04[14] = CessionCode;//Business Type;
                                    _var.dtworkRow04[59] = dclPremPDD;
                                    _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow04);
                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                    _var.dtworkRow03[5] = "PDD";
                                    _var.dtworkRow03[14] = CessionCode;//Business Type;
                                    _var.dtworkRow03[59] = dclPremPDD;
                                    _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow03);
                                }
                                else if (dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "ADB";
                                    dtDataRow[59] = dclPremADB;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;
                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                {
                                    dtDataRow[5] = "WPD";
                                    dtDataRow[59] = dclPremWPD;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                    _var.dtworkRow02[5] = "PDD";
                                    _var.dtworkRow02[14] = CessionCode;//Business Type;
                                    _var.dtworkRow02[59] = dclPremPDD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;

                                    _var.dtworkRow02[5] = "WPD";
                                    _var.dtworkRow02[14] = CessionCode;
                                    _var.dtworkRow02[59] = dclPremWPD;
                                    _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                                    objdt_template.Rows.Add(_var.dtworkRow02);

                                }
                                else if (dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                {
                                    dtDataRow[5] = "LIFE";
                                    dtDataRow[59] = dclPremLife;
                                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                    dtDataRow[14] = CessionCode;//Business Type;

                                }
                            }
                            #endregion
                            objHlpr.fn_getTotalPremiumV2(strCurrency, CessionCode, dclPremLife, dclPremADB, dclPremWPD, dclPremPDD);
                            objHlpr.fn_getTotalSumAtRiskV3(CessionCode, strCurrency, Convert.ToDecimal(strSumAtRisk));
                        }
                    }
                }

            }  //Q3 Workbook
          
            #endregion


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Treaty Total Premium:";
            dtDataRow[1] = Variables.TotalTreatyPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Treaty Total Sum at Risk:";
            dtDataRow[1] = Variables.TotalTreatySAR;
            objdt_template.Rows.Add(dtDataRow);


            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Falcultative Total Premium:";
            dtDataRow[1] = Variables.TotalFaculPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Facultative Total Sum at Risk:";
            dtDataRow[1] = Variables.TotalFaculSAR;
            objdt_template.Rows.Add(dtDataRow);
            #endregion

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            //if (Variables.boogenderfail)
            //{
            //    //objdt_template.Rows.Add(dtDataRow);
            //    _var.dtworkRow01 = objdt_template.NewRow();
            //    _var.dtworkRow01[0] = "Please check for blank genders";
            //    objdt_template.Rows.Add(_var.dtworkRow01);
            //}

           
            string despath = str_saved + @"\BM141-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            Variables.TotalTreatyPremium = 0;
            Variables.TotalTreatySAR = 0;
            Variables.TotalFaculPremium = 0;
            Variables.TotalFaculSAR = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}
