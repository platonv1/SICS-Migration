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
    class BM048
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

            string strSmoker = string.Empty;
            string strCheckSheetName = string.Empty;
            string strOriginalSum = string.Empty;
            string strInitialSum = string.Empty;
            string strRemarksAABBZ = string.Empty;
            string strSumAtRisk = string.Empty;
            string valueTransEffectiveDate = string.Empty;
            string CessionCode = string.Empty;

            decimal dclFaculTotalSumAtRisk = 0;
            decimal dclFaculTotalPremium = 0;
            decimal dclTreatyTotalPremium = 0;
            decimal dclTreatyTotalSumAtRisk = 0;

            decimal dclTotalSumAtRisk = 0;
            decimal dclTotalPremium = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
                newform.ShowDialog();

            }

            DataRow dtDataRow;
            #region NRCP Files
            if (str_raw.ToUpper().Contains("NRCP"))
            {
                if (str_raw.ToUpper().Contains("1Q") || (str_raw.ToUpper().Contains("2Q") || (str_raw.ToUpper().Contains("4Q"))))
                {

                    if (str_sheet.Trim() == "Peso-New Business")
                    {
                        for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                        {
                            string strPolicyNo = wsraw.Cells[intLoop, 2].Text.ToString();
                            if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                            {
                                continue;
                            }
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            dtDataRow[0] = strPolicyNo;
                            string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                            dtDataRow[23] = strCurrency; //  Cession Currency
                            dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                            string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                            dtDataRow[21] = strTcode; // Transcode
                            dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[14] = "T";
                           

                            string strIssueDate = Convert.ToDateTime(wsraw.Cells[intLoop, 1].Value).ToString("MM/dd/yyyy");
                            dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode,Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                            dtDataRow[19] = valueTransEffectiveDate;  // Reinsurance Start Date
                            string strFullName = Convert.ToString(wsraw.Cells[intLoop, 4].Value);
                            dtDataRow[31] = strFullName; //Full Name
                            //objHlpr2.fn_separateLastNameFirstNameV2(strFullName, out strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            objHlpr2.fn_separateLastNameFirstNameV6(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            strLastName = objHlpr2.fn_removeCharacters(strLastName);
                            dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName);
                            dtDataRow[32] = objHlpr2.fn_checkLastname(strLastName);
                            dtDataRow[34] = objHlpr2.fn_removeCharacters(strMiddleInitial);
                            string strDOB = objHlpr2.fn_checkDOB(null);
                            dtDataRow[37] = strDOB; //Birthday
                            string strSex = wsraw.Cells[intLoop, 6].Value;
                            dtDataRow[36] = strSex; //Gender
                            string strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                            dtDataRow[30] = strLifeID;//life ID 

                            dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                            dtDataRow[39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells[intLoop, 7].Value)); //preffered classific
                            dtDataRow[56] = "4000"; // Entry code
                            dtDataRow[79] = wsraw.Cells[intLoop, 5].Value; //life issue age 
                            objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                            dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks
                            
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 8].Value), out decimal dclPremLife);
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 9].Value), out decimal dclPremExtra);
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 10].Value), out decimal dclPremWPD);
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 11].Value), out decimal dclPremADB);
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 12].Value), out decimal dclPremPDD);

                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 14].Value), out decimal dclCededLife); //life
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 15].Value), out decimal dclCededWPD); //wpd
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 16].Value), out decimal dclCedeADB); //adb
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 17].Value), out decimal dclCededPDD); //pdd
                            decimal.TryParse(Convert.ToString(wsraw.Cells[intLoop, 18].Value), out decimal dclSAR); //nar

                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow04 = objdt_template.NewRow();
                            _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow05 = objdt_template.NewRow();
                            _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;


                            #region Premiums and SaR
                            if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [25] = dclCededLife; // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [26] = dclCedeADB;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "WPD"; // Branded Product
                                dtDataRow [57] = dclPremWPD;
                                dtDataRow [26] = dclCededWPD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "PDD"; // Branded Product
                                dtDataRow [57] = dclPremPDD;
                                dtDataRow [26] = dclCededPDD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [26] = 0;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "PDD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;

                                _var.dtworkRow02 [5] = "ADB";
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD";
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// initial sum
                                _var.dtworkRow03 [26] = dclCedeADB;
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow03 [28] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;


                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD; ;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "PDD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremPDD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);


                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);

                                _var.dtworkRow05 [5] = "PDD"; // Branded Product
                                _var.dtworkRow05 [57] = dclPremPDD;
                                _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow05 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow05);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                            }
                            #endregion

                            dclTotalSumAtRisk += dclSAR;
                            dclTotalPremium += dclPremLife + dclPremExtra + dclPremADB + dclPremWPD + dclPremPDD;
                        }

                    }

                    else if (str_sheet.ToUpper().Contains("RENEWALS"))
                        {
                            for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                            {
                                string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                                if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 5].Text.ToString(), wsraw.Cells [intLoop, 6].Text.ToString(), wsraw.Cells [intLoop, 7].Text.ToString()))
                                {
                                    continue;
                                }
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;

                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                dtDataRow [14] = "T";

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                                string strDOB = objHlpr2.fn_checkDOB(null);
                                dtDataRow [37] = strDOB; //Birthday
                                string strFullName = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                                dtDataRow [31] = strFullName; //Full Name
                                //objHlpr.fn_getnamesandlifeID(strFullName, strDOB, out string strFirstName, out string strLastName, out _var.str_outlifeid, "000");
                                objHlpr2.fn_separateLastNameFirstNameV6(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                                dtDataRow [33] = strFirstName.Trim();
                                dtDataRow [32] = strLastName.Trim();
                                dtDataRow [34] = objHlpr2.fn_removeCharacters(strMiddleInitial);



                            string strSex = wsraw.Cells [intLoop, 7].Text;
                                dtDataRow [36] = strSex; //Gender
                                string strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                                dtDataRow [30] = strLifeID;//life ID 

                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                                dtDataRow [39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells [intLoop, 7].Value)); //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [79] = wsraw.Cells [intLoop, 5].Value; //life issue age 
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 9].Value), out decimal dclPremExtra);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 11].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Value), out decimal dclPremPDD);


                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out decimal dclCededLife);//COL N 
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value), out decimal dclCededWPD);//COL O
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 16].Value), out decimal dclCedeADB);//COL P
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 17].Value), out decimal dclCededPDD);//COL Q
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 18].Value), out decimal dclSAR);//COL R


                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow05 = objdt_template.NewRow();
                                _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;


                            #region Premiums and SaR
                            if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [25] = dclCededLife; // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [26] = dclCedeADB;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "WPD"; // Branded Product
                                dtDataRow [59] = dclPremWPD;
                                dtDataRow [26] = dclCededWPD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "PDD"; // Branded Product
                                dtDataRow [59] = dclPremPDD;
                                dtDataRow [26] = dclCededPDD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [26] = 0;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "PDD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;

                                _var.dtworkRow02 [5] = "ADB";
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD";
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// initial sum
                                _var.dtworkRow03 [26] = dclCedeADB;
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow03 [28] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;


                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD; ;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "PDD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremPDD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);


                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);

                                _var.dtworkRow05 [5] = "PDD"; // Branded Product
                                _var.dtworkRow05 [59] = dclPremPDD;
                                _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow05 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow05);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                            }
                            #endregion

                            dclTotalSumAtRisk += dclSAR;
                            dclTotalPremium += dclPremLife + dclPremExtra + dclPremADB + dclPremWPD + dclPremPDD;
                        }

                        }
                }
                else
                {
                    if (str_sheet.Trim() == "Peso-New Business" && str_raw.ToUpper().Contains("REINSURANCE 3Q"))
                    {
                        for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                        {
                            string strPolicyNo = wsraw.Cells[intLoop, 2].Text.ToString();
                            if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                            {
                                continue;
                            }
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            dtDataRow[0] = strPolicyNo;
                            string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                            dtDataRow[23] = strCurrency; //  Cession Currency
                            dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                            string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                            dtDataRow[21] = strTcode; // Transcode
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[14] = "T";
                            string strIssueDate = Convert.ToDateTime(wsraw.Cells[intLoop, 1].Value).ToString("MM/dd/yyyy");
                            dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                            dtDataRow[19] = valueTransEffectiveDate;  // Reinsurance Start Date
                            string strFullName = Convert.ToString(wsraw.Cells[intLoop, 4].Value);
                            dtDataRow[31] = strFullName; //Full Name
                            //objHlpr2.fn_separateLastNameFirstNameV2(strFullName, out strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            objHlpr2.fn_separateLastNameFirstNameV6(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            dtDataRow [33] = objHlpr2.fn_checkFirstname(strFirstName);
                            dtDataRow[32] = objHlpr2.fn_checkLastname(strLastName);
                            dtDataRow[34] = strMiddleInitial;
                            string strDOB = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Cells[intLoop, 5].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = objHlpr2.fn_checkDOB(strDOB); //Birthday
                            string strSex = wsraw.Cells[intLoop, 7].Value;
                            dtDataRow[36] = strSex; //Gender
                            string strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                            dtDataRow[30] = strLifeID;//life ID 
                            dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                            dtDataRow[39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells[intLoop, 8].Value)); //mortality rating
                            dtDataRow[56] = "4000"; // Entry code
                            dtDataRow[79] = wsraw.Cells[intLoop, 6].Value; //life issue age 
                            objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                            dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks
                            decimal dclPremLife = Convert.ToDecimal(wsraw.Cells[intLoop, 9].Value);
                            decimal dclPremExtra = Convert.ToDecimal(wsraw.Cells[intLoop, 10].Value);
                            decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 11].Value);
                            decimal dclPremADB = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);
                            decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);

                            decimal dclCededLife = Convert.ToDecimal(wsraw.Cells [intLoop, 15].Value); //life
                            decimal dclCededWPD = Convert.ToDecimal(wsraw.Cells [intLoop, 16].Value); //wpd
                            decimal dclCedeADB = Convert.ToDecimal(wsraw.Cells [intLoop, 17].Value); //adb
                            decimal dclCededPDD = Convert.ToDecimal(wsraw.Cells [intLoop, 18].Value); //pdd
                            decimal dclSAR = Convert.ToDecimal(wsraw.Cells [intLoop, 19].Value); //nar

                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow04 = objdt_template.NewRow();
                            _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow05 = objdt_template.NewRow();
                            _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;


                            #region Premiums and SaR
                            if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [25] = dclCededLife; // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [26] = dclCedeADB;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "WPD"; // Branded Product
                                dtDataRow [57] = dclPremWPD;
                                dtDataRow [26] = dclCededWPD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "PDD"; // Branded Product
                                dtDataRow [57] = dclPremPDD;
                                dtDataRow [26] = dclCededPDD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [26] = 0;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "PDD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;

                                _var.dtworkRow02 [5] = "ADB";
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD";
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [57] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// initial sum
                                _var.dtworkRow03 [26] = dclCedeADB;
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow03 [28] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;


                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [57] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [57] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD; ;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "PDD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremPDD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);


                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [57] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);

                                _var.dtworkRow05 [5] = "PDD"; // Branded Product
                                _var.dtworkRow05 [57] = dclPremPDD;
                                _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow05 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow05);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [57] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [57] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [57] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                            }
                            #endregion
                            dclTotalSumAtRisk += dclSAR;
                            dclTotalPremium += dclPremLife + dclPremExtra + dclPremADB + dclPremWPD + dclPremPDD;



                        }

                    }
                    else if (str_raw.ToUpper().Contains("REINSURANCE 3Q") || str_sheet.ToUpper().Contains("RENEWALS"))
                    {
                        for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                        {
                            string strPolicyNo = wsraw.Cells[intLoop, 2].Text.ToString();
                            if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                            {
                                continue;
                            }
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            dtDataRow[0] = strPolicyNo;
                            string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                            dtDataRow[23] = strCurrency; //  Cession Currency
                            dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                            string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                            dtDataRow[21] = strTcode; // Transcode
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[14] = "T";

                            string strIssueDate = Convert.ToDateTime(wsraw.Cells[intLoop, 1].Value).ToString("MM/dd/yyyy");
                            dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode,Variables.strBmYear,  strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                            dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                            dtDataRow[19] = valueTransEffectiveDate;  // Reinsurance Start Date
                            string strFullName = Convert.ToString(wsraw.Cells[intLoop, 4].Value);
                            dtDataRow[31] = strFullName; //Full Name
                            objHlpr2.fn_separateLastNameFirstNameV6(strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            //objHlpr2.fn_separateLastNameFirstNameV2(strFullName, out strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            dtDataRow [33] = objHlpr2.fn_checkFirstname(strFirstName);
                            dtDataRow[32] = objHlpr2.fn_checkLastname(strLastName);
                            dtDataRow[34] = strMiddleInitial;
                            string strDOB = "";
                            dtDataRow[37] = objHlpr2.fn_checkDOB(strDOB); //Birthday
                            string strSex = wsraw.Cells[intLoop, 6].Value;
                            dtDataRow[36] = strSex; //Gender
                            string strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                            dtDataRow[30] = strLifeID;//life ID 

                            dtDataRow[38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                            dtDataRow[39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Cells[intLoop, 7].Value)); //mortality rating
                            dtDataRow[58] = "4001"; // Entry code
                            dtDataRow[79] = wsraw.Cells[intLoop, 5].Value; //life issue age 
                            objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                            dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks


                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremLife);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Value), out decimal dclPremExtra);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 11].Value), out decimal dclPremWPD);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Value), out decimal dclPremADB);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out decimal dclPremPDD);


                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out decimal dclCededLife);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value), out decimal dclCededWPD);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 16].Value), out decimal dclCedeADB);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 17].Value), out decimal dclCededPDD);
                            decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 18].Value), out decimal dclSAR);

                         
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow04 = objdt_template.NewRow();
                            _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow05 = objdt_template.NewRow();
                            _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;


                            #region Premiums and SaR
                            if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [25] = dclCededLife; // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [26] = dclCedeADB;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "WPD"; // Branded Product
                                dtDataRow [59] = dclPremWPD;
                                dtDataRow [26] = dclCededWPD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "PDD"; // Branded Product
                                dtDataRow [59] = dclPremPDD;
                                dtDataRow [26] = dclCededPDD;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [26] = 0;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [26] = dclCededLife;//ceded retention;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife)); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); // Initial Sum at Risk

                                _var.dtworkRow02 [5] = "PDD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [26] = dclCededWPD;//ceded retention;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [26] = 0;//ceded retention;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;

                                _var.dtworkRow02 [5] = "ADB";
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [26] = dclCedeADB;//ceded retention;
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Original Sum Assured
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB)); // Initial Sum at Risk
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig SUm;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD";
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "PDD";
                                _var.dtworkRow02 [59] = dclPremPDD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow02 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }

                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// initial sum
                                _var.dtworkRow03 [26] = dclCedeADB;
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// initial sum
                                _var.dtworkRow03 [28] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;


                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// orig sum
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// initial sum
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// orig sum
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// initial sum
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product
                                dtDataRow [59] = dclPremExtra;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = 0;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = "ADB"; // Branded Product
                                dtDataRow [59] = dclPremADB;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCedeADB;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "WPD"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremWPD;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow02 [26] = dclCededWPD; ;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "PDD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremPDD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);

                                _var.dtworkRow04 [5] = "PDD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremPDD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremExtra;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Orig sum;
                                _var.dtworkRow02 [26] = 0;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "ADB"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremADB;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow03 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);


                                _var.dtworkRow04 [5] = "WPD"; // Branded Product
                                _var.dtworkRow04 [59] = dclPremWPD;
                                _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow04 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow04);

                                _var.dtworkRow05 [5] = "PDD"; // Branded Product
                                _var.dtworkRow05 [59] = dclPremPDD;
                                _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// sum at risk
                                _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));// Initial Sum
                                _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededPDD));//Orig sum;
                                _var.dtworkRow05 [26] = dclCededPDD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow05);
                            }
                            else if(dclPremLife != 0 && dclPremExtra == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                                dtDataRow [26] = dclCededLife;//ceded retention;

                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                _var.dtworkRow02 [59] = dclPremADB;
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// sum at risk
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));// Initial Sum
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedeADB));//Orig sum;
                                _var.dtworkRow02 [26] = dclCedeADB;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WPD"; // Branded Product
                                _var.dtworkRow03 [59] = dclPremWPD;
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// sum at risk
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));// Initial Sum
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededWPD));//Orig sum;
                                _var.dtworkRow03 [26] = dclCededWPD;// ceded retention
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }

                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                            {
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [59] = dclPremLife;
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// sum at risk
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));// Initial Sum
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCededLife));//Orig sum;
                            }
                            #endregion
                            dclTotalSumAtRisk += dclSAR;
                            dclTotalPremium += dclPremLife + dclPremExtra + dclPremADB + dclPremWPD + dclPremPDD;
                        }

                    }
                } 

            }
            #endregion

            //UMRE Files/Accounts
            else
            {
                #region Quarter1 ,Quarter2,Quarter3, Quarter4
                if(str_sheet == "Peso Renewals" || (str_sheet == "Renewals-Peso")) //Q1 , Q2
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {

                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);

                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }

                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                objHlpr.fn_checksheetname(str_sheet, out int NarColNumber);
                                decimal dclSARLife = 0;
                                if(NarColNumber == 12)
                                {
                                    dclSARLife = Convert.ToDecimal(wsraw.Cells [intLoop, 12].Value);
                                }
                                else if(NarColNumber == 11)
                                {
                                    dclSARLife = Convert.ToDecimal(wsraw.Cells [intLoop, 11].Value);
                                }

                                string strFullName = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleInitial, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                //objHlpr.fn_getBusinessType(strCessionCode, out strCessionCode);
                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                //dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                //dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleInitial;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [79] = strIssueAge; //Issue Age

                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out decimal dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value),out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion

                            }
                        }
                    }

                }
                else if(str_sheet == "Dollar Renewals" || (str_sheet == "Renewals-Dollar"))  //Q1 and Q2 Workbook
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {
                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {

                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type

                                objHlpr.fn_checksheetname(str_sheet, out int NarColNumber);
                                decimal dclSARLife = 0;
                                if(NarColNumber == 12)
                                {
                                    dclSARLife = Convert.ToDecimal(wsraw.Cells [intLoop, 12].Value);
                                }
                                else if(NarColNumber == 11)
                                {
                                    dclSARLife = Convert.ToDecimal(wsraw.Cells [intLoop, 11].Value);
                                }
                                string strFullName = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                //dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                //dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                //dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [26] = 0;

                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out decimal  dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value), out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion
                            }
                        }
                    }

                }   //Q1 and Q2 Workbook
                else if(str_sheet == "DOLLAR RENEWALS" && str_raw.ToUpper().Contains("3Q")) //Q3 and Q4 Workbook
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {
                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {

                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                //objHlpr.fn_CheckingforA_AB_BZColumn(null, null, Convert.ToString(wsraw.Cells[intLoop, 21].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                                string strFullName = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; //Birthday
                                dtDataRow [36] = strSex; //Gender
                                dtDataRow [30] = strLifeID;//life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [26] = 0; // Entry code

                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Value), out decimal dclSARLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 21].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 22].Value), out decimal dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 23].Value), out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion
                            }
                        }
                    }
                }
                else if(str_sheet == "DOLLAR REN" && str_raw.ToUpper().Contains("4Q")) //Q3 and Q4 Workbook
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {

                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {

                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business 
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                //objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Cells[intLoop, 20].Value), Convert.ToString(wsraw.Cells[intLoop, 20].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                                string strFullName = wsraw.Cells [intLoop, 4].Value;
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; //Birthday
                                dtDataRow [36] = strSex; //Gender
                                dtDataRow [30] = strLifeID;//life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); //Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [26] = 0;

                                dtDataRow [79] = strIssueAge; //Issue Age

                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; //Remarks

                                //decimal dclPremLife = Convert.ToDecimal(wsraw.Cells [intLoop, 5].Value);
                                //decimal dclPremADB = Convert.ToDecimal(wsraw.Cells [intLoop, 6].Value);
                                //decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells [intLoop, 7].Value);
                                //decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells [intLoop, 8].Value);
                                //decimal dclSARLife = Convert.ToDecimal(wsraw.Cells [intLoop, 20].Value);
                                //decimal dclSARADB = Convert.ToDecimal(wsraw.Cells [intLoop, 21].Value);
                                //decimal dclSARACC = Convert.ToDecimal(wsraw.Cells [intLoop, 22].Value);
                                //decimal dclSARPDD = Convert.ToDecimal(wsraw.Cells [intLoop, 23].Value);

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Value), out decimal dclSARLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 21].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 22].Value), out decimal dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 23].Value), out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion
                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion
                            }
                        }
                    }

                }
                else if(str_sheet == "PESO RENEWALS" && (str_raw.ToUpper().Contains("3Q")))
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {

                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                //objHlpr.fn_CheckingforA_AB_BZColumn(null, null, Convert.ToString(wsraw.Cells[intLoop, 20].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                                string strFullName = wsraw.Cells [intLoop, 4].Value;
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [26] = 0;
                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                          
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Value), out decimal dclSARLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 21].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 22].Value), out decimal dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 23].Value), out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion
                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion

                            }
                        }
                    }

                }  //Q3 Workbook
                else if(str_sheet == "PESO RENEWALS" && (str_raw.ToUpper().Contains("4Q"))) //Q4
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {
                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {

                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet);
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                //objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Cells[intLoop, 20].Value), Convert.ToString(wsraw.Cells[intLoop, 20].Value), out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);
                                string strFullName = wsraw.Cells [intLoop, 4].Value;

                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [58] = "4001"; // Entry code
                                dtDataRow [26] = 0;

                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Value), out decimal dclSARLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 21].Value), out decimal dclSARADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 22].Value), out decimal dclSARACC);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 23].Value), out decimal dclSARPDD);

                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }

                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [26] = 0; //ceded sum


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC));
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARACC)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));// sum at risk
                                        dtDataRow [14] = CessionCode;//Business Type;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARPDD)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [14] = CessionCode;
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB));
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Original Sum Assured
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARADB)); // Initial Sum at Risk
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife));// sum at risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Original Sum Assured
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARLife)); // Initial Sum at Risk
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                }
                                #endregion

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += dclSARLife;
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += dclSARLife;
                                }
                                #endregion

                            }
                        }
                    }
                }
                else if(str_sheet == "Adjustments-Peso") //Q2
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {
                        if(wsraw.Cells [intLoop, 4].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 4].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //Cession Currency
                                dtDataRow [23] = strCurrency; //  Cession Currency
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Years
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                objHlpr.fn_CheckingforA_AB_BZColumn(null, null, null, out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);

                                string strFullName = wsraw.Cells [intLoop, 4].Value;
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date
                                                                           //dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [62] = "4002"; // Entry code
                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                                //decimal dclPremLife = Convert.ToDecimal(wsraw.Cells [intLoop, 5].Value);
                                //decimal dclPremADB = Convert.ToDecimal(wsraw.Cells [intLoop, 6].Value);
                                //decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells [intLoop, 7].Value);
                                //decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells [intLoop, 8].Value);

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);


                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                #region Premiums and SaR
                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow04 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }


                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [59] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow04 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [59] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [59] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [59] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [59] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [59] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }

                                }
                                #endregion

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += Convert.ToDecimal(strSumAtRisk);
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += Convert.ToDecimal(strSumAtRisk);
                                }
                                #endregion
                            }
                        }
                    }

                }  // Q1 Workbook
                else if(str_sheet == "Peso Adjustments") //Adjustments
                {
                    for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                    {
                        if(wsraw.Cells [intLoop, 5].Value != null)
                        {
                            string BusinessType = Convert.ToString(wsraw.Cells [intLoop, 5].Value);
                            if(BusinessType.GetType() == typeof(string))
                            {
                                BusinessType = BusinessType.Replace(" ", String.Empty);
                                if(BusinessType == "FACULTATIVE")
                                {
                                    CessionCode = "F";
                                }
                                else if(BusinessType == "AUTOMATIC")
                                {
                                    CessionCode = "T";
                                }
                            }
                            string strPolicyNo = wsraw.Cells [intLoop, 2].Text.ToString();
                            if(Regex.IsMatch(strPolicyNo, @"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"))
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);
                                dtDataRow [0] = strPolicyNo;
                                dtDataRow [5] = wsraw.Cells [intLoop, 3].Value; // Branded Product
                                string strCurrency = objHlpr2.fn_getcurrencyV2(str_sheet); //  Cession Currency
                                dtDataRow [23] = strCurrency;
                                dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Years
                                string strTcode = objHlpr.fn_CheckTransCodeV2(str_sheet);
                                dtDataRow [21] = strTcode; // Transcode
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PAFM"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "YLY"; // Premium Frequency
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                objHlpr.fn_CheckingforA_AB_BZColumn(null, null, null, out strOriginalSum, out strInitialSum, out strSumAtRisk, out strRemarksAABBZ);

                                string strFullName = wsraw.Cells [intLoop, 5].Value;
                                objHlpr2.fn_getFirstFinancialMacrodata(strPolicyNo, strFullName, out string strIssueAge,
                                out string strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                                out string strMiddleName, out string strTitle, out string strDOB, out string strSex, out string strLifeID, out string strRcDummyName,
                                out string strLife);

                                //objHlpr.fn_getBusinessType(strCessionCode, out strCessionCode);
                                string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 1].Value).ToString("MM/dd/yyyy");
                                dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                                dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                                                                                                     //dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                                dtDataRow [31] = strFullName; //Full Name
                                dtDataRow [33] = strFirstName;
                                dtDataRow [32] = strLastName;
                                dtDataRow [34] = strMiddleName;
                                strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strDOB; // Birthday
                                dtDataRow [36] = strSex; // Gender
                                dtDataRow [30] = strLifeID;// life ID 
                                dtDataRow [38] = objHlpr.fn_SmokerCode(strSmoker); // Smoker Status
                                dtDataRow [39] = strMortality; //mortality rating
                                dtDataRow [62] = "4004"; // Entry code
                                dtDataRow [79] = strIssueAge; //Issue Age
                                objHlpr.fn_GetRemarksCode(strDOB, strFullName, strSex, out string strRemarksCode);
                                dtDataRow [76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks

                                //decimal dclPremLife = Convert.ToDecimal(wsraw.Cells [intLoop, 6].Value);
                                //decimal dclPremADB = Convert.ToDecimal(wsraw.Cells [intLoop, 7].Value);
                                //decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells [intLoop, 8].Value);
                                //decimal dclPremPDD = Convert.ToDecimal(wsraw.Cells [intLoop, 9].Value);

                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 5].Value), out decimal dclPremLife);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 6].Value), out decimal dclPremADB);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclPremWPD);
                                decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 8].Value), out decimal dclPremPDD);


                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow03 = objdt_template.NewRow();
                                _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow04 = objdt_template.NewRow();
                                _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;

                                if(CessionCode == "T")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [63] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow04 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }


                                }
                                else if(CessionCode == "F")
                                {
                                    if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;



                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremADB;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremWPD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);


                                        _var.dtworkRow04 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow04 [63] = dclPremPDD;
                                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow04 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;

                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);


                                        _var.dtworkRow03 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow03 [63] = dclPremPDD;
                                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow03 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLife == 0 && dclPremADB != 0 && dclPremWPD == 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremADB;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;
                                    }
                                    else if(dclPremLife == 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD != 0)
                                    {

                                        dtDataRow [63] = dclPremWPD;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;//Business Type;


                                        _var.dtworkRow02 [14] = CessionCode;//Business Type;
                                        _var.dtworkRow02 [63] = dclPremPDD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }
                                    else if(dclPremLife != 0 && dclPremADB == 0 && dclPremWPD != 0 && dclPremPDD == 0)
                                    {

                                        dtDataRow [63] = dclPremLife;
                                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk);// sum at risk
                                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Initial Sum at Risk
                                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        dtDataRow [26] = 0; // ceded sum
                                        dtDataRow [14] = CessionCode;


                                        _var.dtworkRow02 [14] = CessionCode;
                                        _var.dtworkRow02 [63] = dclPremWPD;
                                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Initial Sum at Risk
                                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); // Original Sum Assured
                                        _var.dtworkRow02 [26] = 0; // ceded sum
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                    }

                                }

                                #region HASH TOTALS
                                if(strCurrency.ToUpper() == "PHP" && CessionCode.ToUpper() == "T" || strCurrency.ToUpper() == "USD" && CessionCode.ToUpper() == "T")
                                {
                                    dclTreatyTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclTreatyTotalSumAtRisk += Convert.ToDecimal(strSumAtRisk);
                                }
                                else
                                {
                                    dclFaculTotalPremium += dclPremLife + dclPremADB + dclPremWPD + dclPremPDD;
                                    dclFaculTotalSumAtRisk += Convert.ToDecimal(strSumAtRisk);
                                }
                                #endregion


                            }
                        }
                    }

                }  // Q1 Workbook

                #endregion

            }

            #region Computing Hash 
            if (str_raw.ToUpper().Contains("NRCP"))
            {
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
            }
            else
            {
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Treaty Total Premium:";
                dtDataRow[1] = dclTreatyTotalPremium;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Treaty Total Sum at Risk:";
                dtDataRow[1] = dclTreatyTotalSumAtRisk;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Falcultative Total Premium:";
                dtDataRow[1] = dclFaculTotalPremium;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Facultative Total Sum at Risk:";
                dtDataRow[1] = dclFaculTotalSumAtRisk;
                objdt_template.Rows.Add(dtDataRow);
                #endregion
            }

            if (Variables.boogenderfail)
            {

                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);
                //objdt_template.Rows.Add(dtDataRow);
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Please check for blank genders";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }

            //if (Variables.boomacrofail)
            //{

            //    dtDataRow = objdt_template.NewRow();
            //    objdt_template.Rows.Add(dtDataRow);
            //    //objdt_template.Rows.Add(dtDataRow);
            //    _var.dtworkRow01 = objdt_template.NewRow();
            //    _var.dtworkRow01[0] = "Please check policies with dummy data. these are data that has no record in Macro Database";
            //    objdt_template.Rows.Add(_var.dtworkRow01);
            //}

            string despath = str_saved + @"\BM048-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);
            Variables.strBmYear = string.Empty;

            dclFaculTotalPremium = 0;
            dclFaculTotalSumAtRisk = 0;
            dclTreatyTotalPremium = 0;
            dclTreatyTotalSumAtRisk = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}
