    using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM099
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false, string str_policyYear = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets[str_sheet];

            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 3].End[XlDirection.xlUp].row;
            string valueTransEffectiveDate = string.Empty; string strTcode = string.Empty;
            DataRow dtDataRow;
            double dblTotalPremiumLife = 0,  dblTotalSumAtRiskLife = 0, dbTotalCommissionLife = 0;
            double dblTotalPremiumADB = 0, dblTotalSumAtRiskADB = 0, dbTotalCommissionADB = 0;
            double dblTotalPremiumWP = 0, dblTotalSumAtRiskWP = 0, dbTotalCommissionWP = 0;
            double dblTotalPremiumAEH = 0, dblTotalSumAtRiskAEH = 0, dbTotalCommissionAEH = 0;
            double dblTotalPremiumPA = 0, dblTotalSumAtRiskPA = 0, dbTotalCommissionPA = 0;
            double dblTotalPremiumCIR = 0, dblTotalSumAtRiskCIR = 0, dbTotalCommissionCIR = 0;
            double dblTotalPremiumENCI = 0, dblTotalSumAtRiskENCI = 0, dblTotalCommissionENCI;
            double dblTotalPremiumESCI = 0, dblTotalSumAtRiskESCI = 0, dblTotalCommissionESCI;

            decimal dblTotalPremium = 0, dbltotalSumAtRisk = 0;
            string strFilePath = wbraw.FullName;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
                newform.ShowDialog();

            }
      

            try
            {
                // PESO
                if (strFilePath.ToUpper().Contains("PESO"))
                {
                    //BM099
                    if (str_sheet.ToUpper().Contains("PREM"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["B" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["B" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = Convert.ToString(wsraw.Range ["C" + i].Value);
                                    dtDataRow [32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = objHlpr2.fn_checkDOB(null); // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    dtDataRow[78] = wsraw.Range["E" + i].Value; // Attain Age
                                    dtDataRow[80] = wsraw.Range["J" + i].Value + wsraw.Range["M" + i].Value; // Reinsurance Commission
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "PHP"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // Preferred Classific

                                    string strIssueDate = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date
                                    objHlpr2.fn_getTransReinsuranceDateV4(strIssueDate, Variables.strBmYear, out string transEffectiveDate, out string transCode);
                                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                                    dtDataRow [20] = strIssueDate;//Policy Start Date
                                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date

                                    if (transCode == "TNEWBUS")
                                    {
                                        dtDataRow [21] = transCode;// Transcode
                                        dtDataRow [56] = "4000"; // Entry Code
                                    }
                                    else
                                    {
                                        dtDataRow [21] = transCode; // Transcode
                                        dtDataRow [58] = "4001";
                                    }

                                    //_var.dtworkRow02 = objdt_template.NewRow();
                                    //_var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow03 = objdt_template.NewRow();
                                    //_var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow04 = objdt_template.NewRow();
                                    //_var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow05 = objdt_template.NewRow();
                                    //_var.dtworkRow05.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow06 = objdt_template.NewRow();
                                    //_var.dtworkRow06.ItemArray = dtDataRow.ItemArray;

                                    #region Sum At Risk & Intial Sum
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["K" + i].Value), out double dclADB);

                                    double.TryParse(Convert.ToString(wsraw.Range ["N" + i].Value), out double dclAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["P" + i].Value), out double dclPA);
                                    double.TryParse(Convert.ToString(wsraw.Range ["R" + i].Value), out double dclCIR);
                                    #endregion

                                    #region Premium
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["I" + i].Value), out double dclPremWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["L" + i].Value), out double dclPremADB);
                                    double.TryParse(Convert.ToString(wsraw.Range ["O" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Value), out double dclPremPA);
                                    double.TryParse(Convert.ToString(wsraw.Range ["S" + i].Value), out double dclPremCIR);
                                    #endregion

                                    #region Comission
                                    double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Value), out double dclCommWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["M" + i].Value), out double dclCommADB);
                                    #endregion

                                    if(transCode == "TNEWBUS")
                                    {
                                        #region Premiums
                                        if(dclLIFE != 0 || dclLIFE == 0 && dclPremLIFE != 0)
                                        {
                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremLIFE;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else if(dclWP != 0 || dclWP == 0 && dclPremWP != 0)
                                        {
                                            dtDataRow [5] = "WP/PB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremWP;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                            dtDataRow [80] = dclCommWP;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion


                                        }
                                        else if(dclADB != 0 || dclADB == 0 && dclPremADB != 0)
                                        {
                                            dtDataRow [5] = "ADB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremADB;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                            dtDataRow [80] = dclCommADB;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }
                                        else if(dclAEH != 0 || dclAEH == 0 && dclPremAEH != 0)
                                        {
                                            dtDataRow [5] = "A&H";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremAEH;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion


                                        }
                                        else if(dclPA != 0 || dclPA == 0 && dclPremPA != 0)
                                        {
                                            dtDataRow [5] = "PA";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremPA;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else if(dclCIR != 0 || dclCIR == 0 && dclPremCIR != 0)
                                        {
                                            dtDataRow [5] = "CIR";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremCIR;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else
                                        {

                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremLIFE;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [80] = 0;
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region Premiums
                                        if(dclLIFE != 0 || dclLIFE == 0 && dclPremLIFE != 0)
                                        {
                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremLIFE;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else if(dclWP != 0 || dclWP == 0 && dclPremWP !=0)
                                        {
                                            dtDataRow [5] = "WP/PB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremWP;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                            dtDataRow [80] = dclCommWP;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion


                                        }
                                        else if(dclADB != 0 || dclADB ==0 && dclPremADB != 0)
                                        {
                                            dtDataRow [5] = "ADB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremADB;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                            dtDataRow [80] = dclCommADB;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }
                                        else if(dclAEH != 0 || dclAEH == 0 && dclPremAEH != 0)
                                        {
                                            dtDataRow [5] = "A&H";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremAEH;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion


                                        }
                                        else if(dclPA != 0 || dclPA == 0 && dclPremPA != 0)
                                        {
                                            dtDataRow [5] = "PA";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremPA;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else if(dclCIR != 0 || dclCIR == 0 && dclPremCIR != 0)
                                        {
                                            dtDataRow [5] = "CIR";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremCIR;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));//sum at risk
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));//sum at risk
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));//sum at risk
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            
                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremLIFE;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [80] = 0;
                                        }
                                        #endregion
                                    }


                                    #region HashTotal
                                    if (dclLIFE > 1)
                                    {
                                        dblTotalSumAtRiskLife += dclLIFE;
                                    }
                                    if(dclADB > 1)
                                    {
                                        dblTotalSumAtRiskADB += dclADB;
                                    }
                                    if(dclWP > 1)
                                    {
                                        dblTotalSumAtRiskWP += dclWP;
                                    }
                                    if(dclAEH > 1)
                                    {
                                        dblTotalSumAtRiskAEH += dclAEH;
                                    }
                                    if (dclCIR > 1)
                                    {
                                        dblTotalSumAtRiskCIR += dclCIR;
                                    }
                                    if(dclPA > 1)
                                    {
                                        dblTotalSumAtRiskPA += dclPA; 
                                    }
                                    dblTotalPremiumLife += dclPremLIFE; 
                                    dblTotalPremiumADB += dclPremADB; 
                                    dblTotalPremiumWP += dclPremWP; 
                                    dblTotalPremiumAEH += dclPremAEH; 
                                    dblTotalPremiumCIR += dclPremCIR; 
                                    dblTotalPremiumPA += dclPremPA; 
                                    #endregion
                                }
                            }
                        }
                    }
                    //BM099A
                    else if (str_sheet.ToUpper().Contains("CRQ"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["A" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["A" + i].Value; // Policy Number
                                    //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                    dtDataRow[32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = "07/01/1900"; // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    //dtDataRow[79] = wsraw.Range["H" + i].Value; // Life Issue Age
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business 
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "PHP"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(null); // Preferred Classific

                                    dtDataRow[21] = "ADJUST"; // Transcode
                                    dtDataRow[62] = "4004"; // Entry Code

                                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orig sum
                                    dtDataRow [26] = 1;//ceded sum
                                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//initial sum
                                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//sum at risk
                                    //dtDataRow[63] = wsraw.Range["I" + i].Value; // Premium

                                    dtDataRow [22] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Trans Effective Date
                                    dtDataRow [19] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Reinsurance
                                    dtDataRow [20] = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date
                                    
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclPremCIR);

                                    dclPremLIFE = dclPremLIFE * -1;
                                    dclPremAEH = dclPremAEH * -1;
                                    dclPremCIR = dclPremCIR * -1;

                                    _var.dtworkRow02 = objdt_template.NewRow();
                                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow03 = objdt_template.NewRow();
                                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;

                                    #region premiums
                                    if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [63] = dclPremCIR;
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = 0;
                                    }
                                    #endregion

                                    #region Hashtotals
                                    dblTotalPremiumLife += dclPremLIFE;
                                    dblTotalSumAtRiskLife = 0;
                                    dblTotalPremiumAEH += dclPremAEH;
                                    dblTotalSumAtRiskAEH = 0;
                                    dblTotalPremiumCIR += dclPremCIR;
                                    dblTotalSumAtRiskCIR = 0;
                                    #endregion

                                    
                                }
                            }
                        }
                    }
                    //BM099A
                    else if (str_sheet.ToUpper().Contains("CFQ"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["A" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["A" + i].Value; // Policy Number
                                    //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                    dtDataRow[32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = "07/01/1900"; // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    //dtDataRow[79] = wsraw.Range["H" + i].Value; // Life Issue Age
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "PHP"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow[39] = "STANDARD"; // Preferred Classific
                                    dtDataRow[21] = "ADJUST"; // Transcode
                                    dtDataRow[60] = "4002"; // Entry Code
                                    //dtDataRow[61] = wsraw.Range["I" + i].Value; // Premium
                                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(null); // Preferred Classific

                                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orig sum   
                                    dtDataRow [26] = 1;//ceded sum
                                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//initial sum
                                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//sum at risk

                                    
                                    dtDataRow [22] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Trans Effective Date
                                    dtDataRow [19] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Reinsurance
                                    dtDataRow [20] = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date
                                    //dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["I" + i].Value);
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclPremCIR);

                                    dclPremLIFE = dclPremLIFE * -1;
                                    dclPremAEH = dclPremAEH * -1;
                                    dclPremCIR = dclPremCIR * -1;

                                    _var.dtworkRow02 = objdt_template.NewRow();
                                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow03 = objdt_template.NewRow();
                                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;

                                    #region premiums
                                    if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [61] = dclPremAEH;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [61] = dclPremCIR;
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [61] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = 0;
                                    }
                                    #endregion

                                    #region Hashtotals
                                    dblTotalPremiumLife += dclPremLIFE;
                                    dblTotalSumAtRiskLife = 0;
                                    dblTotalPremiumAEH += dclPremAEH;
                                    dblTotalSumAtRiskAEH = 0;
                                    dblTotalPremiumCIR += dclPremCIR;
                                    dblTotalSumAtRiskCIR = 0;
                                    #endregion
                                }   
                            }
                        }
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 099", "Information");
                        return "";
                    }
                }
                // DolDOLLAR
                else if (strFilePath.ToUpper().Contains("DOLLAR"))
                {
                    //BM099
                    if (str_sheet.ToUpper().Contains("PREM"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["B" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["B" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                                    //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                    Console.WriteLine(strLastName);
                                    dtDataRow[32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = "07/01/1900"; // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    dtDataRow[78] = wsraw.Range["E" + i].Value; // Attain Age
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "USD"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow[39] = "STANDARD"; // Preferred Classific

                                    string strIssueDate = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date
                                    objHlpr2.fn_getTransReinsuranceDateV4(strIssueDate, Variables.strBmYear, out string transEffectiveDate, out string transCode);
                                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                                    dtDataRow [20] = strIssueDate;//Policy Start Date
                                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date

                                    if(transCode == "TNEWBUS")
                                    {
                                        dtDataRow [21] = transCode;// Transcode
                                        dtDataRow [56] = "4000"; // Entry Code
                                    }
                                    else
                                    {
                                        dtDataRow [21] = transCode; // Transcode
                                        dtDataRow [58] = "4001";
                                    }

                                    #region Sum At Risk & Intial Sum
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["K" + i].Value), out double dclADB);
                                    double.TryParse(Convert.ToString(wsraw.Range ["N" + i].Value), out double dclAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["P" + i].Value), out double dclPA);
                                    double.TryParse(Convert.ToString(wsraw.Range ["R" + i].Value), out double dclCIR);
                                    double.TryParse(Convert.ToString(wsraw.Range ["T" + i].Value), out double dclENCI);
                                    double.TryParse(Convert.ToString(wsraw.Range ["V" + i].Value), out double dclESCI);
                                    #endregion

                                    #region Premium
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["I" + i].Value), out double dclPremWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["L" + i].Value), out double dclPremADB);
                                    double.TryParse(Convert.ToString(wsraw.Range ["O" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Value), out double dclPremPA);
                                    double.TryParse(Convert.ToString(wsraw.Range ["S" + i].Value), out double dclPremCIR);
                                    double.TryParse(Convert.ToString(wsraw.Range ["U" + i].Value), out double dclPremENCI);
                                    double.TryParse(Convert.ToString(wsraw.Range ["W" + i].Value), out double dclPremESCI);
                                    #endregion

                                    #region Comission
                                    double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Value), out double dclCommWP);
                                    double.TryParse(Convert.ToString(wsraw.Range ["M" + i].Value), out double dclCommADB);
                                    #endregion

                                    //_var.dtworkRow02 = objdt_template.NewRow();
                                    //_var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow03 = objdt_template.NewRow();
                                    //_var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow04 = objdt_template.NewRow();
                                    //_var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow05 = objdt_template.NewRow();
                                    //_var.dtworkRow05.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow06 = objdt_template.NewRow();
                                    //_var.dtworkRow06.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow07 = objdt_template.NewRow();
                                    //_var.dtworkRow07.ItemArray = dtDataRow.ItemArray;
                                    //_var.dtworkRow08 = objdt_template.NewRow();
                                    //_var.dtworkRow08.ItemArray = dtDataRow.ItemArray;

                                    if(transCode == "TNEWBUS")
                                    {
                                        #region premium validation
                                        if(dclLIFE != 0 || dclLIFE == 0 && dclPremLIFE != 0)
                                        {
                                            #region Subpremiums
                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [57] = dclPremLIFE;
                                            dtDataRow [80] = 0;

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclWP != 0 || dclWP == 0 && dclPremWP != 0)
                                        {
                                            #region Subpremiums
                                            dtDataRow [5] = "WP/PB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                            dtDataRow [57] = dclPremWP;
                                            dtDataRow [80] = dclCommWP;

                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclADB != 0 || dclADB == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "ADB"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremADB;
                                            dtDataRow [80] = dclCommADB;
                                            #region Subpremiums
                                            if(dclLIFE != 0 && dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 && dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclAEH != 0 && dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 && dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 && dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }

                                        else if(dclAEH != 0 || dclAEH == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "A&H"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremAEH;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 && dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclPA != 0 || dclPA == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "PA"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremPA;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }

                                        else if(dclCIR != 0 || dclCIR ==0 && dclPremCIR != 0)
                                        {
                                            dtDataRow [5] = "CIR"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremCIR;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclENCI != 0 || dclENCI == 0 && dclPremENCI != 0)
                                        {
                                            dtDataRow [5] = "EENCI"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremENCI;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclESCI != 0 || dclESCI == 0 && dclPremESCI != 0)
                                        {
                                            dtDataRow [5] = "ESCI"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [57] = dclPremESCI;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [57] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [57] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region premium validation
                                        if(dclLIFE != 0 || dclLIFE == 0 && dclPremLIFE != 0)
                                        {
                                            #region Subpremiums
                                            dtDataRow [5] = "LIFE";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                            dtDataRow [59] = dclPremLIFE;
                                            dtDataRow [80] = 0;

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclWP != 0 || dclWP == 0 && dclPremWP != 0)
                                        {
                                            #region Subpremiums
                                            dtDataRow [5] = "WP/PB";
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));//sum at risk
                                            dtDataRow [59] = dclPremWP;
                                            dtDataRow [80] = dclCommWP;

                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclADB != 0 || dclADB == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "ADB"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremADB;
                                            dtDataRow [80] = dclCommADB;
                                            #region Subpremiums
                                            if(dclLIFE != 0 && dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 && dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclAEH != 0 && dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 && dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 && dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }

                                        else if(dclAEH != 0 || dclAEH == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "A&H"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremAEH;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 && dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclPA != 0 || dclPA == 0 && dclPremADB != 0)
                                        {

                                            dtDataRow [5] = "PA"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremPA;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }

                                        else if(dclCIR != 0 || dclCIR == 0 && dclPremCIR != 0)
                                        {
                                            dtDataRow [5] = "CIR"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremCIR;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclENCI != 0 || dclENCI == 0 && dclPremENCI != 0)
                                        {
                                            dtDataRow [5] = "EENCI"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremENCI;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclESCI != 0 || dclPremESCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ESCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremESCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion
                                        }

                                        else if(dclESCI != 0 || dclESCI == 0 && dclPremESCI != 0)
                                        {
                                            dtDataRow [5] = "ESCI"; // Branded Product
                                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI)); // initial sum
                                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclESCI));// sum at risk
                                            dtDataRow [25] = 1;
                                            dtDataRow [26] = 1;
                                            dtDataRow [59] = dclPremESCI;
                                            dtDataRow [80] = 0;

                                            #region Subpremiums
                                            if(dclLIFE != 0 || dclPremLIFE != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "LIFE";
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclLIFE));//sum at risk
                                                _var.dtworkRow02 [59] = dclPremLIFE;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclWP != 0 || dclPremWP != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclWP));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremWP;
                                                _var.dtworkRow02 [80] = dclCommWP;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclADB != 0 || dclPremADB != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "ADB"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclADB));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremADB;
                                                _var.dtworkRow02 [80] = dclCommADB;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclAEH != 0 || dclPremAEH != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "A&H"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclAEH));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremAEH;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);

                                            }
                                            if(dclPA != 0 || dclPremPA != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "PA"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclPA));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremPA;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            if(dclCIR != 0 || dclPremCIR != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "CIR"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCIR));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremCIR;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }

                                            if(dclENCI != 0 || dclPremENCI != 0)
                                            {
                                                _var.dtworkRow02 = objdt_template.NewRow();
                                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                                _var.dtworkRow02 [5] = "EENCI"; // Branded Product
                                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI)); // initial sum
                                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclENCI));// sum at risk
                                                _var.dtworkRow02 [25] = 1;
                                                _var.dtworkRow02 [26] = 1;
                                                _var.dtworkRow02 [59] = dclPremENCI;
                                                _var.dtworkRow02 [80] = 0;
                                                objdt_template.Rows.Add(_var.dtworkRow02);
                                            }
                                            #endregion

                                        }
                                        #endregion
                                    }

                                    #region HashTotal
                                    dblTotalPremiumLife += dclPremLIFE; dblTotalSumAtRiskLife += dclLIFE;
                                    dblTotalPremiumADB += dclPremADB; dblTotalSumAtRiskADB += dclADB;
                                    dblTotalPremiumWP += dclPremWP; dblTotalSumAtRiskWP += dclWP;
                                    dblTotalPremiumAEH += dclPremAEH; dblTotalSumAtRiskAEH += dclAEH;
                                    dblTotalPremiumCIR += dclPremCIR; dblTotalSumAtRiskCIR += dclCIR;
                                    dblTotalPremiumPA += dclPremPA; dblTotalSumAtRiskPA += dclPA;
                                    dblTotalPremiumENCI += dclENCI; dblTotalSumAtRiskENCI += dclENCI;
                                    dblTotalPremiumESCI += dclESCI; dblTotalSumAtRiskESCI += dclESCI;
                                    #endregion


                                    //dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range ["X" + i].Value);
                                    //dbTotalCommission = dbTotalCommission + (Convert.ToDecimal(wsraw.Range["J" + i].Value) + Convert.ToDecimal(wsraw.Range["M" + i].Value));
                                }
                            }
                        }
                    }
                    //BM099A
                    else if (str_sheet.ToUpper().Contains("CRQ"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["A" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["A" + i].Value; // Policy Number
                                    //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                    dtDataRow[32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = "07/01/1900"; // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    //dtDataRow[79] = wsraw.Range["H" + i].Value; // Life Issue Age
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "USD"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow [39] = "STANDARD"; // Preferred Classific
                                    dtDataRow[21] = "ADJUST"; // Transcode
                                    dtDataRow[62] = "4004"; // Entry Code
                                    //dtDataRow[63] = wsraw.Range["K" + i].Value; // Premium

                                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orig sum   
                                    dtDataRow [26] = 1;//ceded sum
                                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//initial sum
                                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//sum at risk

                                    dtDataRow [22] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Trans Effective Date
                                    dtDataRow [19] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Reinsurance
                                    dtDataRow [20] = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date

                                    //dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range ["K" + i].Value);
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclPremCIR);
                                    double.TryParse(Convert.ToString(wsraw.Range ["I" + i].Value), out double dclPremEENCI);
                                    double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Value), out double dclPremESCI);

                                    dclPremLIFE = dclPremLIFE * -1;
                                    dclPremAEH = dclPremAEH * -1;
                                    dclPremCIR = dclPremCIR * -1;
                                    dclPremEENCI = dclPremEENCI * -1;
                                    dclPremESCI = dclPremESCI * -1;


                                    _var.dtworkRow02 = objdt_template.NewRow();
                                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow03 = objdt_template.NewRow();
                                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow04 = objdt_template.NewRow();
                                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow05 = objdt_template.NewRow();
                                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;

                                    #region Premiums
                                    if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = 0;
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [63] = dclPremCIR;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "EENCI";
                                        dtDataRow [63] = dclPremEENCI;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "ESCI";
                                        dtDataRow [63] = dclPremESCI;
                                    }

                                    //2
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63   ] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "A&H";
                                        dtDataRow [63] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [63] = dclPremCIR;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [63] = dclPremCIR;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "EENCI";
                                        dtDataRow [63] = dclPremEENCI;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    //3
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "EENCI";
                                        _var.dtworkRow03 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "ESCI";
                                        _var.dtworkRow03 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }

                                    //4
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "EENCI";
                                        _var.dtworkRow04 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "ESCI";
                                        _var.dtworkRow04 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }

                                    //5
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [63] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "A&H";
                                        _var.dtworkRow02 [63] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [63] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "EENCI";
                                        _var.dtworkRow04 [63] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);

                                        _var.dtworkRow05 [5] = "ESCI";
                                        _var.dtworkRow05 [63] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow05);
                                    }
                                    #endregion

                                    #region hashtotal
                                    dblTotalPremiumLife += dclPremLIFE;
                                    dblTotalPremiumAEH += dclPremAEH;
                                    dblTotalPremiumCIR += dclPremCIR;
                                    dblTotalPremiumENCI += dclPremEENCI;
                                    dblTotalPremiumESCI += dclPremESCI;
                                    dblTotalSumAtRiskLife = 0;
                                    dblTotalSumAtRiskAEH = 0;
                                    dblTotalSumAtRiskCIR = 0;
                                    dblTotalSumAtRiskENCI = 0;
                                    dblTotalSumAtRiskESCI = 0;
                                    #endregion

                                }
                            }
                        }
                    }
                    //BM099A
                    else if (str_sheet.ToUpper().Contains("CFQ"))
                    {
                        for (int i = 1; i <= intLastRow; i++)
                        {
                            if (wsraw.Range["A" + i].Value != null)
                            {
                                string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                                if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                                {
                                    dtDataRow = objdt_template.NewRow();
                                    objdt_template.Rows.Add(dtDataRow);

                                    dtDataRow[0] = wsraw.Range["A" + i].Value; // Policy Number
                                   //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                    //objHlpr.fn_separatefullnamev3(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                    objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                    dtDataRow[32] = strLastName; // Last Name
                                    dtDataRow[33] = strFirstName; // First Name
                                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                                    string strSex = objHlpr.fn_getgenderv2(strFirstName);
                                    dtDataRow[36] = strSex; // Gender
                                    string strBirthday = "07/01/1900"; // Birthday;
                                    dtDataRow[37] = strBirthday; // Birthday
                                    dtDataRow[29] = "NATREID"; // Life ID Type
                                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                    //dtDataRow[79] = wsraw.Range["H" + i].Value; // Life Issue Age
                                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                                    dtDataRow[9] = "PAFM"; // Type of Business
                                    dtDataRow[10] = "S"; // Reinsurance Methods
                                    dtDataRow[13] = "IND"; // Class of Business
                                    dtDataRow[14] = "T"; // Business Type
                                    dtDataRow[23] = "USD"; //  Currency
                                    dtDataRow[24] = "YLY"; // Premium Frequency
                                    dtDataRow[38] = "NONE"; // Smoker Status
                                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                                    dtDataRow [39] = "STANDARD"; // Preferred Classific
                                    dtDataRow[21] = "ADJUST"; // Transcode
                                    dtDataRow[60] = "4002"; // Entry Code

                                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orig sum   
                                    dtDataRow [26] = 1;//ceded sum
                                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//initial sum
                                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//sum at risk

                                    dtDataRow [22] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Trans Effective Date
                                    dtDataRow [19] = Convert.ToDateTime(wsraw.Range ["E" + i].Value).ToString("MM/dd/yyyy"); //Reinsurance
                                    dtDataRow [20] = Convert.ToDateTime(wsraw.Range ["D" + i].Value).ToString("MM/dd/yyyy");//Policy Start Date

                                  
                                    double.TryParse(Convert.ToString(wsraw.Range ["F" + i].Value), out double dclPremLIFE);
                                    double.TryParse(Convert.ToString(wsraw.Range ["G" + i].Value), out double dclPremAEH);
                                    double.TryParse(Convert.ToString(wsraw.Range ["H" + i].Value), out double dclPremCIR);
                                    double.TryParse(Convert.ToString(wsraw.Range ["I" + i].Value), out double dclPremEENCI);
                                    double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Value), out double dclPremESCI);

                                    dclPremLIFE = dclPremLIFE * -1;
                                    dclPremAEH = dclPremAEH * -1;
                                    dclPremCIR = dclPremCIR * -1;
                                    dclPremEENCI = dclPremEENCI * -1;
                                    dclPremESCI = dclPremESCI * -1;

                           

                                    _var.dtworkRow02 = objdt_template.NewRow();
                                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow03 = objdt_template.NewRow();
                                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow04 = objdt_template.NewRow();
                                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                                    _var.dtworkRow05 = objdt_template.NewRow();
                                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;

                                    #region Premiums
                                    if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = 0;
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "AEH";
                                        dtDataRow [61] = dclPremAEH;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [61] = dclPremCIR;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "EENCI";
                                        dtDataRow [61] = dclPremEENCI;
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "ESCI";
                                        dtDataRow [61] = dclPremESCI;
                                    }

                                    //2
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "AEH";
                                        dtDataRow [61] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "CIR";
                                        _var.dtworkRow02 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "AEH";
                                        dtDataRow [61] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "AEH";
                                        dtDataRow [61] = dclPremAEH;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [61] = dclPremCIR;

                                        _var.dtworkRow02 [5] = "EENCI";
                                        _var.dtworkRow02 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "CIR";
                                        dtDataRow [61] = dclPremCIR;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }
                                    else if(dclPremLIFE == 0 && dclPremAEH == 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "EENCI";
                                        dtDataRow [61] = dclPremEENCI;

                                        _var.dtworkRow02 [5] = "ESCI";
                                        _var.dtworkRow02 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow02);
                                    }

                                    //3
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "EENCI";
                                        _var.dtworkRow03 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR == 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "ESCI";
                                        _var.dtworkRow03 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow03);
                                    }

                                    //4
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI == 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "EENCI";
                                        _var.dtworkRow04 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI == 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "ESCI";
                                        _var.dtworkRow04 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);
                                    }

                                    //5
                                    else if(dclPremLIFE != 0 && dclPremAEH != 0 && dclPremCIR != 0 && dclPremEENCI != 0 && dclPremESCI != 0)
                                    {
                                        dtDataRow [5] = "LIFE";
                                        dtDataRow [61] = dclPremLIFE;

                                        _var.dtworkRow02 [5] = "AEH";
                                        _var.dtworkRow02 [61] = dclPremAEH;
                                        objdt_template.Rows.Add(_var.dtworkRow02);

                                        _var.dtworkRow03 [5] = "CIR";
                                        _var.dtworkRow03 [61] = dclPremCIR;
                                        objdt_template.Rows.Add(_var.dtworkRow03);

                                        _var.dtworkRow04 [5] = "EENCI";
                                        _var.dtworkRow04 [61] = dclPremEENCI;
                                        objdt_template.Rows.Add(_var.dtworkRow04);

                                        _var.dtworkRow05 [5] = "ESCI";
                                        _var.dtworkRow05 [61] = dclPremESCI;
                                        objdt_template.Rows.Add(_var.dtworkRow05);
                                    }
                                    #endregion

                                    #region hashtotal
                                    dblTotalPremiumLife += dclPremLIFE;
                                    dblTotalPremiumAEH += dclPremAEH;
                                    dblTotalPremiumCIR += dclPremCIR;
                                    dblTotalPremiumENCI += dclPremEENCI;
                                    dblTotalPremiumESCI += dclPremESCI;
                                    dblTotalSumAtRiskLife = 0;
                                    dblTotalSumAtRiskAEH = 0;
                                    dblTotalSumAtRiskCIR = 0;
                                    dblTotalSumAtRiskENCI = 0;
                                    dblTotalSumAtRiskESCI = 0;
                                    #endregion 
                                }
                            }
                        }
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 099", "Information");
                        return "";
                    }
                }
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);


                #region PREM SHEET HASHTOTAL
                //if(str_sheet.ToUpper().Contains("PREM"))
                //{
                if(strFilePath.ToUpper().Contains("PESO") && str_sheet.ToUpper().Contains("PREM"))
                {
                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium LIFE:";
                    dtDataRow [1] = dblTotalPremiumLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk LIFE:";
                    dtDataRow [1] = dblTotalSumAtRiskLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium WP/PB:";
                    dtDataRow [1] = dblTotalPremiumWP;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at RiskWP/PB:";
                    dtDataRow [1] = dblTotalSumAtRiskWP;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ADB:";
                    dtDataRow [1] = dblTotalPremiumADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ADB:";
                    dtDataRow [1] = dblTotalSumAtRiskADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium A&H:";
                    dtDataRow [1] = dblTotalPremiumAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk A&H:";
                    dtDataRow [1] = dblTotalSumAtRiskAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium PA:";
                    dtDataRow [1] = dblTotalPremiumPA;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk PA:";
                    dtDataRow [1] = dblTotalSumAtRiskPA;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium CIR:";
                    dtDataRow [1] = dblTotalPremiumCIR;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk CIR:";
                    dtDataRow [1] = dblTotalSumAtRiskCIR;
                    objdt_template.Rows.Add(dtDataRow);
                }
                else if(strFilePath.ToUpper().Contains("PESO") && str_sheet.ToUpper().Contains("CR") || strFilePath.ToUpper().Contains("PESO") && str_sheet.ToUpper().Contains("CF"))
                {
                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium LIFE:";
                    dtDataRow [1] = dblTotalPremiumLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk LIFE:";
                    dtDataRow [1] = dblTotalSumAtRiskLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium A&H:";
                    dtDataRow [1] = dblTotalPremiumAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk A&H:";
                    dtDataRow [1] = dblTotalSumAtRiskAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium CIR:";
                    dtDataRow [1] = dblTotalPremiumCIR;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk CIR:";
                    dtDataRow [1] = dblTotalSumAtRiskCIR;
                    objdt_template.Rows.Add(dtDataRow);
                }
                else if(strFilePath.ToUpper().Contains("DOLLAR") && str_sheet.ToUpper().Contains("PREM"))
                {
                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium LIFE:";
                    dtDataRow [1] = dblTotalPremiumLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk LIFE:";
                    dtDataRow [1] = dblTotalSumAtRiskLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium WP/PB:";
                    dtDataRow [1] = dblTotalPremiumWP;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at RiskWP/PB:";
                    dtDataRow [1] = dblTotalSumAtRiskWP;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ADB:";
                    dtDataRow [1] = dblTotalPremiumADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ADB:";
                    dtDataRow [1] = dblTotalSumAtRiskADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium A&H:";
                    dtDataRow [1] = dblTotalPremiumAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk A&H:";
                    dtDataRow [1] = dblTotalSumAtRiskAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium PA:";
                    dtDataRow [1] = dblTotalPremiumPA;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk PA:";
                    dtDataRow [1] = dblTotalSumAtRiskPA;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium CIR:";
                    dtDataRow [1] = dblTotalPremiumCIR;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk CIR:";
                    dtDataRow [1] = dblTotalSumAtRiskCIR;
                    objdt_template.Rows.Add(dtDataRow);


                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium EENCI:";
                    dtDataRow [1] = dblTotalPremiumENCI;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk EENCI:";
                    dtDataRow [1] = dblTotalSumAtRiskENCI;
                    objdt_template.Rows.Add(dtDataRow);


                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ESCI:";
                    dtDataRow [1] = dblTotalPremiumESCI;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ESCI:";
                    dtDataRow [1] = dblTotalSumAtRiskESCI;
                    objdt_template.Rows.Add(dtDataRow);
                }
                else if(strFilePath.ToUpper().Contains("DOLLAR") && str_sheet.ToUpper().Contains("CR") || strFilePath.ToUpper().Contains("DOLLAR") && str_sheet.ToUpper().Contains("CF"))
                {
                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium LIFE:";
                    dtDataRow [1] = dblTotalPremiumLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk LIFE:";
                    dtDataRow [1] = dblTotalSumAtRiskLife;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium A&H:";
                    dtDataRow [1] = dblTotalPremiumAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk A&H:";
                    dtDataRow [1] = dblTotalSumAtRiskAEH;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium CIR:";
                    dtDataRow [1] = dblTotalPremiumCIR;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk CIR:";
                    dtDataRow [1] = dblTotalSumAtRiskCIR;
                    objdt_template.Rows.Add(dtDataRow);


                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium EENCI:";
                    dtDataRow [1] = dblTotalPremiumENCI;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk EENCI:";
                    dtDataRow [1] = dblTotalSumAtRiskENCI;
                    objdt_template.Rows.Add(dtDataRow);


                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ESCI:";
                    dtDataRow [1] = dblTotalPremiumESCI;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ESCI:";
                    dtDataRow [1] = dblTotalSumAtRiskESCI;
                    objdt_template.Rows.Add(dtDataRow);
                }

                   
                    #endregion

                string despath = str_saved + @"\BM099-" + str_sheet + str_savef + ".xlsx";
                objHlpr.fn_savefile(objdt_template, despath);

                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
