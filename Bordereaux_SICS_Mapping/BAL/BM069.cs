using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM069
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            //try
            //{
                _Global _var = new _Global();
                Helper objHlpr = new Helper();
                System.Data.DataTable objdt_template = new System.Data.DataTable();

                objdt_template = objHlpr.dt_formtemplate(str_sheet);
                Application eapp = new Application();
                Workbook wbraw = eapp.Workbooks.Open(str_raw);
                Worksheet wsraw = wbraw.Worksheets [str_sheet];

                int intLastRow = wsraw.Cells [wsraw.Rows.Count, 1].End [XlDirection.xlUp].row;

                DataRow dtDataRow;
                decimal dblTotalPremiumPHP = 0, dblTotalPremiumUSD = 0, dblTotalSumAtRiskPHP = 0, dblTotalSumAtRiskUSD = 0;
                string strCurrency = "";

                while(string.IsNullOrEmpty(Variables.strBmYear))
                {

                    if(string.IsNullOrEmpty(Variables.strBmYear))
                    {
                        frmPolicyYear newform = new frmPolicyYear();
                        newform.ShowDialog();

                    }
                }

                if(str_sheet.ToUpper().Contains("FAC") || str_sheet.ToUpper().Contains("TRAD"))
                {
                    for(int i = 1; i <= intLastRow; i++)
                    {
                        if(wsraw.Range ["A" + i].Value != null)
                        {
                            var Currency = wsraw.Range ["A" + i].Value;
                            if(Currency.GetType() == typeof(string))
                            {
                                if(Currency == "Peso")
                                {
                                    strCurrency = "PHP";
                                }
                                else if(Currency == "Dollar")
                                {
                                    strCurrency = "USD";
                                }
                            }
                            string strCessionNo = Convert.ToString(wsraw.Range ["A" + i].Value);
                            if(Regex.IsMatch(strCessionNo, @"^\d+$"))
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow [0] = wsraw.Range ["B" + i].Value; // Policy Number
                                string strGender = wsraw.Range ["E" + i].Value; // Gender
                                if(strGender.ToUpper().Contains("FEMALE"))
                                {
                                    dtDataRow [36] = "F";
                                }
                                else
                                {
                                    dtDataRow [36] = "M";
                                }
                                objHlpr.fn_separatefullnamev3(wsraw.Range ["F" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow [32] = strLastName; // Last Name
                                dtDataRow [33] = strFirstName; // First Name
                                dtDataRow [34] = strMiddleInitial; // Middle Initials
                                string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy");
                                dtDataRow [37] = strBirthday; // Birthday
                                dtDataRow [29] = "NATREID"; // Life ID Type
                                dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                dtDataRow [79] = wsraw.Range ["H" + i].Value; // Life Issue Age
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow [9] = "PFO"; // Type of Business
                                dtDataRow [10] = "S"; // Reinsurance Methods
                                dtDataRow [13] = "IND"; // Class of Business
                                dtDataRow [14] = "T"; // Business Type
                                dtDataRow [24] = "MLY"; // Premium Frequency
                                dtDataRow [38] = "NONE"; // Smoker Status
                                dtDataRow [23] = strCurrency; // Cession Currency
                                dtDataRow [41] = Variables.strBmYear; // Policy Year
                                dtDataRow [05] = wsraw.Range ["C" + i].Value; // Branded Product
                                dtDataRow [39] = wsraw.Range ["I" + i].Value; // Preferred Classific
                                dtDataRow [20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["D" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                                dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                                dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                                objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range ["K" + i].Value), null, Convert.ToString(wsraw.Range ["M" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                                dtDataRow [26] = 1; //ceded sum assured
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum) ; // Original Sum Assured
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum) * Convert.ToDecimal(.15); // Initial Sum at Risk
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) * Convert.ToDecimal(.15); // Sum at Risk

                                dtDataRow [21] = "TNEWBUS"; // Transcode
                                dtDataRow [56] = "4000"; // Entry Code
                                dtDataRow [57] = Convert.ToDecimal(wsraw.Range["N" + i].Value) * Convert.ToDecimal(.15); // Premium

                            if (strCurrency == "PHP")
                                {
                                    dblTotalPremiumPHP = dblTotalPremiumPHP + Convert.ToDecimal(wsraw.Range ["N" + i].Value) * Convert.ToDecimal(.15);
                                    dblTotalSumAtRiskPHP = dblTotalSumAtRiskPHP + objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) * Convert.ToDecimal(.15);
                                }
                                else
                                {
                                    dblTotalPremiumUSD = dblTotalPremiumUSD + Convert.ToDecimal(wsraw.Range ["N" + i].Value) * Convert.ToDecimal(.15);
                                    dblTotalSumAtRiskUSD = dblTotalSumAtRiskUSD + objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk) * Convert.ToDecimal(.15);
                                }
                            }
                        }
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 069", "Information");
                    return "";
                }

                #region Computing Hash 
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium PHP:";
                dtDataRow [1] = dblTotalPremiumPHP;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium USD:";
                dtDataRow [1] = dblTotalPremiumUSD;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk PHP:";
                dtDataRow[1] = dblTotalSumAtRiskPHP;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk USD:";
                dtDataRow[1] = dblTotalSumAtRiskUSD;
                objdt_template.Rows.Add(dtDataRow);
                #endregion

                string despath = str_saved + @"\BM069-" + str_sheet + str_savef + ".xlsx";
                objHlpr.fn_savefile(objdt_template, despath);

                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            //}
            //catch(Exception ex)
            //{
            //    return ex.Message;
            //}
        }
    }
}