using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM014_CBL22
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets [str_sheet];

            int intLastRow = wsraw.Cells [wsraw.Rows.Count, 1].End [XlDirection.xlUp].row;

            DataRow dtDataRow;
            double dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            if(str_sheet.ToUpper() != "ADJ RENEWAL")
            {
                for(int i = 7; i <= intLastRow; i++)
                {
                    if(wsraw.Range ["K" + i].Value != null)
                    {
                        string strPolicyNo = Convert.ToString(wsraw.Range ["K" + i].Value);
                        if(Regex.IsMatch(strPolicyNo, @"^\d"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            string bmyear = wsraw.Range ["A" + 4].Text;
                            bmyear = bmyear.Substring(bmyear.Length - 4, 4).Trim();
                            dtDataRow [0] = strPolicyNo; // Policy
                            dtDataRow [1] = wsraw.Range ["A" + i].Value; //cession number
                            dtDataRow [5] = "MRI";
                            string strFirstName = wsraw.Range ["C" + i].Text;
                            string strLastName = wsraw.Range ["B" + i].Text;
                            string strMI = wsraw.Range ["D" + i].Text;
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMI + ".";
                            dtDataRow [32] = strLastName; // Last Name
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow [34] = objHlpr2.fn_removeCharacters(strMI); // Middle Initials
                            objHlpr2.fn_getgenderv2(strFirstName, out string strSex);//gender
                            dtDataRow [36] = strSex;
                            string strBirthday = Convert.ToDateTime(wsraw.Range ["L" + i].Value).ToString("MM/dd/yyyy");
                            dtDataRow [37] = strBirthday; // Birthday
                            dtDataRow [29] = "NATREID"; // Life ID Type
                            dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow [9] = "PAFM"; // Type of Business
                            dtDataRow [10] = "S"; // Reinsurance Methods
                            dtDataRow [13] = "GCL"; // Class of Business
                            dtDataRow [14] = "T"; // Business Type
                            dtDataRow [24] = "YLY"; // Premium Frequency
                            dtDataRow [38] = "NONE"; // Smoker Status
                            dtDataRow [23] = "PHP"; // Cession Currency
                            dtDataRow [41] = bmyear; // Policy Year

                            dtDataRow [39] = objHlpr2.fn_getmortalityrating(wsraw.Range ["G" + i].Text);

                            #region policyYear
                            objHlpr2.fn_getTranscode(str_sheet, out string strTransCode);
                            dtDataRow [21] = strTransCode;
                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["M" + i].Value).ToString("MM/dd/yyyy");
                            objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out string transEffectiveDate, out string policyStartDate); // Policy Start Date

                            if(str_sheet.ToUpper() == "FIRST YEAR" || str_sheet.ToUpper() == "ADJ FIRST YEAR")
                            {
                                dtDataRow [19] = transEffectiveDate;
                                dtDataRow [22] = transEffectiveDate; // Trans Effective Dat e
                                dtDataRow [20] = transEffectiveDate;//Reinsurance Start Date
                            }
                            else
                            {
                                dtDataRow [19] = transEffectiveDate;//Reinsurance Start Date
                                dtDataRow [22] = transEffectiveDate; // Trans Effective Date
                                dtDataRow [20] = policyStartDate;
                            }

                            #endregion

                            double.TryParse(Convert.ToString(wsraw.Range ["P" + i].Text), out double dclOrigSum);
                            double.TryParse(Convert.ToString(wsraw.Range ["R" + i].Text), out double dclInitialSum);
                            double.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Text), out double dclCedentRetion);
                            double.TryParse(Convert.ToString(wsraw.Range ["T" + i].Text), out double dclPremium);
                            dclInitialSum = dclInitialSum * 0.85;
                            dclPremium = dclPremium * 0.85;


                            dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion));
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                            dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(null));
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum));
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum));


                            if(str_sheet.ToUpper().Contains("FIRST YEAR"))
                            {
                                dtDataRow [56] = "4000"; // Entry Code
                                dtDataRow [57] = dclPremium; // Premium

                            }
                            else if(str_sheet.ToUpper().Contains("RENEWAL"))
                            {
                                dtDataRow [58] = "4001"; // Entry Code
                                dtDataRow [59] = dclPremium; // Premium
                            }
                            else if(str_sheet.ToUpper().Contains("ADJUSTMENT FIRST YEAR"))
                            {
                                dtDataRow [60] = "4002"; // Entry Code
                                dtDataRow [61] = dclPremium; // Premium

                            }
                            dblTotalPremium += dclPremium;
                            dblTotalSumAtRisk += dclInitialSum;


                        }
                    }
                }

            }
            else
            {
                for(int i = 7; i <= intLastRow; i++)
                {
                    if(wsraw.Range ["K" + i].Value != null)
                    {
                        string strPolicyNo = Convert.ToString(wsraw.Range ["K" + i].Value);
                        if(Regex.IsMatch(strPolicyNo, @"^\d"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            string bmyear = wsraw.Range ["A" + 4].Text;
                            bmyear = bmyear.Substring(bmyear.Length - 4, 4).Trim();
                            dtDataRow [0] = strPolicyNo; // Policy
                            dtDataRow [1] = wsraw.Range ["A" + i].Value; //cession number
                            dtDataRow [5] = "MRI";
                            string strFirstName = wsraw.Range ["C" + i].Text;
                            string strLastName = wsraw.Range ["B" + i].Text;
                            string strMI = wsraw.Range ["D" + i].Text;
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMI + ".";
                            dtDataRow [32] = strLastName; // Last Name
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow [34] = objHlpr2.fn_removeCharacters(strMI); // Middle Initials
                            objHlpr2.fn_getgenderv2(strFirstName, out string strSex);//gender
                            dtDataRow [36] = strSex;
                            string strBirthday = Convert.ToDateTime(wsraw.Range ["L" + i].Value).ToString("MM/dd/yyyy");
                            dtDataRow [37] = strBirthday; // Birthday
                            dtDataRow [29] = "NATREID"; // Life ID Type
                            dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow [9] = "PAFM"; // Type of Business
                            dtDataRow [10] = "S"; // Reinsurance Methods
                            dtDataRow [13] = "GCL"; // Class of Business
                            dtDataRow [14] = "T"; // Business Type
                            dtDataRow [24] = "YLY"; // Premium Frequency
                            dtDataRow [38] = "NONE"; // Smoker Status
                            dtDataRow [23] = "PHP"; // Cession Currency
                            dtDataRow [41] = bmyear; // Policy Year

                            dtDataRow [39] = objHlpr2.fn_getmortalityrating(wsraw.Range ["G" + i].Text);

                            #region policyYear
                            objHlpr2.fn_getTranscode(str_sheet, out string strTransCode);
                            dtDataRow [21] = strTransCode;
                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["M" + i].Value).ToString("MM/dd/yyyy");
                            objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out string transEffectiveDate, out string policyStartDate); // Policy Start Date

                           
                                dtDataRow [19] = transEffectiveDate;//Reinsurance Start Date
                                dtDataRow [22] = transEffectiveDate; // Trans Effective Date
                                dtDataRow [20] = policyStartDate;
                           

                            #endregion

                            double.TryParse(Convert.ToString(wsraw.Range ["P" + i].Text), out double dclOrigSum);
                            double.TryParse(Convert.ToString(wsraw.Range ["R" + i].Text), out double dclInitialSum);
                            double.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Text), out double dclCedentRetion);
                            double.TryParse(Convert.ToString(wsraw.Range ["T" + i].Text), out double dclPremium);
                            dclInitialSum = dclInitialSum * 0.85;
                            dclPremium = dclPremium * 0.85;


                            dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion));
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum));
                            dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(null));
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum));
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum));


                                dtDataRow [62] = "4004"; // Entry Code
                                dtDataRow [63] = dclPremium; // Premium

                          
                            dblTotalPremium += dclPremium;
                            dblTotalSumAtRisk += dclInitialSum;


                        }
                    }
                }
            }

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM014-CBL22" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
