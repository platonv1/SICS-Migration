using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM013_CGL22
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
            double dclPremiumCGL = 0, dclPremiumADB = 0, dblTotalPremiumCGL = 0, dblTotalSumAtRiskCGL = 0, dblTotalPremiumRider = 0, dblTotalSumAtRiskRider = 0;
            string planCodeRider = string.Empty;


            for(int i = 7; i <= intLastRow; i++)
                {
                    if(wsraw.Range ["I" + i].Value != null)
                    {
                        string strPolicyNo = Convert.ToString(wsraw.Range ["I" + i].Value);
                        if(Regex.IsMatch(strPolicyNo, @"^\d|^FICO"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            string bmyear = wsraw.Range ["A" + 4].Text;
                            bmyear = bmyear.Substring(bmyear.Length - 4, 4).Trim();
                            dtDataRow [0] = strPolicyNo; // Policy
                            dtDataRow [1] = wsraw.Range ["A" + i].Value; //cession number
                            string strFirstName = wsraw.Range ["C" + i].Text;
                            string strLastName = wsraw.Range ["B" + i].Text;
                            string strMI = wsraw.Range ["D" + i].Text;
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMI + ".";
                            dtDataRow [32] = strLastName; // Last Name
                            dtDataRow [33] = strFirstName; // First Name
                            dtDataRow [34] = objHlpr2.fn_removeCharacters(strMI); // Middle Initials
                            objHlpr2.fn_getgenderv2(strFirstName, out string strSex);//gender
                            dtDataRow [36] = strSex; 
                            string strBirthday = Convert.ToDateTime(wsraw.Range ["J" + i].Value).ToString("MM/dd/yyyy");
                            dtDataRow [37] = strBirthday; // Birthday
                            dtDataRow [29] = "NATREID"; // Life ID Type
                            dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow [79] = wsraw.Range ["L" + i].Value; // Life Issue Age
                            //dtDataRow [28] = wsraw.Range ["M" + i].Value; // Cedant Retention
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
                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["K" + i].Value).ToString("MM/dd/yyyy");
                            objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out string transEffectiveDate, out string policyStartDate); // Policy Start Date
                           
                            if (strTransCode == "TNEWBUS" || strTransCode == "ADJUST")
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

                        double.TryParse(Convert.ToString(wsraw.Range ["M" + i].Text), out double dclOrigSum);
                        double.TryParse(Convert.ToString(wsraw.Range ["O" + i].Text), out double dclInitialSum);
                        double.TryParse(Convert.ToString(wsraw.Range ["N" + i].Text), out double dclCedentRetion);
                        double.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Text), out  dclPremiumCGL);
                        double.TryParse(Convert.ToString(wsraw.Range ["R" + i].Text), out  dclPremiumADB);
                            
                        dclInitialSum = dclInitialSum * 0.85;
                        dclPremiumCGL = dclPremiumCGL * 0.85;
                        dclPremiumADB = dclPremiumADB * 0.85;

                        dblTotalPremiumCGL += dclPremiumCGL;
                        dblTotalPremiumRider += dclPremiumADB;
                        dblTotalSumAtRiskCGL += dclInitialSum;

                        string planCode = wsraw.Range ["E" + i].Value;
                        planCodeRider = wsraw.Range ["H" + i].Value;

                        if(str_sheet.ToUpper() == "FIRST YEAR")
                        {
                            
                            if(dclPremiumCGL != 0)
                            {
                                dtDataRow [56] = "4000"; // Entry Code
                                dtDataRow [57] = dclPremiumCGL; // Premium
                                dtDataRow [5] = planCode;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Initial Sum at Risk
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Sum at Risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum)); // Original Sum Assured
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion)); // Cedent Retention
                            }
                            if(dclPremiumADB != 0)
                            {
                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow02 [56] = "4000"; // Entry Code
                                _var.dtworkRow02 [57] = dclPremiumADB; // Premium
                                _var.dtworkRow02 [5] = planCodeRider;
                                _var.dtworkRow02 [27] = 1; // Initial Sum at Risk
                                _var.dtworkRow02 [77] = 1; // Sum at Risk
                                _var.dtworkRow02 [25] = 1; // Original Sum Assured
                                _var.dtworkRow02 [26] = 1; // Cedent Retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                        }
                        else if(str_sheet.ToUpper() == "RENEWAL")
                        {
                            if(dclPremiumCGL != 0)
                            {
                                dtDataRow [58] = "4001"; // Entry Code
                                dtDataRow [59] = dclPremiumCGL; // Premium
                                dtDataRow [5] = planCode;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Initial Sum at Risk
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Sum at Risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum)); // Original Sum Assured
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion)); // Cedent Retention
                            }
                            if(dclPremiumADB != 0)
                            {
                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow02 [58] = "4001"; // Entry Code
                                _var.dtworkRow02 [59] = dclPremiumADB; // Premium
                                _var.dtworkRow02 [5] = planCodeRider;
                                _var.dtworkRow02 [27] = 1; // Initial Sum at Risk
                                _var.dtworkRow02 [77] = 1; // Sum at Risk
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum)); // Original Sum Assured
                                _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion)); // Cedent Retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                        }
                        else if(str_sheet.ToUpper() == "ADJUSTMENT FIRST YEAR")
                        {
                            if(dclPremiumCGL != 0)
                            {
                                dtDataRow [60] = "4002"; // Entry Code
                                dtDataRow [61] = dclPremiumCGL; // Premium
                                dtDataRow [5] = planCode;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Initial Sum at Risk
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Sum at Risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum)); // Original Sum Assured
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion)); // Cedent Retention

                            }
                            if(dclPremiumADB != 0)
                            {
                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow02 [60] = "4002"; // Entry Code
                                _var.dtworkRow02 [61] = dclPremiumADB; // Premium
                                _var.dtworkRow02 [5] = planCodeRider;
                                _var.dtworkRow02 [27] = 1; // Initial Sum at Risk
                                _var.dtworkRow02 [77] = 1; // Sum at Risk
                                _var.dtworkRow02 [25] = 1; // Original Sum Assured
                                _var.dtworkRow02 [26] = 1; // Cedent Retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                        }
                        else if(str_sheet.ToUpper() == "ADJUSTMENT RENEWAL")
                        {
                            if(dclPremiumCGL != 0)
                            {
                                dtDataRow [62] = "4004"; // Entry Code
                                dtDataRow [63] = dclPremiumCGL; // Premium
                                dtDataRow [5] = planCode;
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Initial Sum at Risk
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); // Sum at Risk
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum)); // Original Sum Assured
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentRetion)); // Cedent Retention
                            }
                            if(dclPremiumADB != 0)
                            {
                                _var.dtworkRow02 = objdt_template.NewRow();
                                _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                                _var.dtworkRow02 [62] = "4004"; // Entry Code
                                _var.dtworkRow02 [63] = dclPremiumADB; // Premium
                                _var.dtworkRow02 [5] = planCodeRider;
                                _var.dtworkRow02 [27] = 1; // Initial Sum at Risk
                                _var.dtworkRow02 [77] = 1; // Sum at Risk
                                _var.dtworkRow02 [25] = 1; // Original Sum Assured
                                _var.dtworkRow02 [26] = 1; // Cedent Retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                        }

                       

                    }
                }
            }
            

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if (dblTotalPremiumCGL != 0)
            {
                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium:";
                dtDataRow [1] = dblTotalPremiumCGL;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Sum at Risk:";
                dtDataRow [1] = dblTotalSumAtRiskCGL;
                objdt_template.Rows.Add(dtDataRow);

            }

            if (dblTotalPremiumRider != 0)
            {
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);
                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "ADB Total Premium: ";
                dtDataRow [1] = dblTotalPremiumRider;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "ADB Total Sum at Risk:";
                dtDataRow [1] = 0;
                objdt_template.Rows.Add(dtDataRow);
            }

            string despath = str_saved + @"\BM013-CGL22" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
