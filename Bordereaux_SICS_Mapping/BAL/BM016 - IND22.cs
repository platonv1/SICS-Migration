using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM016_IND22
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
            double dclPremiumLIFE = 0, dclPremiumADB = 0, dblTotalPremiumLIFE = 0, dblTotalSumAtRiskLIFE = 0, dblTotalPremiumRider = 0, dblTotalSumAtRiskRider = 0;
            string planCodeRider = string.Empty; string strTransCode = string.Empty; string planCode = string.Empty;


            for(int i = 7; i <= intLastRow; i++)
            {
                if(wsraw.Range ["G" + i].Value != null)
                {
                    string strPolicyNo = Convert.ToString(wsraw.Range ["G" + i].Value);
                    if(Regex.IsMatch(strPolicyNo, @"^\d"))
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
                        
                        dtDataRow [36] = wsraw.Range ["E" + i].Value;//Gender
                        string strBirthday = Convert.ToDateTime(wsraw.Range ["H" + i].Value).ToString("MM/dd/yyyy");
                        dtDataRow [37] = strBirthday; // Birthday
                        dtDataRow [29] = "NATREID"; // Life ID Type
                        dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                        //dtDataRow [79] = wsraw.Range ["L" + i].Value; // Life Issue Age
                        dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow [9] = "PAFM"; // Type of Business
                        dtDataRow [10] = "S"; // Reinsurance Methods
                        dtDataRow [13] = "GCL"; // Class of Business
                        dtDataRow [14] = "T"; // Business Type
                        dtDataRow [24] = "YLY"; // Premium Frequency
                        dtDataRow [38] = "NONE"; // Smoker Status
                        dtDataRow [23] = "PHP"; // Cession Currency
                        dtDataRow [41] = bmyear; // Policy Year
                        dtDataRow [39] = objHlpr2.fn_getmortalityrating("");

                        #region policyYear
                        //objHlpr2.fn_getTranscode(str_sheet, out string strTransCode);
                        strTransCode = "TRENEW";
                        dtDataRow [21] = strTransCode;
                        string strIssueDate = Convert.ToDateTime(wsraw.Range ["I" + i].Value).ToString("MM/dd/yyyy");
                        objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out string transEffectiveDate, out string policyStartDate); // Policy Start Date

                        //if(strTransCode == "TNEWBUS")
                        //{
                        //    dtDataRow [19] = transEffectiveDate;
                        //    dtDataRow [22] = transEffectiveDate; // Trans Effective Dat e
                        //    dtDataRow [20] = transEffectiveDate;//Reinsurance Start Date
                        //}
                        //else
                        //{
                            dtDataRow [19] = transEffectiveDate;//Reinsurance Start Date
                            dtDataRow [22] = transEffectiveDate; // Trans Effective Date
                            dtDataRow [20] = policyStartDate;
                        //}

                        #endregion
                        double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Text), out double dclOrigSum);
                        double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Text), out double dclInitialSum);
                        double.TryParse(Convert.ToString(wsraw.Range ["J" + i].Text), out double dclCedentRetion);
                        double.TryParse(Convert.ToString(wsraw.Range ["K" + i].Text), out dclPremiumLIFE);
                        double.TryParse(Convert.ToString(wsraw.Range ["L" + i].Text), out dclPremiumADB);

                        dclInitialSum = dclInitialSum * 0.85;
                        dclPremiumLIFE = dclPremiumLIFE * 0.85;
                        dclPremiumADB = dclPremiumADB * 0.85;

                        dblTotalPremiumLIFE += dclPremiumLIFE;
                        dblTotalPremiumRider += dclPremiumADB;
                        dblTotalSumAtRiskLIFE += dclInitialSum;

                        planCode = wsraw.Range ["F" + i].Value;
                        //planCodeRider = wsraw.Range ["H" + i].Value;

                            if(dclPremiumLIFE != 0)
                            {
                                dtDataRow [58] = "4001"; // Entry Code
                                dtDataRow [59] = dclPremiumLIFE; // Premium
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
                                _var.dtworkRow02 [5] = "ADB";
                                _var.dtworkRow02 [27] = 1; // Initial Sum at Risk
                                _var.dtworkRow02 [77] = 1; // Sum at Risk
                                _var.dtworkRow02 [25] = 1; // Original Sum Assured
                                _var.dtworkRow02 [26] = 1; // Cedent Retention
                                objdt_template.Rows.Add(_var.dtworkRow02);
                            }
                       


                    }
                }
            }


            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if(dblTotalPremiumLIFE != 0)
            {
                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium:";
                dtDataRow [1] = dblTotalPremiumLIFE;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Sum at Risk:";
                dtDataRow [1] = dblTotalSumAtRiskLIFE;
                objdt_template.Rows.Add(dtDataRow);

            }

            if(dblTotalPremiumRider != 0)
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

            string despath = str_saved + @"\BM016-IND22" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
