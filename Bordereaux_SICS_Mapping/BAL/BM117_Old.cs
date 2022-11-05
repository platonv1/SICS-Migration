using System;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM117_Old
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
            Worksheet wsraw = wbraw.Worksheets[str_sheet];

            int intLastRow = wsraw.Cells[1, 1].End[XlDirection.xlDown].row;
            intLastRow = wsraw.Cells[intLastRow, 1].End[XlDirection.xlDown].row;
            intLastRow = wsraw.Cells[intLastRow, 1].End[XlDirection.xlDown].row;
            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;
            string valueTransEffectiveDate = "";
            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }
        
            if (str_sheet.ToUpper().Contains("URC_RY") || str_sheet.ToUpper().Contains("URC$_RY"))
            {
                for (int i = 2; i <= intLastRow; i++)
                {
                    if (wsraw.Range["A" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                        if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                            dtDataRow[36] = wsraw.Range["F" + i].Value; // Gender
                            objHlpr.fn_separatefullnamev3(wsraw.Range["D" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow[32] = strLastName; // Last Name
                            dtDataRow[33] = strFirstName; // First Name
                            dtDataRow[34] = objHlpr2.fn_removeCharacters(strMiddleInitial); // Middle Initials
                            string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["E" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow[78] = wsraw.Range["H" + i].Value; // Attained Age
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "IND"; // Class of Business
                            dtDataRow[24] = "YLY"; // Premium Frequency    
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow [14] = objHlpr2.fn_businessTypeV2(Convert.ToString(wsraw.Range ["K" + i].Value)); // Business Type
                            dtDataRow [41] = Variables.strBmYear; // Policy Year
                            dtDataRow[39] = objHlpr.fn_getmortality(Convert.ToString(wsraw.Range["K" + i].Value)); // Preferred Classific)

                            //string strBusinessType = wsraw.Range["J" + i].Value; // Business Type
                            //if (strBusinessType.ToUpper().Contains("A"))
                            //{
                            //    dtDataRow[14] = "T";
                            //}
                            //else
                            //{
                            //    dtDataRow[14] = "F";
                            //}

                            if (str_sheet.ToUpper().Contains("$"))
                            {
                                dtDataRow[23] = "USD"; // Currency
                            }
                            else
                            {
                                dtDataRow[23] = "PHP"; // Currency
                            }
                            //dtDataRow [20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                            //dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                            //dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE
                            //objHlpr.fn_CheckingforA_AB_BZColumn(null, Convert.ToString(wsraw.Range ["N" + i].Value), Convert.ToString(wsraw.Range ["N" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);
                            //dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                            //dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                            //dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Initial Sum at Risk
                            //dtDataRow [26] = strOriginalSum; //ceded summ assured
                            string strTcode = "TRENEW";
                            dtDataRow [21] = strTcode; // Transcode
                            dtDataRow [58] = "4001"; // Entry Code
                            //string strIssueDate = Convert.ToDateTime(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow [22] = Convert.ToDateTime(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy"); //Trans Effective Date
                            dtDataRow [19] = Convert.ToDateTime(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy");//REINSURANCE START DATE
                            dtDataRow [20] = Convert.ToDateTime(Convert.ToString(wsraw.Range ["G" + i].Value)).ToString("MM/dd/yyyy");//Policy Start Date

                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow03 = objdt_template.NewRow();
                            _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;

                            decimal.TryParse(Convert.ToString(wsraw.Range ["P" + i].Value), out decimal dclPremLife);
                            decimal.TryParse(Convert.ToString(wsraw.Range ["Q" + i].Value), out decimal dclPremExtra);
                            decimal.TryParse(Convert.ToString(wsraw.Range ["R" + i].Value), out decimal dclPremWP);

                            decimal.TryParse(Convert.ToString(wsraw.Range ["N" + i]), out decimal dclSAR);

                            if(dclPremLife != 0 && dclPremExtra == 0 && dclPremWP == 0)
                            {
                                dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDecimal(wsraw.Range ["N" + i].Value));
                                dtDataRow [5] = "LIFE"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremLife; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                            }
                            else if(dclPremLife == 0 && dclPremExtra != 0 && dclPremWP == 0)
                            {
                                dtDataRow [5] = "EXTRA"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremExtra; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum

                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremWP != 0)
                            {
                                dtDataRow [5] = "WP/PB"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremWP; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremWP == 0)
                            {
                                dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDecimal(wsraw.Range ["N" + i].Value));
                                dtDataRow [5] = "LIFE"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremLife; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product Cedent Code
                                _var.dtworkRow02 [59] = dclPremExtra; // Premium
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //SAR
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                                objdt_template.Rows.Add(_var.dtworkRow02);

                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremWP == 0)
                            {
                                dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDecimal(wsraw.Range ["N" + i].Value));
                                dtDataRow [5] = "LIFE"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremLife; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum

                                _var.dtworkRow02 [5] = "WP/PB"; // Branded Product Cedent Code
                                _var.dtworkRow02 [59] = dclPremWP; // Premium
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //SAR
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                                objdt_template.Rows.Add(_var.dtworkRow02);

                            }
                            else if(dclPremLife != 0 && dclPremExtra != 0 && dclPremWP != 0)
                            {
                                dblTotalSumAtRisk = dblTotalSumAtRisk + (Convert.ToDecimal(wsraw.Range ["N" + i].Value));
                                dtDataRow [5] = "LIFE"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremLife; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum

                                _var.dtworkRow02 [5] = "EXTRA"; // Branded Product Cedent Code
                                _var.dtworkRow02 [57] = dclPremExtra; // Premium
                                _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //SAR
                                _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial
                                _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                                objdt_template.Rows.Add(_var.dtworkRow02);

                                _var.dtworkRow03 [5] = "WP/PB"; // Branded Product Cedent Code
                                _var.dtworkRow03 [59] = dclPremWP; // Premium
                                _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //SAR
                                _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //Initial
                                _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                                objdt_template.Rows.Add(_var.dtworkRow03);
                            }
                            else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremWP == 0)
                            {
                                dtDataRow [5] = "LIFE"; // Branded Product Cedent Code
                                dtDataRow [59] = dclPremLife; // Premium
                                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //SAR
                                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR)); //Initial
                                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //orig sum
                                dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(null); //ceded sum
                            }
                            #region Hashtotal
                           
                            dblTotalPremium += dclPremLife + dclPremExtra + dclPremWP;
                            #endregion
                        }
                    }
                }
            }

            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 117", "Information");
                return "";
            }

            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);
            #endregion



            string despath = str_saved + @"\BM117-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }
}
