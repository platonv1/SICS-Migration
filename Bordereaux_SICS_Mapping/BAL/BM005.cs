using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    public class BM005
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

            string strFilePath = wbraw.FullName;
            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 3].End[XlDirection.xlUp].row;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0, dblTotalPremiumUL = 0, dblTotalPremiumTRAD = 0, dblTotalSumAtRiskUL= 0, dblTotalSumAtRiskTRAD=0;
            if (str_sheet.ToUpper().Contains("NB") || str_sheet.ToUpper().Contains("IF"))
            {
                for (int i = 5; i <= intLastRow; i++)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["R" + i].Value), null, Convert.ToString(wsraw.Range["W" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                    dtDataRow[0] = wsraw.Range["D" + i].Value; // Policy Number
                    dtDataRow[36] = wsraw.Range["J" + i].Value; // Gender
                    objHlpr.fn_separatefullnamev5(wsraw.Range["C" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                    dtDataRow[32] = strLastName; // Last Name
                    dtDataRow[33] = strFirstName; // First Name
                    dtDataRow[34] = strMiddleInitial; // Middle Initials
                    string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strBirthday; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    dtDataRow[25] = strOriginalSum; // Original Sum Assured
                    dtDataRow[27] = strInitialSum; // Initial Sum at Risk
                    dtDataRow[77] = strSumAtRisk; // Sum at Risk
                    dtDataRow[76] = strRemarksCode; // Remarks
                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(Convert.ToString(wsraw.Range["L" + i].Value)); // Preffered Classific
                    dtDataRow[5] = wsraw.Range["F" + i].Value + "-" + wsraw.Range["Q" + i].Value; // Branded Product Cedent Code

                    dtDataRow[19] = str_sheet.Substring(6, 2) + "/" +  objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["O" + i].Value)).ToString("dd") + "/" + str_sheet.Substring(2, 4); // Reinsurance Start Date
                    dtDataRow[20] = objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["O" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[22] = str_sheet.Substring(6, 2) + "/" + objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["O" + i].Value)).ToString("dd") + "/" + str_sheet.Substring(2, 4); // Reinsurance Start Date

                    if (strFilePath.ToUpper().Contains("PHP"))
                    {
                        dtDataRow[23] = "PHP"; //  Currency
                    }
                    else
                    {
                        dtDataRow[23] = "USD"; //  Currency
                    }

                    if (wsraw.Range["AC" + i].Value == "FY")
                    {
                        dtDataRow[21] = "TNEWBUS"; // Transcode
                        dtDataRow[56] = "4000"; // Entry Code
                        dtDataRow[57] = wsraw.Range["Z" + i].Value - wsraw.Range["AA" + i].Value; // Premium
                    }
                    else
                    {
                        dtDataRow[21] = "TRENEW"; // Transcode
                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["Z" + i].Value - wsraw.Range["AA" + i].Value; // Premium
                    }

                    if (wsraw.Range["AD" + i].Value == "UL")
                    {
                        dblTotalPremiumUL = dblTotalPremiumUL + (Convert.ToDecimal(wsraw.Range["Z" + i].Value) - Convert.ToDecimal(wsraw.Range["AA" + i].Value));
                        dblTotalSumAtRiskUL = dblTotalSumAtRiskUL + Convert.ToDecimal(strSumAtRisk);
                    }
                    else
                    {
                        dblTotalPremiumTRAD = dblTotalPremiumTRAD  +(Convert.ToDecimal(wsraw.Range["Z" + i].Value) - Convert.ToDecimal(wsraw.Range["AA" + i].Value));
                        dblTotalSumAtRiskTRAD = dblTotalSumAtRiskTRAD + Convert.ToDecimal(strSumAtRisk);
                    }
                }
            }
            else if (str_sheet.ToUpper().Contains("ADJ"))
            {
                for (int i = 6; i <= intLastRow; i++)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    objHlpr.fn_CheckingforA_AB_BZColumn(null, null, null, out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                    dtDataRow[0] = wsraw.Range["B" + i].Value; // Policy Number
                    dtDataRow[36] = "M"; // Gender
                    dtDataRow[31] = "DUMMYLASTNAME" + ", " + "DUMMYFIRSTNAME" + " " + "DUMMYMIDDLENAME" + "."; // Full name
                    dtDataRow[32] = "DUMMYLASTNAME"; // Last Name
                    dtDataRow[33] = "DUMMYFIRSTNAME"; // First Name
                    dtDataRow[34] = "DUMMYMIDDLENAME"; // Middle Initials
                    dtDataRow[37] = "07/01/1900"; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID("DUMMYLASTNAME", "DUMMYLASTNAME", "07/01/1900"); // Life ID
                    //string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy");
                    //dtDataRow[37] = strBirthday; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    //dtDataRow[30] = wsraw.Range["B" + i].Value; // Policy Number; // Life ID
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                    dtDataRow[76] = strRemarksCode; // Remarks
                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status

                    dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                    dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["O" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Trans Effective Date

                    if (strFilePath.ToUpper().Contains("PHP"))
                    {
                        dtDataRow[23] = "PHP"; //  Currency
                    }
                    else
                    {
                        dtDataRow[23] = "USD"; //  Currency
                    }

                    string strtranscode = objHlpr.fn_CheckTransCodeV3(wsraw.Range["I" + i].Value);
                    if (strtranscode == "ADJUST")
                    {
                        dtDataRow[21] = strtranscode; // Transcode
                        dtDataRow[62] = "4004"; // Entry Code
                        dtDataRow[63] = wsraw.Range["L" + i].Value; // Premium
                        dtDataRow[76] = wsraw.Range["I" + i].Value; // Remarks

                    }
                    else
                    {
                        dtDataRow[21] = strtranscode; // Transcode
                        dtDataRow[60] = "4004"; // Entry Code
                        dtDataRow[61] = wsraw.Range["L" + i].Value; // Premium
                    }

                    string strRiskType = wsraw.Range["K" + i].Value;
                    string strReasonOfChange = wsraw.Range["I" + i].Value;
                    if (strRiskType.ToUpper().Contains("DEATH") && strReasonOfChange.ToUpper().Contains("WITHDRAWAL"))
                    {
                        dtDataRow[3] = "DEATH"; // Benefit Covered  
                        dtDataRow[4] = "VARIABLELIFE-RE"; // Insurance Product
                        //dtDataRow[5] = "DEATH - VARIABLELIFE-RE"; // Branded Product Cedent Code
                    }
                    else if(strRiskType.ToUpper().Contains("DEATH"))
                    {
                        dtDataRow[3] = "DEATH"; // Benefit Covered  
                        dtDataRow[4] = "TRADITIONALLIFE"; // Insurance Product
                        //dtDataRow[5] = "DEATH - TRADITIONALLIFE"; // Branded Product Cedent Code
                    }
                    else if(strRiskType.ToUpper().Contains("ADB"))
                    {
                        dtDataRow[3] = "DEATH"; // Benefit Covered  
                        dtDataRow[4] = "ADB-I"; // Insurance Product
                        //dtDataRow[5] = "DEATH - ADB-I"; // Branded Product Cedent Code
                    }
                    else if (strRiskType.ToUpper().Contains("ADD"))
                    {
                        dtDataRow[3] = "DEATH"; // Benefit Covered  
                        dtDataRow[4] = "AD&D-IND"; // Insurance Product
                        //dtDataRow[5] = "DEATH - AD&D-IND"; // Branded Product Cedent Code
                    }



                    dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["L" + i].Value);
                    dblTotalSumAtRisk = dblTotalSumAtRisk + Convert.ToDecimal(strSumAtRisk);

                    if (wsraw.Range["AD" + i].Value == "UL")
                    {
                        dblTotalPremiumUL = dblTotalPremiumUL + Convert.ToDecimal(wsraw.Range["L" + i].Value);
                        dblTotalSumAtRiskUL = dblTotalSumAtRiskUL + Convert.ToDecimal(strSumAtRisk);
                    }
                    else
                    {
                        dblTotalPremiumTRAD = dblTotalPremiumTRAD + Convert.ToDecimal(wsraw.Range["L" + i].Value);
                        dblTotalSumAtRiskTRAD = dblTotalSumAtRiskTRAD + Convert.ToDecimal(strSumAtRisk);
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 005", "Information");
                return "";
            }

            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Premium UL:";
                dtDataRow[1] = dblTotalPremiumUL;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Premium TRAD:";
                dtDataRow[1] = dblTotalPremiumTRAD;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk UL:";
                dtDataRow[1] = dblTotalSumAtRiskUL;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Sum at Risk TRAD:";
                dtDataRow[1] = dblTotalSumAtRiskTRAD;
                objdt_template.Rows.Add(dtDataRow);
            #endregion



            string despath = str_saved + @"\BM005-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";


        }
    }
}