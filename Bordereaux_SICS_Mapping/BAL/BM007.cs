using System;
using System.Data;
using System.Data.Odbc;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;


namespace Bordereaux_SICS_Mapping.BAL
{
    public class BM007
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

            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 1].End[XlDirection.xlUp].row;

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            #region CONNECTION TO DATABASE
            //string szConnect = "DSN=SICS_Postgres_DB;" +
            //                       "UID=sics;" +
            //                       "PWD=sics_1";

            //OdbcConnection cnDB = new OdbcConnection(szConnect);

            ////try
            ////{
            //cnDB.Open();
            //string query = "SELECT * FROM dbo_gender";
            //OdbcCommand command = new OdbcCommand(query, cnDB);

            //OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            //while (reader.Read() == true)
            //{
            //    Console.WriteLine("New Row:");
            //    for (int i = 0; i < reader.FieldCount; i++)
            //    {
            //        Console.WriteLine(reader.GetString(i));
            //    }
            //}
            //reader.Close();
            //cnDB.Close();
            #endregion

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            if (Regex.IsMatch(str_sheet, @"^\d+$"))
            {
                for (int i = 3; i <= intLastRow; i++)
                {
                    if (wsraw.Range["C" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range["C" + i].Value);
                        if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            var strPolno = wsraw.Range["C" + i].Value;
                            string strPolicyNumber = strPolno.ToString("0");
                            //objHlpr2.fn_separateLastNameFirstNameV7(wsraw.Range["D" + i].Value, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            objHlpr2.fn_separateLastNameFirstNameV11(wsraw.Range ["D" + i].Value, out string strLastName, out string strFirstName, out string strMiddleInitial);
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow[32] = strLastName; // Last Name
                            dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName); // First Name
                            dtDataRow [34] = strMiddleInitial; // Middle Initials
                            string strSex = objHlpr.fn_getgenderv2(strFirstName);
                            dtDataRow [36] = strSex; // Gender
                            string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["E" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            strBirthday = strBirthday.Replace("/", "");
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                            dtDataRow[1] = wsraw.Range["A" + i].Value; // Cedent Cession Number
                            dtDataRow[7] = wsraw.Range["C" + i].Value; // Group Scheme ID
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "GRP"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow[23] = "PHP"; // Currency
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow[41] = Variables.strBmYear; // Policy Year
                            dtDataRow[28] = wsraw.Range["T" + i].Value; // Cedent Retention
                            if (wsraw.Range["J" + i].Value == "TPD")
                            {
                                dtDataRow[5] = wsraw.Range["J" + i].Value; // BRANDED_PRODUCT_CEDENT_CODE
                            }
                            else
                            {
                                dtDataRow[5] = wsraw.Range["B" + i].Value; // BRANDED_PRODUCT_CEDENT_CODE
                            }
                            dtDataRow[39] = objHlpr.fn_GetMortality(Convert.ToDouble(wsraw.Range["F" + i].Value)); // Preffered Classific
                            dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                            dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                            dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                            objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["K" + i].Value), Convert.ToString(wsraw.Range["L" + i].Value), Convert.ToString(wsraw.Range["L" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode);

                            dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk

                            objHlpr.fn_CheckTransCode(wsraw.Range["S" + i].Value, out string transcode);
                            dtDataRow[21] = transcode; // Transcode

                            if (transcode == "TNEWBUS")
                            {
                                dtDataRow[56] = "4000"; // Entry Code
                                dtDataRow[57] = wsraw.Range["N" + i].Value + wsraw.Range["P" + i].Value; // Premium
                            }
                            else
                            {
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["N" + i].Value + wsraw.Range["P" + i].Value; // Premium
                            }

                            dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["N" + i].Value) + Convert.ToDecimal(wsraw.Range["P" + i].Value);
                            dblTotalSumAtRisk = dblTotalSumAtRisk + Convert.ToDecimal(wsraw.Range["L" + i].Value);
                        }
                    }
                }
            }
            //Adjustment
            else if (str_sheet.ToUpper().Contains("ADJUSTMENT"))
            {
                for (int i = 3; i <= intLastRow; i++)
                {
                    if (wsraw.Range["C" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range["C" + i].Value);
                        if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            var strPolno = wsraw.Range["C" + i].Value;
                            string strPolicyNumber = strPolno.ToString("0");
                            objHlpr.fn_separatefullnamev4(wsraw.Range["D" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow[32] = strLastName; // Last Name
                            dtDataRow[33] = objHlpr2.fn_checkFirstname(strFirstName); // First Name
                            dtDataRow [34] = strMiddleInitial; // Middle Initials
                            string strSex = objHlpr.fn_getgenderv2(strFirstName);
                            dtDataRow[36] = strSex; // Gender
                            string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["E" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            strBirthday = strBirthday.Replace("/", "");
                            dtDataRow[0] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                            dtDataRow[1] = wsraw.Range["A" + i].Value; // Cedent Cession Number
                            dtDataRow[7] = wsraw.Range["C" + i].Value; // Group Scheme ID
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "GRP"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow[23] = "PHP"; // Currency
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow[41] = Variables.strBmYear; // Policy Year
                            dtDataRow[39] = objHlpr.fn_GetMortality(Convert.ToDouble(wsraw.Range ["F" + i].Value));// Preffered Classific
       
                            if (wsraw.Range["H" + i].Value == "TPD")
                            {
                                dtDataRow[5] = wsraw.Range["H" + i].Value; // BRANDED_PRODUCT_CEDENT_CODE
                            }
                            else
                            {
                                dtDataRow[5] = wsraw.Range["B" + i].Value; // BRANDED_PRODUCT_CEDENT_CODE
                            }
                            dtDataRow[20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                            dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                            dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["G" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                            dtDataRow[25] = "1"; // Original Sum Assured
                            dtDataRow[27] = "1"; // Initial Sum at Risk
                            dtDataRow[77] = "1"; // Sum at Risk


                            dtDataRow[21] = "ADJUST"; // Transcode
                            dtDataRow[60] = "4002"; // Entry Code
                            dtDataRow[61] = wsraw.Range["L" + i].Value + wsraw.Range["N" + i].Value + wsraw.Range["O" + i].Value + wsraw.Range["P" + i].Value + wsraw.Range["Q" + i].Value; // Premium
                            dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["L" + i].Value) + Convert.ToDecimal(wsraw.Range["N" + i].Value) + Convert.ToDecimal(wsraw.Range["O" + i].Value) + Convert.ToDecimal(wsraw.Range["P" + i].Value) + Convert.ToDecimal(wsraw.Range["Q" + i].Value);
                        }
                    }
                }
            }

            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 007", "Information");
                return "";
            }
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

            string despath = str_saved + @"\BM007-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";


        }
    }
}