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
    class BM059
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
            double dblTotalPremiumT = 0, dblTotalSumAtRiskT = 0, dblTotalPremiumF = 0, dblTotalSumAtRiskF = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

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

            if (str_sheet.ToUpper().Contains("BFB") || str_sheet.ToUpper().Contains("PVB"))
            {
                for (int i = 1; i <= intLastRow; i++)
                {
                    if (wsraw.Range["D" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range ["D" + i].Value);
                        if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);

                            var strPolno = wsraw.Range["D" + i].Value;
                            string strPolicyNumber; string bmyear = Variables.strBmYear;
                            if (strPolno.GetType() != typeof(string))
                            {
                                strPolicyNumber = strPolno.ToString("0");
                            }
                            else
                            {
                                strPolicyNumber = strPolno;
                            }
                            
                            //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                            objHlpr2.fn_separateLastNameFirstNameV12(wsraw.Range["E" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                            dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow[32] = strLastName; // Last Name
                            dtDataRow[33] = strFirstName; // First Name
                            dtDataRow[34] = strMiddleInitial; // Middle Initials
                            dtDataRow [0] = wsraw.Range ["D" + i].Value; // Cession Number
                            string strBirthday = "07/01/1900"; // Birthday;
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type

                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow[79] = wsraw.Range["F" + i].Value; // Life Issue Age
                            strBirthday = strBirthday.Replace("/", "");
                            if (strPolicyNumber.Length > 7)
                            {
                                strPolicyNumber = strPolicyNumber.Substring(strPolicyNumber.Length - 7);
                            }
                            dtDataRow[1] = strPolicyNumber + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                            dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                            dtDataRow[9] = "PAFM"; // Type of Business
                            dtDataRow[10] = "S"; // Reinsurance Methods
                            dtDataRow[13] = "GCL"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow [5] = "GFSP"; // Branded Product
                            dtDataRow [23] = "PHP"; // Cession Currency
                            dtDataRow[24] = "YLY"; // Premium Frequency
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow[28] = wsraw.Range["I" + i].Value; // Retention
                            dtDataRow[39] = objHlpr.fn_getmortality(Convert.ToString(wsraw.Range["O" + i].Value)); // Preffered Classific
                            dtDataRow[7] = (str_sheet.ToUpper().Contains("BFB")) ? "GFSP_BFB" : "GFSP_PVB"; // Group Scheme ID
                            dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["H" + i].Text)); // Original Sum Assured
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["J" + i].Text)); // Initial Sum at Risk
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["J" + i].Text)); // Sum at Risk
                            dtDataRow[41] = Variables.strBmYear; // Policy Year
                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["G" + i].Value).ToString("MM/dd/yyyy");
                            objHlpr2.fn_getTranscodeV2(str_sheet, out string strTransCode);
                            objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out string transEffectiveDate, out string policyStartDate); // Policy Start Date
                            if (str_sheet.ToUpper().Contains("FY"))
                            {
                                dtDataRow[21] = strTransCode; // Transcode
                                dtDataRow[56] = "4000"; // Entry Code
                                dtDataRow[57] = wsraw.Range["K" + i].Value; // Premium
                                dtDataRow [19] = transEffectiveDate;
                                dtDataRow [22] = transEffectiveDate; // Trans Effective Dat e
                                dtDataRow [20] = transEffectiveDate;//Reinsurance Start Date
                            }
                            else if (str_sheet.ToUpper().Contains("RY"))
                            {
                                dtDataRow[21] = strTransCode; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["K" + i].Value; // Premium
                                dtDataRow [19] = transEffectiveDate;//Reinsurance Start Date
                                dtDataRow [22] = transEffectiveDate; // Trans Effective Date
                                dtDataRow [20] = policyStartDate;
                            }

                            dblTotalPremiumT = dblTotalPremiumT + Convert.ToDouble(wsraw.Range["K" + i].Value);
                            dblTotalSumAtRiskT = dblTotalSumAtRiskT + Convert.ToDouble(wsraw.Range["J" + i].Value);

                            if (wsraw.Range["L" + i].Value != null && wsraw.Range["L" + i].Value != 0)
                            {
                                dtDataRow = objdt_template.NewRow();
                                objdt_template.Rows.Add(dtDataRow);

                                dtDataRow[0] = wsraw.Range["D" + i].Value; // Policy Number
                                //dtDataRow[36] = wsraw.Range["E" + i].Value; // Gender
                                //objHlpr.fn_separatefullnamev7(wsraw.Range["E" + i].Value, out string strFirstName, out string strLastName, out string strMiddleInitial);
                                dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                                dtDataRow[32] = strLastName; // Last Name
                                dtDataRow[33] = strFirstName; // First Name
                                dtDataRow[34] = strMiddleInitial; // Middle Initials
                                strBirthday = "07/01/1900"; // Birthday;
                                dtDataRow[37] = strBirthday; // Birthday
                                dtDataRow[29] = "NATREID"; // Life ID Type
                                dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                                dtDataRow[79] = wsraw.Range["F" + i].Value; // Life Issue Age
                                strBirthday = strBirthday.Replace("/", "");
                                dtDataRow[1] = strPolicyNumber.Substring(strPolicyNumber.Length - 7) + strFirstName.Substring(0, 1) + strLastName.Substring(0, 1) + strBirthday; // PolicyNumber
                                dtDataRow [5] = "GFSP"; // Branded Product
                                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                                dtDataRow[9] = "F"; // Type of Business
                                dtDataRow[10] = "S"; // Reinsurance Methods
                                dtDataRow[13] = "GRP"; // Class of Business
                                dtDataRow[14] = "T"; // Business Type
                                dtDataRow[23] = "PHP"; // Cession Currency
                                dtDataRow[24] = "YLY"; // Premium Frequency
                                dtDataRow[38] = "NONE"; // Smoker Status
                                dtDataRow[28] = wsraw.Range["I" + i].Value; // Retention
                                dtDataRow[39] = objHlpr.fn_getmortality(Convert.ToString(wsraw.Range["P" + i].Value)); // Preffered Classific
                                dtDataRow[7] = (str_sheet.ToUpper().Contains("BFB")) ? "GFSP_BFB" : "GFSP_PVB"; // Group Scheme ID
                                dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["H" + i].Text)); // Original Sum Assured
                                dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["L" + i].Text)); // Initial Sum at Risk
                                dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["L" + i].Text)); // Sum at Risk
                                dtDataRow[41] = Variables.strBmYear; // Policy Year
                                strIssueDate = Convert.ToDateTime(wsraw.Range ["G" + i].Value).ToString("MM/dd/yyyy");
                                objHlpr2.fn_getTranscodeV2(str_sheet, out  strTransCode);
                                objHlpr2.fn_getTransReinsuranceDateV7(strTransCode, bmyear, strIssueDate, out  transEffectiveDate, out  policyStartDate); // Policy Start Date
                                if (str_sheet.ToUpper().Contains("FY"))
                                {
                                    dtDataRow[21] = "TNEWBUS"; // Transcode
                                    dtDataRow[56] = "4000"; // Entry Code
                                    dtDataRow[57] = wsraw.Range["M" + i].Value; // Premium
                                    dtDataRow [19] = transEffectiveDate;
                                    dtDataRow [22] = transEffectiveDate; // Trans Effective Dat e
                                    dtDataRow [20] = transEffectiveDate;//Reinsurance Start Date
                                }
                                else if (str_sheet.ToUpper().Contains("RY"))
                                {
                                    dtDataRow[21] = "TRENEW"; // Transcode
                                    dtDataRow[58] = "4001"; // Entry Code
                                    dtDataRow[59] = wsraw.Range["M" + i].Value; // Premium
                                    dtDataRow [19] = transEffectiveDate;//Reinsurance Start Date
                                    dtDataRow [22] = transEffectiveDate; // Trans Effective Date
                                    dtDataRow [20] = policyStartDate;
                                }

                                dblTotalPremiumF = dblTotalPremiumF + Convert.ToDouble(wsraw.Range["M" + i].Value);
                                dblTotalSumAtRiskF = dblTotalSumAtRiskF + Convert.ToDouble(wsraw.Range["L" + i].Value);
                            }


                        }
                    }
                }
            }


            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium of Treaty:";
            dtDataRow[1] = dblTotalPremiumT;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk of Treaty:";
            dtDataRow[1] = dblTotalSumAtRiskT;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium of Facultative:";
            dtDataRow[1] = dblTotalPremiumF;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk of Facultative:";
            dtDataRow[1] = dblTotalSumAtRiskF;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM059-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }
}
