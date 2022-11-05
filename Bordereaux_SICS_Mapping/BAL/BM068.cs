using System;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM068
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            try
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
                double dblTotalPremium = 0, dblTotalSumAtRisk = 0;

                while (string.IsNullOrEmpty(Variables.strBmYear))
                {

                    if (string.IsNullOrEmpty(Variables.strBmYear))
                    {
                        frmPolicyYear newform = new frmPolicyYear();
                        newform.ShowDialog();

                    }
                }

                for (int i = 1; i <= intLastRow; i++)
                {
                    if (wsraw.Range["A" + i].Value != null)
                    {
                        string strCessionNo = Convert.ToString(wsraw.Range["A" + i].Value);
                        if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                        {
                            dtDataRow = objdt_template.NewRow();
                            objdt_template.Rows.Add(dtDataRow);
                            Console.WriteLine(wsraw.Range ["A" + i].Value);
                            dtDataRow [0] = wsraw.Range["A" + i].Value; // Policy Number
                            dtDataRow [36] = wsraw.Range["K" + i].Value; // Gender
                            string full = Convert.ToString(wsraw.Range ["C" + i].Value);
                            //objHlpr.fn_separatefullnamev6(Convert.ToString(wsraw.Range["C" + i].Value), out string strFirstName, out string strLastName, out string strMiddleInitial);
                            objHlpr2.fn_separateLastNameFirstNameV10(Convert.ToString(wsraw.Range ["C" + i].Value), out string strLastName, out string strFirstName, out string strMiddleInitial);
                            Console.WriteLine(strLastName + ", " + strFirstName + " " + strMiddleInitial + i);
                            dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                            dtDataRow [32] = strLastName; // Last Name
                            dtDataRow[33] = strFirstName; // First Name
                            dtDataRow[34] = strMiddleInitial; // Middle Initials
                            string strBirthday = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["J" + i].Value)).ToString("MM/dd/yyyy");
                            dtDataRow[37] = strBirthday; // Birthday
                            dtDataRow[29] = "NATREID"; // Life ID Type
                            dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                            dtDataRow[79] = wsraw.Range["P" + i].Value; // Life Issue Age
                            dtDataRow[8] = "COMBINE"; // Reinsurance Product
                            dtDataRow [5] = "GCL"; // Branded Product
                            dtDataRow [9] = "PA"; // Type of Business
                            dtDataRow[10] = "Q"; // Reinsurance Methods
                            dtDataRow[13] = "GRP"; // Class of Business
                            dtDataRow[14] = "T"; // Business Type
                            dtDataRow[24] = "MLY"; // Premium Frequency
                            dtDataRow[38] = "NONE"; // Smoker Status
                            dtDataRow[41] = Variables.strBmYear; // Policy Year
                            dtDataRow[39] = objHlpr.fn_getmortality(wsraw.Range["L" + i].Text); // Preferred Classific

                            
                            //dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                            //dtDataRow [20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["H" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                            //dtDataRow [22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["I" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                            string strIssueDate = Convert.ToDateTime(wsraw.Range ["H" + i].Value).ToString("MM/dd/yyyy");
                            objHlpr2.fn_getTransReinsuranceDateV5(strIssueDate, Variables.strBmYear, out string transEffectiveDate, out string transcode);
                            dtDataRow [22] = transEffectiveDate; //Transeffective date
                            dtDataRow [20] = strIssueDate;//Policy Start Date
                            dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date
                            double.TryParse(Convert.ToString(wsraw.Range ["T" + i].Value), out double dblInitialSum); //sum at risk
                            dblInitialSum = dblInitialSum * 0.10;

                            dtDataRow [25] = wsraw.Range ["F" + i].Value;//orig sum
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblInitialSum));//initial sum /sum at risk
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblInitialSum));//initial sum /sum at risk//initial sum /sum at risk

                            if (transcode == "TNEWBUS")
                            {
                                dtDataRow[21] = transcode; // Transcode
                                dtDataRow[56] = "4000"; // Entry Code   
                                dtDataRow[57] = wsraw.Range["V" + i].Value; // Premium
                            }
                            else
                            {
                                dtDataRow[21] = transcode; // Transcode
                                dtDataRow[58] = "4001"; // Entry Code
                                dtDataRow[59] = wsraw.Range["V" + i].Value; // Premium
                            }
                            double.TryParse(Convert.ToString(wsraw.Range ["V" + i].Value), out double dblPrem);
                            if(dblPrem > 0)
                            {
                                dblTotalSumAtRisk += dblInitialSum;
                            }
                            dblPrem = dblPrem * 0.10;
                            dblTotalPremium += dblPrem;
                            
                       
                        }
                    }
                }
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow[0] = "Total Premium:";
                dtDataRow[1] = dblTotalPremium;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Sum at Risk:";
                dtDataRow [1] = dblTotalSumAtRisk;
                objdt_template.Rows.Add(dtDataRow);

                string despath = str_saved + @"\BM068-" + str_sheet + str_savef + ".xlsx";
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