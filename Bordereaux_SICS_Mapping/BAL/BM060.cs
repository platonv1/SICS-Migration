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
    class BM060
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender, bool boo_open = false, bool boo_clean = false, string str_macro = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            System.Data.DataTable dt_macro = new System.Data.DataTable();
            if (!String.IsNullOrEmpty(str_macro))
            {
                dt_macro = objHlpr.fn_Loadmacro(str_macro);
            }

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets[str_sheet];

            int intLastRow = wsraw.Cells[wsraw.Rows.Count, 1].End[XlDirection.xlUp].row;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {

                if (string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            #region CONNECTION TO DATABASE
            string szConnect = "DSN=SICS_Postgres_DB;" +
                                   "UID=sics;" +
                                   "PWD=sics_1";

            OdbcConnection cnDB = new OdbcConnection(szConnect);
            cnDB.Open();
            string query = "SELECT * FROM dbo_gender";
            OdbcCommand command = new OdbcCommand(query, cnDB);

            OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            while (reader.Read() == true)
            {
                Console.WriteLine("New Row:");
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    Console.WriteLine(reader.GetString(i));
                }
            }
            reader.Close();
            cnDB.Close();
            #endregion

            for (int i = 1; i <= intLastRow; i++)
            {
                if (wsraw.Range["C" + i].Value != null)
                {
                    string strCessionNo = Convert.ToString(wsraw.Range["C" + i].Value);
                    if (Regex.IsMatch(strCessionNo, @"^\d+$"))
                    {
                        dtDataRow = objdt_template.NewRow();
                        objdt_template.Rows.Add(dtDataRow);

                        dtDataRow[0] = wsraw.Range["C" + i].Value; // Policy Number
                        objHlpr.fn_searchpolicydb(wsraw.Range["C" + i].Value, out string strAge, out string strFullName, out string strBirthdate, out string strGender);
                        dtDataRow[36] = strGender; // Gender
                        objHlpr.fn_separatefullnamev4(strFullName, out string strFirstName, out string strLastName, out string strMiddleInitial);
                        dtDataRow[31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + ".";
                        dtDataRow[32] = strLastName; // Last Name
                        dtDataRow[33] = strFirstName; // First Name
                        dtDataRow[34] = strMiddleInitial; // Middle Initials
                        dtDataRow[37] = objHlpr.fn_reformatDate(Convert.ToString(strBirthdate)).ToString("MM/dd/yyyy"); // Birthday
                        dtDataRow[29] = "NATREID"; // Life ID Type
                        dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthdate); // Life ID
                        dtDataRow[79] = strAge; // Life Issue Age

                        dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                        dtDataRow[9] = "PAFM"; // Type of Business
                        dtDataRow[10] = "S"; // Reinsurance Methods
                        dtDataRow[13] = "IND"; // Class of Business
                        dtDataRow[14] = "T"; // Business Type
                        dtDataRow[23] = "USD"; // Cession Currency
                        dtDataRow[24] = "YLY"; // Premium Frequency
                        dtDataRow[38] = "NONE"; // Smoker Status
                        dtDataRow[41] = Variables.strBmYear; // Policy Year

                        dtDataRow[21] = "TNEWBUS"; // Transcode
                        dtDataRow[56] = "4000"; // Entry Code
                        dtDataRow[57] = wsraw.Range["F" + i].Value; // Premium

                        if (Regex.IsMatch(Convert.ToString(wsraw.Range["F" + i].Value), @"^\d+$"))
                        {
                            dblTotalPremium = dblTotalPremium + Convert.ToDecimal(wsraw.Range["F" + i].Value);
                        }
                    }
                }
            }

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM060-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
