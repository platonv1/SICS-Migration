using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM090
    {
        public string fn_process( string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false )
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets [str_sheet];

            int intLastRow = wsraw.Cells [wsraw.Rows.Count, 2].End [XlDirection.xlUp].row;

            DataRow dtDataRow;
            decimal dblTotalPremium = 0, dblTotalSumAtRisk = 0;

            while ( string.IsNullOrEmpty(Variables.strBmYear) )
            {

                if ( string.IsNullOrEmpty(Variables.strBmYear) )
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            if ( str_sheet.ToUpper().Contains("NB") )
            {
                for ( int i = 6; i <= intLastRow; i++ )
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow [0] = wsraw.Range ["C" + i].Value; // Policy Number
                    string strGender = wsraw.Range ["H" + i].Value; // Gender
                    if ( strGender.ToUpper().Contains("FEMALE") )
                    {
                        dtDataRow [36] = "F";
                    }
                    else
                    {
                        dtDataRow [36] = "M";
                    }
                    string strFullname = wsraw.Range ["E" + i].Value;
                    if ( strFullname.Contains("0") )
                    {
                        strFullname = strFullname.Substring(0, strFullname.Length - 1);

                    }
                    objHlpr.fn_separatefullnamev3(strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial);
                    dtDataRow [31] = strLastName + ", " + strFirstName + " " + strMiddleInitial + "."; // Full name
                    dtDataRow [32] = strLastName; // Last Name
                    dtDataRow [33] = strFirstName; // First Name
                    dtDataRow [34] = objHlpr2.fn_removeCharacters(strMiddleInitial); // Middle Initials
                    string strBirthday = objHlpr.fn_reformatDatev2(Convert.ToString(wsraw.Range ["I" + i].Value)).ToString("MM/dd/yyyy");
                    dtDataRow [37] = strBirthday; // Birthday
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    dtDataRow [79] = wsraw.Range ["J" + i].Value; // Life Issue Age
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [13] = "GRP"; // Class of Business
                    dtDataRow [14] = "T"; // Business Type
                    dtDataRow [23] = "PHP"; // Cession Currency
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [38] = "NONE"; // Smoker Status
                    dtDataRow [39] = "STANDARD"; // Preferred Clasification
                    dtDataRow [41] = Variables.strBmYear; // Policy Year
                    dtDataRow[5] = "MRI"; // BRANDED_PRODUCT_CEDENT_CODE
                    dtDataRow [20] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range ["D" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // REINSURANCE_START_DATE
                    dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Range["D" + i].Value)).ToString("MM/dd/yyyy"); // TRANS_EFFECTIVE_DATE

                    dtDataRow[26] = wsraw.Range["AC" + i].Value; // Ceded Sum Assured
                    dtDataRow[25] = wsraw.Range["M" + i].Value; // Original Sum Assured
                    dtDataRow[27] = wsraw.Range["AG" + i].Value; // Initial Sum at Risk
                    dtDataRow[77] = wsraw.Range["AG" + i].Value; // Sum at Risk

                    objHlpr.fn_CheckTransCode(wsraw.Range ["K" + i].Value, out string transcode);
                    dtDataRow [21] = transcode; // Transcode

                    if ( transcode == "TNEWBUS" )
                    {
                        dtDataRow [56] = "4000"; // Entry Code
                        dtDataRow [57] = wsraw.Range ["AO" + i].Value; // Premium
                    }
                    else
                    {
                        dtDataRow [58] = "4001"; // Entry Code
                        dtDataRow [59] = wsraw.Range ["AO" + i].Value; // Premium
                    }
                    dblTotalPremium += Convert.ToDecimal(wsraw.Range ["AO" + i].Value);
                    dblTotalSumAtRisk += Convert.ToDecimal(wsraw.Range["AG" + i].Value);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 090", "Information");
                return "";
            }

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dblTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dblTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM090-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}
