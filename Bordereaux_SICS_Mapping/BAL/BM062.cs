using System;
using System.Data;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM062
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);
            Application eapp = new Application();
            Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Worksheet wsraw = wbraw.Worksheets [str_sheet];

            int intLastRow = wsraw.Cells [wsraw.Rows.Count, 1].End [XlDirection.xlUp].row;

            DataRow dtDataRow;
            double dblTotalPremiumGME = 0, dblTotalPremiumGRP = 0, dblTotalEWTPremiumDueGME = 0, dblTotalEWTPremiumDueGRP = 0, dblTotalPTaxGME = 0, dblTotalPTaxGRP = 0, dblTotalCommGME = 0, dblTotalCommGRP = 0, dblTotalERGME = 0, dblTotalERGRP = 0, dblTotalClaimsGME = 0, dblTotalClaimsGRP = 0, dblTotalClaimsIBNRGME = 0, dblTotalClaimsIBNRGRP = 0, dblTotalEWTCommissionGME = 0, dblTotalEWTCommissionGRP = 0;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {

                if(string.IsNullOrEmpty(Variables.strBmYear))
                {
                    frmPolicyYear newform = new frmPolicyYear();
                    newform.ShowDialog();

                }
            }

            if(str_sheet.ToUpper().Contains("RI"))
            {
                for(int i = 3; i <= intLastRow; i++)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);


                    var strPolno = wsraw.Range["C" + i].Value;
                    string strPolicyNumber;
                    if (strPolno.GetType() != typeof(string))
                    {
                        strPolicyNumber = strPolno.ToString("0");
                    }
                    else
                    {
                        strPolicyNumber = strPolno;
                    }
                    strPolicyNumber = strPolicyNumber.TrimEnd();
                    strPolicyNumber = strPolicyNumber.Replace("-", "");
                    if (strPolicyNumber.Length > 7)
                    {
                        strPolicyNumber = strPolicyNumber.Substring(strPolicyNumber.Length - 7);
                    }
                    dtDataRow[0] = strPolicyNumber + "DD07011900"; // Policy Number
                    dtDataRow[36] = "M"; // Gender
                    dtDataRow[31] = "DummyLastName" + ", " + "DummyFirstName" + " " + "DummyMiddleName" + "."; // Full name
                    dtDataRow[32] = "DummyLastName"; // Last Name
                    dtDataRow[33] = "DummyFirstName"; // First Name
                    dtDataRow[34] = "DummyMiddleName"; // Middle Initials
                    dtDataRow[37] = "07/01/1900"; // Birthday
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = "DUMMYDU07011900"; // Life ID
                    dtDataRow[81] = wsraw.Range ["Q" + i].Value; // Number of Lives
                    dtDataRow[82] = wsraw.Range ["B" + i].Value; // Group Policyholder
                    //dtDataRow[25] = wsraw.Range ["R" + i].Value; // Original Sum Assured
                    dtDataRow[7] = wsraw.Range ["C" + i].Value; // Group Scheme ID
                    dtDataRow[41] = Variables.strBmYear; // Policy Year
                    dtDataRow[21] = "TRENEW"; // Transcode
                    dtDataRow[8] = "QA"; // Reinsurance Product
                    dtDataRow[9] = "PA"; // Type of Business
                    dtDataRow[10] = "Q"; // Reinsurance Methods
                    dtDataRow[13] = "GRP"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[23] = "PHP"; // Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[38] = "NONE"; // Smoker Status
                    dtDataRow[39] = "STANDARD"; // Preferred Classific

                    dtDataRow[19] = objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                    dtDataRow[20] = objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["F" + i].Value)).ToString("MM/dd/yyyy"); // Policy Start Date
                    dtDataRow[22] = objHlpr.fn_convertStringtoDateV3(Convert.ToString(wsraw.Range["H" + i].Value)).ToString("MM/dd/yyyy"); // Reinsurance Start Date

                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["R" + i].Value)); // Original Sum Assured
                    dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["R" + i].Value * 0.05)); // Initial Sum at Risk
                    dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range["R" + i].Value * 0.05)); // Sum at Risk

                    dtDataRow[5] = wsraw.Range["E" + i].Value; // Branded Product Cedent Code

                    if (wsraw.Range["E" + i].Value == "GME")
                    {
                        //dtDataRow[4] = "MEDICAL"; // Insurance Product
                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["AL" + i].Value + wsraw.Range["V" + i].Value + wsraw.Range["BD" + i].Value + wsraw.Range["BE" + i].Value + wsraw.Range["BT" + i].Value - wsraw.Range["Y" + i].Value - wsraw.Range["BH" + i].Value;  // Premium
                        dblTotalPremiumGME = dblTotalPremiumGME + (wsraw.Range["AL" + i].Value + wsraw.Range["V" + i].Value + wsraw.Range["BD" + i].Value + wsraw.Range["BE" + i].Value + wsraw.Range["BT" + i].Value - wsraw.Range["Y" + i].Value - wsraw.Range["BH" + i].Value);
                        dblTotalEWTPremiumDueGME = dblTotalEWTPremiumDueGME + wsraw.Range["W" + i].Value;
                        dblTotalPTaxGME = dblTotalPTaxGME + wsraw.Range["AD" + i].Value;
                        dblTotalCommGME = dblTotalCommGME + (wsraw.Range["AA" + i].Value + wsraw.Range["AE" + i].Value);
                        dblTotalERGME = dblTotalERGME + (wsraw.Range["AC" + i].Value     wsraw.Range["BN" + i].Value - wsraw.Range["AX" + i].Value);
                        dblTotalClaimsGME = dblTotalClaimsGME + (wsraw.Range["X" + i].Value + wsraw.Range["BJ" + i].Value - wsraw.Range["AP" + i].Value);
                        dblTotalClaimsIBNRGME = dblTotalClaimsIBNRGME + (wsraw.Range["BL" + i].Value - wsraw.Range["AT" + i].Value);
                        dblTotalEWTCommissionGME = dblTotalEWTCommissionGME + (wsraw.Range["AB" + i].Value + wsraw.Range["AF" + i].Value);
                    }
                    else
                    {
                        //dtDataRow[4] = "Group Life"; // Insurance Product
                        dtDataRow[58] = "4001"; // Entry Code
                        dtDataRow[59] = wsraw.Range["AL" + i].Value + wsraw.Range["V" + i].Value + wsraw.Range["BD" + i].Value + wsraw.Range["BE" + i].Value + wsraw.Range["BT" + i].Value - wsraw.Range["Y" + i].Value - wsraw.Range["BH" + i].Value;  // Premium
                        dblTotalPremiumGRP = dblTotalPremiumGRP + (wsraw.Range["AL" + i].Value + wsraw.Range["V" + i].Value + wsraw.Range["BD" + i].Value + wsraw.Range["BE" + i].Value + wsraw.Range["BT" + i].Value - wsraw.Range["Y" + i].Value - wsraw.Range["BH" + i].Value);
                        dblTotalEWTPremiumDueGRP = dblTotalEWTPremiumDueGRP + wsraw.Range["W" + i].Value;
                        dblTotalPTaxGRP = dblTotalPTaxGRP + wsraw.Range["AD" + i].Value;
                        dblTotalCommGRP = dblTotalCommGRP + (wsraw.Range["AA" + i].Value + wsraw.Range["AE" + i].Value);
                        dblTotalERGRP = dblTotalERGRP + (wsraw.Range["AC" + i].Value + wsraw.Range["BN" + i].Value - wsraw.Range["AX" + i].Value);
                        dblTotalClaimsGRP = dblTotalClaimsGRP + (wsraw.Range["X" + i].Value + wsraw.Range["BJ" + i].Value - wsraw.Range["AP" + i].Value);
                        dblTotalClaimsIBNRGRP = dblTotalClaimsIBNRGRP + (wsraw.Range["BL" + i].Value - wsraw.Range["AT" + i].Value);
                        dblTotalEWTCommissionGRP = dblTotalEWTCommissionGRP + (wsraw.Range["AB" + i].Value + wsraw.Range["AF" + i].Value);
                    }

                    dtDataRow[64] = "1501"; // Entry Code EWT on Premium Due
                    dtDataRow[65] = wsraw.Range["W" + i].Value; // EWT on Premium Due


                    dtDataRow[66] = "5008"; // Entry Code P. Tax
                    dtDataRow[67] = wsraw.Range["AD" + i].Value; // P.Tax
                    

                    dtDataRow[68] = "5005"; // Entry Code Comm
                    dtDataRow[69] =  wsraw.Range["AA" + i].Value + wsraw.Range["AE" + i].Value; // Comm
                    

                    dtDataRow[70] = "4003"; // Entry Code ER
                    dtDataRow[71] = wsraw.Range["AC" + i].Value + wsraw.Range["BN" + i].Value - wsraw.Range["AX" + i].Value; // ER
                    

                    dtDataRow[72] = "5003"; // Entry Code Claims
                    dtDataRow[73] = (wsraw.Range["X" + i].Value + wsraw.Range["BJ" + i].Value - wsraw.Range["AP" + i].Value) + (wsraw.Range["BL" + i].Value - wsraw.Range["AT" + i].Value); // Claims
                    

                    //dtDataRow[72] = "5003"; // Entry Code Claims IBNR
                    dtDataRow[76] = "Claims IBNR " + Convert.ToString(wsraw.Range["BL" + i].Value - wsraw.Range["AT" + i].Value); // Remarks Claims IBNR
                    

                    dtDataRow[74] = "2000"; // Entry Code EWT Comm
                    dtDataRow[75] = wsraw.Range["AB" + i].Value + wsraw.Range["AF" + i].Value; // EWT Comm
                    
                }


            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The sheet is not included in BM 062", "Information");
                return "";
            }

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium of Group Life:";
            dtDataRow [1] = dblTotalPremiumGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total EWT on Premium Due:";
            dtDataRow[1] = dblTotalEWTPremiumDueGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total P. Tax:";
            dtDataRow[1] = dblTotalPTaxGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Commission:";
            dtDataRow[1] = dblTotalCommGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total ER:";
            dtDataRow[1] = dblTotalERGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Claims:";
            dtDataRow[1] = dblTotalClaimsGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Claims IBNR:";
            dtDataRow[1] = dblTotalClaimsIBNRGRP;
            objdt_template.Rows.Add(dtDataRow);
            
            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total EWT on Commission:";
            dtDataRow[1] = dblTotalEWTCommissionGRP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium of Medical:";
            dtDataRow[1] = dblTotalPremiumGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total EWT on Premium Due:";
            dtDataRow[1] = dblTotalEWTPremiumDueGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total P. Tax:";
            dtDataRow[1] = dblTotalPTaxGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Commission:";
            dtDataRow[1] = dblTotalCommGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total ER:";
            dtDataRow[1] = dblTotalERGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Claims:";
            dtDataRow[1] = dblTotalClaimsGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Claims IBNR:";
            dtDataRow[1] = dblTotalClaimsIBNRGME;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total EWT on Commission:";
            dtDataRow[1] = dblTotalEWTCommissionGME;
            objdt_template.Rows.Add(dtDataRow);

            string despath = str_saved + @"\BM062-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";
        }
    }
}