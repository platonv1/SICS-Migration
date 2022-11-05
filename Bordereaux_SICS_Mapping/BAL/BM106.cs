using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM106
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);


            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int intLastRow = wsraw.Range["A1"].End[XlDirection.xlDown].Row;
            string valueTransEffectiveDate = string.Empty;
            decimal TotalPremium = 0;
            decimal TotalSumAtRisk = 0;
            /*string strPolicyYear = strFilePath.Substring(strFilePath.Length - 6);
            strPolicyYear = strPolicyYear.Insert(2, "/");
            DateTime PolicyYear = DateTime.ParseExact(strPolicyYear, "MM/yyyy", CultureInfo.InvariantCulture);
            strPolicyYear = PolicyYear.ToString("MM/yyyy");*/

            DataRow dtDataRow;
       

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
            }

            if (str_sheet =="Inforce")
            {
                for (int i = 5; i <= intLastRow; i++)
                {
                    string strPolicyNo = Convert.ToString(wsraw.Range["B" + i].Value); // Policy Number
                    if (string.IsNullOrEmpty(strPolicyNo)) {
                        break;
                    }
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, Convert.ToString(wsraw.Range["B" + i].Value), Convert.ToString(wsraw.Range["C" + i].Value), Convert.ToString(wsraw.Range["D" + i].Value)))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Range["L" + i].Value), Convert.ToString(wsraw.Range["U" + i].Value), Convert.ToString(wsraw.Range["U" + i].Value), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                    dtDataRow[28] = objHlpr.fn_computecededretention(Convert.ToString(wsraw.Range["L" + i].Value), Convert.ToString(wsraw.Range["M" + i].Value)); //ceded retention
                    dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOriginalSum); // Original Sum Assured
                    decimal dclPremium = Convert.ToDecimal(wsraw.Range ["W" + i].Value);//RY
                    dtDataRow [59] = dclPremium;
                    dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Range ["M" + i].Value)); // ceded sum
                    string strFullName = wsraw.Range["C" + i].Value; //Full Name
                    dtDataRow[31] = strFullName;
                    objHlpr2.fn_separateFirstNameLastNameV1(Convert.ToString(wsraw.Range["C" + i].Value), out string strFirstName, out string strLastName);
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = strLastName;
                    string strSex = wsraw.Range["F" + i].Value;
                    dtDataRow[36] = strSex; //Gender
                    string strBirthday = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Range["D" + i].Value));
                    dtDataRow[37] = strBirthday; // Birthday
                    dtDataRow[5] =  wsraw.Range["I" + i].Value; // Branded Product
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[14] = "T"; // Business Type
                    dtDataRow[24] = "MLY"; // Premium Frequency
                    dtDataRow [23] = Convert.ToString(wsraw.Range ["K" + i].Value.ToUpper());//Currency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); // Life ID
                    dtDataRow[41] = Variables.strBmYear; /*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[79] = Convert.ToString(wsraw.Range["E" + i].Value);//Life Issue Age
                    string strTcode = "TRENEW";
                    dtDataRow[21] = strTcode;
                    string strIssueDate = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Range["G" + i].Value));
                    dtDataRow[22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow[20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    dtDataRow[19] = valueTransEffectiveDate;  // Reinsurance Start Date
                    dtDataRow[38] = "NONE"; // Smoker Status
                    dtDataRow[58] = "4001"; // Entry code

                    #region HashTotal
                    if(dclPremium == 0)
                    {
                        dtDataRow [27] = 1;
                        dtDataRow [77] = 1;
                    }
                    else
                    {
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(strInitialSum); // Initial Sum at Risk
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(strSumAtRisk); // Sum at Risk
                        TotalSumAtRisk += Convert.ToDecimal(strSumAtRisk);
                        TotalPremium += dclPremium;
                    }

                    #endregion
                    objHlpr.fn_GetRemarksCode(strBirthday, strFullName, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksAABBZ + "|" + strRemarksCode; // Remarks
                }

            }
         
            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = TotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = TotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);
            #endregion


            string despath = str_saved + @"\BM106-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}
