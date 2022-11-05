using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM004
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
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets [str_sheet];
            Microsoft.Office.Interop.Excel.Worksheet wssummary = wbraw.Sheets ["Summary"];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;
            Microsoft.Office.Interop.Excel.Range rawrange_ = wssummary.UsedRange;

            int erawrow = rawrange.Rows.Count;
            int erawsummary = rawrange_.Rows.Count;
            string cl100 = "CI 100";
            int row = 1;
            decimal mul = 0;
            for(int i = 1; i <= erawsummary + 1; i++)
            {
                string clcheck = wssummary.Cells [row, 2].Text;
                if(clcheck != cl100)
                {
                    row++;
                    continue;
                }
                else
                {
                    if(string.IsNullOrEmpty(wssummary.Cells [row, 9].Text))
                    {
                        row++;
                        continue;

                    }
                    else
                    {
                        string mulvalue = wssummary.Cells [row, 9].Text;
                        mul = decimal.Parse(mulvalue.TrimEnd(new char [] { '%', ' ' })) / 100M;
                        break;
                    }

                }
            }


            //var mulPHP = decimal.Parse(wssummary.Cells [29, 9].Text.TrimEnd(new char [] { '%', ' ' })) / 100M;
            //var mulUSD = decimal.Parse(wssummary.Cells [36, 9].Text.TrimEnd(new char [] { '%', ' ' })) / 100M;

            //int intLastRow = wsraw.Cells [wsraw.Rows.Count, 10].End [XlDirection.xlDown].row;
            string strPolicyNo = string.Empty;
            string valueTransEffectiveDate = string.Empty;
            string effective = string.Empty;
            string stryear = string.Empty;
            int effective1;
            stryear = wsraw.Cells [4, 8].Text.ToString();
            stryear = stryear.Replace(stryear.Substring(stryear.Length - 3, 3), "-01" + stryear.Substring(stryear.Length - 3, 3));
            DateTime oDate = Convert.ToDateTime(stryear);
            int year = oDate.Year;

            decimal TotalSumAtRiskPHP = 0;
            decimal TotalSumAtRiskUSD = 0;

            decimal TotalPremiumPHP = 0;
            decimal TotalPremiumUSD = 0;


            DataRow dtDataRow;
            //while(string.IsNullOrEmpty(Variables.strBmYear))
            //{
            //    frmPolicyYear newform = new frmPolicyYear();
            //    newform.ShowDialog();

            //}

            for(int i = 1; i <= erawrow + 1; i++)
            {
                strPolicyNo = Convert.ToString(wsraw.Range ["B" + i].Value); // Policy Number
                //if (string.IsNullOrEmpty(strPolicyNo))
                //{
                //    continue;
                //}
                if(string.IsNullOrEmpty(strPolicyNo))
                {
                    continue;
                }
                if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Range ["A" + i].Value.ToString(), wsraw.Range ["C" + i].Value.ToString(), wsraw.Range ["D" + i].Value.ToString()))
                {
                    continue;
                }

                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);
                dtDataRow [0] = strPolicyNo;

                dtDataRow [5] = wsraw.Range ["C" + i].Value; // Branded Product
                dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                dtDataRow [9] = "PAFW"; // Type of Business
                dtDataRow [10] = "S"; // Reinsurance Method
                dtDataRow [13] = "IND"; // Class of Business
                dtDataRow [14] = "T"; // Type of Business

                string strIssueDate = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Range ["D" + i].Value));
                string strIssueDateYear = strIssueDate.Substring(strIssueDate.Length - 4, 4);
                effective1 = Convert.ToInt32(strIssueDateYear);
                string strTcode = objHlpr.fn_gettranscodev2(year, effective1, out bool bolEntry);
                dtDataRow [22] = objHlpr.fn_gettranseffectivedate(strIssueDate, Convert.ToString(year)); // Trans Effective Date
                dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Convert.ToString(year));//REINSURANCE START DATE
                dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(strTcode, strIssueDate, valueTransEffectiveDate);
                dtDataRow [21] = strTcode;

                #region getCurrency
                string getCurrency = wsraw.Range ["C" + i].Value;
                decimal dclosa = Convert.ToDecimal(wsraw.Range ["K" + i].Value); // 
                decimal dclinitialsum = Convert.ToDecimal(wsraw.Range ["L" + i].Value);
                decimal dclprem = Convert.ToDecimal(wsraw.Range ["M" + i].Value);
                decimal dclcededRet = 0;

                if(getCurrency.ToUpper().Contains("PHP"))
                {
                    dtDataRow [23] = "PHP";
                    dclinitialsum = dclinitialsum * mul;
                    dclcededRet = dclosa - dclinitialsum;
                    dclprem = dclprem * mul;
                    TotalPremiumPHP += dclprem;
                    TotalSumAtRiskPHP += dclosa;
                }
                else if(getCurrency.ToUpper().Contains("USD"))
                {
                    dtDataRow [23] = "USD";
                    dclinitialsum = dclinitialsum * mul;
                    dclcededRet = dclosa - dclinitialsum;
                    dclprem = dclprem * mul;
                    TotalPremiumUSD += dclprem;
                    TotalSumAtRiskUSD += dclosa;
                }
                #endregion
                dtDataRow [24] = "MLY";
                //objHlpr.fn_CheckingforA_AB_BZColumn(dclorig.ToString(), null, dclsar.ToString(), out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksAABBZ);
                dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclosa)); //Col Z
                dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclinitialsum));
                dtDataRow [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclcededRet));//cedent retention
                dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclinitialsum));
                dtDataRow [29] = "NATREID";
                string strDOB = objHlpr.fn_convertStringtoDateV2(wsraw.Range ["G" + i].Value.ToString());
                string strFullname = wsraw.Range ["F" + i].Value;// Full Name
                string str_outlifeid = string.Empty;
                if(string.IsNullOrEmpty(strFullname))
                {
                    strFullname = "DummyFullName";
                    str_outlifeid = strPolicyNo;
                }
                dtDataRow [31] = strFullname;

                objHlpr.fn_getnamesandlifeID(strFullname, strDOB, out string strFirstname, out string strLastname, out str_outlifeid, "000");
                //objHlpr2.fn_separateLastNameFirstNameV4(strFullname, out strFullname, out string strLastname, out string strFirstname, out string strMiddlename);
                string str_MI = objHlpr.fn_getMI(strFirstname);
                dtDataRow [34] = str_MI;
                dtDataRow [31] = objHlpr.fn_stringcleanup(strFullname);
                dtDataRow [32] = objHlpr2.fn_checkLastname(strLastname);
                dtDataRow [33] = strFirstname.Replace(" " + str_MI, string.Empty);
                if(string.IsNullOrEmpty(str_outlifeid))
                {
                    dtDataRow [30] = strPolicyNo;
                }
                else
                {
                    dtDataRow [30] = str_outlifeid;
                }
               

                //objHlpr2.fn_separateLastNameFirstNameV3(strFullname, out string strLastname, out string strFirstname, out string strMiddlename);
                //dtDataRow [30] = objHlpr.fn_LifeID(strFirstname, strLastname, strDOB);
                //dtDataRow [32] = strLastname;
                //dtDataRow [33] = strFirstname;
                //dtDataRow [34] = strMiddlename;
                string strgender = objHlpr2.fn_MaleOrFemale(wsraw.Range ["E" + i].Value);
                dtDataRow [36] = strgender;//gender
                dtDataRow [37] = strDOB;
                dtDataRow [38] = "NONE";
                dtDataRow [39] = "STANDARD";
                dtDataRow [41] = oDate.Year;

                if(bolEntry == true)
                {
                    dtDataRow [58] = "4001";
                    dtDataRow [59] = dclprem;
                }
                else
                {
                    dtDataRow [56] = "4000";
                    dtDataRow [57] = dclprem;

                }

                dtDataRow [79] = wsraw.Range ["H" + i].Value; //life issue age
                objHlpr.fn_GetRemarksCode(strDOB, strFullname, strgender, out string strRemarksCode);
                dtDataRow [76] = strRemarksCode; // Remarks


            }

            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium PHP:";
            dtDataRow [1] = TotalPremiumPHP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium USD:";
            dtDataRow [1] = TotalPremiumUSD;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk PHP :";
            dtDataRow [1] = TotalSumAtRiskPHP;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk USD :";
            dtDataRow [1] = TotalSumAtRiskUSD;
            objdt_template.Rows.Add(dtDataRow);
            #endregion


            string despath = str_saved + @"\BM004-" + str_sheet + str_savef + ".xlsx";
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
