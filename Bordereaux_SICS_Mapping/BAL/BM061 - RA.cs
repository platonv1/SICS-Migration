using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;
using Bordereaux_SICS_Mapping.BAL;
namespace Bordereaux_SICS_Mapping.BAL
{
    class BM061_RA
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false, string str_policyYear = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            System.Data.DataTable objdt_template = new System.Data.DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);


            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets [str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            string strFilePath = wbraw.Path;
            string Filename = wbraw.Name.ToUpper().Trim();
            int erawrow = rawrange.Rows.Count;



            decimal dclTotalCommission = 0;
            decimal dclTotalPremium = 0;
            decimal dclTotalSumAtRisk = 0;
            decimal dclPremium = 0;
            decimal dclComission = 0;
            string valueTransEffectiveDate = string.Empty;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }

            DataRow dtDataRow;
            if(str_sheet.ToUpper().Contains("SHEET"))
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 9].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 10].Text.ToString(), wsraw.Cells [intLoop, 11].Text.ToString(), wsraw.Cells [intLoop, 12].Text.ToString()))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [23] = wsraw.Cells [intLoop, 24].Value.ToUpper();//Currency
                    objHlpr2.fn_getcobver2(Filename, out string COB, out string TransCode, out bool bolNegative);
                    dtDataRow [22] = Convert.ToDateTime(wsraw.Cells [intLoop, 49].Value).ToString("MM/dd/yyyy"); // Trans Effective Date
                    dtDataRow [20] = Convert.ToDateTime(wsraw.Cells [intLoop, 20].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    dtDataRow [19] = Convert.ToDateTime(wsraw.Cells [intLoop, 49].Value).ToString("MM/dd/yyyy");
                    dtDataRow [41] = Variables.strBmYear;//Policy Year
                   
                    dtDataRow [21] = TransCode; // Transcode
                    dtDataRow [13] = COB; // Class of Business
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating("");
                    string strRisk = (Convert.ToString(wsraw.Cells [intLoop, 30].Value));
                    dtDataRow [4] = objHlpr2.fn_gettransactionproductV2(Convert.ToString(wsraw.Cells [intLoop, 30].Value), COB);
                    dtDataRow [14] = "T";//Business Type

                    decimal.TryParse(wsraw.Cells [intLoop, 53].Text, out decimal dclPremColBA); //COL BA
                    decimal.TryParse(wsraw.Cells [intLoop, 54].Text, out decimal dclPremColBB); //COL BB
                    decimal.TryParse(wsraw.Cells [intLoop, 55].Text, out decimal dclCOMColBC); //COL BC
                    decimal.TryParse(wsraw.Cells [intLoop, 56].Text, out decimal dclCOMColBD); //COL BD
                    decimal.TryParse(wsraw.Cells [intLoop, 57].Text, out decimal dclCOMColBE); //COL BE

                    if (bolNegative == true)
                    {
                        dclPremColBA = dclPremColBA * -1;
                        dclPremColBB = dclPremColBB * -1;
                        dclPremium = dclPremColBA + dclPremColBB;
                        dclCOMColBC = dclCOMColBC * -1;
                        dclCOMColBD = dclCOMColBD * -1;
                        dclCOMColBE = dclCOMColBE * -1;
                        dclComission = dclCOMColBC + dclCOMColBD + dclCOMColBE;

                        dtDataRow [62] = "4004"; //Entry code
                        dtDataRow [63] = dclPremium;
                        dtDataRow [66] = "5005"; //Entry code
                        dtDataRow [67] = dclComission;
                    }
                    else
                    {
                        dclPremium = dclPremColBA + dclPremColBB;
                        dclComission = dclCOMColBC + dclCOMColBD + dclCOMColBE;
                        dtDataRow [58] = "4001"; //Entry code
                        dtDataRow [59] = dclPremium;
                        dtDataRow [66] = "5005"; //Entry code
                        dtDataRow [67] = dclComission;
                    }
                    dclTotalPremium += dclPremColBA + dclPremColBB;
                    dclTotalCommission += dclCOMColBC + dclCOMColBD + dclCOMColBE;
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty("");//InitialSum
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty("");// sum at risk
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(wsraw.Cells [intLoop, 25].Text); //orignal sum risk
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [38] = objHlpr.fn_SmokerCode(wsraw.Cells [intLoop, 18].Text);
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [36] = wsraw.Cells [intLoop, 15].Value;//Gender
                    string strFullName = wsraw.Cells [intLoop, 10].Value;
                    dtDataRow [31] = strFullName;
                    objHlpr2.fn_seperateforeignamesV2(strFullName, out string strFirstName, out string strLastName, out string strMI);
                    strLastName = objHlpr2.fn_checkLastname(strLastName);
                    dtDataRow [32] = strLastName;
                    strFirstName = objHlpr2.fn_checkFirstname(strFirstName);
                    dtDataRow [33] = strFirstName;
                    dtDataRow [34] = strMI;
                    dtDataRow [79] = Convert.ToString(wsraw.Cells [intLoop, 17].Value);//LIFE ISSUE AGE
                    string strBirthday = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Cells [intLoop, 16].Value));
                    dtDataRow [37] = strBirthday; // Birthday
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strBirthday); //Life ID
                    objHlpr.fn_GetRemarksCode(strBirthday, strFullName, wsraw.Cells [intLoop, 15].Value, out string strRemarksCode);
                    dtDataRow [76] =  strRemarksCode; // Remarks
                }
            }


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Premium:";
            dtDataRow [1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dclTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Commission:";
            dtDataRow [1] = dclTotalCommission;
            objdt_template.Rows.Add(dtDataRow);
            #endregion


            string despath = str_saved + @"\BM061 - RA" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSumAtRisk = 0;
            dclTotalCommission = 0;

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";




        }
    }

}