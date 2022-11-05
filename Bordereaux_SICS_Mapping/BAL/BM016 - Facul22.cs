using System;
using System.Data;
using System.Linq;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Bordereaux_SICS_Mapping.Forms;
using Bordereaux_SICS_Mapping.BAL;
using System.Text.RegularExpressions;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM016_Facul22
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
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            //int intLastRow = wsraw.Range["B12"].End[XlDirection.xlDown].Row;
            int erawrow = rawrange.Rows.Count;
            DataRow dtDataRow;
            string TransEffectiveDate = string.Empty;
            decimal dclTotalPremiumLife = 0;
            decimal dclTotalSumAtRiskLife = 0;
            decimal dclTotalPremiumADB = 0;
            decimal dclTotalSumAtRiskADB = 0;
       
            

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }
            string strTcode = string.Empty;
            string strPremiumYear = string.Empty;


            if(str_sheet.ToUpper() == "FACULTATIVE FIRST YEAR - ADJ" || str_sheet.ToUpper() == "FACULTATIVE REN - ADJ")
            {
                for(int intLoop = 7; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 7].Text.ToString();
                    Regex checkPolicy = new Regex(@"\d");
                    if(!checkPolicy.IsMatch(strPolicyNo))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [1] = wsraw.Cells [intLoop, 1].Value;//cession no
                    dtDataRow [8] = "SURPLUS"; //REINSURANCE PRODUCT
                    dtDataRow [9] = "PAFM"; //TYPE OF BUSINESS
                    dtDataRow [10] = "S"; //REINSURANCE_METHODS
                    dtDataRow [13] = "IND"; //CLASS OF BUSINESS
                    dtDataRow [14] = "F"; //BUSINESS TYPE
                    dtDataRow [23] = "PHP"; //CESSION CURRENCY
                    dtDataRow [24] = "YLY"; //PREMIUM FREQUENCY
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [41] = Variables.strBmYear;//Policy Year
                    string strFirstName = wsraw.Cells [intLoop, 3].Text;
                    string strLastName = wsraw.Cells [intLoop, 2].Text;
                    string strMI = wsraw.Cells [intLoop, 4].Text;
                    dtDataRow [32] = strLastName;
                    dtDataRow [33] = strFirstName;
                    dtDataRow [34] = strMI;
                    dtDataRow [31] = strLastName + " " + strFirstName + " " + strMI; //Full Name
                    dtDataRow [36] = wsraw.Cells [intLoop, 5].Text;
                    string DOB = Convert.ToDateTime(wsraw.Cells [intLoop, 8].Text).ToString("MM/dd/yyyy");//DOB
                    dtDataRow [37] = DOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, DOB); //life ID 
                    dtDataRow [38] = objHlpr.fn_SmokerCode(null);
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(wsraw.Cells [intLoop, 14].Text);//prefered classific
                    dtDataRow [41] = Variables.strBmYear;//Policy Year
                    dtDataRow [79] = wsraw.Cells [intLoop, 18].Text;//issue age

                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 19].Text), out decimal dclOrinalSum);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Text), out decimal dclInitialSum);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Text), out decimal dclCedentReten);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 11].Text), out decimal dclPremLife);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Text), out decimal dclPremRider);

                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 9].Text).ToString("MM/dd/yyyy");
                    string strPeriodCover = wsraw.Cells [intLoop, 22].Text;
                    TransEffectiveDate = objHlpr2.fn_PeriodCover(strPeriodCover, strIssueDate); //Transeffective date
                    #region Transcode Premiums
                    string plancode = objHlpr2.fn_getplanCodeV2(wsraw.Cells [intLoop, 6].Text);
                    string rider = wsraw.Cells [intLoop, 15].Text;
                    if(str_sheet.ToUpper().Contains("FIRST"))
                    {
                        
                        dtDataRow [21] = "ADJUST"; // Transcode
                        dtDataRow [22] = TransEffectiveDate;
                        dtDataRow [20] = strIssueDate;//Policy Start Date
                        dtDataRow [19] = TransEffectiveDate;  // Reinsurance Start Date

                        if(dclPremLife > 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = dclPremLife;
                            dtDataRow [5] = plancode;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                        }
                        if(dclPremRider > 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow02 [60] = "4002";
                            _var.dtworkRow02 [61] = dclPremRider;
                            _var.dtworkRow02 [5] = rider;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }
                    }
                    else
                    {
                        dtDataRow [21] = "ADJUST"; // Transcode
                        dtDataRow [22] = TransEffectiveDate;
                        dtDataRow [20] = strIssueDate;//Policy Start Date
                        dtDataRow [19] = TransEffectiveDate;  // Reinsurance Start Date


                        if(dclPremLife > 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclPremLife;
                            dtDataRow [5] = plancode;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                        }
                        if(dclPremRider > 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow02 [62] = "4004";
                            _var.dtworkRow02 [63] = dclPremRider;
                            _var.dtworkRow02 [5] = rider;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }
                    }
                    #endregion

                    #region HashTotal
                    dclTotalPremiumLife += dclPremLife;
                    dclTotalSumAtRiskLife += dclInitialSum;
                    dclTotalSumAtRiskADB += dclInitialSum;
                    dclTotalPremiumADB += dclPremRider;
                    #endregion

                }
            }
            else if (str_sheet.ToUpper() == "FACULTATIVE RENEWAL")
            {
                for(int intLoop = 7; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 7].Text.ToString();
                    Regex checkPolicy = new Regex(@"\d");
                    if(!checkPolicy.IsMatch(strPolicyNo))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [1] = wsraw.Cells [intLoop, 1].Value;//cession no
                    dtDataRow [8] = "SURPLUS"; //REINSURANCE PRODUCT
                    dtDataRow [9] = "PAFM"; //TYPE OF BUSINESS
                    dtDataRow [10] = "S"; //REINSURANCE_METHODS
                    dtDataRow [13] = "IND"; //CLASS OF BUSINESS
                    dtDataRow [14] = "F"; //BUSINESS TYPE
                    dtDataRow [23] = "PHP"; //CESSION CURRENCY
                    dtDataRow [24] = "YLY"; //PREMIUM FREQUENCY
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [41] = Variables.strBmYear;//Policy Year
                    string strFirstName = wsraw.Cells [intLoop, 3].Text;
                    string strLastName = wsraw.Cells [intLoop, 2].Text;
                    string strMI = wsraw.Cells [intLoop, 4].Text;
                    dtDataRow [32] = strLastName;
                    dtDataRow [33] = strFirstName;
                    dtDataRow [34] = strMI;
                    dtDataRow [31] = strLastName + " " + strFirstName + " " + strMI; //Full Name
                    dtDataRow [36] = wsraw.Cells [intLoop, 5].Text;
                    string DOB = Convert.ToDateTime(wsraw.Cells [intLoop, 8].Text).ToString("MM/dd/yyyy");//DOB
                    dtDataRow [37] = DOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, DOB); //life ID 
                    dtDataRow [38] = objHlpr.fn_SmokerCode(null);
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(wsraw.Cells [intLoop, 14].Text);//prefered classific
                    dtDataRow [41] = Variables.strBmYear;//Policy Year
                    dtDataRow [79] = wsraw.Cells [intLoop, 18].Text;//issue age

                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 19].Text), out decimal dclOrinalSum);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Text), out decimal dclInitialSum);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 20].Text), out decimal dclCedentReten);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 11].Text), out decimal dclPremLife);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Text), out decimal dclPremRider);

                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 9].Text).ToString("MM/dd/yyyy");
                    string strPeriodCover = wsraw.Cells [intLoop, 22].Text;
                    TransEffectiveDate = objHlpr2.fn_PeriodCover(strPeriodCover, strIssueDate); //Transeffective date
                    #region Transcode Premiums
                    string plancode = objHlpr2.fn_getplanCodeV2(wsraw.Cells [intLoop, 6].Text);
                    string rider = wsraw.Cells [intLoop, 15].Text;
                    if(str_sheet.ToUpper() == "FACULTATIVE RENEWAL")
                    {

                        dtDataRow [21] = "TRENEW"; // Transcode
                        dtDataRow [22] = TransEffectiveDate;
                        dtDataRow [20] = strIssueDate;//Policy Start Date
                        dtDataRow [19] = TransEffectiveDate;  // Reinsurance Start Date

                        if(dclPremLife > 0)
                        {
                            dtDataRow [58] = "4001";
                            dtDataRow [59] = dclPremLife;
                            dtDataRow [5] = plancode;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                        }
                        if(dclPremRider > 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            _var.dtworkRow02 [58] = "4001";
                            _var.dtworkRow02 [59] = dclPremRider;
                            _var.dtworkRow02 [5] = rider;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //ISR
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSum)); //SAR
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrinalSum)); //OSA
                            _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCedentReten)); //OSA
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }
                    }                    
                    #endregion

                    #region HashTotal
                    dclTotalPremiumLife += dclPremLife;
                    dclTotalSumAtRiskLife += dclInitialSum;
                    dclTotalSumAtRiskADB += dclInitialSum;
                    dclTotalPremiumADB += dclPremRider;
                    #endregion

                }
            }


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if(str_sheet.ToUpper() == "FACULTATIVE FIRST YEAR - ADJ")
            {
                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium Life:";
                dtDataRow [1] = dclTotalPremiumLife;
                objdt_template.Rows.Add(dtDataRow);

                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Sum at Risk Life:";
                dtDataRow [1] = dclTotalSumAtRiskLife;
                objdt_template.Rows.Add(dtDataRow);

   

                if(dclTotalPremiumADB > 0)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ADB:";
                    dtDataRow [1] = dclTotalPremiumADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ADB:";
                    dtDataRow [1] = dclTotalSumAtRiskADB;
                    objdt_template.Rows.Add(dtDataRow);

                }

            }
            else
            {
                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Premium Life:";
                dtDataRow [1] = dclTotalPremiumLife;
                objdt_template.Rows.Add(dtDataRow);


                dtDataRow = objdt_template.NewRow();
                dtDataRow [0] = "Total Sum at Risk Life:";
                dtDataRow [1] = dclTotalSumAtRiskLife;
                objdt_template.Rows.Add(dtDataRow);

                if (dclTotalPremiumADB > 0)
                {
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Premium ADB:";
                    dtDataRow [1] = dclTotalPremiumADB;
                    objdt_template.Rows.Add(dtDataRow);

                    dtDataRow = objdt_template.NewRow();
                    dtDataRow [0] = "Total Sum at Risk ADB:";
                    dtDataRow [1] = dclTotalSumAtRiskADB;
                    objdt_template.Rows.Add(dtDataRow);

                }

            };
         

            #endregion

            string despath = str_saved + @"\BM016_Facul" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremiumLife = 0;
            dclTotalSumAtRiskLife = 0;
            dclTotalPremiumADB = 0;
            dclTotalSumAtRiskADB = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}