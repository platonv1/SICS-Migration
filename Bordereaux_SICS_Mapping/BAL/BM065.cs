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
    class BM065
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

            //int intLastRow = wsraw.Range["B12"].End[XlDirection.xlDown].Row;
            int erawrow = rawrange.Rows.Count;

            string strFilePath = wbraw.Path;
            string strRemarksAABBZ = string.Empty;
            string valueTransEffectiveDate = string.Empty;
            decimal dclTotalPremium = 0; decimal dclTotalSumAtRisk = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
                
            }
        
            DataRow dtDataRow;

            if (str_sheet.ToUpper().Contains("JAN") || str_sheet.ToUpper().Contains("FEB") || str_sheet.ToUpper().Contains("MARC") || str_sheet.ToUpper().Contains("MAY") || str_sheet.ToUpper().Contains("APR") || str_sheet.ToUpper().Contains("JUNE") || str_sheet.ToUpper().Contains("JULY") ||
            str_sheet.ToUpper().Contains("AUGUST") || str_sheet.ToUpper().Contains("SEPTE") || str_sheet.ToUpper().Contains("OCTO") || str_sheet.ToUpper().Contains("NOV") || str_sheet.ToUpper().Contains("DECEM"))
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 19].Text.ToString();
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 20].Text.ToString(), wsraw.Cells [intLoop, 21].Text.ToString(), wsraw.Cells [intLoop, 22].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [23] = "PHP"; //  Cession Currency
                    dtDataRow [24] = "MLY"; // Premium Frequency
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    string strTcode = "TRENEW";
                    dtDataRow [21] = strTcode; // Transcode
                    dtDataRow [8] = "QA"; // Reinsurance Product
                    dtDataRow [9] = "PA"; // Type of Business
                    dtDataRow [10] = "Q"; // Reinsurance Methods
                    dtDataRow [13] = "GEB"; // Class of Business
                    dtDataRow [14] = "T"; // Business Type
                    dtDataRow [29] = "NATREID"; // Life ID Type

                    string strIssueDate = objHlpr.fn_convertStringtoDateV2(Convert.ToString(wsraw.Cells [intLoop, 3].Value));
                    dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(strTcode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    dtDataRow [19] = objHlpr.fn_gettranseffectivedate(strIssueDate, Variables.strBmYear);//REINSURANCE START DATE
                    dtDataRow [20] = valueTransEffectiveDate;//Policy Start Date
                    string strFullName = Convert.ToString(wsraw.Cells [intLoop, 4].Value); //Full Name
                    objHlpr2.fn_separateLastNameFirstNameV8(strFullName, out string strLastName, out string strFirstName, out string strMiddleName);
                    dtDataRow [31] = objHlpr2.fn_checkFullname(strFullName);
                    dtDataRow [32] = strLastName; // Last Name
                    dtDataRow [33] = strFirstName; // First Name
                    dtDataRow [34] = strMiddleName; // Middle Initials
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(null); //preffered classific
                    string strDOB = "07/01/1900";
                    dtDataRow [37] = strDOB; // Birthday
                    dtDataRow [76] = "BR4"; // Remarks
                    string strSex = Convert.ToString(wsraw.Cells [intLoop, 6].Value); // Gender
                    dtDataRow [36] = strSex;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);// life ID 
                    dtDataRow [58] = "4001";
                    //objHlpr.fn_GetRemarksCode(strDOB, strFirstName, strPolicyNo, strSex, out strRemarksCode);
                    //dtDataRow[78] = objHlpr.fn_getAttainAge(strDOB, Variables.strBmYear);//LIFE ATTAIN AGE AGE;
                    dtDataRow [79] = Convert.ToString(wsraw.Cells [intLoop, 5].Value); //Issue Age

                    #region Sum At Risk
                    decimal dclSumNar_Add = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 34].Value), 0.10M);
                    decimal dclSumNar_Basic = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 35].Value), 0.10M);
                    decimal dclSumNar_Rider = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 36].Value), 0.10M);
                    decimal dclSumNar_Life = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 37].Value), 0.10M);
                    decimal dclSumNar_Tpdi = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 38].Value), 0.10M);
                    decimal dclSumNar_Uma = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 39].Value), 0.10M);
                    #endregion

                    #region Orig Sum
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 7].Value), out decimal dclOrigSum_Add);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 9].Value), out decimal dclOrigSum_Basic);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 11].Value), out decimal dclOrigSum_Rider);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out decimal dclOrigSum_Life);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 15].Value), out decimal dclOrigSum_Tpdi);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 17].Value), out decimal dclOrigSum_Uma);
                    #endregion


                    #region Premium
                    decimal dclPrem_Add = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 41].Value), 0.10M);
                    decimal dclPrem_Basic = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 42].Value), 0.10M);
                    decimal dclPrem_Rider = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 43].Value), 0.10M);
                    decimal dclPrem_Life = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 44].Value), 0.10M);
                    decimal dclPrem_Tpdi = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 45].Value), 0.10M);
                    decimal dclPrem_Uma = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells [intLoop, 46].Value), 0.10M);
                    #endregion

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow05 = objdt_template.NewRow();
                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow06 = objdt_template.NewRow();
                    _var.dtworkRow06.ItemArray = dtDataRow.ItemArray;


                    #region original, sum reinsured, sum at risk Col 25/ 27 / 77
                    if(dclSumNar_Add != 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;
                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic != 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;

                        _var.dtworkRow02 [5] = "AD&D (PA BASIC)";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [59] = dclPrem_Basic;
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [59] = dclPrem_Add;


                        _var.dtworkRow02 [5] = "AD&D (PA BASIC)";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [59] = dclPrem_Basic;
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03 [5] = "AD&D (PA RIDER)";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow03 [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [59] = dclPrem_Rider;
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;


                        _var.dtworkRow02 [5] = "AD&D (PA BASIC)";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [59] = dclPrem_Basic;
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03 [5] = "AD&D (PA RIDER)";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [59] = dclPrem_Rider;
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04 [5] = "GYRT";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [59] = dclPrem_Life;
                        _var.dtworkRow04 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        objdt_template.Rows.Add(_var.dtworkRow04);

                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;


                        _var.dtworkRow02 [5] = "AD&D (PA BASIC)";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [59] = dclPrem_Basic;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "AD&D (PA RIDER)";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow03 [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [59] = dclPrem_Rider;
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04 [5] = "GYRT";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [59] = dclPrem_Life;
                        _var.dtworkRow04 [26] = dclOrigSum_Life;
                        objdt_template.Rows.Add(_var.dtworkRow04);

                        _var.dtworkRow05 [5] = "TPDI";
                        _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow05 [59] = dclPrem_Tpdi;
                        _var.dtworkRow05 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        objdt_template.Rows.Add(_var.dtworkRow05);

                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));

                        _var.dtworkRow02 [5] = "AD&D (PA BASIC)";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        _var.dtworkRow02 [59] = dclPrem_Basic;
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03 [5] = "AD&D (PA RIDER)";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow03 [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow03 [59] = dclPrem_Rider;
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        objdt_template.Rows.Add(_var.dtworkRow03);


                        _var.dtworkRow04 [5] = "GYRT";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow04 [59] = dclPrem_Life;
                        _var.dtworkRow04 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        objdt_template.Rows.Add(_var.dtworkRow04);


                        _var.dtworkRow05 [5] = "TPDI";
                        _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow05 [59] = dclPrem_Tpdi;
                        _var.dtworkRow05 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        objdt_template.Rows.Add(_var.dtworkRow05);

                        _var.dtworkRow06 [5] = "UMA";
                        _var.dtworkRow06 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow06 [28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow06 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow06 [59] = dclPrem_Uma;
                        _var.dtworkRow06 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        objdt_template.Rows.Add(_var.dtworkRow06);

                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "AD&D";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));


                        _var.dtworkRow02 [5] = "UMA";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow02 [59] = dclPrem_Uma;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }

                    else if(dclSumNar_Add == 0 && dclSumNar_Basic != 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D (PA BASIC)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Basic;

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D (PA BASIC)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Basic;

                        _var.dtworkRow02 [5] = "AD&D (PA RIDER))";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [59] = dclPrem_Rider;
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D (PA BASIC)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Basic;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));

                        _var.dtworkRow02 [5] = "AD&D (PA RIDER))";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [59] = dclPrem_Rider;
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "GYRT";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [59] = dclPrem_Life;
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }

                    else if(dclSumNar_Add == 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D (PA BASIC)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Basic;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));

                        _var.dtworkRow02 [5] = "AD&D (PA RIDER))";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [59] = dclPrem_Rider;
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "GYRT";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [59] = dclPrem_Life;
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04 [5] = "TPDI";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow04 [59] = dclPrem_Tpdi;
                        _var.dtworkRow04 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic != 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "AD&D (PA BASIC)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Basic));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Basic;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Basic));

                        _var.dtworkRow02 [5] = "AD&D (PA RIDER))";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        _var.dtworkRow02 [59] = dclPrem_Rider;
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "GYRT";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow03 [59] = dclPrem_Life;
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04 [5] = "TPDI";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow04 [59] = dclPrem_Tpdi;
                        _var.dtworkRow04 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        objdt_template.Rows.Add(_var.dtworkRow04);

                        _var.dtworkRow05 [5] = "UMA";
                        _var.dtworkRow05 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow05 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow05 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow05 [59] = dclPrem_Uma;
                        _var.dtworkRow05 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        objdt_template.Rows.Add(_var.dtworkRow05);
                    }

                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider != 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D (PA RIDER)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Rider;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D (PA RIDER)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Rider;
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));

                        _var.dtworkRow02 [5] = "GYRT";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [59] = dclPrem_Life;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "AD&D (PA RIDER)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Rider;


                        _var.dtworkRow02 [5] = "GYRT";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [59] = dclPrem_Life;
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03 [5] = "TPDI";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow03 [59] = dclPrem_Tpdi;
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider != 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma != 0)
                    {
                        dtDataRow [5] = "AD&D (PA RIDER)";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Rider));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Rider));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Rider;

                        _var.dtworkRow02 [5] = "GYRT";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [59] = dclPrem_Life;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "TPDI";
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow03 [59] = dclPrem_Tpdi;
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04 [5] = "UMA";
                        _var.dtworkRow04 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow04 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow04 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow04 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow04 [59] = dclPrem_Uma;
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }

                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "GYRT";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Life;

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "GYRT";
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Life;

                        _var.dtworkRow02 [5] = "TPDI";
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow02 [59] = dclPrem_Tpdi;
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "GYRT";
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Life;


                        _var.dtworkRow02 [5] = "TPDI";
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        _var.dtworkRow02 [59] = dclPrem_Tpdi;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03 [5] = "UMA";
                        _var.dtworkRow03 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow03 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow03 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow03 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow03 [59] = dclPrem_Uma;
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if(dclSumNar_Add != 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life != 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {

                        dtDataRow [5] = "AD&D";
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Add));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Add));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Add;


                        _var.dtworkRow02 [5] = "GYRT";
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Life));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Life));
                        _var.dtworkRow02 [59] = dclPrem_Life;
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "TPDI";
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Tpdi;

                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi != 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "TPDI";
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Tpdi));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Tpdi));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Tpdi;


                        _var.dtworkRow02 [5] = "UMA";
                        _var.dtworkRow02 [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        _var.dtworkRow02 [59] = dclPrem_Uma;
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }

                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma != 0)
                    {

                        dtDataRow [5] = "UMA";
                        dtDataRow [26] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOrigSum_Uma));
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumNar_Uma));//SUM AT RISK
                        dtDataRow [59] = dclPrem_Uma;


                    }
                    else if(dclSumNar_Add == 0 && dclSumNar_Basic == 0 && dclSumNar_Rider == 0 && dclSumNar_Life == 0 && dclSumNar_Tpdi == 0 && dclSumNar_Uma == 0)
                    {
                        dtDataRow [5] = "GYRT";
                        dtDataRow [26] = 1;
                        dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        dtDataRow [59] = dclPrem_Life;
                    }
                    #endregion
                    dclTotalPremium += dclPrem_Add + dclPrem_Basic + dclPrem_Rider + dclPrem_Life + dclPrem_Tpdi + dclPrem_Uma;

                    #region Sum At Risk
                    if(dclPrem_Add != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Add;
                    }
                    if (dclPrem_Basic != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Basic;
                    }
                    if (dclPrem_Rider != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Rider;
                    }
                    if(dclPrem_Life != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Life;
                    }
                    if(dclPrem_Tpdi != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Tpdi;
                    }
                    if(dclPrem_Uma != 0)
                    {
                        dclTotalSumAtRisk += dclSumNar_Uma;
                    }
                    #endregion

                    //dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(totalSAR));
                    //dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(totalSAR));
                    //dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(totalSAR));//SUM AT RISK
                }
            }


            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Premium:";
            dtDataRow[1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow[0] = "Total Sum at Risk:";
            dtDataRow[1] = dclTotalSumAtRisk;
            objdt_template.Rows.Add(dtDataRow);
            #endregion

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if (Variables.boogenderfail)
            {
                //objdt_template.Rows.Add(dtDataRow);
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Please check for blank genders";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            if (Variables.boomacrofail)
            {
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "MACRO FAIL: Please check those LifeID tagged as policy no, these are data's that has no record in the macro database";
                objdt_template.Rows.Add(_var.dtworkRow01);
            }

            string despath = str_saved + @"\BM065-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSumAtRisk = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}