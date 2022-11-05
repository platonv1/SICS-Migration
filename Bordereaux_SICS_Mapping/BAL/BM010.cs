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
    class BM010
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

            string strRemarksAABBZ = string.Empty;
            string strBrandedProduct = string.Empty;
            decimal dclTotalPremium = 0;
            decimal dclTotalSAR = 0;


            DataRow dtDataRow;

            if (str_sheet.ToUpper().Contains("LIFE"))
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 4].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                   
                    string strPremDueDate = wsraw.Cells[intLoop, 7].Text.ToString();
                    string strbmyear = strPremDueDate.Substring(strPremDueDate.Length - 4, 4);
                    string strURC = Convert.ToString(wsraw.Cells[intLoop, 6].Value);
                    string strCessionNo = Convert.ToString(wsraw.Cells[intLoop, 5].Value);
                    dtDataRow[1] = strCessionNo;
                    string strAge = Convert.ToString(wsraw.Cells[intLoop, 8].Value);
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[21] = "TRENEW"; // Transcode
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[41] = strbmyear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    objHlpr.fn_macrobenlifebm010(strCessionNo, strURC, strPolicyNo, strPremDueDate,
                    out string strCessionTypeCode, out strPolicyNo,
                    out string strFullname, out string strLastName, out string strFirstName, out string strMiddleInitial, out string strSex,
                    out string strDOB, out string strLifeID, out string strIssueDate, out string strMortality,
                    out string strRefunding, out string strCededRetention, out string strOSA, out string strISR, out string strDummyRemarksCode);
                    dtDataRow[0] = strPolicyNo;
                    dtDataRow[14] = objHlpr.fn_checkBusinessTypeV1(strCessionTypeCode); // Business Type
                    dtDataRow[20] = strIssueDate; // Policy Start Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strPremDueDate, strbmyear);//REINSURANCE START DATE
                    dtDataRow[22] = objHlpr.fn_gettranseffectivedate(strPremDueDate, strbmyear);
                    dtDataRow[31] = strFullname; //Full Name
                    Console.WriteLine(strPolicyNo);
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow[34] = strMiddleInitial;
                    dtDataRow[37] = strDOB; // Birthday
                    dtDataRow[36] = strSex;
                    dtDataRow[30] = strLifeID;// life ID 
                    dtDataRow[58] = "4001";// Entry Code
                    dtDataRow[39] = objHlpr2.fn_getmortalityrating(strMortality);//Mortality Rating
                    string strIssueAge = Convert.ToString(wsraw.Cells[intLoop, 8].Value); //Issue Age
                    dtDataRow[79] = strIssueAge;
                    string strSAR = Convert.ToString(wsraw.Cells[intLoop, 9].Value);
                    objHlpr.fn_CheckingforA_AB_BZColumn(strOSA, strISR, strSAR, out strOSA, out strISR, out strSAR, out string strRemakrsABABZ);
                    objHlpr.fn_GetRemarksCode(strDOB, strFullname, strSex, out string strRemarksCode);
                    dtDataRow[76] = strRemarksCode + "|" + strRemakrsABABZ + "|"+ strDummyRemarksCode;

                    decimal dclPremLife = Convert.ToDecimal(wsraw.Cells[intLoop, 10].Value);
                    decimal dclPremExtra = Convert.ToDecimal(wsraw.Cells[intLoop, 11].Value);
                    decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);
                    decimal dclPremADB = Convert.ToDecimal(wsraw.Cells[intLoop, 13].Value);
                    decimal dclPremium = Convert.ToDecimal(wsraw.Cells[intLoop, 15].Value);
                    decimal dclPremSARDI = Convert.ToDecimal(wsraw.Cells[intLoop, 14].Value);
                    decimal dclSAR = Convert.ToDecimal(strSAR);

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow05 = objdt_template.NewRow();
                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;


                    #region Premium
                    if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);


                        _var.dtworkRow04[5] = "ADB";
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[59] = dclPremADB; //Premium
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);


                        _var.dtworkRow04[5] = "ADB";
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[59] = dclPremADB; //Premium
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow04);


                        _var.dtworkRow05[5] = "SARDI";
                        _var.dtworkRow05[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow05[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow05[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow05[59] = dclPremADB; //Premium
                        _var.dtworkRow05[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow05);
                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "SARDI";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremSARDI; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "ADB";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremADB; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "SARDI";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremSARDI; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "WPD";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremWPD; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "WPD";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremWPD; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "SARDI";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremSARDI; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "ADB";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremADB; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "SARDI";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremSARDI; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[59] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03[5] = "ADB";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[59] = dclPremADB; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//cedent retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "ADB";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremADB; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(strOSA); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(strISR);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(strCededRetention);//cedent retention
                    }
                    #endregion


                    dclTotalPremium += dclPremLife + dclPremExtra + dclPremWPD + dclPremADB + dclPremSARDI;
                    dclTotalSAR += dclSAR;


                }
            }
            else
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 3].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                    {
                        continue;
                    }

                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);

                    string strPremDueDate = wsraw.Cells[intLoop, 55].Text.ToString();
                    string strbmyear = strPremDueDate.Substring(strPremDueDate.Length - 4, 4);
                    string strURC = Convert.ToString(wsraw.Cells[intLoop, 6].Value);
                    string strCessionNo = Convert.ToString(wsraw.Cells[intLoop, 2].Value);
                    dtDataRow[1] = strCessionNo;
                    string strAge = Convert.ToString(wsraw.Cells[intLoop, 13].Value);
                    dtDataRow[23] = "PHP"; //  Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    string strCessionTypeCode = "TNEWBUS";
                    dtDataRow[21] = strCessionTypeCode; // Transcode
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    dtDataRow[41] = strbmyear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    
                    dtDataRow[0] = strPolicyNo;
                    dtDataRow[14] = objHlpr.fn_checkBusinessTypeV1(strCessionTypeCode); // Business Type
                    dtDataRow[20] = strPremDueDate; // Policy Start Date
                    dtDataRow[19] = objHlpr.fn_gettranseffectivedate(strPremDueDate, strbmyear);//REINSURANCE START DATE
                    dtDataRow[22] = objHlpr.fn_gettranseffectivedate(strPremDueDate, strbmyear);
                    string strFullname = Convert.ToString(wsraw.Cells[intLoop, 4].Value);
                    dtDataRow[31] = strFullname; //Full Name
                    objHlpr2.fn_separateLastNameFirstNameV2(strFullname, out strFullname, out string strLastName, out string strFirstName, out string strMiddleInitial);
                    dtDataRow[33] = strFirstName;
                    dtDataRow[32] = objHlpr2.fn_removeCharacters(strLastName);
                    dtDataRow[34] = strMiddleInitial;
                    string strDOB = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Cells[intLoop, 10].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthday
                    string strSex = Convert.ToString(wsraw.Cells[intLoop, 9].Value);//sex
                    dtDataRow[36] = strSex;
                    dtDataRow[30] = objHlpr.fn_LifeID(strFirstName,strLastName,strDOB);// life ID 
                    dtDataRow[39] = Convert.ToString(wsraw.Cells[intLoop, 42].Value);//Mortality Rating
                    string strIssueAge = Convert.ToString(wsraw.Cells[intLoop, 13].Value); //Issue Age
                    dtDataRow[79] = strIssueAge;
                    dtDataRow[56] = "4000";

                    objHlpr.fn_GetRemarksCode(strDOB, strFullname, strSex, out string strRemarksCode);
                  
                    objHlpr.fn_CheckingforA_AB_BZColumn(Convert.ToString(wsraw.Cells[intLoop, 15].Value), null, Convert.ToString(wsraw.Cells[intLoop, 64].Value),out string OSA, out string ISA, out string SAR,out string strRemarksCodeAABBBZ);
                    dtDataRow[76] = strRemarksCode + "|" + strRemarksCodeAABBBZ;

                    decimal dclSAR = Convert.ToDecimal(SAR); //NAR
                    decimal dclOSA = Convert.ToDecimal(OSA); //ADB SAR
                    decimal dclSarWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 46].Value);
                    decimal dclSarADB = Convert.ToDecimal(wsraw.Cells[intLoop, 47].Value);
                    decimal dclSARDI = Convert.ToDecimal(wsraw.Cells[intLoop, 48].Value);

                    decimal dclPremLife = Convert.ToDecimal(wsraw.Cells[intLoop, 58].Value);
                    decimal dclPremExtra = Convert.ToDecimal(wsraw.Cells[intLoop, 59].Value);
                    decimal dclPremWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 60].Value);
                    decimal dclPremADB = Convert.ToDecimal(wsraw.Cells[intLoop, 61].Value);
                    decimal dclPremSARDI = Convert.ToDecimal(wsraw.Cells[intLoop, 62].Value);

                    decimal dclCRetentionLife = Convert.ToDecimal(wsraw.Cells[intLoop, 49].Value);
                    decimal dclCRetentionWPD = Convert.ToDecimal(wsraw.Cells[intLoop, 50].Value);
                    decimal dclCRetentionADB = Convert.ToDecimal(wsraw.Cells[intLoop, 51].Value);
                    decimal dclCRetentionSARDI = Convert.ToDecimal(wsraw.Cells[intLoop, 52].Value);

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow05 = objdt_template.NewRow();
                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;

                    if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremExtra; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[57] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[57] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);


                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow03[57] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[57] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow03[57] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);


                        _var.dtworkRow04[5] = "ADB";
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow04[57] = dclPremADB; //Premium
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }
                    else if (dclPremLife != 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[57] = dclPremExtra; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "WPD";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow03[57] = dclPremWPD; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);

                        _var.dtworkRow04[5] = "ADB";
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow04[57] = dclPremADB; //Premium
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow04);

                        _var.dtworkRow05[5] = "SARDI";
                        _var.dtworkRow05[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow05[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow05[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARDI));
                        _var.dtworkRow05[57] = dclPremSARDI; //Premium
                        _var.dtworkRow05[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow05);
                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow02[57] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {

                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow02[57] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                        _var.dtworkRow02[5] = "SARDI";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARDI));
                        _var.dtworkRow02[57] = dclPremSARDI; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow02[57] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "ADB";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow03[57] = dclPremADB; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if (dclPremLife == 0 && dclPremExtra != 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "EXTRA";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                        dtDataRow[57] = dclPremExtra; //Premium

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow02[57] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "SARDI";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARDI));
                        _var.dtworkRow03[57] = dclPremSARDI; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);

                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "WPD";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));//SAR
                        dtDataRow[57] = dclPremWPD; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow02[57] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "WPD";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));//SAR
                        dtDataRow[57] = dclPremWPD; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow02[57] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "SARDI";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSARDI));
                        _var.dtworkRow03[57] = dclPremSARDI; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI != 0)
                    {
                        dtDataRow[5] = "ADB";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));//SAR
                        dtDataRow[57] = dclPremADB; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention

                        _var.dtworkRow02[5] = "SARDI";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));
                        _var.dtworkRow02[57] = dclPremSARDI; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionSARDI));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "ADB";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow02[57] = dclPremADB; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[59] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention


                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow02[57] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dclPremLife != 0 && dclPremExtra == 0 && dclPremWPD != 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention

                        _var.dtworkRow02[5] = "WPD";
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarWPD));
                        _var.dtworkRow02[57] = dclPremWPD; //Premium
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionWPD));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow02);

                        _var.dtworkRow03[5] = "ADB";
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));
                        _var.dtworkRow03[57] = dclPremADB; //Premium
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB != 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "ADB";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSarADB));//SAR
                        dtDataRow[57] = dclPremADB; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionADB));//Cedent Retention
                    }
                    else if (dclPremLife == 0 && dclPremExtra == 0 && dclPremWPD == 0 && dclPremADB == 0 && dclPremSARDI == 0)
                    {
                        dtDataRow[5] = "LIFE";
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclOSA)); //OSA
                        dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(ISA);//ISA
                        dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSAR));//SAR
                        dtDataRow[57] = dclPremLife; //Premium
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclCRetentionLife));//Cedent Retention
                    }

                    dclTotalPremium += dclPremLife + dclPremExtra + dclPremWPD + dclPremADB + dclPremSARDI;
                    dclTotalSAR += dclSAR + dclSarADB + dclSarWPD + dclSARDI;


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
            dtDataRow[1] = dclTotalSAR;
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

            string despath = str_saved + @"\BM010-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotalSAR = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}