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
    class BM033
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
            string strSmoker = string.Empty;
            string strCheckSheetName = string.Empty;
            string strOriginalSum = string.Empty;
            string strInitialSum = string.Empty;
            string strRemarksAABBZ = string.Empty;
            string strSumAtRisk = string.Empty;
            string strBrandedProduct = string.Empty;
            string dbTableName = string.Empty;
            decimal dclTotalSumAtRisk = 0;
            decimal dclTotalPremium = 0;

            while (string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();
            }

            DataRow dtDataRow;



            //if (str_sheet == "URC1Q2021" || str_sheet == "URC1Q2021_R&A_RECAP" || str_sheet == "URC$1Q2021" || str_sheet == "NRe1Q2021" || str_sheet == "NRe$1Q2021" || str_sheet == "NRe1Q2021_R&A_RECAP" ||
            //str_sheet == "URC4Q2020" || str_sheet == "URC4Q2020_R&A_RECAP" || str_sheet == "URC$4Q2020" || str_sheet == "NRe4Q2020" || str_sheet == "NRe4Q2020_R&A_RECAP" ||
            //str_sheet == "NRe$4Q2020") 
            if (str_sheet.ToUpper().Contains("URC") || str_sheet.ToUpper().Contains("R&A_RECAP") || str_sheet.ToUpper().Contains("NRE"))
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells[intLoop, 1].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }
                    _var.dtworkRow01 = _var.objdt_template01.NewRow();
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow[0] = strPolicyNo;
                    string strCertNo = wsraw.Cells[intLoop, 2].Value;
                    dtDataRow[1] = strCertNo; //Cert No
                 
                    dtDataRow[23] = objHlpr2.fn_getcurrencyV2(str_sheet); //  Cession Currency
                    dtDataRow[24] = "YLY"; // Premium Frequency
                    dtDataRow[41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    dtDataRow[21] = objHlpr.fn_gettranscode(Convert.ToString(wsraw.Cells[intLoop, 12].Value),str_sheet); // Transcode
                    dtDataRow[8] = "SURPLUS"; // Reinsurance Product
                    dtDataRow[9] = "PAFM"; // Type of Business
                    dtDataRow[10] = "S"; // Reinsurance Methods
                    dtDataRow[13] = "IND"; // Class of Business
                    dtDataRow[29] = "NATREID"; // Life ID Type
                    string strPolicyStartDate = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Cells[intLoop, 3].Value)).ToString("MM/dd/yyyy");
                    dtDataRow[20] = strPolicyStartDate; // Policy Start Date
                    dtDataRow[19] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Cells[intLoop, 4].Value)).ToString("MM/dd/yyyy");//REINSURANCE START DATE
                    dtDataRow[22] = objHlpr.fn_reformatDate(Convert.ToString(wsraw.Cells[intLoop, 4].Value)).ToString("MM/dd/yyyy"); //TRANS EFFECTIVE DATE
                    
                    objHlpr.fn_getbusinessTypeRefundingCode(Convert.ToString(wsraw.Cells[intLoop, 2].Value), out string strBusinessType, out string strRefundingCode);
                    dtDataRow[14] = strBusinessType;
                    dtDataRow[83] = strRefundingCode;

                    dbTableName = objHlpr2.fn_checkDatabaseTable(str_sheet);
                    objHlpr2.fn_getmacro_prembord_umre(dbTableName, strPolicyNo, strCertNo, out strPolicyNo,
                    out string Volume, out string ADB_Volume, out string SAR_Volume, out string SDI_Volume,
                    out string life_ret, out string rid_ret, out string strADBamt, out string strSARamt, out string strSARDIamt,
                    out string FirstName, out string strMiddileInitial, out string strLastName, out string strFullName, out string strFirstName, out string strSex,
                    out string strDOB, out string strMort, out string strAttainAge, out string Issueage, out string strLifeID, out string strRemarksCode);
                   
                    dtDataRow[31] = strFullName; //Full Name
                    dtDataRow[33] = FirstName;
                    dtDataRow[32] = strLastName;
                    dtDataRow[34] = strMiddileInitial;
                    strDOB = objHlpr.fn_reformatDate(Convert.ToString(strDOB)).ToString("MM/dd/yyyy");
                    dtDataRow[37] = strDOB; // Birthday
                    dtDataRow[36] = strSex; // Gender
                    dtDataRow[30] = strLifeID;// life ID 
                    dtDataRow[39] = strMort;
                   
                    dtDataRow[78] = strAttainAge; // LIFE1_ATTAINED_AGE
                    dtDataRow[79] = Issueage;
                  
                    decimal dblISA = Convert.ToDecimal(wsraw.Cells[intLoop, 11].Value);//Initial Sum Assured

                    decimal dblLIFE = (Convert.ToDecimal(wsraw.Cells[intLoop, 5].Value));
                    decimal dblEXTRA = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells[intLoop, 6].Value),0.90M);
                    decimal dblADB = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells[intLoop, 7].Value), 0.90M);
                    decimal dblSAR = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells[intLoop, 8].Value),0.90M);
                    decimal dblSARDI = objHlpr.fn_multiplier(Convert.ToDecimal(wsraw.Cells[intLoop, 9].Value), 0.90M);

                    decimal dblVolume = Convert.ToDecimal(Volume); // coluumn 25
                    decimal dblADBVolume = Convert.ToDecimal(ADB_Volume);
                    decimal dblSARVolume = Convert.ToDecimal(SAR_Volume);
                    decimal dblSDIVolume = Convert.ToDecimal(SDI_Volume);

                    decimal dblLifeRet =  Convert.ToDecimal(life_ret); //column 28
                    decimal dblRidRet = Convert.ToDecimal(rid_ret);
                    decimal dblADBamt = Convert.ToDecimal(strADBamt);
                    decimal dblSARamt = Convert.ToDecimal(strSARamt);
                    decimal dblSARDIamt = Convert.ToDecimal(strSARDIamt);


                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow03 = objdt_template.NewRow();
                    _var.dtworkRow03.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow04 = objdt_template.NewRow();
                    _var.dtworkRow04.ItemArray = dtDataRow.ItemArray;
                    _var.dtworkRow05 = objdt_template.NewRow();
                    _var.dtworkRow05.ItemArray = dtDataRow.ItemArray;

                    #region Premium
                    if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; // Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//sar
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));//isa
                            dtDataRow[58] = "4001";
                        }
                        #endregion

                    }
                    else if (dblLIFE == 0 && (dblEXTRA != 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "EXTRA";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblEXTRA;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; // Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblEXTRA;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//sar
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));//isa
                            dtDataRow[58] = "4001";
                            
                        }
                        #endregion

                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "ADB";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//sar
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//isa
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; // Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//sar
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));//isa
                            dtDataRow[58] = "4001";
                        }
                        #endregion

                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR != 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "SAR";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblSAR;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; // Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblSAR;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                            
                        }
                        #endregion

                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "SARDI";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblSARDI;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; // Remarks
                        }
                        else
                        {
                            dtDataRow[59] = dblSARDI;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                        }
                        #endregion

                    }
                    else if (dblLIFE != 0 && (dblEXTRA != 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        strBrandedProduct = "EXTRA";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA != 0) && (dblADB != 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        strBrandedProduct = "EXTRA";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);


                        strBrandedProduct = "ADB";
                        _var.dtworkRow03[5] = strBrandedProduct;
                        _var.dtworkRow03[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow03[63] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[62] = "4004";
                            _var.dtworkRow03[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";

                            _var.dtworkRow03[59] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA != 0) && (dblADB != 0) && (dblSAR != 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        strBrandedProduct = "EXTRA";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);

                        strBrandedProduct = "ADB";
                        _var.dtworkRow03[5] = strBrandedProduct;
                        _var.dtworkRow03[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        strBrandedProduct = "SAR";
                        _var.dtworkRow04[5] = strBrandedProduct;
                        _var.dtworkRow04[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow03[63] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[62] = "4004";
                            _var.dtworkRow03[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow04[63] = dblSAR;//Renewal Premium
                            _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow04[62] = "4004";
                            _var.dtworkRow04[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                           
                            _var.dtworkRow02[59] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                           
                            _var.dtworkRow03[59] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[58] = "4001";
                            
                            _var.dtworkRow04[59] = dblADB;//Renewal Premium
                            _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow04[58] = "4001";
                           
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                        objdt_template.Rows.Add(_var.dtworkRow03);
                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        strBrandedProduct = "ADB";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblADB;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                           

                            _var.dtworkRow02[59] = dblADB;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                            
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR == 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));
                      

                        strBrandedProduct = "ADB";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        strBrandedProduct = "SARDI";
                        _var.dtworkRow03[5] = strBrandedProduct;
                        _var.dtworkRow03[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblADB;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow03[63] = dblSARDI;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[62] = "4004";
                            _var.dtworkRow03[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                          

                            _var.dtworkRow02[59] = dblADB;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                            

                            _var.dtworkRow03[59] = dblSARDI;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[58] = "4001";
                           
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR != 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = dblVolume;
                        dtDataRow[28] = dblLifeRet;

                        strBrandedProduct = "SAR";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblSAR;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblSAR;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        strBrandedProduct = "SARDI";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR != 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "ADB";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        strBrandedProduct = "SAR";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblSAR;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblSAR;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR == 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "ADB";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        strBrandedProduct = "SARDI";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR != 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "SAR";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));

                        strBrandedProduct = "SARDI";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblADB;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";

                            _var.dtworkRow02[59] = dblSARDI;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA == 0) && (dblADB != 0) && (dblSAR != 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));


                        strBrandedProduct = "ADB";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));

                        strBrandedProduct = "SAR";
                        _var.dtworkRow03[5] = strBrandedProduct;
                        _var.dtworkRow03[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));


                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblADB;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow03[63] = dblSAR;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[62] = "4004";
                            _var.dtworkRow03[76] = wsraw.Cells[intLoop, 12].Value; //Remarks



                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";


                            _var.dtworkRow02[59] = dblADB ;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";

                            _var.dtworkRow03[59] = dblSAR;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[58] = "4001";

                        

                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }
                    else if (dblLIFE != 0 && (dblEXTRA != 0) && (dblADB != 0) && (dblSAR != 0) && (dblSARDI != 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));
                        

                        strBrandedProduct = "EXTRA";
                        _var.dtworkRow02[5] = strBrandedProduct;
                        _var.dtworkRow02[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow02[25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                        _var.dtworkRow02[28] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                    

                        strBrandedProduct = "ADB";
                        _var.dtworkRow03[5] = strBrandedProduct;
                        _var.dtworkRow03[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow03[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBVolume));
                        _var.dtworkRow03[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblADBamt));
                     

                        strBrandedProduct = "SAR";
                        _var.dtworkRow04[5] = strBrandedProduct;
                        _var.dtworkRow04[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARVolume));
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARamt));
   

                        strBrandedProduct = "SARDI";
                        _var.dtworkRow04[5] = strBrandedProduct;
                        _var.dtworkRow04[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        _var.dtworkRow04[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSDIVolume));
                        _var.dtworkRow04[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblSARDIamt));
                      

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                           dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow02[63] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[62] = "4004";
                            _var.dtworkRow02[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow03[63] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[62] = "4004";
                            _var.dtworkRow03[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow04[63] = dblSAR;//Renewal Premium
                            _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow04[62] = "4004";
                            _var.dtworkRow04[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                             _var.dtworkRow05[63] = dblSARDI;//Renewal Premium
                            _var.dtworkRow05[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow05[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow05[62] = "4004";
                            _var.dtworkRow05[76] = wsraw.Cells[intLoop, 12].Value; //Remarks

                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                          

                            _var.dtworkRow02[59] = dblEXTRA;//Renewal Premium
                            _var.dtworkRow02[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow02[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow02[58] = "4001";
                           

                            _var.dtworkRow03[59] = dblADB;//Renewal Premium
                            _var.dtworkRow03[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow03[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow03[58] = "4001";
                          

                            _var.dtworkRow04[59] = dblSAR;//Renewal Premium
                            _var.dtworkRow04[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow04[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow04[58] = "4001";
                            

                            _var.dtworkRow05[59] = dblSARDI;//Renewal Premium
                            _var.dtworkRow05[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            _var.dtworkRow05[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            _var.dtworkRow05[58] = "4001";
                         

                        }
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow02);
                        objdt_template.Rows.Add(_var.dtworkRow03);
                        objdt_template.Rows.Add(_var.dtworkRow04);
                        objdt_template.Rows.Add(_var.dtworkRow05);
                    }
                    else if (dblLIFE == 0 && (dblEXTRA == 0) && (dblADB == 0) && (dblSAR == 0) && (dblSARDI == 0))
                    {
                        strBrandedProduct = "LIFE";
                        dtDataRow[5] = strBrandedProduct;
                        dtDataRow[6] = objHlpr.fn_bpSicsCode(strBrandedProduct); //BRANDED PRODUCT SICS CODE
                        dtDataRow[25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblVolume));
                        dtDataRow[28] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dblLifeRet));

                        #region RENEWAL OR ADJUSTMENTS
                        if (str_sheet.ToUpper().Contains("R&A_RECAP"))

                        {
                            dtDataRow[63] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);
                            dtDataRow[62] = "4004";
                            dtDataRow[76] = wsraw.Cells[intLoop, 12].Value; //Remarks
                        }
                        else
                        {
                            dtDataRow[59] = dblLIFE;//Renewal Premium
                            dtDataRow[77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 12].Value));//SUM AT RISK
                            dtDataRow[27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells[intLoop, 11].Value));
                            dtDataRow[58] = "4001";
                        }
                        #endregion
                    }
                    #endregion

                    //_var.dtworkRow06 = objdt_template.NewRow();
                    //_var.dtworkRow06.ItemArray = dtDataRow.ItemArray;

                    #region HashTotal
                    if (str_sheet.ToUpper().Contains("R&A_RECAP"))
                    {
                        dclTotalPremium += dblLIFE + dblADB + dblEXTRA + dblSAR + dblSARDI;
                        
                    }
                    else
                    {
                        decimal dblSumAtRisk = Convert.ToDecimal(wsraw.Cells[intLoop, 12].Value);
                        dclTotalSumAtRisk += dblSumAtRisk;
                        dclTotalPremium += dblLIFE + dblADB + dblEXTRA + dblSAR + dblSARDI;
                    }
                    #endregion

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

            string despath = str_saved + @"\BM033-" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            Variables.TotalPremium = 0;
            Variables.TotalSumAtRisk = 0;
            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";

        }
    }

}