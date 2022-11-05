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
    class BM021_A
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
            int erawrow = rawrange.Rows.Count;

            decimal dclTotalTreatyCom = 0;
            decimal dclTotalFaculCom = 0;
            decimal dclTotalTreatyPremium = 0;
            decimal dclTotalFacPremium = 0;
            decimal dclTotalTreatySAR = 0;
            decimal dclTotalFacSAR = 0;
            string valueTransEffectiveDate = string.Empty;
            bool bolTransCode = false;
            string TransCode = string.Empty;
            

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }

            DataRow dtDataRow;
            if(str_sheet.ToUpper().Contains("Q20"))
            {
                for(int intLoop = 18; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 32].Text;
                    bolTransCode = objHlpr2.fn_getTranscode(strPolicyNo, out string withTransCode);


                    if(bolTransCode == true)
                    {
                        TransCode = withTransCode; //Transcode
                    }

                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 33].Text, wsraw.Cells [intLoop, 34].Text, wsraw.Cells [intLoop, 35].Text))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    dtDataRow [0] = strPolicyNo;
                    string brandedProduct = Convert.ToString(wsraw.Cells [intLoop, 45].Text);//Branded Product
                    dtDataRow [5] = Convert.ToString(wsraw.Cells [intLoop, 45].Value);//Branded Product;
                    string strCurrency = "PHP";
                    string strCOB = "IND";
                    dtDataRow [23] = strCurrency; //  Cession Currency
                    dtDataRow [13] = strCOB; // Class of Business
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product   
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [41] = Variables.strBmYear;/*PolicyYear.ToString("MM/yyyy");*/ //Policy Year
                    //string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 39].Value).ToString("MM/dd/yyyy");
                    //dtDataRow [22] = objHlpr2.fn_getTransReinsuranceDate(TransCode, Variables.strBmYear, strIssueDate, out valueTransEffectiveDate); // Trans Effective Date
                    //dtDataRow [20] = objHlpr2.fn_getPolicyStartDate(TransCode, strIssueDate, valueTransEffectiveDate);//Policy Start Date
                    //dtDataRow [19] = valueTransEffectiveDate;  // Reinsurance Start Date

                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 39].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    objHlpr2.fn_getTransReinsuranceDateV3(strIssueDate, Variables.strBmYear, out string transEffectiveDate);
                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                    dtDataRow [20] = strIssueDate;//Policy Start Date
                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date
                    dtDataRow [38] = objHlpr.fn_SmokerCode("");
                    dtDataRow [31] = wsraw.Cells [intLoop, 35].Text; //FULLNAME
                    dtDataRow [32] = wsraw.Cells [intLoop, 36].Text; //LASTNAME
                    dtDataRow [33] = wsraw.Cells [intLoop, 37].Text;//FIRSTNAME
                    dtDataRow [34] = wsraw.Cells [intLoop, 38].Text;//MIDDLENAME
                    dtDataRow [37] = wsraw.Cells [intLoop, 41].Text;//DATE OF BIRTH
                    dtDataRow [36] = wsraw.Cells [intLoop, 43].Value;//GENDER
                    dtDataRow [39] = objHlpr2.fn_getmortalityrating(wsraw.Cells [intLoop, 47].Text);//Preferred Classific
                    dtDataRow [79] = Convert.ToString(wsraw.Cells [intLoop, 42].Text);//ISSUE AGE
                    dtDataRow [30] = objHlpr.fn_LifeID(Convert.ToString(wsraw.Cells [intLoop, 37].Text), Convert.ToString(wsraw.Cells [intLoop, 36].Text), Convert.ToString(wsraw.Cells [intLoop, 41].Text)); //LIFEID



                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 13].Value), out decimal dclShare);
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 51].Value), out decimal dclSumAtRisk); //Sum At Risk
                    decimal.TryParse(Convert.ToString(wsraw.Cells [intLoop, 52].Value), out decimal dclInitialSUm); //InitialSum
                    dclInitialSUm = dclInitialSUm * dclShare;
                    dclSumAtRisk = dclSumAtRisk * dclShare;

                    #region TREATY
                    decimal.TryParse(wsraw.Cells [intLoop, 14].Text, out decimal dclTreatyFY); //COL N
                    decimal.TryParse(wsraw.Cells [intLoop, 16].Text, out decimal dclTreatyRen); //COL P
                    decimal.TryParse(wsraw.Cells [intLoop, 18].Text, out decimal dclTreatyCOM); //COL R
                    #endregion

                    #region FACULTATIVE
                    decimal.TryParse(wsraw.Cells [intLoop, 15].Text, out decimal dclFaculFY); //COL O
                    decimal.TryParse(wsraw.Cells [intLoop, 17].Text, out decimal dclFaculRen); //COL Q
                    decimal.TryParse(wsraw.Cells [intLoop, 19].Text, out decimal dclFaculCOM); //COL S
                    #endregion

                    if(TransCode == "TCONTER" || TransCode == "TLAPSE")
                    {
                        dclTreatyFY = dclTreatyFY * -1;
                        dclTreatyCOM = dclTreatyCOM * -1;
                        dclTreatyRen = dclTreatyRen * -1;

                        dclFaculFY = dclFaculFY * -1;
                        dclFaculCOM = dclFaculCOM * -1;
                        dclFaculRen = dclFaculRen * -1;
                    }

                    decimal dclTreatyColNR = dclTreatyFY - dclTreatyCOM; //TREATY
                    decimal dclFacColOS = dclFaculFY - dclFaculCOM; //FACUL
                    dclTotalTreatyPremium += dclTreatyColNR + dclTreatyRen; //TREATY PREMIUM
                    dclTotalFacPremium += dclFacColOS + dclFaculRen;// FACUL PREMIUM

                    _var.dtworkRow02 = objdt_template.NewRow();
                    _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;

                    if(TransCode == "TNEWBUS" || TransCode == "TRENEW")
                    {
                        if(dclTreatyFY != 0 && dclTreatyRen == 0)
                        {
                            dtDataRow [56] = "4000";
                            dtDataRow [57] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyRen != 0 && dclTreatyFY == 0)
                        {
                            dtDataRow [58] = "4001";
                            dtDataRow [59] = dclTreatyRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyFY != 0 && dclTreatyRen != 0)
                        {
                            dclTotalTreatySAR += dclSumAtRisk;
                            dtDataRow [58] = "4001";
                            dtDataRow [59] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value)) ;//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";


                            _var.dtworkRow02 [58] = "4001";
                            _var.dtworkRow02 [59] = dclTreatyRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "T";
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                        if(dclFaculFY != 0 && dclFaculRen == 0)
                        {
                            dtDataRow [56] = "4000";
                            dtDataRow [57] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculRen != 0 && dclFaculFY == 0)
                        {
                            dtDataRow [58] = "4001";
                            dtDataRow [59] = dclFaculRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculFY != 0 && dclFaculRen != 0)
                        {
                            dtDataRow [56] = "4000";
                            dtDataRow [57] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;

                            _var.dtworkRow02 [58] = "4001";
                            _var.dtworkRow02 [59] = dclFaculRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "F";
                            objdt_template.Rows.Add(_var.dtworkRow02);


                        }
                        else if(dclTreatyFY == 0 && dclFaculFY == 0 && dclTreatyRen == 0 && dclFaculRen == 0)
                        {
                            dtDataRow [56] = "4000";
                            dtDataRow [57] = 0;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }

                    }
                    else if(TransCode == "TCONTER" || TransCode == "TLAPSE")
                    {
                        if(dclTreatyFY != 0 && dclTreatyRen == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [61] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyRen != 0 && dclTreatyFY == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclTreatyRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyFY != 0 && dclTreatyRen != 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;

                            _var.dtworkRow02 [62] = "4004";
                            _var.dtworkRow02 [63] = dclTreatyRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "T";
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                        if(dclFaculFY != 0 && dclFaculRen == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [61] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculRen != 0 && dclFaculFY == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclFaculRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculFY != 0 && dclFaculRen != 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;

                            _var.dtworkRow02 [62] = "4004";
                            _var.dtworkRow02 [63] = dclFaculRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "T";
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                        if(dclFaculRen == 0 && dclFaculFY == 0 && dclTreatyFY == 0 && dclTreatyRen == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = 0;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                    }
                    else if(TransCode == "TADJUST")
                    {
                        if(dclTreatyFY != 0 && dclTreatyRen == 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyRen != 0 && dclTreatyFY == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclTreatyRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                        else if(dclTreatyFY != 0 && dclTreatyRen != 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = dclTreatyColNR;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;

                            _var.dtworkRow02 [62] = "4004";
                            _var.dtworkRow02 [63] = dclTreatyRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "T";
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }

                        if(dclFaculFY != 0 && dclFaculRen == 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculRen != 0 && dclFaculFY == 0)
                        {
                            dtDataRow [62] = "4004";
                            dtDataRow [63] = dclFaculRen;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;
                        }
                        else if(dclFaculFY != 0 && dclFaculRen != 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = dclFacColOS;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "F";
                            dclTotalFacSAR += dclSumAtRisk;

                            _var.dtworkRow02 [62] = "4004";
                            _var.dtworkRow02 [63] = dclTreatyRen;
                            _var.dtworkRow02 [27] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//InitialSum
                            _var.dtworkRow02 [77] = objHlpr.fn_CheckingValueZeroOrEmpty(null);// sum at risk
                            _var.dtworkRow02 [28] = 1; //CEDED RETENTION
                            _var.dtworkRow02 [25] = objHlpr.fn_CheckingValueZeroOrEmpty(null);//orignal sum risk
                            _var.dtworkRow02 [21] = TransCode;
                            _var.dtworkRow02 [14] = "F";
                            objdt_template.Rows.Add(_var.dtworkRow02);


                        }
                        if(dclTreatyFY == 0 && dclFaculFY == 0 && dclTreatyRen == 0 && dclFaculRen == 0)
                        {
                            dtDataRow [60] = "4002";
                            dtDataRow [61] = 0;
                            dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclInitialSUm));//InitialSum
                            dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));// sum at risk
                            dtDataRow [28] = wsraw.Cells [intLoop, 49].Value; //CEDED RETENTION
                            dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(wsraw.Cells [intLoop, 46].Value));//orignal sum risk
                            dtDataRow [21] = TransCode;
                            dtDataRow [14] = "T";
                            dclTotalTreatySAR += dclSumAtRisk;
                        }
                    }
                }
            }
           
            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Treaty Premium:";
            dtDataRow [1] = dclTotalTreatyPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Treaty Sum at Risk:";
            dtDataRow [1] = dclTotalTreatySAR;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Falcultative Premium:";
            dtDataRow [1] = dclTotalFacPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Falcultative Sum at Risk:";
            dtDataRow [1] = dclTotalFacSAR;
            objdt_template.Rows.Add(dtDataRow);
            #endregion


            string despath = str_saved + @"\BM021-A" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalTreatyPremium = 0;
            dclTotalFacPremium = 0;
            dclTotalTreatySAR = 0;
            //dclTotalCommission = 0;

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";




        }
    }

}