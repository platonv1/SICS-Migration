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
    class BM117
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


            double dclTotalPremium = 0;
            double dclTotaSAR = 0;
            string valueTransEffectiveDate = string.Empty;
            bool bolTransCode = false;
            string TransCode = string.Empty;
            string withTransCode = string.Empty;

            while(string.IsNullOrEmpty(Variables.strBmYear))
            {
                frmPolicyYear newform = new frmPolicyYear();
                newform.ShowDialog();

            }

            DataRow dtDataRow;
            if(str_sheet.ToUpper().Contains("URC$") || str_sheet.ToUpper().Contains("URC"))
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string strPolicyNo = wsraw.Cells [intLoop, 2].Text;
                    
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text, wsraw.Cells [intLoop, 3].Text, wsraw.Cells [intLoop, 4].Text))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strCessionNo = wsraw.Cells [intLoop, 3].Text;
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [31] = wsraw.Cells [intLoop, 4].Text; //FULLNAME
                    objHlpr2.fn_separateLastNameFirstNameV8(wsraw.Cells [intLoop, 4].Text, out string strLastName, out string strFirstName, out string strMI);
                    dtDataRow [32] = strLastName; //LASTNAME
                    dtDataRow [33] = strFirstName;//FIRSTNAME
                    dtDataRow [34] = strMI;//MIDDLENAME
                    string strDOB = Convert.ToDateTime(wsraw.Cells [intLoop, 5].Text).ToString("MM/dd/yyyy");//DATE OF BIRTH)
                    dtDataRow [5] = objHlpr2.fn_getplanCode(strPolicyNo, strCessionNo);
                    dtDataRow [37] = strDOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, Convert.ToString(wsraw.Cells [intLoop, 4].Text)); //LIFEID
                    dtDataRow [36] = wsraw.Cells [intLoop, 6].Text; //Gender
                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    objHlpr2.fn_getTransReinsuranceDateV6(strIssueDate, Variables.strBmYear, out string transEffectiveDate);
                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                    dtDataRow [20] = strIssueDate;//Policy Start Date
                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date
                    dtDataRow [78] = Convert.ToString(wsraw.Cells [intLoop, 8].Value);//Attain Age
                    dtDataRow [79] = objHlpr.fn_getIssueAge(strDOB, strIssueDate);//ISSUE AGE
                    //dtDataRow [5] = Convert.ToString(wsraw.Cells [intLoop, 9].Value);//Branded Product;
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product   
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [13] = "IND"; // Class of Business    
                    dtDataRow [23] = objHlpr2.fn_getcurrencyV2(str_sheet); //  Cession Currency
                    dtDataRow [21] = "TRENEW"; // Transaction Code
                    dtDataRow [14] = objHlpr2.fn_businessTypeV2(wsraw.Cells [intLoop, 10].Value); // Business Type
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [39] = objHlpr.fn_getmortality(wsraw.Cells [intLoop, 11].Value); // Preferred Classific
                    dtDataRow [38] = objHlpr.fn_SmokerCode("");
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out double dclSumAtRisk); //Sum At Risk, Original Sum. Initial Sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Initial Sum
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Sum at risk
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk)); //Orig Sum
                    dtDataRow [76] = objHlpr2.fn_RemarksBusinessType(wsraw.Cells [intLoop, 10].Value);//Remarks
                    dtDataRow [58] = "4001";
                    #region Premiums

                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 16].Value), out double dclPremLife);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 17].Value), out double dclPremExtra);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 18].Value), out double dclPremWP);
                    #endregion

                    #region Premium
                    if(dclPremLife != 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclPremLife;

                        if(dclPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            _var.dtworkRow02 [59] = dclPremExtra;
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            _var.dtworkRow02 [59] = dclPremWP;
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }

                    }
                    else if(dclPremExtra != 0)
                    {
                        //dtDataRow [5] = "EXTRA";
                        dtDataRow [59] = dclPremExtra;

                        if(dclPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            _var.dtworkRow02 [59] = dclPremLife;
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            _var.dtworkRow02 [61] = dclPremWP;
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if(dclPremWP != 0)
                    {
                        //dtDataRow [5] = "WP/PB";
                        dtDataRow [59] = dclPremWP;

                        if(dclPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            _var.dtworkRow02 [59] = dclPremLife;
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            _var.dtworkRow02 [59] = dclPremExtra;
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if(dclPremLife == 0 && dclPremExtra == 0 && dclPremWP == 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclPremLife;
                    }
                    #endregion

                    #region hashtotal
                    dclTotalPremium += dclPremLife + dclPremExtra + dclPremWP;
                    if(dclPremLife != 0)
                    {
                        dclTotaSAR += dclSumAtRisk;
                    }
                    #endregion
                }
            }
            else if(str_sheet.ToUpper().Contains("SURR") || str_sheet.ToUpper().Contains("LAPS")  || str_sheet.ToUpper().Contains("TERM") || str_sheet.ToUpper().Contains("MATUR") ||
            str_sheet.ToUpper().Contains("ADJ"))
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    #region lookup for transcode
                    //if(bolTransCode == false)
                    //{
                    //    string strgetTranscode = wsraw.Cells [intLoop, 1].Text;
                    //    bolTransCode = objHlpr2.fn_getTranscode(strgetTranscode, out withTransCode);
                    //}


                    //if(bolTransCode == true)
                    //{
                    //    if(string.IsNullOrEmpty(TransCode))
                    //    {
                    //        TransCode = withTransCode; //Transcode
                    //    }

                    //}
                    #endregion
                    string strPolicyNo = wsraw.Cells [intLoop, 2].Text;
                    
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text, wsraw.Cells [intLoop, 3].Text, wsraw.Cells [intLoop, 4].Text))
                    {
                        continue;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strCessionNo = wsraw.Cells [intLoop, 3].Text;
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [31] = wsraw.Cells [intLoop, 4].Text; //FULLNAME
                    objHlpr2.fn_separateLastNameFirstNameV8(wsraw.Cells [intLoop, 4].Text, out string strLastName, out string strFirstName, out string strMI);
                    dtDataRow [32] = strLastName; //LASTNAME
                    dtDataRow [33] = strFirstName;//FIRSTNAME
                    dtDataRow [34] = strMI;//MIDDLENAME
                    string DOB = objHlpr.fn_getDOB("");//DATE OF BIRTH)
                    dtDataRow [37] = DOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, DOB); //LIFEID
                    dtDataRow [36] = objHlpr.fn_getgenderv2(strFirstName); //Gender
                    string transcode = str_sheet.ToString();
                    objHlpr2.fn_getTranscode(transcode, out transcode);
                    dtDataRow [21] = transcode; // Transaction Code
                    dtDataRow [5] = objHlpr2.fn_getplanCode(strPolicyNo, strCessionNo);
                    if (transcode == "TLAPSE" || transcode == "TFULLREC")
                    {
                        
                        dtDataRow [22] = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");//Transeffective date
                        dtDataRow [20] = Convert.ToDateTime(wsraw.Cells [intLoop, 6].Value).ToString("MM/dd/yyyy");//Policy Start Date;//Policy Start Date
                        dtDataRow [19] = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy"); // Reinsurance Start Date
                    }
                    else
                    {
                        string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 6].Value).ToString("MM/dd/yyyy");//Policy Start Date
                        objHlpr2.fn_getTransReinsuranceDateV6(strIssueDate, Variables.strBmYear, out string transEffectiveDate);
                        dtDataRow [22] = transEffectiveDate; //Transeffective date
                        dtDataRow [20] = strIssueDate;//Policy Start Date
                        dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date
                    }
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product   
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [13] = "IND"; // Class of Business    
                    dtDataRow [23] = "PHP"; //  Cession Currency
                    dtDataRow [14] = objHlpr2.fn_businessTypeV2(wsraw.Cells [intLoop, 8].Value); // Business Type
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [39] = objHlpr.fn_getmortality(""); // Preferred Classific
                    dtDataRow [38] = objHlpr.fn_SmokerCode("");
                    double dclSumAtRisk = 0; //Sum At Risk, Original Sum. Initial Sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Initial Sum
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Sum at risk
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk)); //Orig Sum
                    dtDataRow [76] = objHlpr2.fn_RemarksBusinessType(wsraw.Cells [intLoop, 8].Value);//Remarks
                   
                    #region Premiums
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Value), out double dclRYPremLife);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Value), out double dclRYPremExtra);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out double dclRYPremWP);
                    #endregion

                    #region Premium
                    if(dclRYPremLife != 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        if (transcode == "TLAPSE")
                        {
                            dtDataRow [59] = dclRYPremLife;
                            dtDataRow [58] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremLife;
                            dtDataRow [62] = "4004";
                        }
                      
                        if(dclRYPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [58] = "4001";
                                _var.dtworkRow02 [59] = dclRYPremExtra;
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremExtra;
                            }
                           
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [58] = "4001";
                                _var.dtworkRow02 [59] = dclRYPremWP;
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremWP;
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }

                    }
                    else if(dclRYPremExtra != 0)
                    {
                        //dtDataRow [5] = "EXTRA";
                        if (transcode == "TLAPSE")
                        {
                            dtDataRow [59] = dclRYPremExtra;
                            dtDataRow [58] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremExtra;
                            dtDataRow [62] = "4004";
                        }

                        if(dclRYPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremLife;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremLife;
                            }
                            
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremWP;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremWP;
                            }
                          
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if(dclRYPremWP != 0)
                    {
                        //dtDataRow [5] = "WP/PB";
                        if (transcode == "TLAPSE")
                        {
                            dtDataRow [63] = dclRYPremWP;
                            dtDataRow [62] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremWP;
                            dtDataRow [62] = "4004";
                        }


                        if(dclRYPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremLife;
                                _var.dtworkRow02 [58] = "4001";

                            }
                            else
                            {
                                _var.dtworkRow02 [63] = dclRYPremLife;
                                _var.dtworkRow02 [62] = "4004";
                            }
                          
                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            if (transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremExtra;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [63] = dclRYPremExtra;
                                _var.dtworkRow02 [62] = "4004";
                            }
                            
                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if(dclRYPremLife == 0 && dclRYPremExtra == 0 && dclRYPremWP == 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [63] = dclRYPremLife;
                        dtDataRow [62] = "4004";
                    }
                    #endregion


                    #region hashtotal
                    dclTotalPremium += dclRYPremLife + dclRYPremExtra + dclRYPremWP;
                    if(dclRYPremLife != 0)
                    {
                        dclTotaSAR += dclSumAtRisk;
                    }
                    #endregion
                }
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);
            }
            #region else
            else
            {
                for(int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    #region lookup for transcode
                    //if(bolTransCode == false)
                    //{
                    //    string strgetTranscode = wsraw.Cells [intLoop, 1].Text;
                    //    bolTransCode = objHlpr2.fn_getTranscode(strgetTranscode, out withTransCode);
                    //}


                    //if(bolTransCode == true)
                    //{
                    //    if(string.IsNullOrEmpty(TransCode))
                    //    {
                    //        TransCode = withTransCode; //Transcode
                    //    }

                    //}
                    #endregion
                    string strPolicyNo = wsraw.Cells [intLoop, 2].Text;
                    if(!objHlpr.fn_policyNumChecker(strPolicyNo, wsraw.Cells [intLoop, 2].Text, wsraw.Cells [intLoop, 3].Text, wsraw.Cells [intLoop, 4].Text))
                    {
                        continue;
                    }
                    else if(strPolicyNo == "")
                    {
                        break;
                    }
                    dtDataRow = objdt_template.NewRow();
                    objdt_template.Rows.Add(dtDataRow);
                    string strCessionNo = wsraw.Cells [intLoop, 3].Text;
                    dtDataRow [5] = objHlpr2.fn_getplanCode(strPolicyNo, strCessionNo);
                    dtDataRow [0] = strPolicyNo;
                    dtDataRow [31] = wsraw.Cells [intLoop, 4].Text; //FULLNAME
                    objHlpr2.fn_separateLastNameFirstNameV8(wsraw.Cells [intLoop, 4].Text, out string strLastName, out string strFirstName, out string strMI);
                    dtDataRow [32] = strLastName; //LASTNAME
                    dtDataRow [33] = strFirstName;//FIRSTNAME
                    dtDataRow [34] = strMI;//MIDDLENAME
                    string DOB = objHlpr.fn_getDOB("");//DATE OF BIRTH)
                    dtDataRow [37] = DOB;
                    dtDataRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, DOB); //LIFEID
                    dtDataRow [36] = objHlpr.fn_getgenderv2(strFirstName); //Gender
                    string strIssueDate = Convert.ToDateTime(wsraw.Cells [intLoop, 7].Value).ToString("MM/dd/yyyy");//Policy Start Date
                    objHlpr2.fn_getTransReinsuranceDateV6(strIssueDate, Variables.strBmYear, out string transEffectiveDate);
                    dtDataRow [22] = transEffectiveDate; //Transeffective date
                    dtDataRow [20] = strIssueDate;//Policy Start Date
                    dtDataRow [19] = transEffectiveDate;  // Reinsurance Start Date
                    dtDataRow [8] = "SURPLUS"; // Reinsurance Product   
                    dtDataRow [9] = "PAFM"; // Type of Business
                    dtDataRow [10] = "S"; // Reinsurance Methods
                    dtDataRow [24] = "YLY"; // Premium Frequency
                    dtDataRow [29] = "NATREID"; // Life ID Type
                    dtDataRow [13] = "IND"; // Class of Business    
                    dtDataRow [23] = "PHP"; //  Cession Currency
                    string transcode = str_sheet.ToString();
                    objHlpr2.fn_getTranscode(transcode, out transcode);
                    dtDataRow [21] = transcode; // Transaction Code
                    dtDataRow [14] = objHlpr2.fn_businessTypeV2(wsraw.Cells [intLoop, 8].Value); // Business Type
                    dtDataRow [41] = Variables.strBmYear; //Policy Year
                    dtDataRow [39] = objHlpr.fn_getmortality(""); // Preferred Classific
                    dtDataRow [38] = objHlpr.fn_SmokerCode("");
                    double dclSumAtRisk = 0; //Sum At Risk, Original Sum. Initial Sum
                    dtDataRow [27] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Initial Sum
                    dtDataRow [77] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk));//Sum at risk
                    dtDataRow [25] = objHlpr.fn_CheckingValueZeroOrEmpty(Convert.ToString(dclSumAtRisk)); //Orig Sum
                    dtDataRow [76] = objHlpr2.fn_RemarksBusinessType(wsraw.Cells [intLoop, 8].Value);//Remarks

                    #region Premiums
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 10].Value), out double dclRYPremLife);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 12].Value), out double dclRYPremExtra);
                    double.TryParse(Convert.ToString(wsraw.Cells [intLoop, 14].Value), out double dclRYPremWP);
                    #endregion

                    #region Premium

                    if(dclRYPremLife != 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        if(transcode == "TLAPSE")
                        {
                            dtDataRow [59] = dclRYPremLife;
                            dtDataRow [58] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremLife;
                            dtDataRow [62] = "4004";
                        }

                        if(dclRYPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            if(transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [58] = "4001";
                                _var.dtworkRow02 [59] = dclRYPremExtra;
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremExtra;
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            if(transcode == "TLPASE")
                            {
                                _var.dtworkRow02 [58] = "4001";
                                _var.dtworkRow02 [59] = dclRYPremWP;
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremWP;
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }

                    }
                    else if(dclRYPremExtra != 0)
                    {
                        //dtDataRow [5] = "EXTRA";
                        if(transcode == "TLAPSE")
                        {
                            dtDataRow [59] = dclRYPremExtra;
                            dtDataRow [58] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremExtra;
                            dtDataRow [62] = "4004";
                        }

                        if(dclRYPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            if(transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremLife;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremLife;
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremWP != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "WP/PB";
                            if(transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremWP;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [62] = "4004";
                                _var.dtworkRow02 [63] = dclRYPremWP;
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if(dclRYPremWP != 0)
                    {
                        //dtDataRow [5] = "WP/PB";
                        if(transcode == "TLAPSE")
                        {
                            dtDataRow [59] = dclRYPremWP;
                            dtDataRow [58] = "4001";
                        }
                        else
                        {
                            dtDataRow [63] = dclRYPremWP;
                            dtDataRow [62] = "4004";
                        }


                        if(dclRYPremLife != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "LIFE";
                            if(transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremLife;
                                _var.dtworkRow02 [58] = "4001";

                            }
                            else
                            {
                                _var.dtworkRow02 [63] = dclRYPremLife;
                                _var.dtworkRow02 [62] = "4004";
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);

                        }

                        if(dclRYPremExtra != 0)
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = dtDataRow.ItemArray;
                            //_var.dtworkRow02 [5] = "EXTRA";
                            if(transcode == "TLAPSE")
                            {
                                _var.dtworkRow02 [59] = dclRYPremExtra;
                                _var.dtworkRow02 [58] = "4001";
                            }
                            else
                            {
                                _var.dtworkRow02 [63] = dclRYPremExtra;
                                _var.dtworkRow02 [62] = "4004";
                            }

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }

                    else if(dclRYPremLife == 0 && dclRYPremExtra == 0 && dclRYPremWP == 0)
                    {
                        //dtDataRow [5] = "LIFE";
                        dtDataRow [59] = dclRYPremLife;
                        dtDataRow [58] = "4001";
                    }
                    #endregion


                    #region hashtotal
                    dclTotalPremium += dclRYPremLife + dclRYPremExtra + dclRYPremWP;
                    if(dclRYPremLife != 0)
                    {
                        dclTotaSAR += dclSumAtRisk;
                    }
                    #endregion
                }
                dtDataRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtDataRow);
            }
            #endregion


            dtDataRow = objdt_template.NewRow();
            objdt_template.Rows.Add(dtDataRow);

            #region Computing Hash 
            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total  Premium:";
            dtDataRow [1] = dclTotalPremium;
            objdt_template.Rows.Add(dtDataRow);

            dtDataRow = objdt_template.NewRow();
            dtDataRow [0] = "Total Sum at Risk:";
            dtDataRow [1] = dclTotaSAR;
            objdt_template.Rows.Add(dtDataRow);

            #endregion


            string despath = str_saved + @"\BM117" + str_sheet + str_savef + ".xlsx";
            objHlpr.fn_savefile(objdt_template, despath);

            objHlpr.fn_openfile(despath);

            dclTotalPremium = 0;
            dclTotaSAR = 0;

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;
            return "";




        }
    }

}