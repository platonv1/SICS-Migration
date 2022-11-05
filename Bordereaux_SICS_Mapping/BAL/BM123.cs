using System;
using System.Data;
using System.Globalization;
using System.Linq;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM123
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender, bool boo_open = false, bool boo_clean = false, string str_macro = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            DataTable objdt_template = new DataTable();

            DataTable dt_macro = new DataTable();
            if (!String.IsNullOrEmpty(str_macro))
            {
                dt_macro = objHlpr.fn_Loadmacro(str_macro);
            }

            objdt_template = objHlpr.dt_formtemplate(str_sheet);

            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int erawrow = rawrange.Rows.Count;
            int int_RowCnt = 0;

            double dbl_rate = 1;
            bool boo_invalidIssueDate = false;
            bool boo_FirstYear = false;
            string strbirth = string.Empty;
            string strGenderOutput = string.Empty;
            string strCessionCode = string.Empty;
           

            try
            {

                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string str_PolNum = wsraw.Cells[intLoop, 4].Text.ToString();

                    if (!objHlpr.fn_policyNumChecker(str_PolNum, wsraw.Cells[intLoop, 5].Text.ToString(), wsraw.Cells[intLoop, 6].Text.ToString(), wsraw.Cells[intLoop, 7].Text.ToString()))
                    {
                        continue;
                    }

                    boo_invalidIssueDate = false;

                    string str_CessionNo = wsraw.Cells[intLoop, 5].Text.ToString();
                    string str_PremDueDate = wsraw.Cells[intLoop, 7].Text.ToString();
                    string str_bmyear = str_PremDueDate.Substring(str_PremDueDate.Length - 4, 4);
                    string str_NAAR = wsraw.Cells[intLoop, 9].Text.ToString();
                    string str_PremiumLife = wsraw.Cells[intLoop, 10].Text.ToString();
                    string str_PremiumExtra = wsraw.Cells[intLoop, 11].Text.ToString();
                    string str_Status = wsraw.Cells[intLoop, 14].Text.ToString();
                    string str_Particular = wsraw.Cells[intLoop, 13].Text.ToString();
                    string str_Fullname = wsraw.Cells[intLoop, 15].Text.ToString();

                    if (double.TryParse(str_NAAR, out double dbl_NAAR))
                    {
                        dbl_NAAR = dbl_NAAR * dbl_rate;
                    }
                    else
                    {
                        dbl_NAAR = 1;
                    }

                    if (double.TryParse(str_PremiumLife, out double dbl_PremiumLife))
                    {
                        dbl_PremiumLife = dbl_PremiumLife * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumLife = 0;
                    }

                    if (double.TryParse(str_PremiumExtra, out double dbl_PremiumExtra))
                    {
                        dbl_PremiumExtra = dbl_PremiumExtra * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumExtra = 0;
                    }

                    string strFullName = "";
                    string strDOB = "";
                    string strIssueAge = "";
                    string strMortality = "";
                    string strSex = "";
                    //string str_refunding = "";
                    //string str_IssueDate = "";

                    string str_OSA = "";
                    string str_ISR = "";
                    string str_Ret = "";
                    double dbl_PremiumAmnt = 0;


                    if (double.TryParse(str_OSA, out double dbl_OSA))
                    {
                        dbl_OSA = dbl_OSA * dbl_rate;
                    }
                    else
                    {
                        dbl_OSA = 1;
                    }

                    if (double.TryParse(str_ISR, out double dbl_ISR))
                    {
                        dbl_ISR = dbl_ISR * dbl_rate;
                    }
                    else
                    {
                        dbl_ISR = 1;
                    }

                    if (double.TryParse(str_Ret, out double dbl_Ret))
                    {
                        dbl_Ret = dbl_Ret * dbl_rate;
                    }
                    else
                    {
                        dbl_Ret = 1;
                    }


                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow02 = null;
                    string str_tcode = "TRENEW";
                    _var.dtworkRow01[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow01[1] = str_CessionNo;
                    _var.dtworkRow01[5] = "LIFE";
                    _var.dtworkRow01[8] = "SURPLUS";
                    _var.dtworkRow01[9] = "PAFW";//Type of business
                    _var.dtworkRow01[13] = "IND";//Class of business
                    _var.dtworkRow01[10] = "S"; //Type of business
                    //objHlpr.fn_getbusinesstype(str_CessionNo, out str_CessionNo);
                    //_var.dtworkRow01[14] = str_CessionNo; //Business type
                    _var.dtworkRow01[23] = "USD";
                    _var.dtworkRow01[24] = "YLY";
                    _var.dtworkRow01[29] = "NATREID";

                    //Remove Zeros
                    Console.WriteLine(str_CessionNo);
                    if (str_CessionNo.Contains("NR000") || str_CessionNo.Contains("000"))
                    {
                        str_CessionNo = str_CessionNo.Replace("000", "");
                    }

                    else if (str_CessionNo.Contains("00"))
                    {
                        str_CessionNo = str_CessionNo.Replace("00", "");

                    }

                    //--------- >Add the macro code below before proceeding to DOB condtion
                    objHlpr.fn_macrobenlifebm123(str_CessionNo, str_PolNum, out string strPolNo, out strIssueAge, out string strIssueDate,
                        out strMortality, out string strRefunding, out strFullName, out string strFirstName, out string strLastName,
                        out string strMI, out string strTitle, out strDOB, out strSex, out string strLifeID,
                        out string str_LE_OSA, out string str_LE_ISR, out string str_LE_Ret, out string strRcDummyName, out strCessionCode);

                    _var.dtworkRow01[0] = strPolNo;
                    _var.dtworkRow01[79] = strIssueAge;
                    _var.dtworkRow01[30] = strLifeID;
                    _var.dtworkRow01[31] = strFullName; //Fullname
                    _var.dtworkRow01[41] = str_bmyear;
                    _var.dtworkRow01[32] = strLastName;
                    _var.dtworkRow01[33] = strFirstName;
                    _var.dtworkRow01[34] = strMI;
                    _var.dtworkRow01[35] = strTitle;
                    _var.dtworkRow01[36] = strSex;
                    _var.dtworkRow01[37] = strDOB;
                    _var.dtworkRow01[83] = strRefunding;
                    _var.dtworkRow01[14] = strCessionCode;

                    _var.dtworkRow01[38] = "NONE";
                    _var.dtworkRow01[41] = str_bmyear;//Policy Year
                    #region
                    //if (string.IsNullOrEmpty(str_gender))
                    //{
                    //    objHlpr.fn_getgenderv2(strFirstName.ToUpper(), out strGenderOutput);

                    //    if (string.IsNullOrEmpty(strGenderOutput))
                    //    {

                    //        _var.dtworkRow01[36] = "Add this person to gender database";
                    //        _var.dtworkRow01[37] = strDOB;
                    //        _var.boo_genderfail = true;

                    //    }
                    //    else
                    //    {
                    //        _var.dtworkRow01[36] = strGenderOutput;
                    //        _var.dtworkRow01[37] = strDOB;
                    //    }

                    //}
                    //else   // initialise this condidition if user uploaded a gender db file 
                    //{

                    //    strGenderOutput = objHlpr.fn_getgender(str_gender, strFirstName);
                    //}

                    //if (string.IsNullOrEmpty(str_gender))
                    //    {
                    //    objHlpr.fn_getgenderv2(strFirstNamev2.ToUpper(), out strGenderOutput);
                    //        if (string.IsNullOrEmpty(strGenderOutput))
                    //        {
                    //            _var.boo_genderfail = true;
                    //            _var.dtworkRow01[36] = "Add this person to gender database";
                    //        }
                    //        else
                    //        {
                    //            _var.dtworkRow01[36] = strGenderOutput;
                    //        }

                    //    }
                    //else   //Initialise this condidition if user uploaded a gender db file 
                    //{

                    //    strGenderOutput = objHlpr.fn_getgender(str_gender, strFirstNamev2);
                    //}

                    #endregion

                    //Macro Old code
                    #region
                    //if (!String.IsNullOrEmpty(str_macro)  && (!String.IsNullOrEmpty(str_CessionNo)))
                    //{

                    //    //DataRow[] foundRows = dt_macro.Selec("")
                    //    DataRow[] foundRows = dt_macro.Select("CN = " + "'" + str_CessionNo.ToString() + "'");
                    //    if (foundRows.Length != 0)
                    //    {
                    //        str_age = foundRows[0][5].ToString();
                    //        str_IssueDate = foundRows[0][7].ToString();
                    //        str_Mortality = foundRows[0][8].ToString();
                    //        str_refunding = foundRows[0][9].ToString();
                    //        str_Fullname = foundRows[0][10].ToString();
                    //        str_DOB = foundRows[0][12].ToString();
                    //        str_Sex = foundRows[0][13].ToString();

                    //        //Split Full Name
                    //        objHlpr.fn_getnamesandlifeID(str_Fullname, str_DOB, out string str_outfname, out string str_outlname, out string str_outlifeid, "021");
                    //        _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(str_Fullname);
                    //        string str_MI = objHlpr.fn_getMI(str_outfname);
                    //        _var.dtworkRow01[32] = str_outlname.Trim();
                    //        if (str_MI.Trim() != string.Empty)
                    //        {
                    //            _var.dtworkRow01[33] = str_outfname.Trim().Replace(" " + str_MI.Trim(), "");
                    //            _var.dtworkRow01[34] = str_MI.Trim();
                    //        }
                    //        else
                    //        {
                    //            string[] arr_fname = str_outfname.Split(' ');
                    //            _var.dtworkRow01[33] = str_outfname.Trim().Replace(" " + arr_fname[arr_fname.Length - 1], "");
                    //            _var.dtworkRow01[34] = arr_fname[arr_fname.Length - 1];
                    //        }
                    //        _var.dtworkRow01[30] = str_outlifeid;
                    //        _var.dtworkRow01[37] = str_DOB;



                    //        ////Gender Old Code
                    //        //if (!String.IsNullOrEmpty(str_Sex))
                    //        //{
                    //        //    _var.dtworkRow01[36] = (str_Sex.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                    //        //}
                    //        //else if (String.IsNullOrEmpty(str_Sex) && !String.IsNullOrEmpty(str_gender))
                    //        //{
                    //        //    str_Sex = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                    //        //    _var.dtworkRow01[36] = str_Sex;
                    //        //    _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                    //        //}
                    //        //else if (String.IsNullOrEmpty(str_Sex) && String.IsNullOrEmpty(str_gender))
                    //        //{
                    //        //    _var.dtworkRow01[36] = string.Empty;
                    //        //}
                    //        #endregion
                    //    }
                    //    else
                    //    {

                    //        _var.dtworkRow01[76] = "BR6-1";
                    //        _var.dtworkRow01[30] = str_PolNum;
                    //        _var.dtworkRow01[31] = str_PolNum; //Fullname
                    //        _var.dtworkRow01[32] = "DummyLastName";
                    //        _var.dtworkRow01[33] = "DummyFirstName";
                    //        _var.dtworkRow01[34] = "DummyMiddleName";
                    //        _var.dtworkRow01[36] = "M";
                    //        _var.dtworkRow01[37] = "07/01/1900";
                    //        _var.dtworkRow01[83] = "N";
                    //        _var.boo_genderfail = false;
                    //        //str_OSA = "1";
                    //        //str_ISR = "1";
                    //    }

                    //    foundRows = dt_macro.Select("CN = " + "'" + str_CessionNo.ToString() + "' AND COVER7C = '1'");//life  and extra
                    //    if (foundRows.Length != 0)
                    //    {
                    //        str_OSA = foundRows[0][16].ToString();
                    //        str_ISR = foundRows[0][17].ToString();
                    //        str_Ret = foundRows[0][18].ToString();
                    //    }
                    //}

                    #endregion

                    string[] arr_PremDueDate;
                    arr_PremDueDate = str_PremDueDate.Split('/');

                    DateTime dt_PremiumDate = Convert.ToDateTime(str_PremDueDate);
                    DateTime dt_IssueDate = DateTime.Now;
                    try
                    {
                        dt_IssueDate = Convert.ToDateTime(strIssueDate);
                    }
                    catch { boo_invalidIssueDate = true; }


                    //Mortality
                    _var.dtworkRow01[39] = objHlpr.fn_getmortality(strMortality);
                    if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                    {
                        _var.dtworkRow01[39] = "STANDARD";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                    }
                    _var.dtworkRow01[20] = strIssueDate;
                    _var.dtworkRow01[22] = str_PremDueDate;

                    #region Old Code

                    //DateTime dt = DateTime.ParseExact(str_PremDueDate, "MM/d/yyyy", CultureInfo.InvariantCulture);
                    //DateTime dtRyear = dt.AddYears(1);
                    //string EffectiveDate = dtRyear.ToString("MM/d/yyyy");


                    //if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)//FY 
                    //{
                    //if (str_tcode == "TLAPSE" || (str_tcode == "TREINS") || (str_tcode == "TCONTER") || (str_tcode == "ADJUST") || (str_tcode == "TFULLPU") || (str_tcode == "TFULLSUR"))
                    //{

                    //    _var.dtworkRow01[22] = str_PremDueDate;//Trans Effective Date
                    //    _var.dtworkRow01[19] = str_PremDueDate;//Reinsurance Start Date
                    //    _var.dtworkRow01[20] = str_PremDueDate;//Policy Start Date
                    //    _var.dtworkRow01[62] = "4004"; //BE
                    //    _var.dtworkRow01[63] = dbl_PremiumLife; //BF
                    //    }   

                    ////}
                    //else if (str_tcode == "TRENEW")
                    ////{
                    //_var.dtworkRow01[22] = EffectiveDate;//Trans Effective Date
                    //_var.dtworkRow01[19] = EffectiveDate;//Reinsurance Start Date
                    //_var.dtworkRow01[20] = str_PremDueDate;//Policy Start Date
                    //_var.dtworkRow01[59] = dbl_PremiumLife;
                    //_var.dtworkRow01[58] = "4001";
                    //}
                    //else //RY
                    //{

                    //    _var.dtworkRow01[19] = _var.dtworkRow01[20];
                    //    _var.dtworkRow01[20] = str_PremDueDate; ;//policy Start date
                    //    _var.dtworkRow01[22] = str_PremDueDate;//effective date
                    //    _var.dtworkRow01[19] = str_PremDueDate; ;//reinsurance start date
                    //    _var.dtworkRow01[63] = dbl_PremiumLife;
                    //    _var.dtworkRow01[62] = "4004";
                    //}
                    #endregion


                    if (str_tcode == "TRENEW")
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        _var.dtworkRow01[58] = "4001";
                        _var.dtworkRow01[59] = dbl_PremiumLife;
                    }
                    else
                    {
                        if (boo_invalidIssueDate)
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[20];
                            _var.dtworkRow01[60] = "4002";
                            _var.dtworkRow01[61] = dbl_PremiumLife;
                        }
                        else
                        {
                            if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[20];
                                _var.dtworkRow01[60] = "4002";
                                _var.dtworkRow01[61] = dbl_PremiumLife;
                            }
                            else
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[22];
                                _var.dtworkRow01[62] = "4004";
                                _var.dtworkRow01[63] = dbl_PremiumLife;
                            }
                        }
                    }

                    _var.dtworkRow01[21] = str_tcode;
                    _var.dtworkRow01[25] = str_LE_OSA;
                    _var.dtworkRow01[27] = str_LE_ISR;
                    _var.dtworkRow01[77] = dbl_NAAR;
                    _var.dtworkRow01[28] = dbl_Ret;

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[77] = _var.dtworkRow01[27];
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-1BZ" : _var.dtworkRow01[76].ToString() + "|BR1-1BZ";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[75] = _var.dtworkRow01[25];
                        dbl_PremiumAmnt = Convert.ToDouble(_var.dtworkRow01[75]);
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-2BZ" : _var.dtworkRow01[76].ToString() + "|BR1-2BZ";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[77];
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-1AB" : _var.dtworkRow01[76].ToString() + "|BR2-1AB";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[25];
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-2AB" : _var.dtworkRow01[76].ToString() + "|BR2-2AB";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[27];
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-1Z" : _var.dtworkRow01[76].ToString() + "|BR3-1Z";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[77];
                        //_var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-2Z" : _var.dtworkRow01[76].ToString() + "|BR3-2Z";
                    }


                    objHlpr.fn_GetRemarksCodeBenlife123(strDOB, strFullName, str_PolNum, strSex, strMortality, dbl_ISR, dbl_NAAR, dbl_OSA, dbl_PremiumAmnt, out string strRemakrsCode);
                    _var.dtworkRow01[76] = strRemakrsCode + "|" + strRcDummyName;
               

                    if (!String.IsNullOrEmpty(str_PremiumExtra) && str_PremiumExtra.Trim() != "-")
                    {
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow02[5] = "EXTRA";

                        if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow02[59] = dbl_PremiumExtra;
                        }
                        else
                        {
                            if (boo_invalidIssueDate)
                            {
                                _var.dtworkRow02[61] = dbl_PremiumExtra;

                            }
                            else
                            {
                                if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                                {
                                    _var.dtworkRow02[61] = dbl_PremiumExtra;


                                }
                                else
                                {
                                    _var.dtworkRow02[63] = dbl_PremiumExtra;


                                }
                            }
                        }

                    }

                
                    if (strCessionCode == "T") { 
                        if ((!String.IsNullOrEmpty(str_PremiumLife) && str_PremiumLife.Trim() != "-") || ((String.IsNullOrEmpty(str_PremiumLife) || str_PremiumLife.Trim() == "-") && (String.IsNullOrEmpty(str_PremiumExtra) || str_PremiumExtra.Trim() == "-")))
                        {
                            _var.dbl_BF += decimal.Parse(
                               String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                               );
                            _var.dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                            _var.dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                            _var.dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                            _var.dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );


                            objdt_template.Rows.Add(_var.dtworkRow01);
                        }

                        if (_var.dtworkRow02 != null)
                        {
                            _var.dbl_BF += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                );
                            _var.dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                    );
                            _var.dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                    );
                            _var.dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                    );
                            _var.dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                    );

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }
                    else if (strCessionCode == "F") 
                    {
                        if ((!String.IsNullOrEmpty(str_PremiumLife) && str_PremiumLife.Trim() != "-") || ((String.IsNullOrEmpty(str_PremiumLife) || str_PremiumLife.Trim() == "-") && (String.IsNullOrEmpty(str_PremiumExtra) || str_PremiumExtra.Trim() == "-")))
                        {
                            _var.dbl_FBF += decimal.Parse(
                               String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                               );
                            _var.dbl_FBH += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                            _var.dbl_FBJ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                            _var.dbl_FBL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                            _var.dbl_FBZ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );


                            objdt_template.Rows.Add(_var.dtworkRow01);
                        }

                        if (_var.dtworkRow02 != null)
                        {
                            _var.dbl_FBF += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                );
                            _var.dbl_FBH += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                    );
                            _var.dbl_FBJ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                    );
                            _var.dbl_FBL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                    );
                            _var.dbl_FBZ += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                    );


                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }
                    }

                }


                _var.dtworkRow01 = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow01);


           
                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Treaty Total Premium:";
                    _var.dtworkRow01[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL; 
                    objdt_template.Rows.Add(_var.dtworkRow01);

                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Treaty Total Sum at Risk:";
                    _var.dtworkRow01[1] = _var.dbl_BZ;
                    objdt_template.Rows.Add(_var.dtworkRow01);
            

                    _var.dtworkRow01 = objdt_template.NewRow();
                    objdt_template.Rows.Add(_var.dtworkRow01);

            
                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Facultative Total Premium:";
                    _var.dtworkRow01[1] = _var.dbl_FBF + _var.dbl_FBH + _var.dbl_FBJ + _var.dbl_FBL; 
                    objdt_template.Rows.Add(_var.dtworkRow01);

                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Facultative Total Sum at Risk:";
                    _var.dtworkRow01[1] = _var.dbl_FBZ;
                    objdt_template.Rows.Add(_var.dtworkRow01);

              
                if (_var.boo_genderfail)
                {
                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Please check for blank genders";
                    objdt_template.Rows.Add(_var.dtworkRow01);
                }

                string despath = str_saved + @"\BM123-" + str_savef + ".xlsx";
                objHlpr.fn_savefile(objdt_template, despath);

                if (boo_open)
                {
                    objHlpr.fn_openfile(despath);
                }

                /////
                eapp.DisplayAlerts = false;
                wsraw = null;
                wbraw.SaveAs(str_raw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
                wbraw.Close();
                wbraw = null;
                eapp = null;
                ////
                _var.dtworkRow01 = null;
                _var.dtworkRow02 = null;
                _var.dtworkRow03 = null;
                _var.dtworkRow04 = null;
                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message + Environment.NewLine + " *****On excel row line: " + int_RowCnt + " *****";
            }
        }


    }
}