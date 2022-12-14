using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM031
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false, string str_macro = "")
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

            double dbl_premiumrate = 1;
            double dbl_otherrate = 0.9;
            double dbl_AARrate = 1;

            bool boo_isFacul = false;
            bool boo_isFaculR = false;

            bool boo_isSpecial = false;
            bool boo_isSpecialR = false;
            string str_PolNum = string.Empty;
            try
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    str_PolNum = wsraw.Cells[intLoop, 1].Text.ToString();

                    if (str_PolNum.ToUpper().Contains("FACUL") || str_PolNum.ToUpper().Contains("FACULTATIVE") || str_PolNum.ToUpper().Contains("FAC"))
                    { boo_isFacul = true; }

                    if (str_PolNum.ToUpper().Contains("SPECIAL"))
                    { boo_isSpecial = true; }

                    if (boo_isFacul && str_PolNum.ToUpper().Contains("REFUNDS AND ADJUSTMENTS"))
                    {
                        boo_isFaculR = true;
                    }
                    else if (boo_isSpecial && str_PolNum.ToUpper().Contains("REFUNDS AND ADJUSTMENTS"))
                    {
                        boo_isSpecialR = true;
                    }

                    if (!objHlpr.fn_policyNumChecker(str_PolNum, wsraw.Cells[intLoop, 2].Text.ToString(), wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString()))
                    {
                        continue;
                    }

                    string str_Fullname = "";
                    string str_DOB = "";
                    string str_age = "";
                    string str_Mortality = "";
                    string str_Sex = "";
                    string str_refunding = "";


                    string str_LifeOSA = "";
                    string str_liferet = "";

                    string str_WPDOSA = "";
                    string str_WPDret = "";
                    string str_WPDReinsured = "";

                    string str_ADBOSA = "";
                    string str_ADBret = "";
                    string str_ADBReinsured = "";

                    string str_SAR_SARDIOSA = "";
                    string str_SAR_SARDIret = "";
                    string str_SAR_SARDIReinsured = "";

                    string str_code = "SARDI";
                    //Macro
                    if (!String.IsNullOrEmpty(str_macro))
                    {
                        DataRow[] foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "'");
                        if (foundRows.Length != 0)
                        {
                            str_age = foundRows[0][5].ToString();
                            str_Mortality = foundRows[0][8].ToString();
                            str_refunding = foundRows[0][9].ToString();
                            str_Fullname = foundRows[0][10].ToString();
                            str_DOB = foundRows[0][12].ToString();
                            str_Sex = foundRows[0][13].ToString();
                        }

                        foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "' AND COVER7C = '1'");//life  and extra
                        if (foundRows.Length != 0)
                        {
                            str_LifeOSA = foundRows[0][16].ToString();//25
                            str_liferet = foundRows[0][18].ToString();//28
                        }

                        foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "' AND COVER7C = 'B'");//ADB
                        if (foundRows.Length != 0)
                        {
                            str_ADBOSA = foundRows[0][16].ToString();//25
                            str_ADBReinsured = foundRows[0][17].ToString();//27
                            str_ADBret = foundRows[0][18].ToString();//28
                        }

                        foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "' AND COVER7C = 'A'");//WPD
                        if (foundRows.Length != 0)
                        {
                            str_WPDOSA = foundRows[0][16].ToString();//25
                            str_WPDReinsured = foundRows[0][17].ToString();//27
                            str_WPDret = foundRows[0][18].ToString();//28
                        }

                        foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "' AND COVER7C = 'D'");//SARDI
                        if (foundRows.Length != 0)
                        {
                            str_SAR_SARDIOSA = foundRows[0][16].ToString();//25
                            str_SAR_SARDIReinsured = foundRows[0][17].ToString();//27
                            str_SAR_SARDIret = foundRows[0][18].ToString();//28
                        }
                        else 
                        {
                            foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "' AND COVER7C = 'C'");//SAR
                            if (foundRows.Length != 0)
                            {
                                str_code = "SAR";
                                str_SAR_SARDIOSA = foundRows[0][16].ToString();//25
                                str_SAR_SARDIReinsured = foundRows[0][17].ToString();//27
                                str_SAR_SARDIret = foundRows[0][18].ToString();//28
                            }
                        }
                    }

                    string str_LifeReinsured = wsraw.Cells[intLoop, 11].Text.ToString(); //27
                    string str_AAR = wsraw.Cells[intLoop, 12].Text.ToString(); //76

                    if (!double.TryParse(str_LifeOSA, out double dbl_LifeOSA))
                    {
                        dbl_LifeOSA = 1;
                    }
                    if (!double.TryParse(str_liferet, out double dbl_liferet))
                    {
                        dbl_liferet = 1;
                    }
                    if (!double.TryParse(str_ADBOSA, out double dbl_ADBOSA))
                    {
                        dbl_ADBOSA = 1;
                    }
                    if (!double.TryParse(str_ADBret, out double dbl_ADBret))
                    {
                        dbl_ADBret = 1;
                    }
                    if (!double.TryParse(str_ADBReinsured, out double dbl_ADBReinsured))
                    {
                        dbl_ADBReinsured = 1;
                    }
                    if (!double.TryParse(str_WPDOSA, out double dbl_WPDOSA))
                    {
                        dbl_WPDOSA = 1;
                    }
                    if (!double.TryParse(str_WPDret, out double dbl_WPDret))
                    {
                        dbl_WPDret = 1;
                    }
                    if (!double.TryParse(str_WPDReinsured, out double dbl_WPDReinsured))
                    {
                        dbl_WPDReinsured = 1;
                    }
                    if (!double.TryParse(str_SAR_SARDIOSA, out double dbl_SAR_SARDIOSA))
                    {
                        dbl_SAR_SARDIOSA = 1;
                    }
                    if (!double.TryParse(str_SAR_SARDIret, out double dbl_SAR_SARDIret))
                    {
                        dbl_SAR_SARDIret = 1;
                    }
                    if (!double.TryParse(str_SAR_SARDIReinsured, out double dbl_SAR_SARDIReinsured))
                    {
                        dbl_SAR_SARDIReinsured = 1;
                    }

                    if (!double.TryParse(objHlpr.fn_numbercleanup_negative(str_LifeReinsured), out double dbl_LifeReinsured))
                    {
                        dbl_LifeReinsured = 1;
                    }
                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_AAR), out double dbl_AAR))
                    {
                        dbl_AAR = dbl_AAR * (boo_isFacul || boo_isSpecial ? dbl_AARrate : 1); ;
                    }
                    else
                    {
                        dbl_AAR = 1;
                    }

                    string str_certnum = wsraw.Cells[intLoop, 2].Text.ToString();
                    string str_EffectiveDate = wsraw.Cells[intLoop, 3].Text.ToString();
                    string str_PremiumDate = wsraw.Cells[intLoop, 4].Text.ToString();

                    string str_bmyear = str_PremiumDate.Substring(str_PremiumDate.Length - 4, 4);
                    
                    string str_PremiumLife = wsraw.Cells[intLoop, 5].Text.ToString();
                    string str_PremiumExtra = wsraw.Cells[intLoop, 6].Text.ToString();
                    string str_PremiumADB = wsraw.Cells[intLoop, 7].Text.ToString();
                    string str_PremiumWPD = wsraw.Cells[intLoop, 8].Text.ToString();
                    string str_PremiumSAR = wsraw.Cells[intLoop, 9].Text.ToString();

                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_PremiumLife), out double dbl_PremiumLife))
                    {
                        dbl_PremiumLife = dbl_PremiumLife * dbl_premiumrate;
                    }
                    else
                    {
                        dbl_PremiumLife = 0;
                    }

                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_PremiumExtra), out double dbl_PremiumExtra))
                    {
                        dbl_PremiumExtra = dbl_PremiumExtra * dbl_otherrate;
                    }
                    else
                    {
                        dbl_PremiumExtra = 0;
                    }

                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_PremiumADB), out double dbl_PremiumADB))
                    {
                        dbl_PremiumADB = dbl_PremiumADB * dbl_otherrate;
                    }
                    else
                    {
                        dbl_PremiumADB = 0;
                    }

                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_PremiumWPD), out double dbl_PremiumWPD))
                    {
                        dbl_PremiumWPD = dbl_PremiumWPD * dbl_otherrate;
                    }
                    else
                    {
                        dbl_PremiumWPD = 0;
                    }

                    if (double.TryParse(objHlpr.fn_numbercleanup_negative(str_PremiumSAR), out double dbl_PremiumSAR))
                    {
                        dbl_PremiumSAR = dbl_PremiumSAR * dbl_otherrate;
                    }
                    else
                    {
                        dbl_PremiumSAR = 0;
                    }

                    DateTime dt_PremiumDate = Convert.ToDateTime(str_PremiumDate),
                        dt_IssueDate = Convert.ToDateTime(str_EffectiveDate);

                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow02 = null;
                    _var.dtworkRow03 = null;
                    _var.dtworkRow04 = null;
                    _var.dtworkRow05 = null;

                    string str_tcode = string.Empty;

                    if (str_sheet.ToUpper().Contains("RENEWAL") || (boo_isFacul & !boo_isFaculR) || (boo_isSpecial & !boo_isSpecialR))
                    {
                        str_tcode = "TRENEW";
                    }
                    else
                    {
                        if (str_AAR.ToUpper().Contains("S"))
                        {
                            str_tcode = "TFULLSUR";
                        }
                        else if (str_AAR.ToUpper().Contains("L"))
                        {
                            str_tcode = "TLAPSE";
                        }
                        else
                        {
                            str_tcode = "TADJUST";
                        }

                        if (!double.TryParse(objHlpr.fn_numbercleanup_negative(str_AAR), out double dbl_AAR_1))
                        {
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? str_AAR : _var.dtworkRow01[76].ToString() + "|" + str_AAR;
                        }
                    }
                            

                    _var.dtworkRow01[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow01[1] = str_certnum;
                    _var.dtworkRow01[5] = "LIFE";
                    _var.dtworkRow01[6] = "BP289";
                    _var.dtworkRow01[8] = "SURPLUS";
                    _var.dtworkRow01[9] = "PAFM";
                    _var.dtworkRow01[13] = "IND";
                    _var.dtworkRow01[10] = "S";
                    _var.dtworkRow01[14] = (boo_isFacul ? "F" : "T");
                    _var.dtworkRow01[23] = "PHP";
                    _var.dtworkRow01[24] = "YLY";
                    _var.dtworkRow01[29] = "NATREID";

                    //DOB
                    if (String.IsNullOrEmpty(str_DOB))
                    {
                        str_DOB = "07/01/1900";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                    }
                    _var.dtworkRow01[37] = str_DOB;

                    //Smoker
                    _var.dtworkRow01[38] = "NONE";

                    //Mortality
                    _var.dtworkRow01[39] = objHlpr.fn_getmortality(str_Mortality);
                    if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                    {
                        _var.dtworkRow01[39] = "STANDARD";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                    }

                    _var.dtworkRow01[41] = str_bmyear;
                    _var.dtworkRow01[79] = str_age;

                    _var.dtworkRow01[20] = str_EffectiveDate;
                    _var.dtworkRow01[22] = str_PremiumDate;

                    

                    if (str_tcode == "TRENEW")
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        _var.dtworkRow01[58] = "4001";
                        _var.dtworkRow01[59] = dbl_PremiumLife;
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

                    _var.dtworkRow01[21] = str_tcode;

                    _var.dtworkRow01[25] = dbl_LifeOSA;
                    _var.dtworkRow01[27] = dbl_LifeReinsured;
                    _var.dtworkRow01[77] = dbl_AAR;
                    _var.dtworkRow01[28] = dbl_liferet;

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[77] = _var.dtworkRow01[27];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-1BZ" : _var.dtworkRow01[76].ToString() + "|BR1-1BZ";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[75] = _var.dtworkRow01[25];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-2BZ" : _var.dtworkRow01[76].ToString() + "|BR1-2BZ";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[77];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-1AB" : _var.dtworkRow01[76].ToString() + "|BR2-1AB";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[25];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-2AB" : _var.dtworkRow01[76].ToString() + "|BR2-2AB";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[27];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-1Z" : _var.dtworkRow01[76].ToString() + "|BR3-1Z";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[77];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-2Z" : _var.dtworkRow01[76].ToString() + "|BR3-2Z";
                    }

                    //Name
                    if (String.IsNullOrEmpty(str_Fullname))
                    {
                        str_Fullname = str_PolNum;
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                    }

                    objHlpr.fn_getnamesandlifeID(str_Fullname, str_DOB, out string str_outfname, out string str_outlname, out string str_outlifeid, "021");
                    _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(str_Fullname);
                    string str_MI = objHlpr.fn_getMI(str_outfname);
                    _var.dtworkRow01[32] = str_outlname.Trim();
                    if (str_MI.Trim() != string.Empty)
                    {
                        _var.dtworkRow01[33] = str_outfname.Trim().Replace(" " + str_MI.Trim(), "");
                        _var.dtworkRow01[34] = str_MI.Trim();
                    }
                    else
                    {
                        string[] arr_fname = str_outfname.Split(' ');
                        _var.dtworkRow01[33] = str_outfname.Trim().Replace(" " + arr_fname[arr_fname.Length -1], "");
                        _var.dtworkRow01[34] = arr_fname[arr_fname.Length - 1];
                    }
                    _var.dtworkRow01[30] = str_outlifeid;

                    //Gender
                    if (!String.IsNullOrEmpty(str_Sex))
                    {
                        _var.dtworkRow01[36] = (str_Sex.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && !String.IsNullOrEmpty(str_gender))
                    {
                        str_Sex = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                        _var.dtworkRow01[36] = str_Sex;
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && String.IsNullOrEmpty(str_gender))
                    {
                        _var.dtworkRow01[36] = string.Empty;
                    }

                    if (String.IsNullOrEmpty(_var.dtworkRow01[36].ToString()))
                    {
                        _var.boo_genderfail = true;
                    }

                    _var.dtworkRow01[83] = str_refunding;

                    //Group Accounts
                    DateTime parsedDOB = Convert.ToDateTime(str_DOB);
                    string initialNR = string.Empty;
                    if (!String.IsNullOrEmpty(_var.str_outfname))
                    {
                        initialNR = _var.str_outfname.Substring(0, 1);
                    }
                    if (!String.IsNullOrEmpty(_var.str_outlname))
                    {
                        initialNR += _var.str_outlname.Substring(0, 1);
                    }

                    if (_var.dtworkRow01[13].ToString() == "GRP" || _var.dtworkRow01[13].ToString() == "GCL" || _var.dtworkRow01[13].ToString() == "GEB")
                    {
                        if (_var.dtworkRow01[0].ToString().Length >= 7)
                        {
                            _var.dtworkRow01[0] = _var.dtworkRow01[0].ToString().Substring(_var.dtworkRow01[0].ToString().Length - 7, 7) +
                                initialNR +
                                parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                        }
                        else
                        {
                            _var.dtworkRow01[0] = _var.dtworkRow01[0].ToString() +
                                initialNR +
                                parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                        }
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR5-1A" : _var.dtworkRow01[76].ToString() + "|BR5-1A";


                        if (!string.IsNullOrEmpty(str_Sex))
                        {
                            _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + str_Sex.Substring(0, 1);
                        }
                        else
                        {
                            _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + "-";
                        }

                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR5-2B" : _var.dtworkRow01[76].ToString() + "|BR5-2B";

                        _var.dtworkRow01[7] = str_PolNum;
                    }
                    else
                    {
                        _var.dtworkRow01[7] = string.Empty;
                    }

                    if (!String.IsNullOrEmpty(str_PremiumExtra) && str_PremiumExtra.Trim() != "-")
                    {
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow02[5] = "EXTRA";
                        _var.dtworkRow02[6] = "BP293";

                        if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow02[59] = dbl_PremiumExtra;
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

                    if (!String.IsNullOrEmpty(str_PremiumADB) && str_PremiumADB.Trim() != "-")
                    {
                        _var.dtworkRow03 = objdt_template.NewRow();
                        _var.dtworkRow03.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow03[5] = "ADB";
                        _var.dtworkRow03[6] = "BP291";

                        _var.dtworkRow03[25] = dbl_ADBOSA;
                        _var.dtworkRow03[27] = dbl_ADBReinsured;
                        _var.dtworkRow03[77] = 1;
                        _var.dtworkRow03[28] = dbl_ADBret;

                        if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow03[59] = dbl_PremiumADB;
                        }
                        else
                        {
                            if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                            {
                                _var.dtworkRow03[61] = dbl_PremiumADB;
                            }
                            else
                            {
                                _var.dtworkRow03[63] = dbl_PremiumADB;
                            }
                        }
                    }

                    if (!String.IsNullOrEmpty(str_PremiumWPD) && str_PremiumWPD.Trim() != "-")
                    {
                        _var.dtworkRow04 = objdt_template.NewRow();
                        _var.dtworkRow04.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow04[5] = "WPD";
                        _var.dtworkRow04[6] = "BP292";

                        _var.dtworkRow04[25] = dbl_WPDOSA;
                        _var.dtworkRow04[27] = dbl_WPDReinsured;
                        _var.dtworkRow04[77] = 1;
                        _var.dtworkRow04[28] = dbl_WPDret;

                        if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow04[59] = dbl_PremiumWPD;
                        }
                        else
                        {
                            if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                            {
                                _var.dtworkRow04[61] = dbl_PremiumWPD;
                            }
                            else
                            {
                                _var.dtworkRow04[63] = dbl_PremiumWPD;
                            }
                        }
                    }

                    if (!String.IsNullOrEmpty(str_PremiumSAR) && str_PremiumSAR.Trim() != "-")
                    {
                        _var.dtworkRow05 = objdt_template.NewRow();
                        _var.dtworkRow05.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow05[5] = str_code;
                        _var.dtworkRow05[6] = "BP290";

                        _var.dtworkRow05[25] = dbl_SAR_SARDIOSA;
                        _var.dtworkRow05[27] = dbl_SAR_SARDIReinsured;
                        _var.dtworkRow05[77] = 1;
                        _var.dtworkRow05[28] = dbl_SAR_SARDIret;

                        if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow05[59] = dbl_PremiumSAR;
                        }
                        else
                        {
                            if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                            {
                                _var.dtworkRow05[61] = dbl_PremiumSAR;
                            }
                            else
                            {
                                _var.dtworkRow05[63] = dbl_PremiumSAR;
                            }
                        }
                    }

                    if ((!String.IsNullOrEmpty(str_PremiumLife) && str_PremiumLife.Trim() != "-")
                         ||
                         (
                         (String.IsNullOrEmpty(str_PremiumLife) || str_PremiumLife.Trim() == "-") &&
                         (String.IsNullOrEmpty(str_PremiumExtra) || str_PremiumExtra.Trim() == "-") &&
                         (String.IsNullOrEmpty(str_PremiumADB) || str_PremiumADB.Trim() == "-") &&
                         (String.IsNullOrEmpty(str_PremiumWPD) || str_PremiumWPD.Trim() == "-") &&
                         (String.IsNullOrEmpty(str_PremiumSAR) || str_PremiumSAR.Trim() == "-")
                         )
                         )
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

                    if (_var.dtworkRow03 != null)
                    {
                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                            );
                        _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                );
                        _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                );
                        _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                );
                        _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                );

                        objdt_template.Rows.Add(_var.dtworkRow03);
                    }

                    if (_var.dtworkRow04 != null)
                    {
                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow04[57].ToString()) ? "0" : _var.dtworkRow04[57].ToString()
                            );
                        _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[59].ToString()) ? "0" : _var.dtworkRow04[59].ToString()
                                );
                        _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[61].ToString()) ? "0" : _var.dtworkRow04[61].ToString()
                                );
                        _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[63].ToString()) ? "0" : _var.dtworkRow04[63].ToString()
                                );
                        _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[77].ToString()) ? "0" : _var.dtworkRow04[77].ToString()
                                );

                        objdt_template.Rows.Add(_var.dtworkRow04);
                    }

                    if (_var.dtworkRow05 != null)
                    {
                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow05[57].ToString()) ? "0" : _var.dtworkRow05[57].ToString()
                            );
                        _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[59].ToString()) ? "0" : _var.dtworkRow05[59].ToString()
                                );
                        _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[61].ToString()) ? "0" : _var.dtworkRow05[61].ToString()
                                );
                        _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[63].ToString()) ? "0" : _var.dtworkRow05[63].ToString()
                                );
                        _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[77].ToString()) ? "0" : _var.dtworkRow05[77].ToString()
                                );

                        objdt_template.Rows.Add(_var.dtworkRow05);
                    }
                }

                _var.dtworkRow01 = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium:";
                _var.dtworkRow01[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk:";
                _var.dtworkRow01[1] = _var.dbl_BZ;
                objdt_template.Rows.Add(_var.dtworkRow01);

                if (_var.boo_genderfail)
                {
                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Please check for blank genders";
                    objdt_template.Rows.Add(_var.dtworkRow01);
                }

                string despath = str_saved + @"\BM031-" + str_savef + ".xlsx";
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
                _var.dtworkRow05 = null;
                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message + Environment.NewLine + " *****On excel row line: " + int_RowCnt + " *****" + Environment.NewLine + str_PolNum;
            }
        }


    }
}