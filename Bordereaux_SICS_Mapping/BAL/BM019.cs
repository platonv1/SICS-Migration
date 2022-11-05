using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM019
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            DataTable objdt_template = new DataTable();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);

            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int erawrow = rawrange.Rows.Count;
            int int_RowCnt = 0;

            _var.objdt_template01 = objHlpr.dt_formtemplate("JAN");
            _var.objdt_template02 = objHlpr.dt_formtemplate("FEB");
            _var.objdt_template03 = objHlpr.dt_formtemplate("MAR");
            _var.objdt_template04 = objHlpr.dt_formtemplate("APR");
            _var.objdt_template05 = objHlpr.dt_formtemplate("MAY");
            _var.objdt_template06 = objHlpr.dt_formtemplate("JUN");
            _var.objdt_template07 = objHlpr.dt_formtemplate("JUL");
            _var.objdt_template08 = objHlpr.dt_formtemplate("AUG");
            _var.objdt_template09 = objHlpr.dt_formtemplate("SEP");
            _var.objdt_template10 = objHlpr.dt_formtemplate("OCT");
            _var.objdt_template11 = objHlpr.dt_formtemplate("NOV");
            _var.objdt_template12 = objHlpr.dt_formtemplate("DEC");

            string str_bmyear = wsraw.Cells[4, 2].Text.ToString().Substring(wsraw.Cells[4, 2].Text.ToString().Length - 4);
            double dbl_rate = 0.15;
   

            //if (str_raw.ToUpper().Contains("FACULTATIVE"))
            //{
            //    dbl_rate = 1.00;
            //}
            
            int int_columnvariation = 0;
            if (boo_clean)
            {
                wsraw = objHlpr.fn_extendwidth(wsraw);
            }

            try 
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    int_RowCnt = intLoop;
                    string str_PolNum = wsraw.Cells[intLoop, 2].Text.ToString();


                    if (!objHlpr.fn_policyNumChecker(str_PolNum, wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString(), wsraw.Cells[intLoop, 5].Text.ToString()))
                    {
                        continue;
                    }

                    string str_Plan = wsraw.Cells[intLoop, 3].Text.ToString();
                    string str_Rtype = wsraw.Cells[intLoop, 4 + int_columnvariation].Text.ToString();
                    string str_IssueDate = wsraw.Cells[intLoop, 5 + int_columnvariation].Text.ToString();
                    string str_Sex = wsraw.Cells[intLoop, 7 + int_columnvariation].Text.ToString();
                    string str_Fullname = wsraw.Cells[intLoop, 8 + int_columnvariation].Text.ToString();
                    string str_Risk = wsraw.Cells[intLoop, 9 + int_columnvariation].Text.ToString();
                    string str_DOB = wsraw.Cells[intLoop, 11 + int_columnvariation].Text.ToString();
                    string str_age = wsraw.Cells[intLoop, 12 + int_columnvariation].Text.ToString();
                    string str_Mortality = wsraw.Cells[intLoop, 13 + int_columnvariation].Text.ToString();

                    string str_OSILife = wsraw.Cells[intLoop, 15 + int_columnvariation].Text.ToString();//COL 16
                    string str_OSIWP = wsraw.Cells[intLoop, 16 + int_columnvariation].Text.ToString();//COL 17
                    string str_OSIADB = wsraw.Cells[intLoop, 17 + int_columnvariation].Text.ToString();//COL 18
                    string str_OSIRider = wsraw.Cells[intLoop, 18 + int_columnvariation].Text.ToString();//COL 19
                    string str_RetentionLife = wsraw.Cells[intLoop, 19 + int_columnvariation].Text.ToString();
                    string str_RetentionWP = wsraw.Cells[intLoop, 20 + int_columnvariation].Text.ToString();
                    string str_RetentionADB = wsraw.Cells[intLoop, 21 + int_columnvariation].Text.ToString();
                    string str_RetentionRider = wsraw.Cells[intLoop, 22 + int_columnvariation].Text.ToString();
                    string str_ISRLife = wsraw.Cells[intLoop, 23 + int_columnvariation].Text.ToString(); //COL 24
                    string str_ISRWP = wsraw.Cells[intLoop, 24 + int_columnvariation].Text.ToString(); //COL 25
                    string str_ISRADB = wsraw.Cells[intLoop, 25 + int_columnvariation].Text.ToString();//COl 26
                    string str_ISRRider = wsraw.Cells[intLoop, 26 + int_columnvariation].Text.ToString();//COL 27
                    string str_Premium = wsraw.Cells[intLoop, 27 + int_columnvariation].Text.ToString(); 

                    
                    if (double.TryParse(str_OSILife, out double dbl_OSILife))
                    {
                        dbl_OSILife = dbl_OSILife * 1;
                    }
                    else 
                    {
                        dbl_OSILife = 1;
                    }
                    if (double.TryParse(str_OSIWP, out double dbl_OSIWP))
                    {
                        dbl_OSIWP = dbl_OSIWP * 1;
                    }
                    else
                    {
                        dbl_OSIWP = 1;
                    }
                    if (double.TryParse(str_OSIADB, out double dbl_OSIADB))
                    {
                        dbl_OSIADB = dbl_OSIADB * 1;
                    }
                    else
                    {
                        dbl_OSIADB = 1;
                    }
                    if (double.TryParse(str_OSIRider, out double dbl_OSIRider))
                    {
                        dbl_OSIRider = dbl_OSIRider * 1;
                    }
                    else
                    {
                        dbl_OSIRider = 1;
                    }
                    if (double.TryParse(str_RetentionLife, out double dbl_RetentionLife))
                    {
                        dbl_RetentionLife = dbl_RetentionLife * 1;
                    }
                    else
                    {
                        dbl_RetentionLife = 1;
                    }
                    if (double.TryParse(str_RetentionWP, out double dbl_RetentionWP))
                    {
                        dbl_RetentionWP = dbl_RetentionWP * 1;
                    }
                    else
                    {
                        dbl_RetentionWP = 1;
                    }
                    if (double.TryParse(str_RetentionADB, out double dbl_RetentionADB))
                    {
                        dbl_RetentionADB = dbl_RetentionADB * 1;
                    }
                    else
                    {
                        dbl_RetentionADB = 1;
                    }
                    if (double.TryParse(str_RetentionRider, out double dbl_RetentionRider))
                    {
                        dbl_RetentionRider = dbl_RetentionRider * 1;
                    }
                    else
                    {   
                        dbl_RetentionRider = 1;
                    }
                    if (double.TryParse(str_ISRLife, out double dbl_ISRLife)) //COL 24
                    {
                        dbl_ISRLife = dbl_ISRLife * dbl_rate;
                    }
                    else
                    {
                        dbl_ISRLife = 1;
                    }
                    if (double.TryParse(str_ISRWP, out double dbl_ISRWP)) //COL 25
                    {
                        dbl_ISRWP = dbl_ISRWP * dbl_rate;
                    }
                    else
                    {
                        dbl_ISRWP = 1;
                    }
                    if (double.TryParse(str_ISRADB, out double dbl_ISRADB)) //COL 26
                    {
                        dbl_ISRADB = dbl_ISRADB * dbl_rate;
                    }
                    else
                    {
                        dbl_ISRADB = 1;
                    }
                    if (double.TryParse(str_ISRRider, out double dbl_ISRRider)) //COL 27
                    {
                        dbl_ISRRider = dbl_ISRRider * dbl_rate;
                    }
                    else
                    {
                        dbl_ISRRider = 1;
                    }
                    if (double.TryParse(str_Premium, out double dbl_Premium))
                    {
                        dbl_Premium = dbl_Premium * dbl_rate;
                    }
                    else
                    {
                        dbl_Premium = 0;
                    }

                    string str_PremiumDate = wsraw.Cells[intLoop, 28].Text.ToString();

                    bool boo_1styr = false;
                    DateTime dt_PremiumDate = Convert.ToDateTime(str_PremiumDate), dt_IssueDate = Convert.ToDateTime(str_IssueDate);
                    DateTime dt_sheetmonth = Convert.ToDateTime(str_PremiumDate);

                    switch (dt_sheetmonth.Month)
                    {
                        case 1:
                            _var.dtworkRow = _var.objdt_template01.NewRow();
                            break;
                        case 2:
                            _var.dtworkRow = _var.objdt_template02.NewRow();
                            break;
                        case 3:
                            _var.dtworkRow = _var.objdt_template03.NewRow();
                            break;
                        case 4:
                            _var.dtworkRow = _var.objdt_template04.NewRow();
                            break;
                        case 5:
                            _var.dtworkRow = _var.objdt_template05.NewRow();
                            break;
                        case 6:
                            _var.dtworkRow = _var.objdt_template06.NewRow();
                            break;
                        case 7:
                            _var.dtworkRow = _var.objdt_template07.NewRow();
                            break;
                        case 8:
                            _var.dtworkRow = _var.objdt_template08.NewRow();
                            break;
                        case 9:
                            _var.dtworkRow = _var.objdt_template09.NewRow();
                            break;
                        case 10:
                            _var.dtworkRow = _var.objdt_template10.NewRow();
                            break;
                        case 11:
                            _var.dtworkRow = _var.objdt_template11.NewRow();
                            break;
                        case 12:
                            _var.dtworkRow = _var.objdt_template12.NewRow();
                            break;
                        default:
                            break;
                    }

                    _var.dtworkRow[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow[5] = str_Plan;
                    _var.dtworkRow[8] = "SURPLUS";
                    _var.dtworkRow[9] = "PAFW";
                    _var.dtworkRow[13] = "IND";
                    _var.dtworkRow[10] = "S";
                    
                    
                    if (str_Rtype.ToUpper().Contains("FACUL"))
                    {
                        _var.dtworkRow[14] = "F";
                    }
                    else
                    {
                        _var.dtworkRow[14] = "T";
                    }

                    if (str_raw.ToUpper().Contains("DOLLAR") || str_raw.ToUpper().Contains("USD") || str_raw.ToUpper().Contains("$"))
                    {
                        _var.dtworkRow[23] = "USD";
                    }
                    else
                    {
                        _var.dtworkRow[23] = "PHP";
                    }

                    _var.dtworkRow[24] = "MLY";
                    _var.dtworkRow[29] = "NATREID";

                    //DOB
                    if (String.IsNullOrEmpty(str_DOB))
                    {
                        str_DOB = "7/1/1900";
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR4AL" : _var.dtworkRow[76].ToString() + "|BR4AL";
                    }
                    _var.dtworkRow[37] = str_DOB;

                    //Smoker
                    _var.dtworkRow[38] = objHlpr.fn_smokercode(str_Risk, "019");

                    //Mortality
                    _var.dtworkRow[39] = objHlpr.fn_getmortality(str_Mortality);
                    if (objHlpr.fn_isDMort(_var.dtworkRow[39].ToString()))
                    {
                        _var.dtworkRow[39] = "STANDARD";
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR8AN" : _var.dtworkRow[76].ToString() + "|BR8AN";
                    }
                    
                    _var.dtworkRow[41] = str_bmyear;
                    _var.dtworkRow[79] = str_age;

                    _var.dtworkRow[20] = str_IssueDate;
                    _var.dtworkRow [22] = str_PremiumDate; /* str_IssueDate.Substring(0, str_IssueDate.Length - 4) + str_bmyear;*/

                    if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                    {
                        boo_1styr = true;
                        _var.dtworkRow[19] = _var.dtworkRow[20];
                    }
                    else 
                    {
                        boo_1styr = false;
                        _var.dtworkRow [19] = str_PremiumDate;/*_var.dtworkRow[22];*/
                    }

                    string str_tcode = string.Empty;
                    if (str_OSILife.ToUpper().Contains("TERMINATED") ||
                        str_OSIWP.ToUpper().Contains("TERMINATED") ||
                        str_OSIADB.ToUpper().Contains("TERMINATED") ||
                        str_OSIRider.ToUpper().Contains("TERMINATED") ||
                        str_RetentionLife.ToUpper().Contains("TERMINATED") ||
                        str_RetentionWP.ToUpper().Contains("TERMINATED") ||
                        str_RetentionADB.ToUpper().Contains("TERMINATED") ||
                        str_RetentionRider.ToUpper().Contains("TERMINATED") ||
                        str_ISRLife.ToUpper().Contains("TERMINATED") ||
                        str_ISRWP.ToUpper().Contains("TERMINATED") ||
                        str_ISRADB.ToUpper().Contains("TERMINATED") ||
                        str_ISRRider.ToUpper().Contains("TERMINATED") ||
                        str_Premium.ToUpper().Contains("TERMINATED"))
                    {
                        str_tcode = "TCONTER";
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "TERMINATED" : _var.dtworkRow[76].ToString() + "|TERMINATED";
                    }
                    else if (str_OSILife.ToUpper().Contains("LAPSE") ||
                        str_OSIWP.ToUpper().Contains("LAPSE") ||
                        str_OSIADB.ToUpper().Contains("LAPSE") ||
                        str_OSIRider.ToUpper().Contains("LAPSE") ||
                        str_RetentionLife.ToUpper().Contains("LAPSE") ||
                        str_RetentionWP.ToUpper().Contains("LAPSE") ||
                        str_RetentionADB.ToUpper().Contains("LAPSE") ||
                        str_RetentionRider.ToUpper().Contains("LAPSE") ||
                        str_ISRLife.ToUpper().Contains("LAPSE") ||
                        str_ISRWP.ToUpper().Contains("LAPSE") ||
                        str_ISRADB.ToUpper().Contains("LAPSE") ||
                        str_ISRRider.ToUpper().Contains("LAPSE") ||
                        str_Premium.ToUpper().Contains("LAPSE"))
                    {
                        str_tcode = "TLAPSE";
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "LAPSE" : _var.dtworkRow[76].ToString() + "|LAPSE";
                    }


                    if (str_tcode == "TCONTER" && boo_1styr)
                    {
                        _var.dtworkRow[60] = "4002";
                        _var.dtworkRow[61] = dbl_Premium;
                    }
                    else if (str_tcode == "TCONTER" && !boo_1styr)
                    {
                        _var.dtworkRow[62] = "4004";
                        _var.dtworkRow[63] = dbl_Premium;
                    }
                    else if (str_tcode == "TLAPSE" && boo_1styr)
                    {
                        _var.dtworkRow[60] = "4002";
                        _var.dtworkRow[61] = dbl_Premium;
                    }
                    else if (str_tcode == "TLAPSE" && !boo_1styr)
                    {
                        _var.dtworkRow[62] = "4004";
                        _var.dtworkRow[63] = dbl_Premium;
                    }
                    else if (str_tcode == string.Empty && boo_1styr)
                    {
                        str_tcode = "TNEWBUS";
                        _var.dtworkRow[56] = "4000";
                        _var.dtworkRow[57] = dbl_Premium;
                    }
                    else if (str_tcode == string.Empty && !boo_1styr)
                    {
                        str_tcode = "TRENEW";
                        _var.dtworkRow[58] = "4001";
                        _var.dtworkRow[59] = dbl_Premium;
                    }

                    _var.dtworkRow[21] = str_tcode;

                    if (str_OSILife.Trim() != string.Empty)
                    {
                        _var.dtworkRow[25] = dbl_OSILife.Equals(0.00) ? 1 : dbl_OSILife;
                        _var.dtworkRow[27] = dbl_ISRLife.Equals(0.00) ? 1 : dbl_ISRLife;
                        _var.dtworkRow[28] = dbl_RetentionLife.Equals(0.00) ? 1 : dbl_RetentionLife;
                    }
                    else if (str_OSIWP.Trim() != string.Empty)
                    {
                        _var.dtworkRow[25] = dbl_OSIWP.Equals(0.00) ? 1 : dbl_OSIWP;
                        _var.dtworkRow[27] = dbl_ISRWP.Equals(0.00) ? 1 : dbl_ISRWP;
                        _var.dtworkRow[28] = dbl_RetentionWP.Equals(0.00) ? 1 : dbl_RetentionWP;
                    }
                    else if (str_OSIADB.Trim() != string.Empty)
                    {
                        _var.dtworkRow[25] = dbl_OSIADB.Equals(0.00) ? 1 : dbl_OSIADB;
                        _var.dtworkRow[27] = dbl_ISRADB.Equals(0.00) ? 1 : dbl_ISRADB;
                        _var.dtworkRow[28] = dbl_RetentionADB.Equals(0.00) ? 1 : dbl_RetentionADB;
                    }
                    else if (str_OSIRider.Trim() != string.Empty)
                    {
                        _var.dtworkRow[25] = dbl_OSIRider.Equals(0.00) ? 1 : dbl_OSIRider;
                        _var.dtworkRow[27] = dbl_ISRRider.Equals(0.00) ? 1 : dbl_ISRRider;
                        _var.dtworkRow[28] = dbl_RetentionRider.Equals(0.00) ? 1 : dbl_RetentionRider;
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()))
                    {
                        _var.dtworkRow[77] = _var.dtworkRow[27];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR1-1BZ" : _var.dtworkRow[76].ToString() + "|BR1-1BZ";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow[77].ToString()))
                    {
                        _var.dtworkRow[75] = _var.dtworkRow[25];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR1-2BZ" : _var.dtworkRow[76].ToString() + "|BR1-2BZ";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow[27].ToString()))
                    {
                        _var.dtworkRow[27] = _var.dtworkRow[77];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR2-1AB" : _var.dtworkRow[76].ToString() + "|BR2-1AB";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow[27].ToString()))
                    {
                        _var.dtworkRow[27] = _var.dtworkRow[25];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR2-2AB" : _var.dtworkRow[76].ToString() + "|BR2-2AB";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow[27].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow[25].ToString()))
                    {
                        _var.dtworkRow[25] = _var.dtworkRow[27];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR3-1Z" : _var.dtworkRow[76].ToString() + "|BR3-1Z";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow[25].ToString()))
                    {
                        _var.dtworkRow[25] = _var.dtworkRow[77];
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR3-2Z" : _var.dtworkRow[76].ToString() + "|BR3-2Z";
                    }

                    //Name
                    if (String.IsNullOrEmpty(str_Fullname))
                    {
                        str_Fullname = str_PolNum;
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR6AF" : _var.dtworkRow[76].ToString() + "|BR6AF";
                    }

                    objHlpr.fn_getnamesandlifeID(str_Fullname, str_DOB, out string str_outfname, out string str_outlname, out string str_outlifeid, "019");

                    _var.dtworkRow[31] = objHlpr.fn_stringcleanup(str_Fullname);
                    
                    string str_fullnoSuffix = str_Fullname.ToUpper().Replace("JR", "").Replace("SR", "").Replace(".", "");
                    string str_FnoSuffix = str_outfname.Replace("JR", "").Replace("SR", "");

                    string str_MI = str_fullnoSuffix.ToUpper().Replace(str_FnoSuffix, "").Replace(str_outlname, "").Trim();
                    _var.dtworkRow[32] = str_outlname.Trim();
                    _var.dtworkRow[33] = str_outfname.Trim();
                    _var.dtworkRow[34] = str_MI.Trim();
                    _var.dtworkRow[30] = str_outlifeid;

                    //Gender
                    if (!String.IsNullOrEmpty(str_Sex))
                    {
                        _var.dtworkRow[36] = (str_Sex.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && !String.IsNullOrEmpty(str_gender))
                    {
                        str_Sex = objHlpr.fn_getgender(str_gender, _var.dtworkRow[33].ToString());
                        _var.dtworkRow[36] = str_Sex;
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR7AK" : _var.dtworkRow[76].ToString() + "|BR7AK";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && String.IsNullOrEmpty(str_gender))
                    {
                        _var.dtworkRow[36] = string.Empty;
                    }

                    if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                    {
                        _var.boo_genderfail = true;
                    }

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

                    if (_var.dtworkRow[13].ToString() == "GRP" || _var.dtworkRow[13].ToString() == "GCL" || _var.dtworkRow[13].ToString() == "GEB")
                    {
                        if (_var.dtworkRow[0].ToString().Length >= 7)
                        {
                            _var.dtworkRow[0] = _var.dtworkRow[0].ToString().Substring(_var.dtworkRow[0].ToString().Length - 7, 7) +
                                initialNR +
                                parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                        }
                        else
                        {
                            _var.dtworkRow[0] = _var.dtworkRow[0].ToString() +
                                initialNR +
                                parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                        }
                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR5-1A" : _var.dtworkRow[76].ToString() + "|BR5-1A";


                        if (!string.IsNullOrEmpty(str_Sex))
                        {
                            _var.dtworkRow[1] = _var.dtworkRow[0].ToString() + str_Sex.Substring(0, 1);
                        }
                        else
                        {
                            _var.dtworkRow[1] = _var.dtworkRow[0].ToString() + "-";
                        }

                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR5-2B" : _var.dtworkRow[76].ToString() + "|BR5-2B";

                        _var.dtworkRow[7] = str_PolNum;
                    }
                    else
                    {
                        _var.dtworkRow[1] = string.Empty;
                        _var.dtworkRow[7] = string.Empty;
                    }

                    switch (dt_sheetmonth.Month)
                    {
                        case 1:
                            _var.objdt_template01.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF01 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH01 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ01 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL01 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ01 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );
                            break;
                        case 2:
                            _var.objdt_template02.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF02 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH02 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ02 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL02 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ02 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 3:
                            _var.objdt_template03.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF03 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH03 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ03 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL03 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ03 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 4:
                            _var.objdt_template04.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF04 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH04 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ04 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL04 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ04 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 5:
                            _var.objdt_template05.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF05 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH05 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ05 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL05 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ05 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 6:
                            _var.objdt_template06.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF06 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH06 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ06 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL06 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ06 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 7:
                            _var.objdt_template07.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF07 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH07 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ07 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL07 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ07 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 8:
                            _var.objdt_template08.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF08 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH08 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ08 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL08 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ08 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 9:
                            _var.objdt_template09.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF09 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH09 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ09 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL09 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ09 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 10:
                            _var.objdt_template10.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF10 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH10 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ10 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL10 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ10 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 11:
                            _var.objdt_template11.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF11 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH11 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ11 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL11 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ11 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        case 12:
                            _var.objdt_template12.Rows.Add(_var.dtworkRow);

                            _var.dbl_BF12 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                                );
                            _var.dbl_BH12 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                                );
                            _var.dbl_BJ12 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                                );
                            _var.dbl_BL12 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                                );
                            _var.dbl_BZ12 += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                                );

                            break;
                        default:
                            break;
                    }
                }

                if (_var.objdt_template01.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template01.NewRow();
                    _var.objdt_template01.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template01.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF01 + _var.dbl_BH01 + _var.dbl_BJ01 + _var.dbl_BL01;
                    _var.objdt_template01.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template01.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ01;
                    _var.objdt_template01.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template01, boo_open, str_saved + @"\BM019-JAN-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template02.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template02.NewRow();
                    _var.objdt_template02.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template02.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF02 + _var.dbl_BH02 + _var.dbl_BJ02 + _var.dbl_BL02;
                    _var.objdt_template02.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template02.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ02;
                    _var.objdt_template02.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template02, boo_open, str_saved + @"\BM019-FEB-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template03.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template03.NewRow();
                    _var.objdt_template03.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template03.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF03 + _var.dbl_BH03 + _var.dbl_BJ03 + _var.dbl_BL03;
                    _var.objdt_template03.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template03.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ03;
                    _var.objdt_template03.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template03, boo_open, str_saved + @"\BM019-MAR-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template04.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template04.NewRow();
                    _var.objdt_template04.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template04.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF04 + _var.dbl_BH04 + _var.dbl_BJ04 + _var.dbl_BL04;
                    _var.objdt_template04.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template04.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ04;
                    _var.objdt_template04.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template04, boo_open, str_saved + @"\BM019-APR-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template05.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template05.NewRow();
                    _var.objdt_template05.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template05.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF05 + _var.dbl_BH05 + _var.dbl_BJ05 + _var.dbl_BL05;
                    _var.objdt_template05.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template05.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ05;
                    _var.objdt_template05.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template05, boo_open, str_saved + @"\BM019-MAY-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template06.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template06.NewRow();
                    _var.objdt_template06.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template06.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF06 + _var.dbl_BH06 + _var.dbl_BJ06 + _var.dbl_BL06;
                    _var.objdt_template06.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template06.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ06;
                    _var.objdt_template06.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template06, boo_open, str_saved + @"\BM019-JUN-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template07.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template07.NewRow();
                    _var.objdt_template07.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template07.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF07 + _var.dbl_BH07 + _var.dbl_BJ07 + _var.dbl_BL07;
                    _var.objdt_template07.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template07.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ07;
                    _var.objdt_template07.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template07, boo_open, str_saved + @"\BM019-JUL-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template08.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template08.NewRow();
                    _var.objdt_template08.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template08.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF08 + _var.dbl_BH08 + _var.dbl_BJ08 + _var.dbl_BL08;
                    _var.objdt_template08.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template08.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ08;
                    _var.objdt_template08.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template08, boo_open, str_saved + @"\BM019-AUG-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template09.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template09.NewRow();
                    _var.objdt_template09.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template09.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF09 + _var.dbl_BH09 + _var.dbl_BJ09 + _var.dbl_BL09;
                    _var.objdt_template09.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template09.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ09;
                    _var.objdt_template09.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template09, boo_open, str_saved + @"\BM019-SEP-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template10.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template10.NewRow();
                    _var.objdt_template10.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template10.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF10 + _var.dbl_BH10 + _var.dbl_BJ10 + _var.dbl_BL10;
                    _var.objdt_template10.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template10.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ10;
                    _var.objdt_template10.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template10, boo_open, str_saved + @"\BM019-OCT-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template11.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template11.NewRow();
                    _var.objdt_template11.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template11.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF11 + _var.dbl_BH11 + _var.dbl_BJ11 + _var.dbl_BL11;
                    _var.objdt_template11.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template11.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ11;
                    _var.objdt_template11.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template11, boo_open, str_saved + @"\BM019-NOV-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (_var.objdt_template12.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_template12.NewRow();
                    _var.objdt_template12.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template12.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF12 + _var.dbl_BH12 + _var.dbl_BJ12 + _var.dbl_BL12;
                    _var.objdt_template12.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_template12.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ12;
                    _var.objdt_template12.Rows.Add(_var.dtworkRow);
                    #endregion
                    objHlpr.fn_savemultiple(_var.objdt_template12, boo_open, str_saved + @"\BM019-DEC-" + str_savef + ".xlsx");
                }
                //if (_var.boo_genderfail)
                //{
                //    _var.dtworkRow = objdt_template.NewRow();
                //    _var.dtworkRow[0] = "Please check for blank genders";
                //    _var.objdt_template01.Rows.Add(_var.dtworkRow);
                //}
                
                eapp.DisplayAlerts = false;
                wsraw = null;
                wbraw.SaveAs(str_raw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
                wbraw.Close();
                wbraw = null;
                eapp = null;
                ////
                _var.dtworkRow = null; //Dispose datarow


                _var.objdt_template01.Dispose();
                _var.objdt_template01 = null;
                _var.objdt_template02.Dispose();
                _var.objdt_template02 = null;
                _var.objdt_template03.Dispose();
                _var.objdt_template03 = null;
                _var.objdt_template04.Dispose();
                _var.objdt_template04 = null;
                _var.objdt_template05.Dispose();
                _var.objdt_template05 = null;
                _var.objdt_template06.Dispose();
                _var.objdt_template06 = null;
                _var.objdt_template07.Dispose();
                _var.objdt_template07 = null;
                _var.objdt_template08.Dispose();
                _var.objdt_template08 = null;
                _var.objdt_template09.Dispose();
                _var.objdt_template09 = null;
                _var.objdt_template10.Dispose();
                _var.objdt_template10 = null;
                _var.objdt_template11.Dispose();
                _var.objdt_template11 = null;
                _var.objdt_template12.Dispose();
                _var.objdt_template12 = null;

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