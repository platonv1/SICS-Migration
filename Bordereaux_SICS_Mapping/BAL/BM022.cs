using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM022
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

            string str_bmyear = wsraw.Cells[2, 2].Text.ToString().Substring(wsraw.Cells[2, 2].Text.ToString().Length - 4, 4);
            double dbl_rate = 1.00;

            try
            {
                for (int intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string str_PolNum = wsraw.Cells[intLoop, 2].Text.ToString();
                    if (!objHlpr.fn_policyNumChecker(str_PolNum, wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString(), wsraw.Cells[intLoop, 5].Text.ToString()))
                    {
                        continue;
                    }

                    string str_Plan = wsraw.Cells[intLoop, 3].Text.ToString();
                    string str_Mortality = wsraw.Cells[intLoop, 4].Text.ToString();
                    //string str_OCC = wsraw.Cells[intLoop, 5].Text.ToString();
                    string str_Rtype = wsraw.Cells[intLoop, 6].Text.ToString();
                    string str_Fullname = wsraw.Cells[intLoop, 7].Text.ToString();
                    string str_DOB = wsraw.Cells[intLoop, 8].Text.ToString();
                    string str_Sex = wsraw.Cells[intLoop, 9].Text.ToString();
                    string str_IssueDate = wsraw.Cells[intLoop, 10].Text.ToString();
                    string str_age = wsraw.Cells[intLoop, 11].Text.ToString();

                    string str_FaceAmt = wsraw.Cells[intLoop, 13].Text.ToString(); //25
                    string str_Retention = wsraw.Cells[intLoop, 14].Text.ToString(); //28
                    string str_Reinsured = wsraw.Cells[intLoop, 15].Text.ToString(); //27
                    string str_NAAR = wsraw.Cells[intLoop, 28].Text.ToString(); //76

                    if (double.TryParse(str_FaceAmt, out double dbl_FaceAmt))
                    {
                        dbl_FaceAmt = dbl_FaceAmt * dbl_rate;
                    }
                    else
                    {
                        dbl_FaceAmt = 1;
                    }

                    if (double.TryParse(str_Retention, out double dbl_Retention))
                    {
                        dbl_Retention = dbl_Retention * dbl_rate;
                    }
                    else
                    {
                        dbl_Retention = 1;
                    }

                    if (double.TryParse(str_Reinsured, out double dbl_Reinsured))
                    {
                        dbl_Reinsured = dbl_Reinsured * dbl_rate;
                    }
                    else
                    {
                        dbl_Reinsured = 1;
                    }

                    if (double.TryParse(str_NAAR, out double dbl_NAAR))
                    {
                        dbl_NAAR = dbl_NAAR * dbl_rate;
                    }
                    else
                    {
                        dbl_NAAR = 1;
                    }

                    string str_PremiumLife = wsraw.Cells[intLoop, 40].Text.ToString();
                    string str_PremiumWP = wsraw.Cells[intLoop, 41].Text.ToString();
                    string str_PremiumADB = wsraw.Cells[intLoop, 42].Text.ToString();
                    if (double.TryParse(str_PremiumLife, out double dbl_PremiumLife))
                    {
                        dbl_PremiumLife = dbl_PremiumLife * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumLife = 0;
                    }
                    if (double.TryParse(str_PremiumWP, out double dbl_PremiumWP))
                    {
                        dbl_PremiumWP = dbl_PremiumWP * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumWP = 0;
                    }
                    if (double.TryParse(str_PremiumADB, out double dbl_PremiumADB))
                    {
                        dbl_PremiumADB = dbl_PremiumADB * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumADB = 0;
                    }

                    DateTime dt_PremiumDate = Convert.ToDateTime(str_IssueDate.Substring(0, str_IssueDate.Length - 4) + str_bmyear), 
                        dt_IssueDate = Convert.ToDateTime(str_IssueDate);

                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow02 = null;
                    _var.dtworkRow03 = null;

                    _var.dtworkRow01[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow01[5] = str_Plan;
                    _var.dtworkRow01[8] = "SURPLUS";
                    _var.dtworkRow01[9] = "PAFW";
                    _var.dtworkRow01[13] = "IND";
                    _var.dtworkRow01[10] = "S";

                    if (str_Rtype.ToUpper().Contains("FACUL"))
                    {
                        _var.dtworkRow01[14] = "F";
                    }
                    else
                    {
                        _var.dtworkRow01[14] = "T";
                    }

                    _var.dtworkRow01[23] = "PHP";
                    _var.dtworkRow01[24] = "YLY";
                    _var.dtworkRow01[29] = "NATREID";
                    //_var.dtworkRow01[44] = "'" + str_OCC;

                    //DOB
                    if (String.IsNullOrEmpty(str_DOB))
                    {
                        str_DOB = "7/1/1900";
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

                    _var.dtworkRow01[20] = str_IssueDate;
                    _var.dtworkRow01[22] = str_IssueDate.Substring(0, str_IssueDate.Length - 4) + str_bmyear;

                    string str_tcode = string.Empty;
                    if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[20];
                        str_tcode = "TNEWBUS";
                        _var.dtworkRow01[56] = "4000";
                        _var.dtworkRow01[57] = dbl_PremiumLife;
                    }
                    else
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        str_tcode = "TRENEW";
                        _var.dtworkRow01[58] = "4001";
                        _var.dtworkRow01[59] = dbl_PremiumLife;
                    }

                    _var.dtworkRow01[21] = str_tcode;

                    _var.dtworkRow01[25] = dbl_FaceAmt;
                    _var.dtworkRow01[27] = dbl_Reinsured;
                    _var.dtworkRow01[77] = dbl_NAAR;
                    _var.dtworkRow01[28] = dbl_Retention;

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

                    objHlpr.fn_getnamesandlifeID(str_Fullname, str_DOB, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");
                    _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(str_Fullname);
                    string str_MI = objHlpr.fn_getMI(str_outfname);
                    _var.dtworkRow01[32] = str_outlname.Trim();
                    _var.dtworkRow01[33] = str_outfname.Trim().Replace(" " + str_MI.Trim(),"");
                    _var.dtworkRow01[34] = str_MI.Trim();
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

                    //Macro
                    if (!String.IsNullOrEmpty(str_macro))
                    {
                        DataRow[] foundRows = dt_macro.Select("SPN = " + "'" + str_PolNum.ToString() + "'");
                        if (foundRows.Length != 0)
                        { _var.dtworkRow01[83] = foundRows[0][9]; }
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
                        _var.dtworkRow01[1] = string.Empty;
                        _var.dtworkRow01[7] = string.Empty;
                    }

                    if (!String.IsNullOrEmpty(str_PremiumWP) && str_PremiumWP.Trim() != "-")
                    {
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow02[5] = "WP";
                        if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                        {
                            _var.dtworkRow02[57] = dbl_PremiumWP;
                        }
                        else
                        {
                            _var.dtworkRow02[59] = dbl_PremiumWP;
                        }
                    }

                    if (!String.IsNullOrEmpty(str_PremiumADB) && str_PremiumADB.Trim() != "-")
                    {
                        _var.dtworkRow03 = objdt_template.NewRow();
                        _var.dtworkRow03.ItemArray = _var.dtworkRow01.ItemArray;
                        _var.dtworkRow03[5] = "ADB";
                        if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                        {
                            _var.dtworkRow03[57] = dbl_PremiumADB;
                        }
                        else
                        {
                            _var.dtworkRow03[59] = dbl_PremiumADB;
                        }
                    }

                    if (_var.dtworkRow01 != null)
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

                string despath = str_saved + @"\BM022-" + str_savef + ".xlsx";
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