using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM030
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

            string str_bmyear = wsraw.Cells[3, 1].Text.ToString().Substring(0,4);
            double dbl_rate = 0.15;
            double dbl_xOne = 1;
            

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
                    string str_Rtype = wsraw.Cells[intLoop, 4].Text.ToString();
                    string str_IssueDate = wsraw.Cells[intLoop, 5].Text.ToString();

                    string str_ExpiryDate = wsraw.Cells[intLoop, 8].Text.ToString();
                    string str_Sex = wsraw.Cells[intLoop, 9].Text.ToString();
                    string str_Fullname = wsraw.Cells[intLoop, 10].Text.ToString();
                    string str_DOB = wsraw.Cells[intLoop, 11].Text.ToString();
                    string str_age = wsraw.Cells[intLoop, 12].Text.ToString();
                    string str_Mortality = wsraw.Cells[intLoop, 13].Text.ToString();


                    string str_OSILife = wsraw.Cells[intLoop, 16].Text.ToString();
                    string str_OSITPD = wsraw.Cells[intLoop, 17].Text.ToString();
                    string str_OSIADD = wsraw.Cells[intLoop, 18].Text.ToString();
                    string str_RetentionLife = wsraw.Cells[intLoop, 19].Text.ToString();
                    string str_RetentionTPD = wsraw.Cells[intLoop, 20].Text.ToString();
                    string str_RetentionADD = wsraw.Cells[intLoop, 21].Text.ToString();
                    string str_NARLife = wsraw.Cells[intLoop, 22].Text.ToString();
                    string str_NARTPD = wsraw.Cells[intLoop, 23].Text.ToString();
                    string str_NARADD = wsraw.Cells[intLoop, 24].Text.ToString();

                    string str_PremiumLife = wsraw.Cells[intLoop, 25].Text.ToString();
                    string str_PremiumRating = wsraw.Cells[intLoop, 26].Text.ToString();
                    string str_PremiumExtra = wsraw.Cells[intLoop, 27].Text.ToString();

                    string str_PremiumTPD = wsraw.Cells[intLoop, 28].Text.ToString();
                    string str_PremiumADD = wsraw.Cells[intLoop, 29].Text.ToString();

                    if (double.TryParse(str_OSILife, out double dbl_OSILife))
                    {
                        dbl_OSILife = dbl_OSILife * dbl_xOne;
                    }
                    else
                    {
                        dbl_OSILife = 1;
                    }
                    if (double.TryParse(str_OSITPD, out double dbl_OSITPD))
                    {
                        dbl_OSITPD = dbl_OSITPD * dbl_xOne;
                    }
                    else
                    {
                        dbl_OSITPD = 1;
                    }
                    if (double.TryParse(str_OSIADD, out double dbl_OSIADD))
                    {
                        dbl_OSIADD = dbl_OSIADD * dbl_xOne;
                    }
                    else
                    {
                        dbl_OSIADD = 1;
                    }
                    if (double.TryParse(str_RetentionLife, out double dbl_RetentionLife))
                    {
                        dbl_RetentionLife = dbl_RetentionLife * dbl_xOne;
                    }
                    else
                    {
                        dbl_RetentionLife = 1;
                    }
                    if (double.TryParse(str_RetentionTPD, out double dbl_RetentionTPD))
                    {
                        dbl_RetentionTPD = dbl_RetentionTPD * dbl_xOne;
                    }
                    else
                    {
                        dbl_RetentionTPD = 1;
                    }
                    if (double.TryParse(str_RetentionADD, out double dbl_RetentionADD))
                    {
                        dbl_RetentionADD = dbl_RetentionADD * dbl_xOne;
                    }
                    else
                    {
                        dbl_RetentionADD = 1;
                    }
                    if (double.TryParse(str_NARLife, out double dbl_NARLife))
                    {
                        dbl_NARLife = dbl_NARLife * dbl_rate;
                    }
                    else
                    {
                        dbl_NARLife = 1;
                    }
                    if (double.TryParse(str_NARTPD, out double dbl_NARTPD))
                    {
                        dbl_NARTPD = dbl_NARTPD * dbl_rate;
                    }
                    else
                    {
                        dbl_NARTPD = 1;
                    }
                    if (double.TryParse(str_NARADD, out double dbl_NARADD))
                    {
                        dbl_NARADD = dbl_NARADD * dbl_rate;
                    }
                    else
                    {
                        dbl_NARADD = 1;
                    }

                    if (double.TryParse(str_PremiumLife, out double dbl_PremiumLife))
                    {
                        dbl_PremiumLife = dbl_PremiumLife * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumLife = 0;
                    }

                    if (double.TryParse(str_PremiumRating, out double dbl_PremiumRating))
                    {
                        dbl_PremiumRating = dbl_PremiumRating * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumRating = 0;
                    }

                    if (double.TryParse(str_PremiumExtra, out double dbl_PremiumExtra))
                    {
                        dbl_PremiumExtra = dbl_PremiumExtra * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumExtra = 0;
                    }
                    dbl_PremiumLife = dbl_PremiumLife + dbl_PremiumRating + dbl_PremiumExtra; ///////////////////////////////////////////////////////

                    if (double.TryParse(str_PremiumTPD, out double dbl_PremiumTPD))
                    {
                        dbl_PremiumTPD = dbl_PremiumTPD * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumTPD = 0;
                    }
                    if (double.TryParse(str_PremiumADD, out double dbl_PremiumADD))
                    {
                        dbl_PremiumADD = dbl_PremiumADD * dbl_rate;
                    }
                    else
                    {
                        dbl_PremiumADD = 0;
                    }

                    if (str_Plan.ToUpper().Trim() == "GCLI")
                    {
                        str_Plan = "GCL2";
                    }
                    else if(str_Plan.ToUpper().Trim() == "GMRI")
                    {
                        str_Plan = "GMRI2";
                    }
                    string str_tcode = wsraw.Cells[intLoop, 32].Text.ToString().Trim().ToUpper().Equals("NEW") ? "TNEWBUS" : "TRENEW";


                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow02 = null;
                    _var.dtworkRow03 = null;

                    _var.dtworkRow01[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow01[5] = str_Plan;
                    _var.dtworkRow01[8] = "SURPLUS";
                    _var.dtworkRow01[9] = "PAFM";
                    _var.dtworkRow01[13] = "GRP";
                    _var.dtworkRow01[10] = "S";
                    _var.dtworkRow01[14] = str_Rtype.ToUpper().Contains("FACUL") ? "F" : "T";
                    _var.dtworkRow01[23] = "PHP";
                    _var.dtworkRow01[24] = "YLY";
                    _var.dtworkRow01[29] = "NATREID";
                    _var.dtworkRow01[40] = str_ExpiryDate;

                    //DOB
                    if (String.IsNullOrEmpty(str_DOB))
                    {
                        str_DOB = "7/1/1900";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                    }
                    _var.dtworkRow01[37] = str_DOB;
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

                    if (str_tcode == "TNEWBUS")
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[20];
                        _var.dtworkRow01[56] = "4000";
                        _var.dtworkRow01[57] = dbl_PremiumLife;
                    }
                    else if (str_tcode == "TRENEW")
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        _var.dtworkRow01[58] = "4001";
                        _var.dtworkRow01[59] = dbl_PremiumLife;
                    }

                    _var.dtworkRow01[21] = str_tcode;

                    
                    _var.dtworkRow01[25] = dbl_OSILife.Equals(0.00) ? 1 : dbl_OSILife;
                    _var.dtworkRow01[77] = dbl_NARLife.Equals(0.00) ? 1 : dbl_NARLife;
                    _var.dtworkRow01[28] = dbl_RetentionLife.Equals(0.00) ? 1 : dbl_RetentionLife;

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
                    _var.dtworkRow01[32] = str_outlname;
                    _var.dtworkRow01[33] = str_outfname.Replace(" " + str_MI, string.Empty); 
                    _var.dtworkRow01[34] = str_MI;
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

                    if (!String.IsNullOrEmpty(str_OSITPD))
                    {
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;

                        if (str_tcode == "TNEWBUS")
                        {
                            _var.dtworkRow02[57] = dbl_PremiumTPD;
                        }
                        else if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow02[59] = dbl_PremiumTPD;
                        }

                        _var.dtworkRow02[5] = "TPD";

                        _var.dtworkRow02[25] = dbl_OSITPD.Equals(0.00) ? 1 : dbl_OSITPD;
                        _var.dtworkRow02[77] = dbl_NARTPD.Equals(0.00) ? 1 : dbl_NARTPD;
                        _var.dtworkRow02[28] = dbl_RetentionTPD.Equals(0.00) ? 1 : dbl_RetentionTPD;
                        _var.dtworkRow02[27] = _var.dtworkRow02[77];
                    }

                    if (!String.IsNullOrEmpty(str_OSIADD))
                    {
                        _var.dtworkRow03 = objdt_template.NewRow();
                        _var.dtworkRow03.ItemArray = _var.dtworkRow01.ItemArray;

                        if (str_tcode == "TNEWBUS")
                        {
                            _var.dtworkRow03[57] = dbl_PremiumADD;
                        }
                        else if (str_tcode == "TRENEW")
                        {
                            _var.dtworkRow03[59] = dbl_PremiumADD;
                        }

                        _var.dtworkRow03[5] = "ADD";

                        _var.dtworkRow03[25] = dbl_OSIADD.Equals(0.00) ? 1 : dbl_OSIADD;
                        _var.dtworkRow03[77] = dbl_NARADD.Equals(0.00) ? 1 : dbl_NARADD;
                        _var.dtworkRow03[28] = dbl_RetentionADD.Equals(0.00) ? 1 : dbl_RetentionADD;
                        _var.dtworkRow03[27] = _var.dtworkRow03[77];
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

                string despath = str_saved + @"\BM030-" + str_savef + ".xlsx";
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