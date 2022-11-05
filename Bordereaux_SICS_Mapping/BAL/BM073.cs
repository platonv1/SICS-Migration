using System;
using System.Data;
using System.Linq;
using System.Globalization;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM073
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            #region NOTES
            //Declaration for exception line debugging on excel
            #endregion
            int rowcount = 1;

            try
            {
                _Global _var = new _Global();
                Helper objHlpr = new Helper();
                DataTable objdt_template = new DataTable();

                objdt_template = objHlpr.dt_formtemplate(str_sheet);

                Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);

                Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets["SUMMARY"];
                string byear = wsraw.Cells[2, 1].Text.ToString();
                byear = byear.Trim();
                byear = byear.Substring(byear.Length - 4, 4);


                wsraw = wbraw.Sheets[str_sheet];
                Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

                int erawrow = rawrange.Rows.Count;
                int erawcol = rawrange.Columns.Count;
                int prawrow = 1;
                string curr = wsraw.Cells[prawrow, 1].Text.ToString();
                string polnum = wsraw.Cells[prawrow, 2].Text.ToString();
                string code = wsraw.Cells[prawrow, 3].Text.ToString();
                string effective = wsraw.Cells[prawrow, 4].Text.ToString();
                string dob = wsraw.Cells[prawrow, 7].Text.ToString();
                string gender = wsraw.Cells[prawrow, 5].Text.ToString();
                string rating = wsraw.Cells[prawrow, 9].Text.ToString();
                string orig = wsraw.Cells[prawrow, 11].Text.ToString(); //Original Sum Insured Z
                string fullname = wsraw.Cells[prawrow, 6].Text.ToString();
                string prem = wsraw.Cells[prawrow, 14].Text.ToString();// Premium
                string SR = wsraw.Cells[prawrow, 13].Text.ToString();//Sum Reinsured BZ
                string age = wsraw.Cells[prawrow, 8].Text.ToString();
                string year = wsraw.Cells[4, 8].Text.ToString();
                string strRate = wsraw.Cells[prawrow, 15].Text.ToString();
                string strtype = string.Empty;

                year = year.Replace(year.Substring(year.Length - 3, 3), "-01" + year.Substring(year.Length - 3, 3));

                DateTime oDate = Convert.ToDateTime(year);

                if (boo_clean)
                {
                    wsraw = objHlpr.fn_extendwidth(wsraw);
                }

                string currency = string.Empty;
                string year12 = string.Empty;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                int storee;
                bool chck;
                double rate = 0.00;
                polnum = objHlpr.fn_stringcleanup(polnum);
                fullname = objHlpr.fn_stringcleanup(fullname);
                gender = objHlpr.fn_stringcleanup(gender);
                dob = objHlpr.fn_stringcleanup(dob);
                curr = objHlpr.fn_stringcleanup(curr);
                code = objHlpr.fn_stringcleanup(code);
                effective = objHlpr.fn_stringcleanup(effective);
                rating = objHlpr.fn_stringcleanup(rating);
                orig = objHlpr.fn_stringcleanup(orig);
                prem = objHlpr.fn_stringcleanup(prem);
                age = objHlpr.fn_stringcleanup(age);
                year = objHlpr.fn_stringcleanup(year);
                #region Data Processing

                while (rowcount != erawrow + 2)
                {
                    chck = int.TryParse(polnum, out storee);
                    polnum = objHlpr.fn_stringcleanup(polnum);

                    if (polnum == string.Empty && chck == false)
                    {
                        if (curr.ToUpper().Contains("PESO"))
                        {
                            currency = "PHP";
                        }
                        else if ((curr.ToUpper().Contains("DOLLAR")) || (curr.ToUpper().Contains("$")))
                        {
                            currency = "USD";
                        }
                    }
                    else if (polnum != string.Empty && chck == true)
                    {

                        if (strRate.Trim().Replace("%", "") == string.Empty)
                        {
                            rate = 0.15;
                        }
                        else
                        {
                            rate = double.Parse(strRate.Trim().Replace("%", "")) / 100;
                        }

                        orig = (double.Parse(orig) * rate).ToString();
                        prem = (double.Parse(prem) * rate).ToString();
                        SR = (double.Parse(SR) * rate).ToString();

                        _var.dtworkRow01 = objdt_template.NewRow();

                        _var.dtworkRow01[0] = polnum;
                        _var.dtworkRow01[5] = code.ToString();
                        _var.dtworkRow01[8] = "SURPLUS";
                        _var.dtworkRow01[9] = "PAFW";
                        _var.dtworkRow01[10] = "S";
                        _var.dtworkRow01[13] = "IND";
                        _var.dtworkRow01[14] = "T";

                        _var.dtworkRow01[20] = effective.ToString();
                        _var.dtworkRow01[22] = effective.Substring(0, effective.Length - 4) + byear;
                        _var.dtworkRow01[24] = "MLY";
                        _var.dtworkRow01[25] = orig;
                        _var.dtworkRow01[28] = orig;
                        _var.dtworkRow01[77] = SR;
                        _var.dtworkRow01[29] = "NATREID";
                        _var.dtworkRow01[79] = age.ToString();
                        _var.dtworkRow01[41] = byear;

                        if (gender.ToUpper().Contains("FEMALE"))
                        {
                            _var.dtworkRow01[36] = "F";
                        }
                        else
                        {
                            _var.dtworkRow01[36] = "M";
                        }
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        _var.dtworkRow01[37] = dob.ToString();

                        _var.dtworkRow01[38] = "NONE";


                        _var.dtworkRow01[39] = objHlpr.fn_getmortality(rating);
                        if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                        }

                        _var.dtworkRow01[23] = currency;


                        int effective1;
                        effective = effective.Substring(effective.Length - 4, 4);
                        effective1 = Convert.ToInt32(effective);

                        if (oDate.Year > effective1)
                        {
                            _var.dtworkRow01[21] = "TRENEW";
                            _var.dtworkRow01[58] = "4001";
                            _var.dtworkRow01[59] = prem.ToString();
                        }

                        else if (oDate.Year == effective1)
                        {
                            _var.dtworkRow01[21] = "TNEWBUS";
                            _var.dtworkRow01[56] = "4000";
                            _var.dtworkRow01[57] = prem.ToString();
                        }

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        _var.dtworkRow01[34] = str_MI;

                        _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(fullname);
                        _var.dtworkRow01[32] = str_outlname;

                        _var.dtworkRow01[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        _var.dtworkRow01[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow01[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                        }

                        #region "New Requirements"
                        _var.dtworkRow01[26] = string.Empty;

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
                            _var.dtworkRow01[77] = _var.dtworkRow01[25];
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

                        var parsedDOB = DateTime.Parse(dob);
                        string initialNR = string.Empty;
                        if (!String.IsNullOrEmpty(str_outfname))
                        {
                            initialNR = str_outfname.Substring(0, 1);
                        }
                        if (!String.IsNullOrEmpty(str_outlname))
                        {
                            initialNR += str_outlname.Substring(0, 1);
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

                            _var.dtworkRow01[1] = polnum.ToString() + gender.Substring(0, 1);
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR5-2B" : _var.dtworkRow01[76].ToString() + "|BR5-2B";

                            _var.dtworkRow01[7] = polnum.ToString();
                        }
                        else
                        {
                            _var.dtworkRow01[1] = string.Empty;
                            _var.dtworkRow01[7] = string.Empty;
                        }




                        //ISSUE#010-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow01[19].ToString()))
                        {
                            if (_var.dtworkRow01[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[20];
                            }
                            else if (_var.dtworkRow01[21].ToString().ToUpper() == "TRENEW")
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[22];
                            }
                            //else
                            //{
                            //    _var.dtworkRow01[19] = premium;
                            //    _var.dtworkRow01[22] = premium;
                            //}
                        }
                        //ISSUE#010-End-----------

                        if (_var.dtworkRow01[25].ToString() == "0")
                        {
                            _var.dtworkRow01[25] = "1";
                        }
                        if (_var.dtworkRow01[26].ToString() == "0")
                        {
                            _var.dtworkRow01[26] = "1";
                        }
                        if (_var.dtworkRow01[27].ToString() == "0")
                        {
                            _var.dtworkRow01[27] = "1";
                        }
                        if (_var.dtworkRow01[28].ToString() == "0")
                        {
                            _var.dtworkRow01[28] = "1";
                        }
                        if (_var.dtworkRow01[77].ToString() == "0")
                        {
                            _var.dtworkRow01[77] = "1";
                        }
                        #endregion


                        if (strtype == "UL")
                        {
                            if (_var.dtworkRow01[23].ToString() == "PHP")
                            {
                                _var.dbl_BF_PHP_UL += decimal.Parse(
                                       String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                       );
                                _var.dbl_BH_PHP_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                                _var.dbl_BJ_PHP_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                                _var.dbl_BL_PHP_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                                _var.dbl_BZ_PHP_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );
                            }
                            else if (_var.dtworkRow01[23].ToString() == "USD")
                            {
                                _var.dbl_BF_USD_UL += decimal.Parse(
                                       String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                       );
                                _var.dbl_BH_USD_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                                _var.dbl_BJ_USD_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                                _var.dbl_BL_USD_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                                _var.dbl_BZ_USD_UL += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );
                            }
                        }
                        else if (strtype == "TRAD")
                        {
                            if (_var.dtworkRow01[23].ToString() == "PHP")
                            {
                                _var.dbl_BF_PHP += decimal.Parse(
                                       String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                       );
                                _var.dbl_BH_PHP += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                                _var.dbl_BJ_PHP += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                                _var.dbl_BL_PHP += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                                _var.dbl_BZ_PHP += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );
                            }
                            else if (_var.dtworkRow01[23].ToString() == "USD")
                            {
                                _var.dbl_BF_USD += decimal.Parse(
                                       String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                       );
                                _var.dbl_BH_USD += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                    );
                                _var.dbl_BJ_USD += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                    );
                                _var.dbl_BL_USD += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                    );
                                _var.dbl_BZ_USD += decimal.Parse(
                                    String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                    );
                            }
                        }

                        

                        objdt_template.Rows.Add(_var.dtworkRow01);// inpu8trow+++
                    }

                    prawrow++;
                    curr = wsraw.Cells[prawrow, 1].Text.ToString();
                    polnum = wsraw.Cells[prawrow, 2].Text.ToString();
                    code = wsraw.Cells[prawrow, 3].Text.ToString();

                    if (code.ToUpper().StartsWith("UL"))
                    {
                        strtype = "UL";
                    }
                    else
                    {
                        strtype = "TRAD";
                    }

                    strRate = wsraw.Cells[prawrow, 15].Text.ToString();
                   
                    

                    effective = wsraw.Cells[prawrow, 4].Text.ToString();
                    dob = wsraw.Cells[prawrow, 7].Text.ToString();
                    gender = wsraw.Cells[prawrow, 5].Text.ToString();
                    rating = wsraw.Cells[prawrow, 9].Text.ToString();
                    age = wsraw.Cells[prawrow, 8].Text.ToString();


                    fullname = wsraw.Cells[prawrow, 6].Text.ToString();
                    orig = wsraw.Cells[prawrow, 11].Text.ToString().Replace("P", "").Replace("$", "").Replace(",", "").Trim(); //Original Sum Insured Z
                    prem = wsraw.Cells[prawrow, 14].Text.ToString().Replace("P", "").Replace("$", "").Replace(",", "").Trim();// Premium
                    SR = wsraw.Cells[prawrow, 13].Text.ToString().Replace("P", "").Replace("$", "").Replace(",", "").Trim();//Sum Reinsured BZ


                    rowcount++;
                }


                #endregion

                #region "Compute Hash Total"
                _var.dtworkRow01 = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium PHP (TRAD):";
                _var.dtworkRow01[1] = _var.dbl_BF_PHP + _var.dbl_BH_PHP + _var.dbl_BJ_PHP + _var.dbl_BL_PHP;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium USD (TRAD):";
                _var.dtworkRow01[1] = _var.dbl_BF_USD + _var.dbl_BH_USD + _var.dbl_BJ_USD + _var.dbl_BL_USD;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk PHP (TRAD):";
                _var.dtworkRow01[1] = _var.dbl_BZ_PHP;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk USD (TRAD):";
                _var.dtworkRow01[1] = _var.dbl_BZ_USD;
                objdt_template.Rows.Add(_var.dtworkRow01);

                //=================================================
                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium PHP (UL):";
                _var.dtworkRow01[1] = _var.dbl_BF_PHP_UL + _var.dbl_BH_PHP_UL + _var.dbl_BJ_PHP_UL + _var.dbl_BL_PHP_UL;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium USD (UL):";
                _var.dtworkRow01[1] = _var.dbl_BF_USD_UL + _var.dbl_BH_USD_UL + _var.dbl_BJ_USD_UL + _var.dbl_BL_USD_UL;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk PHP (UL):";
                _var.dtworkRow01[1] = _var.dbl_BZ_PHP_UL;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk USD (UL):";
                _var.dtworkRow01[1] = _var.dbl_BZ_USD_UL;
                objdt_template.Rows.Add(_var.dtworkRow01);
                #endregion

                string despath = str_saved + @"\BM073-" + str_savef + ".xlsx";
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
                _var.dtworkRow01 = null; //Dispose datarow
                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message + Environment.NewLine + " *****On excel row line: " + rowcount + " *****";
            }

        }
    }
}
