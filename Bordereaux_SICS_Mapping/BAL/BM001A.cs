using System;
using System.Data;
using System.Linq;
using System.Globalization;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM001A
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            int rowcount = 1;

            try
            {
                _Global _var = new _Global();
                Helper objHlpr = new Helper();
                DataTable objdt_template = new DataTable();

                objdt_template = objHlpr.dt_formtemplate(str_sheet);

                Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
                Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
                Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

                if (boo_clean)
                {
                    wsraw = objHlpr.fn_extendwidth(wsraw);
                }

                int erawrow = rawrange.Rows.Count;
                int erawcol = rawrange.Columns.Count;
                int prawrow = 1;
                string cession = wsraw.Cells[prawrow, 1].Text;
                string curr = wsraw.Cells[prawrow, 1].Text;
                string polnum = wsraw.Cells[prawrow, 2].Text;
                string eff = wsraw.Cells[prawrow, 4].Text;
                string gender = wsraw.Cells[prawrow, 5].Text;
                string code = wsraw.Cells[prawrow, 3].Text;
                string fullname = wsraw.Cells[prawrow, 6].Text;
                string dob = wsraw.Cells[prawrow, 7].Text;
                string age = wsraw.Cells[prawrow, 8].Text;
                string rating = wsraw.Cells[prawrow, 9].Text;
                string orig = wsraw.Cells[prawrow, 10].Text;
                string sum = wsraw.Cells[prawrow, 13].Text;
                string origs = wsraw.Cells[prawrow, 11].Text;

                string premium = wsraw.Cells[prawrow, 14].Text;

                string[] arr_bmyear = wsraw.Cells[4, 8].Text.Split('-');
                string bmyear = objHlpr.fn_getMonthNumber(arr_bmyear[0]) + "/01/20" + arr_bmyear[1];

                double mul;
                string currency = string.Empty;
                string year12 = string.Empty;
                string[] comparestring = new string[] { "" };
                mul = 1;
                int storee;
                bool chck;

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

                        _var.dtworkRow = objdt_template.NewRow();
                        origs = origs.TrimStart(' ').TrimEnd(' ').Replace("P", String.Empty).Replace("$", String.Empty);
                        sum = sum.TrimStart(' ').TrimEnd(' ').Replace("$", String.Empty).Replace("P", String.Empty);
                        premium = premium.TrimStart(' ').TrimEnd(' ').Replace("$", String.Empty).Replace("P", String.Empty);

                        if (str_sheet.ToUpper().IndexOf("FAC") > -1)
                        {
                            _var.dtworkRow[9] = "PFO";
                            _var.dtworkRow[14] = "F";
                        }
                        else
                        {
                            _var.dtworkRow[9] = "PAFM";
                            _var.dtworkRow[14] = "T";
                        }

                        _var.dtworkRow[0] = polnum;
                        _var.dtworkRow[5] = code;
                        _var.dtworkRow[8] = "SURPLUS";

                        _var.dtworkRow[10] = "S";
                        _var.dtworkRow[13] = "IND";

                        _var.dtworkRow[20] = eff.ToString();
                        DateTime oDate = Convert.ToDateTime(bmyear);
                        _var.dtworkRow[22] = bmyear.ToString();
                        _var.dtworkRow[20] = eff.ToString();
                        _var.dtworkRow[23] = currency;
                        _var.dtworkRow[25] = origs;
                        double summ;
                        summ = Convert.ToDouble(sum);
                        _var.dtworkRow[26] = summ * mul;
                        _var.dtworkRow[27] = summ * mul;
                        _var.dtworkRow[29] = "NATREID";
                        _var.dtworkRow[78] = age;
                        _var.dtworkRow[31] = fullname;
                        _var.dtworkRow[24] = "MLY";

                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            //ISSUE#009-Start---------
                            dob = "07/01/1900";
                            //ISSUE#009-End-----------

                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR4AL" : _var.dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion
                        _var.dtworkRow[37] = dob;

                        int bmy;
                        string bmyear1;
                        int effective1;
                        string eff1;
                        oDate = Convert.ToDateTime(eff);
                        eff1 = oDate.Year.ToString();
                        effective1 = Convert.ToInt32(eff1);
                        oDate = Convert.ToDateTime(bmyear);
                        bmyear1 = oDate.Year.ToString();
                        bmy = Convert.ToInt32(bmyear1);

                        if (bmy >= effective1)
                        {
                            _var.dtworkRow[21] = "TRENEW";
                            _var.dtworkRow[58] = "4001";
                            double premiumm;
                            //premiumm = Convert.ToDouble(premium);
                            //premiumm = premiumm * mul;

                            double.TryParse(premium, out premiumm);
                            premiumm = premiumm * mul;
                            _var.dtworkRow[59] = premiumm;
                        }
                        else if (bmy < effective1)
                        {
                            _var.dtworkRow[21] = "TNEWBUS";
                            _var.dtworkRow[56] = "4000";
                            double premiumm;
                            premiumm = Convert.ToDouble(premium);
                            premiumm = premiumm * mul;
                            _var.dtworkRow[57] = premiumm;
                        }

                        string classific = rating;

                        //ISSUE# Bug on mortality-Start---------
                        _var.dtworkRow[39] = objHlpr.fn_getmortality(classific);
                        if (objHlpr.fn_isDMort(_var.dtworkRow[39].ToString()))
                        {
                            _var.dtworkRow[39] = "STANDARD";
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR8AN" : _var.dtworkRow[76].ToString() + "|BR8AN";
                        }
                        //ISSUE# Bug on mortality-End-----------

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR6AF" : _var.dtworkRow[76].ToString() + "|BR6AF";
                        }
                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out _var.str_outfname, out _var.str_outlname, out _var.str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(_var.str_outfname);
                        _var.dtworkRow[34] = str_MI;

                        _var.dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        _var.dtworkRow[32] = _var.str_outlname;

                        //ISSUE#-Start022---------
                        string[] arr_fname;
                        arr_fname = _var.str_outfname.Split(' ');

                        if (!String.IsNullOrEmpty(str_MI.Trim()))
                        {
                            for (int i = 0; i <= arr_fname.Length - 1; i++)
                            {
                                if (arr_fname[i] != str_MI)
                                {
                                    _var.dtworkRow[33] = String.IsNullOrEmpty(_var.dtworkRow[33].ToString()) ? arr_fname[i] : _var.dtworkRow[33].ToString() + " " + arr_fname[i];
                                }
                            }
                        }
                        else
                        {
                            //NIGNES 20200818
                            //Correct First name and Middlename out
                            //Start
                            string[] arr_mname;
                            arr_mname = _var.str_outfname.Split(' ');
                            if (arr_mname.Length > 1)
                            {
                                string[] str_suffix = {
                                            "JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "VI"
                                        };

                                if (str_suffix.Any(arr_mname[arr_mname.Length - 1].Contains))
                                {
                                    _var.dtworkRow[34] = arr_mname[arr_mname.Length - 2];
                                }
                                else
                                {
                                    _var.dtworkRow[34] = arr_mname[arr_mname.Length - 1];
                                }
                                _var.dtworkRow[33] = _var.str_outfname.Replace(" " + _var.dtworkRow[34].ToString(), string.Empty);
                            }
                            else
                            {
                                _var.dtworkRow[33] = _var.str_outfname;
                            }
                            arr_mname = null;
                            //NIGNES 20200818
                        }
                        //ISSUE#-End022-----------

                        _var.dtworkRow[30] = _var.str_outlifeid;



                        //ISSUE#020-Start---------
                        if (!String.IsNullOrEmpty(gender))
                        {
                            _var.dtworkRow[36] = (gender.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                        }
                        //ISSUE#020-End-----------
                        else if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow[33].ToString());
                            //ISSUE#003-Start---------
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR7AK" : _var.dtworkRow[76].ToString() + "|BR7AK";
                            //ISSUE#003-End-----------
                        }
                        else if (String.IsNullOrEmpty(gender) && String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow[36] = string.Empty;
                        }

                        //ISSUE#013-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                        {
                            _var.str_GFailLines = String.IsNullOrEmpty(_var.str_GFailLines) ? prawrow.ToString() : _var.str_GFailLines + "," + prawrow.ToString();
                        }
                        //ISSUE#013-End-----------

                        #region "New Requirements"
                        _var.dtworkRow[26] = string.Empty;

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

                        //ISSUE#009-Start---------
                        var parsedDOB = DateTime.ParseExact(dob, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                        //ISSUE#009-End-----------

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

                            //ISSUE#019-Start---------
                            _var.dtworkRow[1] = _var.dtworkRow[0].ToString() + gender.Substring(0, 1);
                            //ISSUE#019-End-----------

                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR5-2B" : _var.dtworkRow[76].ToString() + "|BR5-2B";

                            _var.dtworkRow[7] = polnum.ToString();
                        }
                        else
                        {
                            _var.dtworkRow[1] = string.Empty;
                            _var.dtworkRow[7] = string.Empty;
                        }

                        //ISSUE#010-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow[19].ToString()))
                        {
                            if (_var.dtworkRow[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                _var.dtworkRow[19] = _var.dtworkRow[20];
                            }
                            else
                            {
                                _var.dtworkRow[19] = _var.dtworkRow[22];
                            }
                        }
                        //ISSUE#010-End-----------

                        //ISSUE#017-Start---------
                        if (_var.dtworkRow[25].ToString() == "0")
                        {
                            _var.dtworkRow[25] = "1";
                        }
                        if (_var.dtworkRow[26].ToString() == "0")
                        {
                            _var.dtworkRow[26] = "1";
                        }
                        if (_var.dtworkRow[27].ToString() == "0")
                        {
                            _var.dtworkRow[27] = "1";
                        }
                        if (_var.dtworkRow[28].ToString() == "0")
                        {
                            _var.dtworkRow[28] = "1";
                        }
                        if (_var.dtworkRow[77].ToString() == "0")
                        {
                            _var.dtworkRow[77] = "1";
                        }
                        //ISSUE#017-End-----------

                        #endregion

                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                            );
                        _var.dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                            );
                        _var.dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                            );
                        _var.dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                            );
                        _var.dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(_var.dtworkRow);
                    }

                    prawrow++;
                    curr = wsraw.Cells[prawrow, 1].Text;
                    cession = wsraw.Cells[prawrow, 1].Text;
                    polnum = wsraw.Cells[prawrow, 2].Text;
                    eff = wsraw.Cells[prawrow, 4].Text;
                    gender = wsraw.Cells[prawrow, 5].Text;
                    code = wsraw.Cells[prawrow, 3].Text;
                    fullname = wsraw.Cells[prawrow, 6].Text;
                    dob = wsraw.Cells[prawrow, 7].Text;
                    age = wsraw.Cells[prawrow, 8].Text;
                    rating = wsraw.Cells[prawrow, 9].Text;
                    orig = wsraw.Cells[prawrow, 10].Text;
                    sum = wsraw.Cells[prawrow, 13].Text;
                    premium = wsraw.Cells[prawrow, 14].Text;
                    origs = wsraw.Cells[prawrow, 11].Text;
                    rowcount++;
                }
                #endregion

                #region "Compute Hash Total"
                _var.dtworkRow = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow);

                _var.dtworkRow = objdt_template.NewRow();
                _var.dtworkRow[0] = "Total Premium:";
                _var.dtworkRow[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL;
                objdt_template.Rows.Add(_var.dtworkRow);

                _var.dtworkRow = objdt_template.NewRow();
                _var.dtworkRow[0] = "Total Sum at Risk:";
                _var.dtworkRow[1] = _var.dbl_BZ;
                objdt_template.Rows.Add(_var.dtworkRow);
                #endregion

                //ISSUE#013-Start---------
                #region "List all failed genders"
                if (_var.str_GFailLines != string.Empty)
                {
                    _var.dtworkRow = objdt_template.NewRow();
                    _var.dtworkRow[0] = "Gender Fail Lines on RAW:";
                    _var.dtworkRow[1] = _var.str_GFailLines;
                    objdt_template.Rows.Add(_var.dtworkRow);
                }
                #endregion
                //ISSUE#013-End-----------

                string despath = str_saved + @"\BM001A-" + str_savef + ".xlsx";
                objHlpr.fn_savefile(objdt_template, despath);

                if (boo_open)
                {
                    objHlpr.fn_openfile(despath);
                }

                eapp.DisplayAlerts = false;
                wsraw = null;
                wbraw.SaveAs(str_raw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
                wbraw.Close();
                wbraw = null;
                eapp = null;
                _var.dtworkRow = null;
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
