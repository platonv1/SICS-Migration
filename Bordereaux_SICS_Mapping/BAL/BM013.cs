using System;
using System.Data;
using System.Linq;
using System.Globalization;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM013
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
                Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
                Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

                if (boo_clean)
                {
                    wsraw = objHlpr.fn_extendwidth(wsraw);
                }

                int erawrow = rawrange.Rows.Count;
                int erawcol = rawrange.Columns.Count;
                int prawrow = 1;
                string polnum = wsraw.Cells[prawrow, 7].Text.ToString();
                polnum = polnum.Replace(" ", string.Empty);
                string cession = wsraw.Cells[prawrow, 1].Text.ToString();
                string branded = wsraw.Cells[prawrow, 3].Text.ToString();
                string reins = wsraw.Cells[prawrow, 9].Text.ToString();
                string ceded = wsraw.Cells[prawrow, 13].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                string dob = wsraw.Cells[prawrow, 8].Text.ToString();
                string premium = wsraw.Cells[prawrow,15].Text.ToString();
                string age = wsraw.Cells[prawrow, 10].Text.ToString();
                string pref = wsraw.Cells[prawrow, 5].Text.ToString();
                string rider = wsraw.Cells[prawrow, 6].Text.ToString();
                string rprem = wsraw.Cells[prawrow, 16].Text.ToString();
                string retention = wsraw.Cells[prawrow, 12].Text.ToString();
                string byear = wsraw.Cells[4, 1].Text.ToString();

                string[] arr_fullbyear;
                arr_fullbyear = byear.TrimEnd().Split(' ');

                byear = byear.Trim();
                byear = byear.Substring(byear.Length - 4, 4);
                string iyear = string.Empty;

                string currency = string.Empty;
                string year12 = string.Empty;
                string[] comparestring = new string[] { "" };
                string gender = string.Empty;
                bool chck;
                //decimal classific;


                string fullbyear = arr_fullbyear[arr_fullbyear.Length - 3] + " " + arr_fullbyear[arr_fullbyear.Length - 2] + " " + arr_fullbyear[arr_fullbyear.Length - 1];
                DateTime parsedfullbyear = Convert.ToDateTime(fullbyear);
                
                double rate = 0;
                if (parsedfullbyear.Year <= 2017)
                {
                    rate = .15;
                }
                else if ((parsedfullbyear.Year == 2018) && (parsedfullbyear.Month <= 6))
                {
                    rate = 1;
                }
                else
                {
                    rate = .85;
                }

                #region Data Processing

                while (rowcount != erawrow + 2)
                {
                    polnum = objHlpr.fn_stringcleanup(polnum);

                    if (cession.ToUpper().IndexOf("H-") == 0)
                    {
                        chck = true;
                    }
                    else
                    {
                        chck = false;
                        if (wsraw.Cells[prawrow, 1].Text.ToString().Contains("quarter ending"))
                        {
                            byear = wsraw.Cells[prawrow, 1].Text.ToString();

                            arr_fullbyear = byear.TrimEnd().Split(' ');

                            byear = byear.Trim();
                            byear = byear.Substring(byear.Length - 4, 4);
                        }
                    }

                    if (cession != string.Empty && chck == true)
                    {

                        _var.dtworkRow01 = objdt_template.NewRow();

                        _var.dtworkRow01[0] = polnum;
                        _var.dtworkRow01[1] = cession;
                        _var.dtworkRow01[5] = branded;
                        _var.dtworkRow01[8] = "SURPLUS";
                        _var.dtworkRow01[9] = "PAFM";
                        _var.dtworkRow01[10] = "S";
                        _var.dtworkRow01[13] = "GCL";
                        _var.dtworkRow01[14] = "T";

                        _var.dtworkRow01[20] = reins;
                        _var.dtworkRow01[22] = reins;
                        //_var.dtworkRow01[19] = reins;

                        _var.dtworkRow01[23] = "PHP";
                        _var.dtworkRow01[24] = "YLY";
                        _var.dtworkRow01[27] = ceded;
                        _var.dtworkRow01[25] = ceded;
                        _var.dtworkRow01[26] = ceded;
                        _var.dtworkRow01[77] = ceded;
                        _var.dtworkRow01[28] = double.Parse(objHlpr.fn_dashtozero(retention));
                        _var.dtworkRow01[79] = age;
                        _var.dtworkRow01[31] = fullname;
                        _var.dtworkRow01[41] = byear;
                        _var.dtworkRow01[29] = "NATREID";
                        
                        //_var.dtworkRow01[78] = (int.Parse(byear) - int.Parse(iyear)) + int.Parse(age);

                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            //ISSUE#009-Start---------
                            dob = "07/01/1900";
                            //ISSUE#009-End-----------

                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                        }
                        #endregion
                        _var.dtworkRow01[37] = dob;


                        if (str_sheet.ToUpper() == ("FIRST YEAR"))
                        {
                            _var.dtworkRow01[21] = "TNEWBUS";
                            _var.dtworkRow01[56] = "4000";
                            _var.dtworkRow01[57] = Math.Round(double.Parse(premium.ToString()) * rate, 2);
                        }
                        else if (str_sheet.ToUpper() == ("RENEWAL"))
                        {
                            _var.dtworkRow01[21] = "TRENEW";
                            _var.dtworkRow01[58] = "4001";
                            _var.dtworkRow01[59] = Math.Round(double.Parse(premium.ToString()) * rate, 2);
                            
                        }
                        else if (str_sheet.ToUpper() == ("ADJUSTMENT FIRST YEAR"))
                        {
                            _var.dtworkRow01[21] = "ADJUST";
                            _var.dtworkRow01[60] = "4002";
                            _var.dtworkRow01[61] = Math.Round(double.Parse(premium.ToString()) * rate, 2);
                        }
                        else if (str_sheet.ToUpper().Contains("ADJUSTMENT RENEWAL"))
                        {
                            _var.dtworkRow01[21] = "ADJUST";
                            _var.dtworkRow01[62] = "4004";
                            _var.dtworkRow01[63] = Math.Round(double.Parse(premium.ToString()) * rate, 2);
                        }

                        _var.dtworkRow01[38] = "NONE";
                        //ISSUE# Bug on mortality-Start---------
                        _var.dtworkRow01[39] = objHlpr.fn_getmortality(pref);
                        if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                        }
                        //ISSUE# Bug on mortality-End-----------

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                        }
                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out _var.str_outfname, out _var.str_outlname, out _var.str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(_var.str_outfname);
                        _var.dtworkRow01[34] = str_MI;

                        _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(fullname);
                        _var.dtworkRow01[32] = _var.str_outlname;
                        _var.dtworkRow01[33] = _var.str_outfname;

                        _var.dtworkRow01[30] = _var.str_outlifeid;

                        //ISSUE#020-Start---------
                        if (!String.IsNullOrEmpty(gender))
                        {
                            _var.dtworkRow01[36] = (gender.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                        }
                        //ISSUE#020-End-----------
                        else if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            //_var.dtworkRow01[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                            _var.dtworkRow01 [36] = objHlpr.fn_getgenderv2(_var.dtworkRow01 [33].ToString());
                            //ISSUE#003-Start---------
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                            //ISSUE#003-End-----------
                        }
                        else if (String.IsNullOrEmpty(gender) && String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow01 [36] = objHlpr.fn_getgenderv2(_var.dtworkRow01 [33].ToString());
                            //_var.dtworkRow01[36] = string.Empty;
                        }
                        gender = _var.dtworkRow01[36].ToString();

                        //ISSUE#013-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow01[36].ToString()))
                        {
                            _var.str_GFailLines = String.IsNullOrEmpty(_var.str_GFailLines) ? prawrow.ToString() : _var.str_GFailLines + "," + prawrow.ToString();
                        }
                        //ISSUE#013-End-----------

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

                            //ISSUE#019-Start---------
                            //if (!string.IsNullOrEmpty(gender))
                            //{
                            //    _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + gender.Substring(0, 1);
                            //}
                            //else
                            //{
                            //    _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + "-";
                            //}

                            _var.dtworkRow01[1] = cession;
                            //ISSUE#019-End-----------

                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR5-2B" : _var.dtworkRow01[76].ToString() + "|BR5-2B";

                            _var.dtworkRow01[7] = polnum.ToString();
                        }
                        else
                        {
                            _var.dtworkRow01[1] = string.Empty;
                            _var.dtworkRow01[7] = string.Empty;
                        }

                        if (_var.dtworkRow01[21].ToString().ToUpper() == "TRENEW")
                        {
                            DateTime parsedreins = Convert.ToDateTime(reins);
                            _var.dtworkRow01[22] = parsedreins.Month + "/" + parsedreins.Day + "/" + byear;
                        }
                        else
                        {
                            _var.dtworkRow01[22] = reins;
                        }

                        //ISSUE#010-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow01[19].ToString()))
                        {
                            if (_var.dtworkRow01[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[20];
                            }
                            else
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[22];
                            }
                        }
                        //ISSUE#010-End-----------

                        //ISSUE#017-Start---------
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
                        //ISSUE#017-End-----------
                        _var.dtworkRow01[25] = Math.Round(double.Parse(_var.dtworkRow01[25].ToString()) * rate, 2);
                        _var.dtworkRow01[27] = Math.Round(double.Parse(_var.dtworkRow01[27].ToString()) * rate, 2);
                        _var.dtworkRow01[77] = Math.Round(double.Parse(_var.dtworkRow01[77].ToString()) * rate, 2);
                        #endregion

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

                        objdt_template.Rows.Add(_var.dtworkRow01);// inpu8trow+++

                        if (!string.IsNullOrEmpty(rider))
                        {
                            _var.dtworkRow02 = objdt_template.NewRow();
                            _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;
                            _var.dtworkRow02[5] = rider;
                            
                            _var.dtworkRow02[25] = "1";
                            _var.dtworkRow02[27] = "1";
                            _var.dtworkRow02[28] = "1";
                            _var.dtworkRow02[77] = "1";

                            if (str_sheet.ToUpper() == ("FIRST YEAR"))
                            {
                                _var.dtworkRow02[57] = Math.Round(double.Parse(rprem.ToString()) * rate, 2); 
                            }
                            else if (str_sheet.ToUpper() == ("RENEWAL"))
                            {
                                _var.dtworkRow02[59] = Math.Round(double.Parse(rprem.ToString()) * rate, 2);
                            }
                            else if (str_sheet.ToUpper() == ("ADJUSTMENT FIRST YEAR"))
                            {
                                _var.dtworkRow02[61] = Math.Round(double.Parse(rprem.ToString()) * rate, 2);
                            }
                            else if (str_sheet.ToUpper().Contains("ADJUSTMENT RENEWAL"))
                            {
                                _var.dtworkRow02[63] = Math.Round(double.Parse(rprem.ToString()) * rate, 2);
                            }


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

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 7].Text.ToString();
                    polnum = polnum.Replace(" ", string.Empty);
                    cession = wsraw.Cells[prawrow, 1].Text.ToString();
                    branded = wsraw.Cells[prawrow, 3].Text.ToString();
                    reins = wsraw.Cells[prawrow, 9].Text.ToString();
                    ceded = wsraw.Cells[prawrow, 13].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                    dob = wsraw.Cells[prawrow, 8].Text.ToString();
                    premium = wsraw.Cells[prawrow, 15].Text.ToString();
                    age = wsraw.Cells[prawrow, 10].Text.ToString();
                    pref = wsraw.Cells[prawrow, 5].Text.ToString();
                    rider = wsraw.Cells[prawrow, 6].Text.ToString();
                    rprem = wsraw.Cells[prawrow, 16].Text.ToString();
                    retention = wsraw.Cells[prawrow, 12].Text.ToString();

                    if ((!String.IsNullOrEmpty(reins)) && (reins.Length >= 4)) {
                        reins = reins.Trim();
                        iyear = reins.Substring(reins.Length - 4, 4);
                    }
                    

                    gender = string.Empty;
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

                string despath = str_saved + @"\BM013-" + str_savef + ".xlsx";
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
                _var.dtworkRow = null; //Dispose datarow
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
