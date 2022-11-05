using System;
using System.Data;
using System.Linq;
using System.Globalization;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM025
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

                _var.objdt_templateADJ = objHlpr.dt_formtemplate("ADJ");

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

                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                
                string issue = wsraw.Cells[prawrow, 3].Text.ToString();
                string issueyear = string.Empty;

                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string age = wsraw.Cells[prawrow, 5].Text.ToString();
                string month = wsraw.Cells[prawrow, 15].Text.ToString();
                string orig = wsraw.Cells[prawrow, 6].Text.ToString();
                string risk = wsraw.Cells[prawrow, 8].Text.ToString();
                string ret = wsraw.Cells[prawrow, 7].Text.ToString();
                string premium = wsraw.Cells[prawrow, 21].Text.ToString();
                string sum = wsraw.Cells[prawrow, 17].Text.ToString();

                string year = wsraw.Cells[1][4].Text.ToString();
                int qtr = int.Parse(year.Trim().Substring(0, 1));
                int iyear = int.Parse(year.Substring(year.Length -5, 4));
                year = year.Substring(year.Length - 5, 4);

                string gender = string.Empty;
                string currency = string.Empty;
                string[] comparestring = new string[] { "" };
                int storee;
                bool chck;

                #region Data Processing
                while (rowcount != erawrow + 2)
                {
                    chck = int.TryParse(polnum, out storee);
                    polnum = objHlpr.fn_stringcleanup(polnum);

                    if (polnum != string.Empty && chck == true)
                    {
                        bool is1stRow = true;
                        _var.dtworkRow = _var.objdt_template01.NewRow();
                        _var.dtworkRow01 = _var.objdt_template01.NewRow();
                        _var.dtworkRow02 = _var.objdt_template01.NewRow();
                        _var.dtworkRow03 = _var.objdt_template01.NewRow();

                        _var.dtworkRow01[0] = polnum;
                        _var.dtworkRow01[31] = fullname;
                        _var.dtworkRow01[20] = issue;
                        _var.dtworkRow01[37] = dob;
                        _var.dtworkRow01[79] = age;
                        //_var.dtworkRow01[3] = "DEATH";
                        //_var.dtworkRow01[4] = "VARIABLE-RE";
                        _var.dtworkRow01[5] = "MOUL";
                        _var.dtworkRow01[8] = "SURPLUS";
                        _var.dtworkRow01[9] = "PFO";
                        _var.dtworkRow01[10] = "S";
                        _var.dtworkRow01[13] = "IND";
                        _var.dtworkRow01[14] = "F";
                        _var.dtworkRow01[23] = "PHP";
                        _var.dtworkRow01[24] = "MLY";
                        _var.dtworkRow01[26] = "1.00";
                        _var.dtworkRow01[29] = "NATREID";
                        _var.dtworkRow01[38] = "NONE";
                        _var.dtworkRow01[25] = orig;
                        _var.dtworkRow01[27] = sum;
                        _var.dtworkRow01[28] = ret;
                        _var.dtworkRow01[77] = risk;
                        _var.dtworkRow01[78] = (int.Parse(year) - int.Parse(issueyear)) + int.Parse(age);
                        _var.dtworkRow01[41] = year;

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

                        string year1;
                        DateTime oDate = Convert.ToDateTime(issue);
                        year1 = oDate.Year.ToString();
                        int year11;
                        year11 = Convert.ToInt32(year);
                        int year2;
                        year2 = Convert.ToInt32(year1);
                        
                        month = wsraw.Cells[prawrow + 1, 15].Text.ToString();
                        DateTime rec_month = Convert.ToDateTime(month);

                        string code = string.Empty;

                        string isTerm = wsraw.Cells[prawrow + 1, 17].Text.ToString().ToUpper();
                        
                        if (isTerm.Trim() == "TERMINATED")
                        {
                            _var.dtworkRow01[21] = "TCONTER";
                            _var.dtworkRow01[62] = "4004";
                            _var.dtworkRow01[63] = "0";
                        }
                        else if (!(rec_month.Year == iyear && objHlpr.fn_isinQTR(qtr, rec_month.Month)))
                        {
                            if (objHlpr.MonthDiff(oDate, rec_month) <= 12)
                            {
                                _var.dtworkRow01[21] = "ADJUST";
                                _var.dtworkRow01[60] = "4002";
                                _var.dtworkRow01[61] = premium.ToString();

                                code = "4002";
                            }
                            else 
                            {
                                _var.dtworkRow01[21] = "ADJUST";
                                _var.dtworkRow01[62] = "4004";
                                _var.dtworkRow01[63] = premium.ToString();

                                code = "4004";
                            }
                        }
                        else if (year11 >= year2)
                        {
                            
                            _var.dtworkRow01[21] = "TRENEW";
                            _var.dtworkRow01[58] = "4001";
                            _var.dtworkRow01[59] = premium.ToString();
                        }
                        else if (year11 < year2)
                        {
                            _var.dtworkRow01[21] = "TNEWBUS";
                            _var.dtworkRow01[56] = "4000";
                            _var.dtworkRow01[57] = premium.ToString();
                        }

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                        }
                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out _var.str_outfname, out _var.str_outlname, out _var.str_outlifeid, "025");

                        string str_MI = objHlpr.fn_getMI(_var.str_outfname);
                        _var.dtworkRow01[34] = str_MI;

                        _var.dtworkRow01[31] = objHlpr.fn_stringcleanup(fullname);
                        _var.dtworkRow01[32] = _var.str_outlname;

                        //ISSUE#-Start022---------
                        string[] arr_fname;
                        arr_fname = _var.str_outfname.Split(' ');

                        if (!String.IsNullOrEmpty(str_MI.Trim()))
                        {
                            for (int i = 0; i <= arr_fname.Length - 1; i++)
                            {
                                if (arr_fname[i] != str_MI)
                                {
                                    _var.dtworkRow01[33] = String.IsNullOrEmpty(_var.dtworkRow01[33].ToString()) ? arr_fname[i] : _var.dtworkRow01[33].ToString() + " " + arr_fname[i];
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
                                    _var.dtworkRow01[34] = arr_mname[arr_mname.Length - 2];
                                }
                                else
                                {
                                    _var.dtworkRow01[34] = arr_mname[arr_mname.Length - 1];
                                }
                                _var.dtworkRow01[33] = _var.str_outfname.Replace(" " + _var.dtworkRow01[34].ToString(), string.Empty);
                            }
                            else
                            {
                                _var.dtworkRow01[33] = _var.str_outfname;
                            }
                            arr_mname = null;
                            //NIGNES 20200818
                        }
                        //ISSUE#-End022-----------

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

                            //New Method for Gender 05/12/2022
                            _var.dtworkRow01 [36] = objHlpr.fn_getgenderv2(_var.dtworkRow01 [33].ToString());
                            //ISSUE#003-Start---------
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                            //ISSUE#003-End-----------
                        }
                        else if (String.IsNullOrEmpty(gender) && String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow01[36] = objHlpr.fn_getgenderv2(_var.dtworkRow01 [33].ToString());
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
                            _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + gender.Substring(0, 1);
                            //ISSUE#019-End-----------

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
                            else
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[22];
                            }
                        }
                        //ISSUE#010-End-----------

                        _var.dtworkRow.ItemArray = _var.dtworkRow01.ItemArray;
                        
                        double d_out = 0;

                        for (int i = 1; i <= 3; i++)
                        {
                            prawrow++;
                            month = wsraw.Cells[prawrow, 15].Text.ToString();
                            rec_month = Convert.ToDateTime(month);

                            if (rec_month.Year == iyear && objHlpr.fn_isinQTR(qtr, rec_month.Month))
                            {
                                switch (rec_month.Month)
                                {
                                    case 1:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines01 = String.IsNullOrEmpty(_var.str_GFailLines01) ? prawrow.ToString() : _var.str_GFailLines01 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template01.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template01.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template01.NewRow();
                                        break;
                                    case 2:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines02 = String.IsNullOrEmpty(_var.str_GFailLines02) ? prawrow.ToString() : _var.str_GFailLines02 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template02.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template02.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template02.NewRow();
                                        break;
                                    case 3:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines03 = String.IsNullOrEmpty(_var.str_GFailLines03) ? prawrow.ToString() : _var.str_GFailLines03 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template03.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template03.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template03.NewRow();
                                        break;
                                    case 4:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines04 = String.IsNullOrEmpty(_var.str_GFailLines04) ? prawrow.ToString() : _var.str_GFailLines04 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template04.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template04.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template04.NewRow();
                                        break;
                                    case 5:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines05 = String.IsNullOrEmpty(_var.str_GFailLines05) ? prawrow.ToString() : _var.str_GFailLines05 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template05.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template05.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template05.NewRow();
                                        break;
                                    case 6:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines06 = String.IsNullOrEmpty(_var.str_GFailLines06) ? prawrow.ToString() : _var.str_GFailLines06 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template06.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template06.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template06.NewRow();
                                        break;
                                    case 7:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines07 = String.IsNullOrEmpty(_var.str_GFailLines07) ? prawrow.ToString() : _var.str_GFailLines07 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template07.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template07.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template07.NewRow();
                                        break;
                                    case 8:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines08 = String.IsNullOrEmpty(_var.str_GFailLines08) ? prawrow.ToString() : _var.str_GFailLines08 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template08.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template08.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template08.NewRow();
                                        break;
                                    case 9:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines09 = String.IsNullOrEmpty(_var.str_GFailLines09) ? prawrow.ToString() : _var.str_GFailLines09 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template09.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template09.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template09.NewRow();
                                        break;
                                    case 10:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines10 = String.IsNullOrEmpty(_var.str_GFailLines10) ? prawrow.ToString() : _var.str_GFailLines10 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template10.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template10.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template10.NewRow();
                                        break;
                                    case 11:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines11 = String.IsNullOrEmpty(_var.str_GFailLines11) ? prawrow.ToString() : _var.str_GFailLines11 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template11.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template11.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template11.NewRow();
                                        break;
                                    case 12:
                                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                        {
                                            _var.str_GFailLines12 = String.IsNullOrEmpty(_var.str_GFailLines12) ? prawrow.ToString() : _var.str_GFailLines12 + "," + prawrow.ToString();
                                        }
                                        _var.dtworkRow01 = _var.objdt_template12.NewRow();
                                        _var.dtworkRow02 = _var.objdt_template12.NewRow();
                                        _var.dtworkRow03 = _var.objdt_template12.NewRow();
                                        break;
                                    default:
                                        break;
                                }
                            }
                            else 
                            {
                                if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                                {
                                    _var.str_GFailLines_adj = String.IsNullOrEmpty(_var.str_GFailLines_adj) ? prawrow.ToString() : _var.str_GFailLines_adj + "," + prawrow.ToString();
                                }
                                _var.dtworkRow01 = _var.objdt_templateADJ.NewRow();
                                _var.dtworkRow02 = _var.objdt_templateADJ.NewRow();
                                _var.dtworkRow03 = _var.objdt_templateADJ.NewRow();
                            }

                            _var.dtworkRow01.ItemArray = _var.dtworkRow.ItemArray;
                            _var.dtworkRow02.ItemArray = _var.dtworkRow.ItemArray;
                            _var.dtworkRow03.ItemArray = _var.dtworkRow.ItemArray;

                            string moul = objHlpr.fn_numbercleanup(wsraw.Cells[prawrow, 18].Text.ToString());
                            if (!double.TryParse(moul, out d_out))
                            {
                                moul = "0";
                            }

                            string extra = objHlpr.fn_numbercleanup(wsraw.Cells[prawrow, 19].Text.ToString());
                            if (!double.TryParse(extra, out d_out))
                            {
                                extra = "0";
                            }

                            string wp = objHlpr.fn_numbercleanup(wsraw.Cells[prawrow, 20].Text.ToString());
                            if (!double.TryParse(wp, out d_out))
                            {
                                wp = "0";
                            }

                            string NAAR = objHlpr.fn_numbercleanup(wsraw.Cells[prawrow, 17].Text.ToString());
                            if (!double.TryParse(NAAR, out d_out))
                            {
                                NAAR = "0";
                            }

                            string mort_moul = objHlpr.fn_getmortality(wsraw.Cells[prawrow, 9].Text.ToString());
                            string mort_extra = objHlpr.fn_getmortality(wsraw.Cells[prawrow, 9].Text.ToString());
                            string mort_wp = objHlpr.fn_getmortality(wsraw.Cells[prawrow, 10].Text.ToString());

                            //1st Line
                            if (objHlpr.fn_isDMort(mort_moul))
                            {
                                mort_moul = "STANDARD";
                                _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                            }
                            _var.dtworkRow01[39] = mort_moul;

                            //2nd Line
                            if (objHlpr.fn_isDMort(mort_extra))
                            {
                                mort_extra = "STANDARD";
                                _var.dtworkRow02[76] = String.IsNullOrEmpty(_var.dtworkRow02[76].ToString()) ? "BR8AN" : _var.dtworkRow02[76].ToString() + "|BR8AN";
                            }
                            _var.dtworkRow02[39] = mort_extra;

                            //3rd Line
                            if (objHlpr.fn_isDMort(mort_wp))
                            {
                                mort_wp = "STANDARD";
                                _var.dtworkRow03[76] = String.IsNullOrEmpty(_var.dtworkRow03[76].ToString()) ? "BR8AN" : _var.dtworkRow03[76].ToString() + "|BR8AN";
                            }
                            _var.dtworkRow03[39] = mort_wp;

                            if (_var.dtworkRow01[21].ToString() == "TRENEW")
                            {
                                _var.dtworkRow01[59] = moul;
                                _var.dtworkRow02[59] = extra;
                                _var.dtworkRow03[59] = wp;
                            }
                            else if (_var.dtworkRow01[21].ToString() == "TNEWBUS")
                            {
                                _var.dtworkRow01[57] = moul;
                                _var.dtworkRow02[57] = extra;
                                _var.dtworkRow03[57] = wp;
                            }
                            else if ((_var.dtworkRow01[21].ToString() == "ADJUST") && (code == "4002"))
                            {
                                _var.dtworkRow01[61] = moul;
                                _var.dtworkRow02[61] = extra;
                                _var.dtworkRow03[61] = wp;
                            }
                            else if ((_var.dtworkRow01[21].ToString() == "ADJUST") && (code == "4004"))
                            {
                                _var.dtworkRow01[63] = moul;
                                _var.dtworkRow02[63] = extra;
                                _var.dtworkRow03[63] = wp;
                            }

                            _var.dtworkRow01[5] = "MOUL";
                            _var.dtworkRow02[5] = "EXTRA";
                            _var.dtworkRow03[5] = "WP";

                            _var.dtworkRow01[22] = month;
                            _var.dtworkRow02[22] = month;
                            _var.dtworkRow03[22] = month;


                            //ISSUE#010-Start---------
                            
                            if (_var.dtworkRow01[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[20];
                                _var.dtworkRow02[19] = _var.dtworkRow02[20];
                                _var.dtworkRow03[19] = _var.dtworkRow03[20];
                            }
                            else
                            {
                                _var.dtworkRow01[19] = _var.dtworkRow01[22];
                                _var.dtworkRow02[19] = _var.dtworkRow02[22];
                                _var.dtworkRow03[19] = _var.dtworkRow03[22];
                            }


                            
                            //ISSUE#010-End-----------

                            _var.dtworkRow01[77] = NAAR;
                            _var.dtworkRow02[77] = NAAR;
                            _var.dtworkRow03[77] = NAAR;

                            #region "computation and substitution"
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

                            _var.dtworkRow02[25] = "1";
                            //_var.dtworkRow02[26] = "1";
                            _var.dtworkRow02[27] = "1";
                            _var.dtworkRow02[28] = "1";
                            _var.dtworkRow02[77] = "1";

                            _var.dtworkRow03[25] = "1";
                            //_var.dtworkRow03[26] = "1";
                            _var.dtworkRow03[27] = "1";
                            _var.dtworkRow03[28] = "1";
                            _var.dtworkRow03[77] = "1";

                            #endregion
                            if (rec_month.Year == iyear && objHlpr.fn_isinQTR(qtr, rec_month.Month))
                            {
                                switch (rec_month.Month)
                                {
                                    case 1:
                                        _var.objdt_template01.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template01.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template01.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                            );
                                        _var.dbl_BH01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ01 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 2:
                                        _var.objdt_template02.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template02.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template02.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ02 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 3:
                                        _var.objdt_template03.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template03.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template03.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ03 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 4:
                                        _var.objdt_template04.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template04.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template04.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ04 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 5:
                                        _var.objdt_template05.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template05.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template05.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ05 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 6:
                                        _var.objdt_template06.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template06.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template06.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ06 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 7:
                                        _var.objdt_template07.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template07.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template07.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ07 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 8:
                                        _var.objdt_template08.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template08.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template08.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ08 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 9:
                                        _var.objdt_template09.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template09.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template09.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ09 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 10:
                                        _var.objdt_template10.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template10.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template10.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ10 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 11:
                                        _var.objdt_template11.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template11.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template11.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ11 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    case 12:
                                        _var.objdt_template12.Rows.Add(_var.dtworkRow01);
                                        _var.objdt_template12.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_template12.Rows.Add(_var.dtworkRow03);

                                        _var.dbl_BF12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                            );
                                        _var.dbl_BH12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                            );
                                        _var.dbl_BJ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                            );
                                        _var.dbl_BL12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                            );
                                        _var.dbl_BZ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                            );

                                        _var.dbl_BF12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                        );
                                        _var.dbl_BH12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                        );
                                        _var.dbl_BH12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ12 += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                        break;
                                    default:
                                        break;
                                }
                            }
                            else 
                            {
                                if (is1stRow) 
                                {
                                    _var.objdt_templateADJ.Rows.Add(_var.dtworkRow01);
                                
                                    _var.dbl_BF_adj += decimal.Parse(
                                        String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                                        );
                                    _var.dbl_BH_adj += decimal.Parse(
                                        String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                                        );
                                    _var.dbl_BJ_adj += decimal.Parse(
                                        String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                                        );
                                    _var.dbl_BL_adj += decimal.Parse(
                                        String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                                        );
                                    _var.dbl_BZ_adj += decimal.Parse(
                                        String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                                        );

                                    string var = wsraw.Cells[prawrow, 17].Text.ToString().Trim().ToUpper();
                                    if (var != "TERMINATED")
                                    {
                                        _var.objdt_templateADJ.Rows.Add(_var.dtworkRow02);
                                        _var.objdt_templateADJ.Rows.Add(_var.dtworkRow03);
                                        _var.dbl_BF_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : _var.dtworkRow02[57].ToString()
                                            );
                                        _var.dbl_BH_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : _var.dtworkRow02[59].ToString()
                                            );
                                        _var.dbl_BJ_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : _var.dtworkRow02[61].ToString()
                                            );
                                        _var.dbl_BL_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : _var.dtworkRow02[63].ToString()
                                            );
                                        _var.dbl_BZ_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : _var.dtworkRow02[77].ToString()
                                            );

                                        _var.dbl_BF_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : _var.dtworkRow03[57].ToString()
                                            );
                                        _var.dbl_BH_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : _var.dtworkRow03[59].ToString()
                                            );
                                        _var.dbl_BJ_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : _var.dtworkRow03[61].ToString()
                                            );
                                        _var.dbl_BL_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : _var.dtworkRow03[63].ToString()
                                            );
                                        _var.dbl_BZ_adj += decimal.Parse(
                                            String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : _var.dtworkRow03[77].ToString()
                                            );
                                    }
                                    else
                                    {
                                        is1stRow = false;
                                    }
                                }
                            }
                            
                            #endregion "computation and substitution"
                        }

                    }

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                    
                    issue = wsraw.Cells[prawrow, 3].Text.ToString();

                    if (!String.IsNullOrEmpty(issue))
                    {
                        issue = issue.Trim();
                        issueyear = issue.Substring(issue.Length - 4, 4);
                    }

                    dob = wsraw.Cells[prawrow, 4].Text.ToString();
                    age = wsraw.Cells[prawrow, 5].Text.ToString();
                    month = wsraw.Cells[prawrow, 15].Text.ToString();
                    orig = wsraw.Cells[prawrow, 6].Text.ToString();
                    risk = wsraw.Cells[prawrow, 8].Text.ToString();
                    ret = wsraw.Cells[prawrow, 7].Text.ToString();
                    premium = wsraw.Cells[prawrow, 21].Text.ToString();
                    sum = wsraw.Cells[prawrow, 17].Text.ToString();
                    gender = string.Empty;
                    rowcount++;
                }
                #endregion
                #region "Compute Hash Total"
                if (_var.objdt_templateADJ.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    _var.dtworkRow = _var.objdt_templateADJ.NewRow();
                    _var.objdt_templateADJ.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_templateADJ.NewRow();
                    _var.dtworkRow[0] = "Total Premium:";
                    _var.dtworkRow[1] = _var.dbl_BF_adj + _var.dbl_BH_adj + _var.dbl_BJ_adj + _var.dbl_BL_adj;
                    _var.objdt_templateADJ.Rows.Add(_var.dtworkRow);

                    _var.dtworkRow = _var.objdt_templateADJ.NewRow();
                    _var.dtworkRow[0] = "Total Sum at Risk:";
                    _var.dtworkRow[1] = _var.dbl_BZ_adj;
                    _var.objdt_templateADJ.Rows.Add(_var.dtworkRow);
                    #endregion

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines_adj != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_templateADJ.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines_adj;
                        _var.objdt_templateADJ.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_templateADJ, boo_open, str_saved + @"\BM025-ADJ-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines01 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template01.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines01;
                        _var.objdt_template01.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template01, boo_open, str_saved + @"\BM025-JAN-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines02 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template02.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines02;
                        _var.objdt_template02.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template02, boo_open, str_saved + @"\BM025-FEB-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines03 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template03.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines03;
                        _var.objdt_template03.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template03, boo_open, str_saved + @"\BM025-MAR-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines04 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template04.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines04;
                        _var.objdt_template04.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template04, boo_open, str_saved + @"\BM025-APR-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines05 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template05.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines05;
                        _var.objdt_template05.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template05, boo_open, str_saved + @"\BM025-MAY-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines06 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template06.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines06;
                        _var.objdt_template06.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template06, boo_open, str_saved + @"\BM025-JUN-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines07 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template07.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines07;
                        _var.objdt_template07.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template07, boo_open, str_saved + @"\BM025-JUL-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines08 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template08.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines08;
                        _var.objdt_template08.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template08, boo_open, str_saved + @"\BM025-AUG-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines09 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template09.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines09;
                        _var.objdt_template09.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template09, boo_open, str_saved + @"\BM025-SEP-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines10 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template10.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines10;
                        _var.objdt_template10.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template10, boo_open, str_saved + @"\BM025-OCT-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines11 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template11.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines11;
                        _var.objdt_template11.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template11, boo_open, str_saved + @"\BM025-NOV-" + str_savef + ".xlsx");
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

                    //ISSUE#013-Start---------
                    #region "List all failed genders"
                    if (_var.str_GFailLines12 != string.Empty)
                    {
                        _var.dtworkRow01 = _var.objdt_template12.NewRow();
                        _var.dtworkRow01[0] = "Gender Fail Lines on RAW:";
                        _var.dtworkRow01[1] = _var.str_GFailLines12;
                        _var.objdt_template12.Rows.Add(_var.dtworkRow01);
                    }
                    #endregion
                    //ISSUE#013-End-----------

                    objHlpr.fn_savemultiple(_var.objdt_template12, boo_open, str_saved + @"\BM025-DEC-" + str_savef + ".xlsx");
                }
                #endregion

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
                return ex.Message + Environment.NewLine + " *****On excel row line: " + rowcount + " *****";
            }

        }
    }
}
