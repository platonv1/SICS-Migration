using System;
using System.Data;
using System.Globalization;
using System.Linq;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM020
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

                int reinss;
                int rett;
                long origg;
                int byear2;
                decimal oriig;
                decimal oriig1;
                decimal oriig2;
                decimal reett1;
                decimal reett2;
                decimal reett;
                decimal inss1;
                decimal reiin1;
                decimal reinn1;
                decimal premm;
                decimal premm11;
                decimal premm22;
                decimal prem4m;
                decimal prem1m;
                long premm1;
                long premm2;
                long premm3;
                long pre1mm;
                long pre4mm;

                string polnum = wsraw.Cells[prawrow, 2].Text.ToString();      
                string branded = wsraw.Cells[prawrow, 3].Text.ToString();
                string businessType = wsraw.Cells [prawrow, 4].Text.ToString();

                string mortality = wsraw.Cells[prawrow, 13].Text.ToString();
                string age = wsraw.Cells[prawrow, 12].Text.ToString();
                string reins = wsraw.Cells[prawrow, 7].Text.ToString();
                string comdate = wsraw.Cells[prawrow, 6].Text.ToString();
                //string trans = wsraw.Cells[prawrow, 6].Text.ToString();
                string rating = wsraw.Cells[prawrow, 14].Text.ToString();
                string occ = wsraw.Cells [prawrow, 15].Text.ToString();
                string orig = wsraw.Cells [prawrow, 16].Text.ToString();
                string orig1 = wsraw.Cells[prawrow, 17].Text.ToString();
                string orig2 = wsraw.Cells[prawrow, 18].Text.ToString();
                string ret2 = wsraw.Cells[prawrow, 19].Text.ToString();
                string ret = wsraw.Cells[prawrow, 20].Text.ToString();
                string ret1 = wsraw.Cells[prawrow, 21].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 10].Text.ToString();
                string gender = wsraw.Cells[prawrow, 9].Text.ToString();
                string dob = wsraw.Cells[prawrow, 11].Text.ToString();
                string prem = wsraw.Cells[prawrow, 25].Text.ToString();
                string ins = wsraw.Cells[prawrow, 23].Text.ToString();
                string rein = wsraw.Cells[prawrow, 22].Text.ToString();
                string reins1 = wsraw.Cells[prawrow, 24].Text.ToString();
                string prem1 = wsraw.Cells[prawrow, 26].Text.ToString();
                string prem2 = wsraw.Cells[prawrow, 27].Text.ToString();
                string prem3 = wsraw.Cells[prawrow, 28].Text.ToString();
                string prem4 = wsraw.Cells[prawrow, 29].Text.ToString();
                string premium = wsraw.Cells[prawrow, 30].Text.ToString();
                string type = wsraw.Cells[prawrow, 4].Text.ToString();
                string expiry = wsraw.Cells[prawrow, 8].Text.ToString();
                string byear = wsraw.Cells[prawrow, 1][4].Text.ToString();

                byear = byear.Substring(byear.Length - 7, 4);
                byear2 = Convert.ToInt32(byear);

                //string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };

                int storee;
                bool chck;
                double mul = 0.15;
                long origg2;
                long origg1;
                long reinss1;
                long rett1;
                long ins1;
                long reet2;
                int transs;
                string trans11;
                string byears;
                int byears1;

                #region Data Processing
                while (rowcount != erawrow + 1)
                {
                    chck = false;
                    string[] polnum1 = polnum.Split('-');

                    if (polnum1.Length == 2)
                    {
                        if (polnum1[0].ToUpper().Contains("G") && int.TryParse(polnum1[1], out storee))
                        {
                            chck = true;
                        }
                    }

                    if (chck == true)
                    {
                        _var.dtworkRow01 = objdt_template.NewRow();
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow03 = objdt_template.NewRow();

                        _var.dtworkRow01[0] = polnum;
                        if (branded.ToUpper() == "GCLI")
                        {
                            branded = "GMRI";
                        }
                        _var.dtworkRow01[5] = branded;
                        _var.dtworkRow01[7] = polnum;
                        _var.dtworkRow01[8] = "SURPLUS";
                        _var.dtworkRow01[9] = "PAFM"; //PAFW
                        _var.dtworkRow01[13] = "GRP";

                        _var.dtworkRow01 [14] = objHlpr.fn_checkBusinessTypeV2(businessType);
                        _var.dtworkRow01[24] = "Q";
                        _var.dtworkRow01[29] = "NATREID";
                        _var.dtworkRow01[23] = "PHP";
                        _var.dtworkRow01[24] = "YLY";
                        //_var.dtworkRow01[19] = comdate;
                        _var.dtworkRow01[20] = comdate;
                        _var.dtworkRow01[28] = ret;
                        _var.dtworkRow01[41] = byear;
                        _var.dtworkRow01[29] = "NATREID";
                        _var.dtworkRow01[31] = fullname;
                        _var.dtworkRow01[38] = "NONE";

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

                        _var.dtworkRow01[10] = "S";
                        _var.dtworkRow01[26] = "1.00";
                        _var.dtworkRow01[78] = age;
                        _var.dtworkRow01[40] = expiry;

                        mortality = mortality.ToUpper();
                        rating = rating.ToUpper();

                        if (String.IsNullOrEmpty(mortality))
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                        }
                        else if (mortality == "STANDARD")
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                        }
                        else if (mortality.Contains("SUBSTANDARD"))
                        {
                            if (String.IsNullOrEmpty(rating))
                            {
                                _var.dtworkRow01[39] = "SUBSTANDARD";
                                _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                            }
                            else if (rating.Contains("CLASS A"))
                            {
                                _var.dtworkRow01[39] = "CLASSA";
                            }
                            else if (rating.Contains("CLASS AA"))
                            {
                                _var.dtworkRow01[39] = "CLASSAA";
                            }
                            else if (rating.Contains("CLASS B"))
                            {
                                _var.dtworkRow01[39] = "CLASSB";
                            }
                            else if (rating.Contains("CLASS C"))
                            {
                                _var.dtworkRow01[39] = "CLASSC";
                            }
                            else if (rating.Contains("CLASS D"))
                            {
                                _var.dtworkRow01[39] = "CLASSD";
                            }
                            else if (rating.Contains("CLASS E"))
                            {
                                _var.dtworkRow01[39] = "CLASSE";
                            }
                            else if (rating.Contains("CLASS F"))
                            {
                                _var.dtworkRow01[39] = "CLASSF";
                            }
                            else if (rating.Contains("CLASS G"))
                            {
                                _var.dtworkRow01[39] = "CLASSG";
                            }
                            else if (rating.Contains("CLASS H"))
                            {
                                _var.dtworkRow01[39] = "CLASSH";
                            }
                            else if (rating.Contains("CLASS I"))
                            {
                                _var.dtworkRow01[39] = "CLASSI";
                            }
                            else if (rating.Contains("CLASS J"))
                            {
                                _var.dtworkRow01[39] = "CLASSJ";
                            }
                            else if (rating.Contains("CLASS K"))
                            {
                                _var.dtworkRow01[39] = "CLASSK";
                            }
                            else if (rating.Contains("CLASS L"))
                            {
                                _var.dtworkRow01[39] = "CLASSL";
                            }
                            else if (rating.Contains("CLASS M"))
                            {
                                _var.dtworkRow01[39] = "CLASSM";
                            }
                            else if (rating.Contains("CLASS N"))
                            {
                                _var.dtworkRow01[39] = "CLASSN";
                            }
                            else if (rating.Contains("CLASS O"))
                            {
                                _var.dtworkRow01[39] = "CLASSO";
                            }
                            else if (rating.Contains("CLASS P"))
                            {
                                _var.dtworkRow01[39] = "CLASSP";
                            }
                            else
                            {
                                _var.dtworkRow01[39] = "SUBSTANDARD";
                                _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                            }
                        }
                        else 
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                        }

                        if (!String.IsNullOrEmpty(occ))
                        {
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? occ : _var.dtworkRow01[76].ToString() + "|" + occ;
                        }

                        premm = Convert.ToDecimal(prem);
                        premm1 = Convert.ToInt64(premm);

                        prem1m = Convert.ToDecimal(prem1);
                        pre1mm = Convert.ToInt64(prem1m);

                        prem4m = Convert.ToDecimal(prem4);
                        pre4mm = Convert.ToInt64(prem4m);

                        premm11 = Convert.ToDecimal(prem2);
                        premm2 = Convert.ToInt64(premm11);

                        premm22 = Convert.ToDecimal(prem3);
                        premm3 = Convert.ToInt64(premm22);

                        
                        if (comdate.Length == 7)
                        {
                            _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 7, 5) + byear2;
                        }
                        if (comdate.Length == 6)
                        {
                            _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 6, 4) + byear2;
                        }
                        if (comdate.Length == 8)
                        {
                            if (comdate.Contains("20"))
                            {
                                _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 8, 4) + byear2;
                            }
                            else
                            {
                                _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 8, 6) + byear2;
                            }
                        }

                        if (comdate.Length == 9)
                        {
                            _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 9, 5) + byear2;
                        }
                        if (comdate.Length == 10)
                        {
                            _var.dtworkRow01[22] = comdate.Substring(comdate.Length - 10, 6) + byear2;
                        }
                        double sum;
                        double sum1;
                        trans11 = comdate.Substring(comdate.Length - 2, 2);
                        transs = Convert.ToInt32(trans11);
                        byears = Convert.ToString(byear2);
                        byears = byears.Substring(byears.Length - 2, 2);
                        byears1 = Convert.ToInt32(byears);
                        double premium1;
                        premium1 = Convert.ToDouble(premium);
                        string str_transcode = "";

                        if (transs >= byears1) 
                        {
                            _var.dtworkRow01[21] = "TNEWBUS";
                            str_transcode = "TNEWBUS";
                            _var.dtworkRow01[56] = "4000";
                            sum = premium1;
                            sum1 = sum * mul;
                            _var.dtworkRow01[57] = sum1;
                        }
                        else
                        {
                            _var.dtworkRow01[21] = "TRENEW";
                            str_transcode = "TRENEW";
                            _var.dtworkRow01[58] = "4001";
                            sum = premium1 * mul;
                            _var.dtworkRow01[59] = sum;
                        }

                        if (branded.Contains("GYRT"))
                        {
                            if (transs >= byears1)
                            {
                                _var.dtworkRow01[21] = "TNEWBUS";
                                str_transcode = "TNEWBUS";
                                _var.dtworkRow01[56] = "4000";
                                sum = premium1;
                                sum1 = premm1 + pre1mm + pre4mm;
                                sum1 = sum1 * mul;
                                _var.dtworkRow01[57] = sum1;
                            }
                            else
                            {
                                _var.dtworkRow01[21] = "TRENEW";
                                str_transcode = "TRENEW";
                                _var.dtworkRow01[58] = "4001";
                                sum = premm1 + pre1mm + pre4mm;
                                sum = sum * mul;
                                _var.dtworkRow01[59] = sum;
                            }
                        }

                        if (str_transcode == "TNEWBUS")
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[20];
                        }
                        else if (str_transcode == "TRENEW")
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        }
                            #region "New Requirements - No Name"
                            if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                        }
                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out _var.str_outfname, out _var.str_outlname, out _var.str_outlifeid);

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
                            _var.dtworkRow01[33] = _var.str_outfname;
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
                            gender = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                            _var.dtworkRow01[36] = gender;
                            //ISSUE#003-Start---------
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                            //ISSUE#003-End-----------
                        }
                        else if (String.IsNullOrEmpty(gender) && String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow01[36] = string.Empty;
                        }

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
                       
                            if (!string.IsNullOrEmpty(gender))
                            {
                                _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + gender.Substring(0, 1);
                            }
                            else
                            {
                                _var.dtworkRow01[1] = _var.dtworkRow01[0].ToString() + "-";
                            }
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

                        


                        #endregion

                        orig = orig.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(orig = string.IsNullOrEmpty(orig) ? "0" : orig);
                        oriig = Convert.ToDecimal(orig);
                        origg = Convert.ToInt64(oriig);

                        orig1 = orig1.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(orig1 = string.IsNullOrEmpty(orig1) ? "0" : orig1);
                        oriig1 = Convert.ToDecimal(orig1);
                        origg1 = Convert.ToInt64(oriig1);

                        orig2 = orig2.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(orig2 = string.IsNullOrEmpty(orig2) ? "0" : orig2);
                        oriig2 = Convert.ToDecimal(orig2);
                        origg2 = Convert.ToInt64(oriig2);

                        rein = rein.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(rein = string.IsNullOrEmpty(rein) ? "0" : rein);
                        reinn1 = Convert.ToDecimal(rein);
                        reinss = Convert.ToInt32(reinn1);

                        reins1 = reins1.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0").Replace("(", String.Empty).Replace(")", String.Empty);
                        Convert.ToDecimal(reins1 = string.IsNullOrEmpty(reins1) ? "0" : reins1);
                        reiin1 = Convert.ToDecimal(reins1);
                        reinss1 = Convert.ToInt32(reiin1);

                        ins = ins.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0").Replace("(", String.Empty).Replace(")", String.Empty);
                        Convert.ToDecimal(ins = string.IsNullOrEmpty(ins) ? "0" : ins);
                        inss1 = Convert.ToDecimal(ins);
                        ins1 = Convert.ToInt64(inss1);

                        ret2 = ret2.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(ret2 = string.IsNullOrEmpty(ret2) ? "0" : ret2);
                        reett2 = Convert.ToDecimal(ret2);
                        rett = Convert.ToInt32(reett2);

                        ret = ret.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(ret = string.IsNullOrEmpty(ret) ? "0" : ret);
                        reett = Convert.ToDecimal(ret);
                        rett1 = Convert.ToInt64(reett);

                        ret1 = ret1.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace(".00", String.Empty).Replace("-", "0");
                        Convert.ToDecimal(ret1 = string.IsNullOrEmpty(ret1) ? "0" : ret1);
                        reett1 = Convert.ToDecimal(ret1);
                        reet2 = Convert.ToInt32(reett1);

                        _var.dtworkRow01[25] = origg * 1;

                        if (ins1 == (0) && (reinss1 == (0)))
                        {
                            _var.dtworkRow01[27] = reinss * mul;
                            _var.dtworkRow01[28] = rett * 1;
                            _var.dtworkRow01[77] = reinss * mul;
                            _var.dtworkRow02 = null;
                            _var.dtworkRow03 = null;
                        } 
                        else
                        {
                            _var.dtworkRow01[27] = reinss * mul;
                            _var.dtworkRow01[28] = rett * 1;
                            _var.dtworkRow01[77] = reinss * mul;
                            _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;
                            _var.dtworkRow02[5] = "TPD";
                            _var.dtworkRow02[25] = origg1 * 1;
                            _var.dtworkRow02[26] = "1.00";
                            _var.dtworkRow02[28] = rett1 * 1;

                            if (comdate.Length == 7)
                            {
                                _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 7, 5) + byear2;
                            }
                            if (comdate.Length == 6)
                            {
                                _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 6, 4) + byear2;
                            }
                            if (comdate.Length == 8)
                            {
                                if (comdate.Contains("20"))
                                {
                                    _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 8, 4) + byear2;
                                }
                                else
                                {
                                    _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 8, 6) + byear2;
                                }
                            }
                            if (comdate.Length == 9)
                            {
                                _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 9, 5) + byear2;
                            }
                            if (comdate.Length == 10)
                            {
                                _var.dtworkRow02[22] = comdate.Substring(comdate.Length - 10, 6) + byear2;
                            }

                            trans11 = comdate.Substring(comdate.Length - 2, 2);
                            transs = Convert.ToInt32(trans11);
                            byears = Convert.ToString(byear2);
                            byears = byears.Substring(byears.Length - 2, 2);
                            byears1 = Convert.ToInt32(byears);
                            if (transs >= byears1)
                            {
                                _var.dtworkRow02[21] = "TNEWBUS";
                                _var.dtworkRow02[56] = "4000";
                                _var.dtworkRow02[57] = premm2 * mul;
                            }
                            else
                            {
                                _var.dtworkRow02[21] = "TRENEW";
                                _var.dtworkRow02[58] = "4001";
                                _var.dtworkRow02[59] = premm2 * mul;
                            }

                            if (branded.Contains("GYRT"))
                            {
                                if (transs >= byears1)
                                {
                                    _var.dtworkRow02[21] = "TNEWBUS";
                                    _var.dtworkRow02[56] = "4000";
                                    sum = premm2;
                                    sum = sum * mul;
                                    _var.dtworkRow02[57] = sum;
                                }
                                else
                                {
                                    _var.dtworkRow02[21] = "TRENEW";
                                    _var.dtworkRow02[58] = "4001";
                                    sum = premm2;
                                    sum = sum * mul;
                                    _var.dtworkRow02[59] = sum;

                                }
                            }
                            _var.dtworkRow02[77] = ins1 * mul;

                            _var.dtworkRow03.ItemArray = _var.dtworkRow01.ItemArray;
                            _var.dtworkRow03[5] = "ADD";
                            _var.dtworkRow03[27] = reinss1 * mul;
                            _var.dtworkRow03[28] = reet2 * 1;

                            if (comdate.Length == 7)
                            {
                                _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 7, 5) + byear2;
                            }
                            if (comdate.Length == 6)
                            {
                                _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 6, 4) + byear2;
                            }
                            if (comdate.Length == 8)
                            {
                                if (comdate.Contains("20"))
                                {
                                    _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 8, 4) + byear2;
                                }
                                else
                                {
                                    _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 8, 6) + byear2;
                                }
                            }
                            if (comdate.Length == 9)
                            {
                                _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 9, 5) + byear2;
                            }
                            if (comdate.Length == 10)
                            {
                                _var.dtworkRow03[22] = comdate.Substring(comdate.Length - 10, 6) + byear2;
                            }

                            trans11 = comdate.Substring(comdate.Length - 2, 2);
                            transs = Convert.ToInt32(trans11);
                            byears = Convert.ToString(byear2);
                            byears = byears.Substring(byears.Length - 2, 2);
                            byears1 = Convert.ToInt32(byears);

                            if (transs >= byears1)
                            {
                                _var.dtworkRow03[21] = "TNEWBUS";
                                _var.dtworkRow03[56] = "4000";
                                _var.dtworkRow03[57] = premm3 * mul;
                            }
                            else
                            {
                                _var.dtworkRow03[21] = "TRENEW";
                                _var.dtworkRow03[58] = "4001";
                                _var.dtworkRow03[59] = premm3 * mul;
                            }

                            if (branded.Contains("GYRT"))
                            {
                                if (transs >= byears1)
                                {
                                    _var.dtworkRow03[21] = "TNEWBUS";
                                    _var.dtworkRow03[56] = "4000";
                                    sum = premm3;
                                    sum = sum * mul;
                                    _var.dtworkRow03[57] = sum;
                                }
                                else
                                {
                                    _var.dtworkRow03[21] = "TRENEW";
                                    _var.dtworkRow03[58] = "4001";
                                    sum = premm3;
                                    sum = sum * mul;
                                    _var.dtworkRow03[59] = sum;
                                }

                            }
                            _var.dtworkRow03[77] = reinss1 * mul;
                        }

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

                        if (_var.dtworkRow02 != null)
                        {
                            if (_var.dtworkRow02[25].ToString() == "0")
                            {
                                _var.dtworkRow02[25] = "1";
                            }
                            if (_var.dtworkRow02[26].ToString() == "0")
                            {
                                _var.dtworkRow02[26] = "1";
                            }
                            if (_var.dtworkRow02[27].ToString() == "0")
                            {
                                _var.dtworkRow02[27] = "1";
                            }
                            if (_var.dtworkRow02[28].ToString() == "0")
                            {
                                _var.dtworkRow02[28] = "1";
                            }
                            if (_var.dtworkRow02[77].ToString() == "0")
                            {
                                _var.dtworkRow02[77] = "1";
                            }
                        }

                        if (_var.dtworkRow03 != null)
                        {
                            if (_var.dtworkRow03[25].ToString() == "0")
                            {
                                _var.dtworkRow03[25] = "1";
                            }
                            if (_var.dtworkRow03[26].ToString() == "0")
                            {
                                _var.dtworkRow03[26] = "1";
                            }
                            if (_var.dtworkRow03[27].ToString() == "0")
                            {
                                _var.dtworkRow03[27] = "1";
                            }
                            if (_var.dtworkRow03[28].ToString() == "0")
                            {
                                _var.dtworkRow03[28] = "1";
                            }
                            if (_var.dtworkRow03[77].ToString() == "0")
                            {
                                _var.dtworkRow03[77] = "1";
                            }
                        }
                        //ISSUE#017-End-----------

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

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 2].Text.ToString();
                    branded = wsraw.Cells[prawrow, 3].Text.ToString();
                    
                    reins = wsraw.Cells[prawrow, 7].Text.ToString();
                    comdate = wsraw.Cells[prawrow, 5].Text.ToString();
                    orig = wsraw.Cells[prawrow, 16].Text.ToString();
                    occ = wsraw.Cells[prawrow, 15].Text.ToString();
                    orig1 = wsraw.Cells[prawrow, 17].Text.ToString();
                    orig2 = wsraw.Cells[prawrow, 18].Text.ToString();
                    ret2 = wsraw.Cells[prawrow, 19].Text.ToString();
                    ret = wsraw.Cells[prawrow, 20].Text.ToString();
                    ret1 = wsraw.Cells[prawrow, 21].Text.ToString();
                    rein = wsraw.Cells[prawrow, 22].Text.ToString();
                    reins1 = wsraw.Cells[prawrow, 24].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 10].Text.ToString();
                    //trans = wsraw.Cells[prawrow, 6].Text.ToString();
                    gender = wsraw.Cells[prawrow, 9].Text.ToString();
                    dob = wsraw.Cells[prawrow, 11].Text.ToString();
                    prem = wsraw.Cells[prawrow, 25].Text.ToString();
                    prem1 = wsraw.Cells[prawrow, 26].Text.ToString();
                    prem2 = wsraw.Cells[prawrow, 27].Text.ToString();
                    prem3 = wsraw.Cells[prawrow, 28].Text.ToString();
                    prem4 = wsraw.Cells[prawrow, 29].Text.ToString();
                    age = wsraw.Cells[prawrow, 12].Text.ToString();
                    ret1 = wsraw.Cells[prawrow, 21].Text.ToString();
                    rating = wsraw.Cells[prawrow, 14].Text.ToString();
                    mortality = wsraw.Cells[prawrow, 13].Text.ToString();
                    ins = wsraw.Cells[prawrow, 23].Text.ToString();
                    expiry = wsraw.Cells[prawrow, 8].Text.ToString();
                    premium = wsraw.Cells[prawrow, 30].Text.ToString();
                    rowcount++;
                }
                #endregion
                _var.dtworkRow01 = null; //Dispose datarow
                _var.dtworkRow02 = null; //Dispose datarow
                _var.dtworkRow03 = null; //Dispose datarow

                #region "Compute Hash Total"
                System.Data.DataRow dtworkRow;
                dtworkRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Premium:";
                dtworkRow[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL;
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Sum at Risk:";
                dtworkRow[1] = _var.dbl_BZ;
                objdt_template.Rows.Add(dtworkRow);
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

                string despath = str_saved + @"\BM020-" + str_sheet + "-" + str_savef + ".xlsx";
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
