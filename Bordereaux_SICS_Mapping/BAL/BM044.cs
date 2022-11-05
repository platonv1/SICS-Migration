using System;
using System.Data;
using System.Linq;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM044
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
                HelperV21 objHlpr2 = new HelperV21();
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

                string strFirstName = string.Empty;
                string strLastName = string.Empty;
                string strMiddleInitial = string.Empty;

                int erawrow = rawrange.Rows.Count;
                int erawcol = rawrange.Columns.Count;
                int prawrow = 1;

                int storee;
               
                string busmean = "";
                string newPol = string.Empty;
                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                string gender = wsraw.Cells[prawrow, 3].Text.ToString();
                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string smoker = wsraw.Cells[prawrow, 5].Text.ToString();
                string age = wsraw.Cells[prawrow, 6].Text.ToString(); // Issue age
                string extra = wsraw.Cells[prawrow, 7].Text.ToString(); // multiple extra
                string bustype = wsraw.Cells[prawrow, 8].Text.ToString(); //treaty
                string sum = wsraw.Cells[prawrow, 13].Text.ToString(); //current
                string sum1 = wsraw.Cells[prawrow, 14].Text.ToString(); //current
                string orig = wsraw.Cells[prawrow, 10].Text.ToString(); //orig face
                string curr = wsraw.Cells[prawrow, 11].Text.ToString(); //curr face
                string term = wsraw.Cells[prawrow, 16].Text.ToString(); // termination date
                string paid = wsraw.Cells[prawrow, 14].Text.ToString(); // paid to date
                string code = wsraw.Cells[prawrow, 18].Text.ToString(); // plan code 
                string premium = wsraw.Cells[prawrow, 17].Text.ToString(); // plan code 
                string period = wsraw.Cells[prawrow, 27].Text.ToString(); // premuim for the period
                string remarks = wsraw.Cells[prawrow, 28].Text.ToString(); //comment
                string trans = wsraw.Cells[prawrow, 15].Text.ToString(); //trans
                string lsci = wsraw.Cells[prawrow, 19].Text.ToString(); //lsci
                string esci = wsraw.Cells[prawrow, 20].Text.ToString(); //esci
                string angio = wsraw.Cells[prawrow, 21].Text.ToString(); //angio
                string cancer = wsraw.Cells[prawrow, 22].Text.ToString(); //cancer
                string hospc = wsraw.Cells[prawrow, 23].Text.ToString(); //hosp
                string posth = wsraw.Cells[prawrow, 24].Text.ToString(); //posth
                string homer = wsraw.Cells[prawrow, 25].Text.ToString(); //homer
                string pall = wsraw.Cells[prawrow, 26].Text.ToString(); //pall
                string bmyear = wsraw.Cells[4, 2].Text.ToString();
                bmyear = bmyear.Substring(bmyear.Length - 4);

                string type = wsraw.Cells[prawrow, 31].Text.ToString();
                string fyry = wsraw.Cells[prawrow, 29].Text.ToString();

                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;


                bool chck;
                string prefstore;
                int namectr = 0;
                int lnamectr = 0;
                bool lst;


                #region Data Processing
                while (rowcount != erawrow + 1) //loop
                {
                    //chck = int.TryParse(polnum, out storee);
                    chck = objHlpr.fn_policyNumChecker(polnum, wsraw.Cells[prawrow, 2].Text.ToString(), wsraw.Cells[prawrow, 3].Text.ToString(), wsraw.Cells[prawrow, 4].Text.ToString());
                    if (chck == true)
                    {
                        _var.dtworkRow01 = objdt_template.NewRow();
                        _var.dtworkRow02 = objdt_template.NewRow();
                        _var.dtworkRow03 = objdt_template.NewRow();
                        _var.dtworkRow04 = objdt_template.NewRow();
                        _var.dtworkRow05 = objdt_template.NewRow();
                        _var.dtworkRow06 = objdt_template.NewRow();
                        _var.dtworkRow07 = objdt_template.NewRow();
                        _var.dtworkRow08 = objdt_template.NewRow();

                        _var.dtworkRow01[0] = "'" + polnum;
                        _var.dtworkRow01[1] = "'" + polnum;
                        _var.dtworkRow01[3] = "CI";
                        _var.dtworkRow01[4] = "CIRAND";
                        _var.dtworkRow01[8] = "QA";
                        _var.dtworkRow01[9] = "PAFM";
                        _var.dtworkRow01[10] = "Q";
                        _var.dtworkRow01[13] = "IND";
                        _var.dtworkRow01[24] = "YLY";
                        _var.dtworkRow01[29] = "NATREID";
                        _var.dtworkRow01[5] = code;
                        _var.dtworkRow01[14] = busmean;

                        _var.dtworkRow01[23] = "PHP";
                        _var.dtworkRow01[31] = fullnames;
                        _var.dtworkRow01[36] = gender;


                        _var.dtworkRow01[41] = bmyear;
                        _var.dtworkRow01[20] = trans;
                        _var.dtworkRow01[22] = trans.Substring(0, trans.Length - 4) + bmyear;


                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        _var.dtworkRow01[37] = dob.ToString();
                        _var.dtworkRow01[38] = smoker;
                        _var.dtworkRow01[76] = remarks;
                        _var.dtworkRow01[40] = term;
                        _var.dtworkRow01[79] = age;
                        _var.dtworkRow01[77] = objHlpr.fn_numbercleanup_negative(sum);

                        _var.dtworkRow01[25] = orig;

                        _var.dtworkRow01[83] = "NR";

                        if (type.ToUpper() == "NB")
                        {
                            TRANCODE = "TNEWBUS";
                        }
                        else if ((type.ToUpper() == "REN") & (fyry.Trim() != "FY"))
                        {
                            TRANCODE = "TRENEW";
                        }
                        else if ((type.ToUpper() == "REN") & (fyry.Trim() == "FY"))
                        {
                            TRANCODE = "TNEWBUS";
                        }
                        else if (type.ToUpper() == "RECOV")
                        {
                            TRANCODE = "ADJUST";
                        }

                        _var.dtworkRow01[21] = TRANCODE;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow01[56] = "4000";
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow01[58] = "4001";
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow01[60] = "4002";
                            }
                            else
                            {
                                _var.dtworkRow01[62] = "4004";
                            }
                        }


                        //business type
                        if (bustype == "F")
                        {
                            busmean = "F";
                        }
                        else if (bustype == "A")
                        {
                            busmean = "T";
                        }
                        _var.dtworkRow01[14] = busmean;

                        if (smoker.ToString() == "N")
                        {
                            _var.dtworkRow01[38] = "NSMOK";
                        }
                        else
                        {
                            _var.dtworkRow01[38] = "SMOK";
                        }


                        prefstore = extra.Replace("%", string.Empty);
                        _var.dtworkRow01[39] = objHlpr.fn_getmortality((int.Parse(prefstore) + 100).ToString());
                        if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                        {
                            _var.dtworkRow01[39] = "STANDARD";
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                        }

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullnames))
                        {
                            fullnames = polnum.ToString();
                            _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        //objHlpr.fn_getnamesandlifeID(fullnames, dob, out string str_outfname, out string str_outlname, out string str_outlifeid);

                        //string str_MI = string.Empty;
                        //string[] arr_fullname;
                        //arr_fullname = fullnames.Split(',');
                        //str_outlname = arr_fullname[0];

                        //if (arr_fullname.Count() > 1)
                        //{
                        //    str_outfname = arr_fullname[1];
                        //}

                        //if (arr_fullname.Count() > 2)
                        //{
                        //    str_MI = arr_fullname[2];
                        //    _var.dtworkRow01[34] = str_MI;
                        //}fn_separateLastNameFirstNameV2

                        _var.dtworkRow01 [31] = fullnames; /*objHlpr.fn_stringcleanup(fullnames);*/

                        objHlpr2.fn_separateLastNameFirstNameV4(fullnames, out fullnames, out  strLastName, out  strFirstName, out  strMiddleInitial);
                        
                        _var.dtworkRow01 [32] = objHlpr2.fn_removeCharacters(strLastName);/*str_outlname;*/

                        _var.dtworkRow01 [33] = objHlpr2.fn_removeCharacters(strFirstName);/*str_outfname.Replace(" " + str_MI, string.Empty);*/
                        
                        _var.dtworkRow01 [30] = objHlpr.fn_LifeID(strFirstName, strLastName, dob);/*str_outlifeid;*/
                        _var.dtworkRow01 [34] = strMiddleInitial;
                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow01[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                        }

                        _var.dtworkRow01[27] = sum.ToString();

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow01[57] = lsci;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow01[59] = lsci;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow01[61] = lsci;
                            }
                            else
                            {
                                _var.dtworkRow01[63] = lsci;
                            }
                        }


                        _var.dtworkRow01[76] = "LSCI PREMIUM";
                        _var.dtworkRow01[3] = "CI";
                        _var.dtworkRow01[4] = "SACIENDSTAPIND";
                        _var.dtworkRow01[6] = "BP279";

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

                        var parsedDOB = DateTime.Parse(dob);
                        string initialNR = string.Empty;
                        if (!String.IsNullOrEmpty(strFirstName))
                        {
                            initialNR = strFirstName.Substring(0, 1);
                        }
                        if (!String.IsNullOrEmpty(strLastName))
                        {
                            initialNR += strLastName.Substring(0, 1);
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
                            else
                            {
                                _var.dtworkRow01[19] = premium;
                                _var.dtworkRow01[22] = premium;
                            }
                        }
                        //ISSUE#010-End-----------

                        

                        #endregion

                        _var.dtworkRow02.ItemArray = _var.dtworkRow01.ItemArray;


                        _var.dtworkRow02[27] = sum.ToString();

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow02[57] = esci;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow02[59] = esci;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow02[61] = esci;
                            }
                            else
                            {
                                _var.dtworkRow02[63] = esci;
                            }
                        }

                        _var.dtworkRow02[76] = "ESCI PREMIUM";
                        _var.dtworkRow02[3] = "CI";
                        _var.dtworkRow02[4] = "SACIESPIND";
                        _var.dtworkRow02[6] = "BP278";

                        _var.dtworkRow03.ItemArray = _var.dtworkRow01.ItemArray;

                        _var.dtworkRow03[27] = sum.ToString();

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow03[57] = angio;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow03[59] = angio;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow03[61] = angio;
                            }
                            else
                            {
                                _var.dtworkRow03[63] = angio;
                            }
                        }


                        _var.dtworkRow03[76] = "ANGIOPLASTY PREMIUM";
                        _var.dtworkRow03[3] = "CI";
                        _var.dtworkRow03[4] = "SACIESPIND";
                        _var.dtworkRow03[6] = "BP277";

                        _var.dtworkRow04.ItemArray = _var.dtworkRow01.ItemArray;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow04[57] = cancer;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow04[59] = cancer;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow04[61] = cancer;
                            }
                            else
                            {
                                _var.dtworkRow04[63] = cancer;
                            }
                        }

                        _var.dtworkRow04[76] = "CANCER BOOSTER PREMIUM";
                        _var.dtworkRow04[3] = "CI";
                        _var.dtworkRow04[4] = "STANDALONEENH";
                        _var.dtworkRow04[6] = "BP280";
                        _var.dtworkRow04[27] = sum.ToString();

                        _var.dtworkRow05.ItemArray = _var.dtworkRow01.ItemArray;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow05[57] = hospc;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow05[59] = hospc;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow05[61] = hospc;
                            }
                            else
                            {
                                _var.dtworkRow05[63] = hospc;
                            }
                        }

                        _var.dtworkRow05[76] = "HOSPITAL CONFINEMENT PREMIUM";
                        _var.dtworkRow05[3] = "MEDICAL";
                        _var.dtworkRow05[4] = "DHIBILIND";
                        _var.dtworkRow05[6] = "BP281";
                        _var.dtworkRow05[27] = objHlpr.fn_numbercleanup_negative(sum1);

                        _var.dtworkRow06.ItemArray = _var.dtworkRow01.ItemArray;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow06[57] = posth;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow06[59] = posth;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow06[61] = posth;
                            }
                            else
                            {
                                _var.dtworkRow06[63] = posth;
                            }
                        }

                        _var.dtworkRow06[76] = "POST-HOSPITAL PREMIUM";
                        _var.dtworkRow06[3] = "MEDICAL";
                        _var.dtworkRow06[4] = "DHIBILIND";
                        _var.dtworkRow06[6] = "BP282";
                        _var.dtworkRow06[27] = objHlpr.fn_numbercleanup_negative(sum1);


                        _var.dtworkRow07.ItemArray = _var.dtworkRow01.ItemArray;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow07[57] = homer;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow07[59] = homer;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow07[61] = homer;
                            }
                            else
                            {
                                _var.dtworkRow07[63] = homer;
                            }
                        }

                        _var.dtworkRow07[76] = "HOME RECOVERY PREMIUM";
                        _var.dtworkRow07[3] = "MEDICAL";
                        _var.dtworkRow07[4] = "MEDICALREIMBURS";
                        _var.dtworkRow07[6] = "BP283";
                        _var.dtworkRow07[27] = objHlpr.fn_numbercleanup_negative(sum1);


                        _var.dtworkRow08.ItemArray = _var.dtworkRow01.ItemArray;


                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            _var.dtworkRow08[57] = pall;
                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            _var.dtworkRow08[59] = pall;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow08[61] = pall;
                            }
                            else
                            {
                                _var.dtworkRow08[63] = pall;
                            }
                        }


                        _var.dtworkRow08[76] = "PALLIATIVE CARE PREMIUM";
                        _var.dtworkRow08[3] = "MEDICAL";
                        _var.dtworkRow08[4] = "MEDICALREIMBURS";
                        _var.dtworkRow08[6] = "BP284";
                        _var.dtworkRow08[27] = objHlpr.fn_numbercleanup_negative(sum1);


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

                        if (_var.dtworkRow04[25].ToString() == "0")
                        {
                            _var.dtworkRow04[25] = "1";
                        }
                        if (_var.dtworkRow04[26].ToString() == "0")
                        {
                            _var.dtworkRow04[26] = "1";
                        }
                        if (_var.dtworkRow04[27].ToString() == "0")
                        {
                            _var.dtworkRow04[27] = "1";
                        }
                        if (_var.dtworkRow04[28].ToString() == "0")
                        {
                            _var.dtworkRow04[28] = "1";
                        }
                        if (_var.dtworkRow04[77].ToString() == "0")
                        {
                            _var.dtworkRow04[77] = "1";
                        }

                        if (_var.dtworkRow05[25].ToString() == "0")
                        {
                            _var.dtworkRow05[25] = "1";
                        }
                        if (_var.dtworkRow05[26].ToString() == "0")
                        {
                            _var.dtworkRow05[26] = "1";
                        }
                        if (_var.dtworkRow05[27].ToString() == "0")
                        {
                            _var.dtworkRow05[27] = "1";
                        }
                        if (_var.dtworkRow05[28].ToString() == "0")
                        {
                            _var.dtworkRow05[28] = "1";
                        }
                        if (_var.dtworkRow05[77].ToString() == "0")
                        {
                            _var.dtworkRow05[77] = "1";
                        }

                        if (_var.dtworkRow06[25].ToString() == "0")
                        {
                            _var.dtworkRow06[25] = "1";
                        }
                        if (_var.dtworkRow06[26].ToString() == "0")
                        {
                            _var.dtworkRow06[26] = "1";
                        }
                        if (_var.dtworkRow06[27].ToString() == "0")
                        {
                            _var.dtworkRow06[27] = "1";
                        }
                        if (_var.dtworkRow06[28].ToString() == "0")
                        {
                            _var.dtworkRow06[28] = "1";
                        }
                        if (_var.dtworkRow06[77].ToString() == "0")
                        {
                            _var.dtworkRow06[77] = "1";
                        }

                        if (_var.dtworkRow07[25].ToString() == "0")
                        {
                            _var.dtworkRow07[25] = "1";
                        }
                        if (_var.dtworkRow07[26].ToString() == "0")
                        {
                            _var.dtworkRow07[26] = "1";
                        }
                        if (_var.dtworkRow07[27].ToString() == "0")
                        {
                            _var.dtworkRow07[27] = "1";
                        }
                        if (_var.dtworkRow07[28].ToString() == "0")
                        {
                            _var.dtworkRow07[28] = "1";
                        }
                        if (_var.dtworkRow07[77].ToString() == "0")
                        {
                            _var.dtworkRow07[77] = "1";
                        }

                        if (_var.dtworkRow08[25].ToString() == "0")
                        {
                            _var.dtworkRow08[25] = "1";
                        }
                        if (_var.dtworkRow08[26].ToString() == "0")
                        {
                            _var.dtworkRow08[26] = "1";
                        }
                        if (_var.dtworkRow08[27].ToString() == "0")
                        {
                            _var.dtworkRow08[27] = "1";
                        }
                        if (_var.dtworkRow08[28].ToString() == "0")
                        {
                            _var.dtworkRow08[28] = "1";
                        }
                        if (_var.dtworkRow08[77].ToString() == "0")
                        {
                            _var.dtworkRow08[77] = "1";
                        }
                        //ISSUE#017-End-----------

                        _var.dtworkRow01[77] = _var.dtworkRow01[27];
                        _var.dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow01[57].ToString())
                           );
                        _var.dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow01[59].ToString())
                            );
                        _var.dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow01[61].ToString())
                            );
                        _var.dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow01[63].ToString())
                            );
                        _var.dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow01[77].ToString())
                            );

                        objdt_template.Rows.Add(_var.dtworkRow01);

                        if (_var.dtworkRow02 != null)
                        {
                            _var.dtworkRow02[77] = _var.dtworkRow02[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow02[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow02[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow02[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow02[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow02[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow02[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow02[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow02);
                        }

                        if (_var.dtworkRow03 != null)
                        {
                            _var.dtworkRow03[77] = _var.dtworkRow03[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow03[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow03[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow03[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow03[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow03[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow03[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow03[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow03);
                        }

                        if (_var.dtworkRow04 != null)
                        {
                            _var.dtworkRow04[77] = _var.dtworkRow04[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow04[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow04[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow04[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow04[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow04[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow04[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow04[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow04);
                        }

                        if (_var.dtworkRow05 != null)
                        {
                            _var.dtworkRow05[77] = _var.dtworkRow05[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow05[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow05[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow05[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow05[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow05[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow05[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow05[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow05);
                        }

                        if (_var.dtworkRow06 != null)
                        {
                            _var.dtworkRow06[77] = _var.dtworkRow06[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow06[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow06[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow06[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow06[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow06[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow06[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow06[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow06[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow06[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow06[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow06);
                        }

                        if (_var.dtworkRow07 != null)
                        {
                            _var.dtworkRow07[77] = _var.dtworkRow07[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow07[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow07[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow07[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow07[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow07[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow07[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow07[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow07[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow07[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow07[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow07);
                        }

                        if (_var.dtworkRow08 != null)
                        {
                            _var.dtworkRow08[77] = _var.dtworkRow08[27];
                            _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow08[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow08[57].ToString())
                            );
                            _var.dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow08[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow08[59].ToString())
                                );
                            _var.dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow08[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow08[61].ToString())
                                );
                            _var.dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow08[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow08[63].ToString())
                                );
                            _var.dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(_var.dtworkRow08[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow08[77].ToString())
                                );

                            objdt_template.Rows.Add(_var.dtworkRow08);
                        }
                    }

                    prawrow++;

                    polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                    fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                    gender = wsraw.Cells[prawrow, 3].Text.ToString();
                    dob = wsraw.Cells[prawrow, 4].Text.ToString();
                    smoker = wsraw.Cells[prawrow, 5].Text.ToString();
                    age = wsraw.Cells[prawrow, 6].Text.ToString(); // Issue age
                    extra = wsraw.Cells[prawrow, 7].Text.ToString(); // multiple extra
                    bustype = wsraw.Cells[prawrow, 8].Text.ToString(); //treaty
                    sum = wsraw.Cells[prawrow, 13].Text.ToString(); //current
                    sum1 = wsraw.Cells[prawrow, 14].Text.ToString(); //current
                    orig = wsraw.Cells[prawrow, 10].Text.ToString(); //orig face
                    curr = wsraw.Cells[prawrow, 11].Text.ToString(); //curr face
                    term = wsraw.Cells[prawrow, 16].Text.ToString(); // termination date
                    paid = wsraw.Cells[prawrow, 14].Text.ToString(); // paid to date
                    code = wsraw.Cells[prawrow, 18].Text.ToString(); // plan code 
                    premium = wsraw.Cells[prawrow, 17].Text.ToString(); // plan code 
                    period = wsraw.Cells[prawrow, 27].Text.ToString(); // premuim for the period
                    remarks = wsraw.Cells[prawrow, 28].Text.ToString(); //comment
                    trans = wsraw.Cells[prawrow, 15].Text.ToString(); //trans
                    lsci = wsraw.Cells[prawrow, 19].Text.ToString(); //lsci
                    esci = wsraw.Cells[prawrow, 20].Text.ToString(); //esci
                    angio = wsraw.Cells[prawrow, 21].Text.ToString(); //angio
                    cancer = wsraw.Cells[prawrow, 22].Text.ToString(); //cancer
                    hospc = wsraw.Cells[prawrow, 23].Text.ToString(); //hosp
                    posth = wsraw.Cells[prawrow, 24].Text.ToString(); //posth
                    homer = wsraw.Cells[prawrow, 25].Text.ToString(); //homer
                    pall = wsraw.Cells[prawrow, 26].Text.ToString(); //pall

                    type = wsraw.Cells[prawrow, 31].Text.ToString();
                    fyry = wsraw.Cells[prawrow, 29].Text.ToString();
                    bool isNumber = false;
                    double Doutput;
                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(lsci.Trim()), out Doutput);
                    if (!isNumber) { lsci = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(esci.Trim()), out Doutput);
                    if (!isNumber) { esci = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(angio.Trim()), out Doutput);
                    if (!isNumber) { angio = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(cancer.Trim()), out Doutput);
                    if (!isNumber) { cancer = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(hospc.Trim()), out Doutput);
                    if (!isNumber) { hospc = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(posth.Trim()), out Doutput);
                    if (!isNumber) { posth = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(homer.Trim()), out Doutput);
                    if (!isNumber) { homer = "0"; }

                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(pall.Trim()), out Doutput);
                    if (!isNumber) { pall = "0"; }
                    
                    rowcount++;
                }
                #endregion
                #region "Compute Hash Total"
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
                #endregion
                string despath = str_saved + @"\BM044-" + str_sheet + str_savef + ".xlsx";
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
                _var.dtworkRow02 = null; //Dispose datarow
                _var.dtworkRow03 = null; //Dispose datarow
                _var.dtworkRow04 = null; //Dispose datarow
                _var.dtworkRow05 = null; //Dispose datarow
                _var.dtworkRow06 = null; //Dispose datarow
                _var.dtworkRow07 = null; //Dispose datarow
                _var.dtworkRow08 = null; //Dispose datarow

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
