using System;
using System.Data;
using System.Linq;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM043
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            #region NOTES
            //Declaration for exception line debugging on excel
            #endregion
            int rowcount = 1;
            HelperV21 objHlpr2 = new HelperV21();
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
                int prawrow = 1;
                string insurance = "";
                string bcover = "";

                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                string gender = wsraw.Cells[prawrow, 3].Text.ToString();
                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string smoker = wsraw.Cells[prawrow, 5].Text.ToString();
                string age = wsraw.Cells[prawrow, 6].Text.ToString(); // Issue age
                string extra = wsraw.Cells[prawrow, 7].Text.ToString(); // multiple extra
                string bustype = wsraw.Cells[prawrow, 8].Text.ToString(); //treaty
                string issue = wsraw.Cells[prawrow, 10].Text.ToString(); //issue
                string current = wsraw.Cells[prawrow, 11].Text.ToString(); //current
                string term = wsraw.Cells[prawrow, 13].Text.ToString(); // termination date
                string paid = wsraw.Cells[prawrow, 14].Text.ToString(); // paid to date
                string premium = wsraw.Cells[prawrow, 17].Text.ToString(); // prem
                string period = wsraw.Cells[prawrow, 18].Text.ToString(); // premuim for the period
                string remarks = wsraw.Cells[prawrow, 19].Text.ToString(); //comment
                string trans = wsraw.Cells[prawrow, 12].Text.ToString();
                string benefit = wsraw.Cells[prawrow, 16].Text.ToString();
                string plan = wsraw.Cells[prawrow, 15].Text.ToString();
                string bmyear = wsraw.Cells[4, 2].Text.ToString();
                bmyear = bmyear.Substring(bmyear.Length - 4);


                string type = wsraw.Cells[prawrow, 22].Text.ToString();
                string fyry = wsraw.Cells[prawrow, 20].Text.ToString();

                polnum = objHlpr.fn_stringcleanup(polnum);
                fullname= objHlpr.fn_stringcleanup(fullname);
                gender= objHlpr.fn_stringcleanup(gender);
                dob = objHlpr.fn_stringcleanup(dob);
                smoker= objHlpr.fn_stringcleanup(smoker);
                age= objHlpr.fn_stringcleanup(age);
                extra= objHlpr.fn_stringcleanup(extra);
                bustype= objHlpr.fn_stringcleanup(bustype);
                issue= objHlpr.fn_stringcleanup(issue);
                current= objHlpr.fn_stringcleanup(current);
                term= objHlpr.fn_stringcleanup(term);
                paid= objHlpr.fn_stringcleanup(paid);
                premium= objHlpr.fn_stringcleanup(premium);
                period= objHlpr.fn_stringcleanup(period);
                remarks= objHlpr.fn_stringcleanup(remarks);
                trans= objHlpr.fn_stringcleanup(trans);
                benefit= objHlpr.fn_stringcleanup(benefit);

                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;


                int storee;
                bool chck;
                string prefstore;
             


                #region Data Processing
                while (rowcount != erawrow + 1) //loop
                {
                    //chck = int.TryParse(polnum, out storee);
                    chck = objHlpr.fn_policyNumChecker(polnum, wsraw.Cells[prawrow, 2].Text.ToString(), wsraw.Cells[prawrow, 3].Text.ToString(), wsraw.Cells[prawrow, 4].Text.ToString());

                    if (chck == true)
                    {
                        _var.dtworkRow = objdt_template.NewRow();

                        _var.dtworkRow[0] = "'" + polnum;
                        _var.dtworkRow[1] = "'" + polnum;
                        _var.dtworkRow[8] = "QA";
                        _var.dtworkRow[9] = "PAFM";
                        _var.dtworkRow[10] = "Q";
                        _var.dtworkRow[13] = "IND";
                        _var.dtworkRow[14] = "T";
                        _var.dtworkRow[24] = "YLY";
                        _var.dtworkRow[29] = "NATREID";
                        _var.dtworkRow[3] = bcover;
                        _var.dtworkRow[4] = insurance;

                        _var.dtworkRow[41] = bmyear;
                        _var.dtworkRow[20] = trans;
                        _var.dtworkRow[22] = trans.Substring(0, trans.Length - 4) + bmyear;

                        _var.dtworkRow[23] = "PHP";
                        _var.dtworkRow[31] = fullname;
                        _var.dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR4AL" : _var.dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        _var.dtworkRow[37] = dob.ToString();
                        _var.dtworkRow[38] = smoker;
                        //_var.dtworkRow[19] = paid;
                        
                        _var.dtworkRow[76] = benefit;
                        _var.dtworkRow[40] = term;
                        _var.dtworkRow[79] = age;
                        _var.dtworkRow[25] = issue;
                        _var.dtworkRow[27] = objHlpr.fn_numbercleanup_negative(current);
                        _var.dtworkRow[77] = objHlpr.fn_numbercleanup_negative(current);
                        _var.dtworkRow[38] = paid;
                        _var.dtworkRow[83] = "NR";

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
                        
                        _var.dtworkRow[21] = TRANCODE;

                        if (TRANCODE.Contains("TNEWBUS"))
                        {
                            comparestring = new string[] { "FIRST, First" };

                            _var.dtworkRow[56] = "4000";
                            _var.dtworkRow[57] = period;

                        }
                        else if (TRANCODE.Contains("TRENEW"))
                        {
                            comparestring = new string[] { "RENEWAL", "Renewal" };

                            _var.dtworkRow[58] = "4001";
                            _var.dtworkRow[59] = period;
                        }
                        else if (TRANCODE.Contains("ADJUST"))
                        {
                            if (fyry.Trim() == "FY")
                            {
                                _var.dtworkRow[60] = "4002";
                                _var.dtworkRow[61] = period;
                            }
                            else
                            {
                                _var.dtworkRow[62] = "4004";
                                _var.dtworkRow[63] = period;
                            }
                        }

                        //insurance
                        if (benefit == "FEMALE CI W/ MATERNITY")
                        {
                            insurance = "CIRACIND";
                            _var.dtworkRow[6] = "BP273";
                        }
                        else if (benefit == "FEMALE CI")
                        {
                            insurance = "CIRACIND";
                            _var.dtworkRow[6] = "BP272";
                        }
                        else if (benefit == "SUN MAIDEN")
                        {
                            insurance = "TRADITIONALLIFE";
                            _var.dtworkRow[6] = "BP274";
                        }
                        else if (benefit == "SUN MAIDEN W/ MATERNITY")
                        {
                            insurance = "TRADITIONALLIFE";
                            _var.dtworkRow[6] = "BP288";
                        }

                        _var.dtworkRow[5] = plan;
                        _var.dtworkRow[4] = insurance;

                        //benefit cover
                        if (insurance == "CIRACIND")
                        {
                            bcover = "CI";
                        }
                        else if (insurance == "TRADITIONALLIFE")
                        {
                            bcover = "DEATH";
                        }
                        _var.dtworkRow[3] = bcover;

                        //smok/nsmok

                        if (smoker.ToUpper().Contains("N"))
                        {
                            _var.dtworkRow[38] = "NSMOK";
                        }
                        else if (smoker.ToUpper().Contains("S"))
                        {
                            _var.dtworkRow[38] = "SMOK";
                        }

                        prefstore = extra.Replace("%", string.Empty);

                        _var.dtworkRow[39] = objHlpr.fn_getmortality((int.Parse(prefstore) + 100).ToString());
                        if (objHlpr.fn_isDMort(_var.dtworkRow[39].ToString()))
                        {
                            _var.dtworkRow[39] = "STANDARD";
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR8AN" : _var.dtworkRow[76].ToString() + "|BR8AN";
                        }

                        //Buss Code 
                        
                           
                        

                        _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? remarks : _var.dtworkRow[76].ToString() + "|" + remarks;

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR6AF" : _var.dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion
                        //Updated logic 05/20/2022
                        _var.dtworkRow [31] = fullname; /*objHlpr.fn_stringcleanup(fullnames);*/

                        objHlpr2.fn_separateLastNameFirstNameV4(fullname, out fullname, out string strLastName, out string strFirstName, out string strMiddleInitial);

                        _var.dtworkRow [32] = objHlpr2.fn_removeCharacters(strLastName);/*str_outlname;*/

                        _var.dtworkRow [33] = objHlpr2.fn_removeCharacters(strFirstName);/*str_outfname.Replace(" " + str_MI, string.Empty);*/

                        _var.dtworkRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, dob);/*str_outlifeid;*/
                        _var.dtworkRow [34] = strMiddleInitial;


                        //objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid);

                        //string str_MI = string.Empty;
                        //string[] arr_fullname;
                        //arr_fullname = fullname.Split(',');
                        //str_outlname = arr_fullname[0];

                        //if (arr_fullname.Count() > 1)
                        //{
                        //    str_outfname = arr_fullname[1];
                        //}

                        //if (arr_fullname.Count() > 2)
                        //{
                        //    str_MI = arr_fullname[2];
                        //    _var.dtworkRow[34] = str_MI;
                        //}


                        //_var.dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        //_var.dtworkRow[32] = str_outlname;

                        //_var.dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        //_var.dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow[33].ToString());
                        }
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

                            _var.dtworkRow[1] = polnum.ToString() + gender.Substring(0, 1);
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
                            else if (_var.dtworkRow[21].ToString().ToUpper() == "TRENEW")
                            {
                                _var.dtworkRow[19] = _var.dtworkRow[22];
                            }
                            else
                            {
                                _var.dtworkRow[19] = paid;
                                _var.dtworkRow[22] = paid;
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

                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow[57].ToString())
                            );
                        _var.dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow[59].ToString())
                            );
                        _var.dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow[61].ToString())
                            );
                        _var.dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow[63].ToString())
                            );
                        _var.dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(_var.dtworkRow[77].ToString())
                            );
                        #endregion
                        objdt_template.Rows.Add(_var.dtworkRow);
                    }
                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                    gender = wsraw.Cells[prawrow, 3].Text.ToString();
                    dob = wsraw.Cells[prawrow, 4].Text.ToString();
                    smoker = wsraw.Cells[prawrow, 5].Text.ToString();
                    age = wsraw.Cells[prawrow, 6].Text.ToString(); // Issue age
                    extra = wsraw.Cells[prawrow, 7].Text.ToString(); // multiple extra
                    bustype = wsraw.Cells[prawrow, 8].Text.ToString(); //treaty
                    issue = wsraw.Cells[prawrow, 10].Text.ToString(); //issue
                    current = wsraw.Cells[prawrow, 11].Text.ToString(); //current
                    term = wsraw.Cells[prawrow, 13].Text.ToString(); // termination date
                    paid = wsraw.Cells[prawrow, 14].Text.ToString(); // paid to date
                    benefit = wsraw.Cells[prawrow, 16].Text.ToString(); // paid to date
                    premium = wsraw.Cells[prawrow, 17].Text.ToString(); // prem
                    period = wsraw.Cells[prawrow, 18].Text.ToString(); // premuim for the period
                    remarks = wsraw.Cells[prawrow, 19].Text.ToString(); //comment
                    trans = wsraw.Cells[prawrow, 12].Text.ToString();
                    plan = wsraw.Cells[prawrow, 15].Text.ToString();

                    type = wsraw.Cells[prawrow, 22].Text.ToString();
                    fyry = wsraw.Cells[prawrow, 20].Text.ToString();

                    polnum = objHlpr.fn_stringcleanup(polnum);
                    fullname = objHlpr.fn_stringcleanup(fullname);
                    gender = objHlpr.fn_stringcleanup(gender);
                    dob = objHlpr.fn_stringcleanup(dob);
                    smoker = objHlpr.fn_stringcleanup(smoker);
                    age = objHlpr.fn_stringcleanup(age);
                    extra = objHlpr.fn_stringcleanup(extra);
                    bustype = objHlpr.fn_stringcleanup(bustype);
                    issue = objHlpr.fn_stringcleanup(issue);
                    current = objHlpr.fn_stringcleanup(current);
                    term = objHlpr.fn_stringcleanup(term);
                    paid = objHlpr.fn_stringcleanup(paid);
                    premium = objHlpr.fn_stringcleanup(premium);
                    period = objHlpr.fn_stringcleanup(period);
                    remarks = objHlpr.fn_stringcleanup(remarks);
                    trans = objHlpr.fn_stringcleanup(trans);
                    benefit = objHlpr.fn_stringcleanup(benefit);

                    bool isNumber = false;
                    double Doutput;
                    isNumber = double.TryParse(objHlpr.fn_parenthesistoNegative(period.Trim()), out Doutput);
                    if (!isNumber) { period = "0"; }

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
                string despath = str_saved + @"\BM043-" + str_sheet + str_savef + ".xlsx";
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
