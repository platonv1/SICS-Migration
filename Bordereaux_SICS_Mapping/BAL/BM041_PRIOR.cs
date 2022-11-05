using System;
using System.Data;
namespace Bordereaux_SICS_Mapping.BAL
{
    class BM041_PRIOR
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {
            #region NOTES
            //Declaration for exception line debugging on excel
            #endregion
            int rowcount = 1;
            try
            {
                #region "HASH Total"
                decimal dbl_BF = 0, dbl_BH = 0, dbl_BJ = 0, dbl_BL = 0, dbl_BZ = 0;
                #endregion

                str_sheet = str_sheet.ToUpper();

                Helper objHlpr = new Helper();
                DataTable objdt_template = new DataTable();
                System.Data.DataRow dtworkRow;

                

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
                int prawrow2 = 1;

                string busmean = "";
                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string gender = wsraw.Cells[prawrow, 5].Text.ToString();
                string smoker = wsraw.Cells[prawrow, 6].Text.ToString();
                string bustype = wsraw.Cells[prawrow, 8].Text.ToString();
                string age = wsraw.Cells[prawrow, 7].Text.ToString();
                string paid = wsraw.Cells[prawrow, 9].Text.ToString();
                string premyr = wsraw.Cells[prawrow, 10].Text.ToString();
                string polnum2 = wsraw.Cells[prawrow2, 1].Text.ToString();
                string branded = wsraw.Cells[prawrow2, 2].Text.ToString();
                string curr = wsraw.Cells[prawrow2, 14].Text.ToString();
                string cededsum = wsraw.Cells[prawrow2, 4].Text.ToString();
                string inisum = wsraw.Cells[prawrow2, 8].Text.ToString();
                string inisum2 = wsraw.Cells[prawrow2, 8].Text.ToString();
                string effdt = wsraw.Cells[prawrow2, 3].Text.ToString();
                string prem = wsraw.Cells[prawrow2, 10].Text.ToString();
                string classpref = wsraw.Cells[prawrow2, 9].Text.ToString();
                string premyr1 = wsraw.Cells[prawrow, 12].Text.ToString();
                string total = wsraw.Cells[prawrow, 15].Text.ToString();
                string adjs = wsraw.Cells[prawrow, 10].Text.ToString();
                string remarks = wsraw.Cells[prawrow, 11].Text.ToString();
                string adjtc = wsraw.Cells[prawrow, 12].Text.ToString();
                string adjbt = wsraw.Cells[prawrow, 13].Text.ToString();
                string adjcur = wsraw.Cells[prawrow, 14].Text.ToString();
                string adjpl = wsraw.Cells[prawrow, 2].Text.ToString();
                string adjex = wsraw.Cells[prawrow, 5].Text.ToString();
                string adjpol = wsraw.Cells[prawrow, 3].Text.ToString();
                string adjprem = wsraw.Cells[prawrow, 6].Text.ToString();
                string tot = wsraw.Cells[prawrow2, 13].Text.ToString();

                string polstore;
                string bcstore;
                string ipstore;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                long storee;
                double store;
                double prefstore;
                bool chck;
                bool cla;

        
                #region Data Processing
                if (str_sheet.Contains("NB") || str_sheet.Contains("REN"))
                {
                    while (rowcount != erawrow + 2)
                    {
                        chck = long.TryParse(polnum, out storee);

                        if (polnum != string.Empty && chck == false)
                        {
                            findboo = false;

                            if (str_sheet.Contains("NB"))
                            {
                                TRANCODE = "TNEWBUS";
                                findboo = true;

                            }
                            else if (str_sheet.Contains("REN"))
                            {
                                TRANCODE = "TRENEW";
                                findboo = true;
                            }
                            comparestring = new string[] { "FIRST" };

                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TNEWBUS";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "RENEWAL" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TRENEW";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TREINS";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "TERMINATION", "TERMINATED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TCONTER";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "CANCELLED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TCANCINC";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "EXPIRY, EXPIRED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TEXPIRY";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "EXTENDED TERM", "ETI" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TEXTTER";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "MATURITY", "MATURED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLMAT";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "FULL PAID-UP", "FULL PAID UP", "PAID UP", "FULLY PAID-UP", "FULLY PAID-UP" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLPU";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "FULL RECAPTURE", "RECAPTURED", "RECAP", "PARTIAL RECAP", "FULL RECAP", "PARTIAL RECAPTURED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLREC";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "SURRENDERED", "SURRENDER", "FULL SURRENDERED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLSUR";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "LAPSE", "LAPSED", "LAPSES/SURRENDERS" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TLAPSE";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "RECOVERIES", "OTHERS" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "ADJUST";
                                        findboo = true;
                                        break;
                                }
                            }

                            if (!findboo)
                            {
                                TRANCODE = "ADJUST";
                            }
                        }
                        else if (chck == true)
                        {
                            prawrow2++;
                            polnum2 = wsraw.Cells[prawrow2, 1].Text.ToString();
                            branded = wsraw.Cells[prawrow2, 2].Text.ToString();
                            effdt = wsraw.Cells[prawrow2, 3].Text.ToString();
                            cededsum = wsraw.Cells[prawrow2, 4].Text.ToString();
                            inisum = wsraw.Cells[prawrow2, 8].Text.ToString();
                            inisum2 = wsraw.Cells[prawrow2, 8].Text.ToString();
                            classpref = wsraw.Cells[prawrow2, 9].Text.ToString();
                            curr = wsraw.Cells[prawrow2, 14].Text.ToString();
                            prem = wsraw.Cells[prawrow2, 10].Text.ToString();
                            age = wsraw.Cells[prawrow, 7].Text.ToString();
                            adjs = wsraw.Cells[prawrow, 10].Text.ToString();     //premium amount(adj)
                            remarks = wsraw.Cells[prawrow, 11].Text.ToString();     //remarks(adj)
                            adjtc = wsraw.Cells[prawrow, 12].Text.ToString();     //trans code(adj)
                            adjbt = wsraw.Cells[prawrow, 13].Text.ToString();     //business type(adj)
                            adjcur = wsraw.Cells[prawrow, 14].Text.ToString();     //currency(adj)\
                            adjpl = wsraw.Cells[prawrow, 2].Text.ToString();
                            tot = wsraw.Cells[prawrow2, 13].Text.ToString();
                            while (!string.IsNullOrEmpty(branded))
                            {
                                dtworkRow = objdt_template.NewRow();

                                if (bustype == "F")
                                {
                                    busmean = "F";
                                }
                                else
                                {
                                    busmean = "T";
                                }

                                if (curr.ToString() == "PESO")
                                {
                                    dtworkRow[23] = "PHP";
                                }
                                else
                                {
                                    dtworkRow[23] = "USD";
                                }

                                if (polnum.StartsWith("08"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "VARLIFE-GU";
                                }
                                else if (!polnum.StartsWith("08") && !(branded == "ADB" || branded == "TDB"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "TRADITIONALLIFE";
                                }
                                else if (!polnum.StartsWith("08") && (branded == "ADB"))
                                {
                                    dtworkRow[3] = "DISAB";
                                    dtworkRow[4] = "ADB-IND";
                                }
                                else if (!polnum.StartsWith("08") && (branded == "TDB"))
                                {
                                    dtworkRow[3] = "DISAB";
                                    dtworkRow[4] = "WOPDIIND";
                                }
                                if (smoker.ToUpper().Contains("N"))
                                {
                                    dtworkRow[38] = "NSMOK";
                                }
                                else if (smoker.ToUpper().Contains("S"))
                                {
                                    dtworkRow[38] = "SMOK";
                                }

                                dtworkRow[0] = "'" + polnum.ToString().Trim(new char[0]);
                                dtworkRow[1] = "'" + polnum.ToString().Trim(new char[0]);
                                dtworkRow[5] = branded.ToString();
                                dtworkRow[9] = "PAFW";
                                dtworkRow[8] = "SURPLUS";
                                dtworkRow[10] = "S";
                                dtworkRow[13] = "IND";
                                dtworkRow[14] = busmean.ToString();
                                dtworkRow[83] = objHlpr.fn_getrefcode(busmean);
                                dtworkRow[41] = premyr.ToString();   //for paid date
                                dtworkRow[20] = paid.ToString();
                                dtworkRow[24] = "YLY";
                                dtworkRow[25] = cededsum.ToString();
                                if (inisum.ToString() != "0")
                                {
                                    dtworkRow[27] = inisum2.ToString();
                                    dtworkRow[77] = inisum2.ToString();
                                }
                                else
                                {
                                    dtworkRow[27] = "1";
                                    dtworkRow[77] = "1";
                                }
                                dtworkRow[29] = "NATREID";
                                dtworkRow[31] = fullnames.ToString();
                                dtworkRow[36] = gender.ToString();
                                #region "New Requirements - No DOB"
                                if (String.IsNullOrEmpty(dob))
                                {
                                    dob = "7/1/1900";
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                                }
                                #endregion

                                dtworkRow[37] = dob.ToString();

                                if (paid.Length == 9)
                                {
                                    dtworkRow[19] = paid.Substring(paid.Length - 9, 5) + premyr;
                                    dtworkRow[22] = paid.Substring(paid.Length - 9, 5) + premyr;
                                }
                                if (paid.Length == 10)
                                {
                                    dtworkRow[19] = paid.Substring(paid.Length - 10, 6) + premyr;
                                    dtworkRow[22] = paid.Substring(paid.Length - 10, 6) + premyr;
                                }
                                if (paid.Length == 8)
                                {
                                    dtworkRow[19] = paid.Substring(paid.Length - 8, 4) + premyr;
                                    dtworkRow[22] = paid.Substring(paid.Length - 8, 4) + premyr;
                                }

                                dtworkRow[79] = age.ToString();

                                if (TRANCODE.Contains("TNEWBUS"))
                                {
                                    comparestring = new string[] { "FIRST", "First" };
                                    {
                                        dtworkRow[56] = "4000";
                                        dtworkRow[57] = tot.ToString();
                                    }

                                }
                                else if (TRANCODE.Contains("TRENEW"))
                                {
                                    comparestring = new string[] { "RENEWAL", "Renewal" };
                                    dtworkRow[58] = "4001";
                                    dtworkRow[59] = tot.ToString();
                                }
                                else if (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                        (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                        (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                        (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))
                                {
                                    comparestring = new string[] { "FIRST", "First,Recoveries", "Others" };
                                    dtworkRow[60] = "4002";
                                    dtworkRow[61] = tot.ToString();

                                }
                                else if (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                         (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                         (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                         (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))
                                {
                                    comparestring = new string[] { "Renewal", "RENEWAL", "Recoveries", "Others" };
                                    dtworkRow[62] = "4004";
                                    dtworkRow[63] = tot.ToString();

                                }

                                if (inisum.ToString() != "0")
                                {
                                    dtworkRow[27] = inisum.ToString();
                                }
                                else
                                {
                                    dtworkRow[27] = "1";
                                }

                                #region "New Requirements - No Name"
                                if (String.IsNullOrEmpty(fullnames))
                                {
                                    fullnames = polnum.ToString();
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                                }

                                #endregion

                                objHlpr.fn_getnamesandlifeID(fullnames, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                                string str_MI = objHlpr.fn_getMI(str_outfname);
                                dtworkRow[34] = str_MI;

                                dtworkRow[31] = objHlpr.fn_stringcleanup(fullnames);
                                dtworkRow[32] = str_outlname;

                                dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                                dtworkRow[30] = str_outlifeid;

                                if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                                {
                                    dtworkRow[36] = objHlpr.fn_getgender(str_gender, dtworkRow[33].ToString());
                                }

                                #region "New Requirements"
                                dtworkRow[26] = string.Empty;

                                if (!String.IsNullOrEmpty(dtworkRow[27].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[77].ToString()))
                                {
                                    dtworkRow[77] = dtworkRow[27];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR1-1BZ" : dtworkRow[76].ToString() + "|BR1-1BZ";
                                }
                                else if (!String.IsNullOrEmpty(dtworkRow[25].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[77].ToString()))
                                {
                                    dtworkRow[75] = dtworkRow[25];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR1-2BZ" : dtworkRow[76].ToString() + "|BR1-2BZ";
                                }

                                if (!String.IsNullOrEmpty(dtworkRow[77].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[27].ToString()))
                                {
                                    dtworkRow[27] = dtworkRow[77];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR2-1AB" : dtworkRow[76].ToString() + "|BR2-1AB";
                                }
                                else if (!String.IsNullOrEmpty(dtworkRow[25].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[27].ToString()))
                                {
                                    dtworkRow[27] = dtworkRow[25];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR2-2AB" : dtworkRow[76].ToString() + "|BR2-2AB";
                                }

                                if (!String.IsNullOrEmpty(dtworkRow[27].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[25].ToString()))
                                {
                                    dtworkRow[25] = dtworkRow[27];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR3-1Z" : dtworkRow[76].ToString() + "|BR3-1Z";
                                }
                                else if (!String.IsNullOrEmpty(dtworkRow[77].ToString())
                                    &&
                                    String.IsNullOrEmpty(dtworkRow[25].ToString()))
                                {
                                    dtworkRow[25] = dtworkRow[77];
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR3-2Z" : dtworkRow[76].ToString() + "|BR3-2Z";
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

                                if (dtworkRow[13].ToString() == "GRP" || dtworkRow[13].ToString() == "GCL" || dtworkRow[13].ToString() == "GEB")
                                {
                                    if (dtworkRow[0].ToString().Length >= 7)
                                    {
                                        dtworkRow[0] = dtworkRow[0].ToString().Substring(dtworkRow[0].ToString().Length - 7, 7) +
                                            initialNR +
                                            parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                                    }
                                    else
                                    {
                                        dtworkRow[0] = dtworkRow[0].ToString() +
                                            initialNR +
                                            parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                                    }
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR5-1A" : dtworkRow[76].ToString() + "|BR5-1A";

                                    dtworkRow[1] = polnum.ToString() + gender.Substring(0, 1);
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR5-2B" : dtworkRow[76].ToString() + "|BR5-2B";

                                    dtworkRow[7] = polnum.ToString();
                                }
                                else
                                {
                                    dtworkRow[1] = string.Empty;
                                    dtworkRow[7] = string.Empty;
                                }

                                dtworkRow[19] = dtworkRow[20];

                                
                                #endregion

                                cla = double.TryParse(classpref.ToString(), out prefstore);
                                if (cla == true)
                                {
                                    if (prefstore == 1.00)
                                    {
                                        dtworkRow[39] = "STANDARD";
                                    }
                                    if (prefstore == 0)
                                    {
                                        dtworkRow[39] = "STANDARD";
                                    }
                                    else if (prefstore == 1.25)
                                    {
                                        dtworkRow[39] = "CLASSA";
                                    }
                                    else if (prefstore == 1.375)
                                    {
                                        dtworkRow[39] = "CLASSAA";
                                    }
                                    else if (prefstore == 1.5)
                                    {
                                        dtworkRow[39] = "CLASSB";
                                    }
                                    else if (prefstore == 1.75)
                                    {
                                        dtworkRow[39] = "CLASSC";
                                    }
                                    else if (prefstore == 2)
                                    {
                                        dtworkRow[39] = "CLASSD";
                                    }
                                    else if (prefstore == 2.25)
                                    {
                                        dtworkRow[39] = "CLASSE";
                                    }
                                    else if (prefstore == 2.5)
                                    {
                                        dtworkRow[39] = "CLASSF";
                                    }
                                    else if (prefstore == 2.75)
                                    {
                                        dtworkRow[39] = "CLASSG";
                                    }
                                    else if (prefstore == 3)
                                    {
                                        dtworkRow[39] = "CLASSH";
                                    }
                                    else if (prefstore == 3.25)
                                    {
                                        dtworkRow[39] = "CLASSI";
                                    }
                                    else if (prefstore == 3.5)
                                    {
                                        dtworkRow[39] = "CLASSJ";
                                    }
                                    else if (prefstore == 3.75)
                                    {
                                        dtworkRow[39] = "CLASSK";
                                    }
                                    else if (prefstore == 4)
                                    {
                                        dtworkRow[39] = "CLASSL";
                                    }
                                    else if (prefstore == 4.25)
                                    {
                                        dtworkRow[39] = "CLASSM";
                                    }
                                    else if (prefstore == 4.5)
                                    {
                                        dtworkRow[39] = "CLASSN";
                                    }
                                    else if (prefstore == 4.75)
                                    {
                                        dtworkRow[39] = "CLASSO";
                                    }
                                    else if (prefstore == 5)
                                    {
                                        dtworkRow[39] = "CLASSP";
                                    }
                                }

                                if (str_sheet.Contains("NB"))
                                {
                                    dtworkRow[21] = "FY";

                                }
                                else if (str_sheet.Contains("REN"))
                                {
                                    dtworkRow[21] = "RY";
                                }

                                if (remarks != string.Empty)
                                {
                                    findboo = false;

                                    comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (remarks.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TREINS";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "TERMINATION", "TERMINATED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (remarks.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TCONTER";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "CANCELLED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TCANCINC";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "EXPIRY, EXPIRED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TEXPIRY";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "EXTENDED TERM", "ETI" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TEXTTER";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "MATURITY", "MATURED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLMAT";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "FULL PAID-UP", "FULL PAID UP", "PAID UP", "FULLY PAID-UP", "FULLY PAID-UP" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLPU";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "FULL RECAPTURE, RECAPTURED, RECAP, PARTIAL RECAP, FULL RECAP, PARTIAL RECAPTURED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLREC";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "SURRENDERED", "SURRENDER", "FULL SURRENDERED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLSUR";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "LAPSE", "LAPSED", "LAPSES/SURRENDERS" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TLAPSE";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "RECOVERIES", "OTHERS" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "ADJUST";
                                                findboo = true;
                                                break;
                                        }
                                    }

                                    if (!findboo)
                                    {
                                        dtworkRow[21] = TRANCODE;
                                    }
                                }

                                else
                                {
                                    dtworkRow[21] = TRANCODE;
                                }

                                prawrow2++;
                                tot = wsraw.Cells[prawrow2, 13].Text.ToString();
                                polnum2 = wsraw.Cells[prawrow2, 1].Text.ToString();
                                branded = wsraw.Cells[prawrow2, 2].Text.ToString();
                                effdt = wsraw.Cells[prawrow2, 3].Text.ToString();
                                cededsum = wsraw.Cells[prawrow2, 4].Text.ToString();
                                inisum = wsraw.Cells[prawrow2, 8].Text.ToString();
                                classpref = wsraw.Cells[prawrow2, 9].Text.ToString();
                                prem = wsraw.Cells[prawrow2, 10].Text.ToString();
                                curr = wsraw.Cells[prawrow2, 14].Text.ToString();

                                dbl_BF += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[57].ToString())
                                    );
                                dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[59].ToString())
                                    );
                                dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[61].ToString())
                                    );
                                dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[63].ToString())
                                    );
                                dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[77].ToString())
                                    );
                                objdt_template.Rows.Add(dtworkRow);
                            }
                            prawrow++;
                        }
                        prawrow++;
                        prawrow2 = prawrow;
                        polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                        fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                        dob = wsraw.Cells[prawrow, 4].Text.ToString();
                        gender = wsraw.Cells[prawrow, 5].Text.ToString();
                        smoker = wsraw.Cells[prawrow, 6].Text.ToString();
                        bustype = wsraw.Cells[prawrow, 8].Text.ToString();
                        curr = wsraw.Cells[prawrow2, 14].Text.ToString();
                        paid = wsraw.Cells[prawrow, 9].Text.ToString(); //issue date
                        premyr = wsraw.Cells[prawrow, 10].Text.ToString(); // prem yr
                        tot = wsraw.Cells[prawrow2, 13].Text.ToString();
                        rowcount++;
                    }
                }
                else if (str_sheet.Contains("ADJ"))
                {
                    while (rowcount != erawrow + 2)
                    {
                        if (polnum.ToUpper().Contains("TOTAL"))
                        {
                            break;
                        }
                        chck = long.TryParse(polnum, out storee);

                        if ((polnum == string.Empty && chck == false) || remarks != string.Empty)

                        {
                            findboo = false;

                            if (str_sheet.Contains("NB"))
                            {
                                TRANCODE = "TNEWBUS";
                                findboo = true;

                            }
                            else if (str_sheet.Contains("REN"))
                            {
                                TRANCODE = "TRENEW";
                                findboo = true;

                            }
                            comparestring = new string[] { "FIRST" };

                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TNEWBUS";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "RENEWAL" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))

                                {
                                    case true:
                                        TRANCODE = "TRENEW";
                                        findboo = true;
                                        break;

                                }
                            }
                            //insert code
                            comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TREINS";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "TERMINATION", "TERMINATED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TCONTER";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "CANCELLED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TCANCINC";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "EXPIRY, EXPIRED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TEXPIRY";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "EXTENDED TERM", "ETI" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TEXTTER";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "MATURITY", "MATURED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLMAT";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "FULL PAID-UP", "FULL PAID UP", "PAID UP", "FULLY PAID-UP", "FULLY PAID-UP" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLPU";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "FULL RECAPTURE", "RECAPTURED", "RECAP", "PARTIAL RECAP", "FULL RECAP", "PARTIAL RECAPTURED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLREC";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "SURRENDERED", "SURRENDER", "FULL SURRENDERED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TFULLSUR";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "LAPSE", "LAPSED", "LAPSES/SURRENDERS" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "TLAPSE";
                                        findboo = true;
                                        break;
                                }
                            }

                            comparestring = new string[] { "RECOVERIES", "OTHERS" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Trim().ToUpper().Contains(s))
                                {
                                    case true:
                                        TRANCODE = "ADJUST";
                                        findboo = true;
                                        break;
                                }
                            }
                            if (!findboo)
                            {
                                TRANCODE = "ADJUST";
                            }

                        }


                        if (chck == true)
                        {
                            dtworkRow = objdt_template.NewRow();

                            dtworkRow[0] = polnum.ToString();
                            dtworkRow[1] = polnum.ToString();
                            dtworkRow[9] = "PAFW";
                            dtworkRow[13] = "IND";
                            dtworkRow[8] = "SURPLUS";
                            dtworkRow[10] = "S";
                            dtworkRow[24] = "YLY";
                            dtworkRow[29] = "NATREID";
                            dtworkRow[40] = adjex.ToString();
                            dtworkRow[76] = remarks.ToString();
                            dtworkRow[5] = branded.ToString();
                            dtworkRow[79] = age.ToString();
                            dtworkRow[60] = prem.ToString();
                            dtworkRow[22] = adjprem.ToString();
                            dtworkRow[19] = adjprem.ToString();
                            dtworkRow[20] = adjprem.ToString();
                            dtworkRow[23] = adjcur.ToString();

                            if (adjcur.ToString() == "PESO")
                            {
                                dtworkRow[23] = "PHP";
                            }
                            else
                            {
                                dtworkRow[23] = "USD";
                            }


                            if (TRANCODE.Contains("TNEWBUS"))
                            {
                                comparestring = new string[] { "FIRST, First" };
                                dtworkRow[56] = "4000";
                                dtworkRow[57] = adjs.ToString();
                            }
                            else if (TRANCODE.Contains("TRENEW"))
                            {
                                comparestring = new string[] { "RENEWAL", "Renewal" };
                                dtworkRow[58] = "4001";
                                dtworkRow[59] = adjs.ToString();
                            }

                            else if ((adjtc.ToUpper().Contains("FY") && (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                    (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                    (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                    (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))))
                            {
                                comparestring = new string[] { "FIRST", "First", "Recoveries", "Others" };

                                dtworkRow[60] = "4002";
                                dtworkRow[61] = adjs.ToString();
                            }

                            else if ((adjtc.ToUpper().Contains("RY") && (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                     (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                     (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                     (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))))
                            {
                                comparestring = new string[] { "Renewal", "RENEWAL", "Recoveries", "Others" };

                                dtworkRow[62] = "4004";
                                dtworkRow[63] = adjs.ToString();
                            }


                            if (adjbt == "F")
                            {
                                busmean = "F";
                            }
                            else
                            {
                                busmean = "T";
                            }
                            dtworkRow[14] = busmean.ToString();
                            dtworkRow[83] = objHlpr.fn_getrefcode(busmean);

                            polstore = polnum.ToString();

                            if (polnum.StartsWith("08"))
                            {
                                dtworkRow[3] = "DEATH";
                                dtworkRow[4] = "VARLIFE-GU";

                                ipstore = "VARLIFE-GU";
                                bcstore = "DEATH";

                            }
                            else if (!polnum.StartsWith("08") && !(branded == "ADB" || branded == "TDB"))
                            {
                                dtworkRow[3] = "DEATH";
                                dtworkRow[4] = "TRADITIONAL";

                                ipstore = "DEATH";
                                bcstore = "TRADITIONAL";
                            }
                            else if (!polnum.StartsWith("08") && (branded == "ADB"))
                            {
                                dtworkRow[3] = "DISAB";
                                dtworkRow[4] = "ADB-IND";

                                ipstore = "DISAB";
                                bcstore = "ADB-IND";
                            }
                            else if (!polnum.StartsWith("08") && (branded == "TDB"))
                            {
                                dtworkRow[3] = "DISAB";
                                dtworkRow[4] = "WOPDIIND";

                                ipstore = "DISAB";
                                bcstore = "WOPDIIND";
                            }
                            else
                            {
                                ipstore = string.Empty;
                                bcstore = string.Empty;
                            }

                            if (remarks != string.Empty)
                            {
                                findboo = false;

                                comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                                foreach (string s in comparestring)
                                {
                                    switch (remarks.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            dtworkRow[21] = "TREINS";
                                            findboo = true;
                                            break;
                                    }
                                }
                                comparestring = new string[] { "TERMINATION", "TERMINATED" };
                                foreach (string s in comparestring)
                                {
                                    switch (remarks.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            dtworkRow[21] = "TCONTER";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "CANCELLED" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TCANCINC";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "EXPIRY, EXPIRED" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TEXPIRY";
                                            findboo = true;
                                            break;
                                    }
                                }
                                comparestring = new string[] { "EXTENDED TERM", "ETI" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TEXTTER";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "MATURITY", "MATURED" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TFULLMAT";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "FULL PAID-UP", "FULL PAID UP", "PAID UP", "FULLY PAID-UP", "FULLY PAID-UP" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TFULLPU";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "FULL RECAPTURE, RECAPTURED, RECAP, PARTIAL RECAP, FULL RECAP, PARTIAL RECAPTURED" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TFULLREC";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "SURRENDERED", "SURRENDER", "FULL SURRENDERED" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TFULLSUR";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "LAPSE", "LAPSED", "LAPSES/SURRENDERS" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "TLAPSE";
                                            findboo = true;
                                            break;
                                    }
                                }

                                comparestring = new string[] { "RECOVERIES", "OTHERS" };
                                foreach (string s in comparestring)
                                {
                                    switch (polnum.Trim().ToUpper().Contains(s))
                                    {
                                        case true:
                                            TRANCODE = "ADJUST";
                                            findboo = true;
                                            break;
                                    }

                                }

                                if (!findboo)
                                {
                                    dtworkRow[21] = TRANCODE;
                                }
                            }

                            else
                            {
                                dtworkRow[21] = TRANCODE;
                            }

                            dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[57].ToString())
                           );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[59].ToString())
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[61].ToString())
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[63].ToString())
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[77].ToString())
                                );

                            objdt_template.Rows.Add(dtworkRow);
                        }


                        else if ((string.IsNullOrEmpty(polnum)) && !(string.IsNullOrEmpty(remarks)))
                        {
                            if (string.IsNullOrEmpty(bustype))
                            {
                                bustype = "0";
                            }
                            if (string.IsNullOrEmpty(paid))
                            {
                                paid = "0";
                            }
                            if (string.IsNullOrEmpty(adjs))
                            {
                                adjs = "0";
                            }
                            bustype = bustype.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace("-", String.Empty).Replace("(", "-").Replace(")", String.Empty);
                            double bustype1;
                            bustype1 = Convert.ToDouble(bustype);
                            paid = paid.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace("-", String.Empty).Replace("(", "-").Replace(")", String.Empty);
                            double paid1;
                            paid1 = Convert.ToDouble(paid);
                            adjs = adjs.TrimStart(' ').TrimEnd(' ').Replace(",", String.Empty).Replace("-", String.Empty).Replace("(", "-").Replace(")", String.Empty);
                            double adjs1;
                            adjs1 = Convert.ToDouble(adjs);

                            if ((string.IsNullOrEmpty(polnum)) && (double.TryParse(bustype, out store)) && (double.TryParse(paid, out store)) && (double.TryParse(adjs, out store)) && !(string.IsNullOrEmpty(remarks)))
                            {
                                dtworkRow = objdt_template.NewRow();

                                dtworkRow[1] = objdt_template.Rows[objdt_template.Rows.Count - 1][1].ToString();
                                dtworkRow[0] = objdt_template.Rows[objdt_template.Rows.Count - 1][0].ToString();
                                dtworkRow[9] = "PAFW";
                                dtworkRow[8] = "SURPLUS";
                                dtworkRow[10] = "S";
                                dtworkRow[13] = "IND";
                                dtworkRow[24] = "YLY";
                                dtworkRow[29] = "NATREID";
                                dtworkRow[40] = adjex.ToString();
                                dtworkRow[76] = remarks.ToString();
                                dtworkRow[5] = branded.ToString();
                                dtworkRow[5] = objdt_template.Rows[objdt_template.Rows.Count - 1][5].ToString();
                                dtworkRow[79] = age.ToString();
                                dtworkRow[60] = prem.ToString();
                                dtworkRow[22] = adjprem.ToString();
                                dtworkRow[19] = adjprem.ToString();
                                dtworkRow[20] = adjprem.ToString();
                                dtworkRow[23] = adjcur.ToString();

                                if (adjcur.ToString() == "PESO")
                                {
                                    dtworkRow[23] = "PHP";
                                }
                                else
                                {
                                    dtworkRow[23] = "USD";
                                }

                                if (TRANCODE.Contains("TNEWBUS"))
                                {
                                    comparestring = new string[] { "FIRST, First" };
                                    dtworkRow[56] = "4000";
                                    dtworkRow[57] = adjs.ToString();
                                }
                                else if (TRANCODE.Contains("TRENEW"))
                                {
                                    comparestring = new string[] { "RENEWAL", "Renewal" };
                                    dtworkRow[58] = "4001";
                                    dtworkRow[59] = adjs.ToString();
                                }
                                else if ((adjtc.ToUpper().Contains("FY") && (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                        (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                        (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                        (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))))
                                {
                                    comparestring = new string[] { "FIRST", "First", "Recoveries", "Others" };
                                    dtworkRow[60] = "4002";
                                    dtworkRow[61] = adjs.ToString();
                                }
                                else if ((adjtc.ToUpper().Contains("RY") && (TRANCODE.Contains("ADJUST") || (TRANCODE.Contains("TCANCINC")) || (TRANCODE.Contains("TEXPIRY")) ||
                                         (TRANCODE.Contains("TEXTTER")) || (TRANCODE.Contains("TFULLMAT")) || (TRANCODE.Contains("TFULLPU") ||
                                         (TRANCODE.Contains("TFULLREC")) || (TRANCODE.Contains("TFULLSUR") || (TRANCODE.Contains("TLAPSE") ||
                                         (TRANCODE.Contains("TREINS")) || (TRANCODE.Contains("TCONTER"))))))))
                                {
                                    comparestring = new string[] { "Renewal", "RENEWAL", "Recoveries", "Others" };
                                    dtworkRow[62] = "4004";
                                    dtworkRow[63] = adjs.ToString();
                                }

                                if (adjbt == "F")
                                {
                                    busmean = "F";
                                }
                                else
                                {
                                    busmean = "T";
                                }

                                dtworkRow[14] = busmean.ToString();
                                dtworkRow[83] = objHlpr.fn_getrefcode(busmean);
                                polstore = polnum.ToString();

                                if (polnum.StartsWith("08"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "VARLIFE-GU";
                                    ipstore = "VARLIFE-GU";
                                    bcstore = "DEATH";
                                }
                                else if (!polnum.StartsWith("08") && !(branded == "ADB" || branded == "TDB"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "TRADITIONAL";
                                    ipstore = "DEATH";
                                    bcstore = "TRADITIONAL";
                                }
                                else if (!polnum.StartsWith("08") && (branded == "ADB"))
                                {
                                    dtworkRow[3] = "DISAB";
                                    dtworkRow[4] = "ADB-IND";
                                    ipstore = "DISAB";
                                    bcstore = "ADB-IND";
                                }
                                else if (!polnum.StartsWith("08") && (branded == "TDB"))
                                {
                                    dtworkRow[3] = "DISAB";
                                    dtworkRow[4] = "WOPDIIND";
                                    ipstore = "DISAB";
                                    bcstore = "WOPDIIND";
                                }
                                else
                                {
                                    ipstore = string.Empty;
                                    bcstore = string.Empty;
                                }

                                if (remarks != string.Empty)
                                {
                                    findboo = false;

                                    comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (remarks.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                dtworkRow[21] = "TREINS";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "TERMINATION", "TERMINATED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (remarks.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                dtworkRow[21] = "TCONTER";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "CANCELLED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TCANCINC";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "EXPIRY, EXPIRED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TEXPIRY";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "EXTENDED TERM", "ETI" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TEXTTER";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "MATURITY", "MATURED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLMAT";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "FULL PAID-UP", "FULL PAID UP", "PAID UP", "FULLY PAID-UP", "FULLY PAID-UP" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLPU";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "FULL RECAPTURE, RECAPTURED, RECAP, PARTIAL RECAP, FULL RECAP, PARTIAL RECAPTURED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLREC";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "SURRENDERED", "SURRENDER", "FULL SURRENDERED" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TFULLSUR";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "LAPSE", "LAPSED", "LAPSES/SURRENDERS" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "TLAPSE";
                                                findboo = true;
                                                break;
                                        }
                                    }
                                    comparestring = new string[] { "RECOVERIES", "OTHERS" };
                                    foreach (string s in comparestring)
                                    {
                                        switch (polnum.Trim().ToUpper().Contains(s))
                                        {
                                            case true:
                                                TRANCODE = "ADJUST";
                                                findboo = true;
                                                break;
                                        }
                                    }

                                    if (!findboo)
                                    {
                                        dtworkRow[21] = TRANCODE;
                                    }
                                }

                                else
                                {
                                    dtworkRow[21] = TRANCODE;
                                }

                                dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[57].ToString())
                           );
                                dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[59].ToString())
                                    );
                                dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[61].ToString())
                                    );
                                dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[63].ToString())
                                    );
                                dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : objHlpr.fn_numbercleanup_negative(dtworkRow[77].ToString())
                                    );
                                objdt_template.Rows.Add(dtworkRow);

                            }
                        }

                        rowcount++;
                        prawrow++;
                        prawrow2 = prawrow;
                        polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                        adjs = wsraw.Cells[prawrow, 10].Text.ToString();
                        remarks = wsraw.Cells[prawrow, 11].Text.ToString();
                        adjtc = wsraw.Cells[prawrow, 12].Text.ToString();
                        adjbt = wsraw.Cells[prawrow, 13].Text.ToString();
                        adjcur = wsraw.Cells[prawrow, 14].Text.ToString();
                        branded = wsraw.Cells[prawrow, 2].Text.ToString();
                        adjs = wsraw.Cells[prawrow, 10].Text.ToString();
                        remarks = wsraw.Cells[prawrow, 11].Text.ToString();
                        adjtc = wsraw.Cells[prawrow, 12].Text.ToString();
                        adjbt = wsraw.Cells[prawrow, 13].Text.ToString();
                        adjcur = wsraw.Cells[prawrow, 14].Text.ToString();
                        adjpl = wsraw.Cells[prawrow, 2].Text.ToString();
                        adjex = wsraw.Cells[prawrow, 5].Text.ToString();
                        adjprem = wsraw.Cells[prawrow, 6].Text.ToString();
                        adjpol = wsraw.Cells[prawrow, 3].Text.ToString();
                        bustype = wsraw.Cells[prawrow, 8].Text.ToString();
                        paid = wsraw.Cells[prawrow, 9].Text.ToString();

                    }
                }
                #endregion
                #region "Compute Hash Total"
                dtworkRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Premium:";
                dtworkRow[1] = dbl_BF + dbl_BH + dbl_BJ + dbl_BL;
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Sum at Risk:";
                dtworkRow[1] = dbl_BZ;
                objdt_template.Rows.Add(dtworkRow);
                #endregion
                string despath = str_saved + @"\BM041-PRIOR-" + str_sheet + str_savef + ".xlsx";
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
                dtworkRow = null;
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
