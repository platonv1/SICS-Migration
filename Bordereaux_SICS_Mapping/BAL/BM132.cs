using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM132
    {
        public string fn_process(string str_raw, string str_sicstemp, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
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

                System.Data.DataRow dtworkRow;
                System.Data.DataRow dtworkRow02;
                System.Data.DataRow dtworkRow03;
                System.Data.DataRow dtworkRow04;
                System.Data.DataRow dtworkRow05;
                System.Data.DataRow dtworkRow06;
                System.Data.DataRow dtworkRow07;
                System.Data.DataRow dtworkRow08;
                Helper objHlpr = new Helper();
                DataTable objdt_template = new DataTable();

                objdt_template = objHlpr.dt_formtemplate(str_sicstemp, str_sheet);

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
                string polnum = wsraw.Cells[prawrow, 4].Text.ToString();
                string lifeplan = wsraw.Cells[prawrow,6 ].Text.ToString();
                string sdate = wsraw.Cells[prawrow, 12].Text.ToString();
                string comdate = wsraw.Cells[prawrow, 7].Text.ToString();
                string life = wsraw.Cells[prawrow, 17].Text.ToString();
                string life1 = wsraw.Cells[prawrow, 21].Text.ToString();
                string life2 = wsraw.Cells[prawrow, 25].Text.ToString();
                string tpd= wsraw.Cells[prawrow, 18].Text.ToString();
                string tpd1 = wsraw.Cells[prawrow, 22].Text.ToString();
                string tpd2 = wsraw.Cells[prawrow, 26].Text.ToString();
                string ti = wsraw.Cells[prawrow, 19].Text.ToString();
                string ti1 = wsraw.Cells[prawrow, 23].Text.ToString();
                string ti2 = wsraw.Cells[prawrow, 27].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                string gender = wsraw.Cells[prawrow, 8].Text.ToString();
                string dob = wsraw.Cells[prawrow, 9].Text.ToString();
                string classific = wsraw.Cells[prawrow, 15].Text.ToString();
                string premiumlife = wsraw.Cells[prawrow, 34].Text.ToString();
                string premiumtpd = wsraw.Cells[prawrow, 35].Text.ToString();
                string premiumti = wsraw.Cells[prawrow, 36].Text.ToString();
                string age = wsraw.Cells[prawrow, 11].Text.ToString();


                //string[] arr_bmyear = wsraw.Cells[4, 8].Text.ToString().Split('-');
                //string bmyear = arr_bmyear[0] + "-01-" + arr_bmyear[1];
                //DateTime date_bmyear = Convert.ToDateTime(bmyear);
                // year = year.Replace(year.Substring(year.Length - 3, 3), "-01" + year.Substring(year.Length - 3, 3));


                string currency = string.Empty;
                string year12 = string.Empty;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                int storee;
                bool chck;
                //decimal classific;
               
                #region Data Processing


                while (rowcount != erawrow + 2)
                {
                    life = life.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    life1 = life1.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    life2 = life2.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);

                    tpd = tpd.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    tpd1 = tpd1.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    tpd2 = tpd2.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);

                    ti = ti.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    ti1 = ti1.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    ti2 = ti2.TrimStart(' ').TrimEnd(' ').Replace("-", String.Empty);
                    chck = int.TryParse(polnum, out storee);
                    polnum = objHlpr.fn_stringcleanup(polnum);


                    if (polnum == string.Empty && chck == false)
                    {
                        findboo = false;

                        comparestring = new string[] { "FIRST", "NEW" };
                        foreach (string s in comparestring)
                        {
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
                            {
                                case true:
                                    TRANCODE = "TCANCINC";
                                    findboo = true;
                                    break;
                            }
                        }
                        comparestring = new string[] { "EXPIRY", "EXPIRED" };
                        foreach (string s in comparestring)
                        {
                            switch (polnum.Contains(s))
                            {
                                case true:
                                    TRANCODE = "TEXPIRY";
                                    findboo = true;
                                    break;
                            }
                        }
                        comparestring = new string[] { "EXTENDED", "TERM", "ETI" };
                        foreach (string s in comparestring)
                        {
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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
                            switch (polnum.Contains(s))
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


                else if ((polnum != string.Empty && chck == true) && !(life == String.Empty) && !(life1 == String.Empty) && !(life2 == String.Empty)
                && !(tpd == String.Empty) && !(tpd1 == String.Empty) && !(tpd2 == String.Empty) && !(ti == String.Empty) && !(ti1 == String.Empty)
                && !(ti2 == String.Empty))
                     {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow03 = objdt_template.NewRow();
                        dtworkRow04 = objdt_template.NewRow();
                        dtworkRow05 = objdt_template.NewRow();
                        dtworkRow06 = objdt_template.NewRow();
                        dtworkRow07 = objdt_template.NewRow();
                        dtworkRow08 = objdt_template.NewRow();
                        dtworkRow[0] = polnum;
                        
                        dtworkRow[5] = lifeplan;
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        dtworkRow[31] = fullname;

                        if (fullname == String.Empty)
                        {
                           fullname= objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;

                        }
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = life;
                        dtworkRow[26] = life1;
                        dtworkRow[27] =life2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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

                        dtworkRow02.ItemArray = dtworkRow.ItemArray;

                        dtworkRow02[5] = "TPD";
                        dtworkRow02[25] = tpd;
                        dtworkRow02[26] = tpd1;
                        dtworkRow02[27] = tpd2;

                        dtworkRow03.ItemArray = dtworkRow02.ItemArray;

                        dtworkRow03[5] = "TI";
                        dtworkRow03[25] = ti;
                        dtworkRow03[26] = ti1;
                        dtworkRow03[27] = ti2;

                        dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                           );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(dtworkRow);

                        if (dtworkRow02 != null)
                        {
                            dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                            );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                );

                            objdt_template.Rows.Add(dtworkRow02);
                        }
                        if (dtworkRow03 != null)
                        {
                            dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow03[57].ToString()) ? "0" : dtworkRow03[57].ToString()
                            );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow03[59].ToString()) ? "0" : dtworkRow03[59].ToString()
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow03[61].ToString()) ? "0" : dtworkRow03[61].ToString()
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow03[63].ToString()) ? "0" : dtworkRow03[63].ToString()
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow03[77].ToString()) ? "0" : dtworkRow03[77].ToString()
                                );

                            objdt_template.Rows.Add(dtworkRow03);
                        }
                    }
                    else if ((polnum != string.Empty && chck == true) && !(life == String.Empty) && !(life1 == String.Empty) && !(life2 == String.Empty))
                   
                    {

                        dtworkRow = objdt_template.NewRow();
          
                        dtworkRow[0] = polnum;
                        dtworkRow[5] = lifeplan;
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        dtworkRow[31] = fullname;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = life;
                        dtworkRow[26] = life1;
                        dtworkRow[27] = life2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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

                        dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                            );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );
                        #endregion

                        objdt_template.Rows.Add(dtworkRow);// inpu8trow+++


                    }

                        else if ((polnum != string.Empty && chck == true) && !(tpd == String.Empty) && !(tpd1 == String.Empty) && !(tpd2 == String.Empty))
                        {

                        dtworkRow = objdt_template.NewRow();

                        dtworkRow[0] = polnum;
                       
                        dtworkRow[5] = "TPD";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[31] = fullname;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = tpd;
                        dtworkRow[26] = tpd1;
                        dtworkRow[27] = tpd2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
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

                        dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                            );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );
                        #endregion
                        objdt_template.Rows.Add(dtworkRow);// inpu8trow+++


                    }
                    else if ((polnum != string.Empty && chck == true) && !(ti == String.Empty) && !(ti1 == String.Empty)&& !(ti2 == String.Empty))
                    {

                        dtworkRow = objdt_template.NewRow();

                        dtworkRow[0] = polnum;
                        
                        dtworkRow[5] = "TI";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[31] = fullname;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = ti;
                        dtworkRow[26] = ti1;
                        dtworkRow[27] = ti2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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

                        dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                            );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );
                        #endregion
                        objdt_template.Rows.Add(dtworkRow);// inpu8trow+++


                    }
                    else if ((polnum != string.Empty && chck == true) && !(life == String.Empty) && !(life1 == String.Empty) && !(life2 == String.Empty)
                           && !(tpd == String.Empty) && !(tpd1 == String.Empty) && !(tpd2 == String.Empty))
                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow[0] = polnum;
                        dtworkRow[0] = objdt_template.Rows[objdt_template.Rows.Count - 1][1].ToString();
                        dtworkRow[5] = lifeplan;
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[31] = fullname;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = life;
                        dtworkRow[26] = life1;
                        dtworkRow[27] = life2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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
                        


                        dtworkRow02.ItemArray = dtworkRow.ItemArray;

                        dtworkRow02[5] = "TPD";
                        dtworkRow02[25] = tpd;
                        dtworkRow02[26] = tpd1;
                        dtworkRow02[27] = tpd2;

                        dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                           );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(dtworkRow);

                        if (dtworkRow02 != null)
                        {
                            dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                            );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                );

                            objdt_template.Rows.Add(dtworkRow02);
                        }

                    }
                    else if ((polnum != string.Empty && chck == true) && !(life == String.Empty) && !(life1 == String.Empty) && !(life2 == String.Empty)
                        && !(ti == String.Empty) && !(ti1 == String.Empty) && !(ti2 == String.Empty))
                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow[0] = polnum;
                        
                        dtworkRow[5] = lifeplan;
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[31] = fullname;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = life;
                        dtworkRow[26] = life1;
                        dtworkRow[27] = life2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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
                        


                        dtworkRow02.ItemArray = dtworkRow.ItemArray;

                        dtworkRow02[5] = "TI";
                        dtworkRow02[25] = ti;
                        dtworkRow02[26] = ti1;
                        dtworkRow02[27] = ti2;

                        dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                           );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(dtworkRow);

                        if (dtworkRow02 != null)
                        {
                            dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                            );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                );

                            objdt_template.Rows.Add(dtworkRow02);
                        }

                    }
                    else if ((polnum != string.Empty && chck == true) && !(tpd == String.Empty) && !(tpd1 == String.Empty) && !(tpd2 == String.Empty) && !(ti == String.Empty) && !(ti1 == String.Empty)
      && !(ti2 == String.Empty))
                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow[0] = polnum;
                       
                        dtworkRow[5] = "TPD";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = sdate;
                        dtworkRow[22] = sdate;
                        dtworkRow[20] = comdate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[78] = age;
                        if (fullname == String.Empty)
                        {
                            fullname = objdt_template.Rows[objdt_template.Rows.Count - 1][31].ToString();
                            dtworkRow[31] = fullname;
                        }
                        dtworkRow[31] = fullname;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[25] = tpd;
                        dtworkRow[26] = tpd1;
                        dtworkRow[27] = tpd2;
                        dtworkRow[38] = "NONE";


                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow[34] = str_MI;

                        dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow[32] = str_outlname;

                        dtworkRow[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow.Table.Columns[36].ColumnName);
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
                        


                        dtworkRow02.ItemArray = dtworkRow.ItemArray;

                        dtworkRow02[5] = "TI";
                        dtworkRow02[25] = ti;
                        dtworkRow02[26] = ti1;
                        dtworkRow02[27] = ti2;

                        dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                           );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(dtworkRow);

                        if (dtworkRow02 != null)
                        {
                            dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                            );
                            dbl_BH += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                );
                            dbl_BJ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                );
                            dbl_BL += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                );
                            dbl_BZ += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                );

                            objdt_template.Rows.Add(dtworkRow02);
                        }

                    }

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 4].Text.ToString();
                    lifeplan = wsraw.Cells[prawrow, 6].Text.ToString();
                    sdate = wsraw.Cells[prawrow, 12].Text.ToString();
                    comdate = wsraw.Cells[prawrow, 7].Text.ToString();
                    life = wsraw.Cells[prawrow, 17].Text.ToString();
                    life1 = wsraw.Cells[prawrow, 21].Text.ToString();
                    life2 = wsraw.Cells[prawrow, 25].Text.ToString();
                    tpd = wsraw.Cells[prawrow, 18].Text.ToString();
                    tpd1 = wsraw.Cells[prawrow, 22].Text.ToString();
                    tpd2 = wsraw.Cells[prawrow, 26].Text.ToString();
                    ti = wsraw.Cells[prawrow, 19].Text.ToString();
                    ti1 = wsraw.Cells[prawrow, 23].Text.ToString();
                    ti2 = wsraw.Cells[prawrow, 27].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                    gender = wsraw.Cells[prawrow, 8].Text.ToString();
                    dob = wsraw.Cells[prawrow, 9].Text.ToString();
                    classific = wsraw.Cells[prawrow, 15].Text.ToString();
                    premiumlife = wsraw.Cells[prawrow, 34].Text.ToString();
                    premiumtpd = wsraw.Cells[prawrow, 35].Text.ToString();
                    premiumti = wsraw.Cells[prawrow, 36].Text.ToString();
                    age = wsraw.Cells[prawrow, 11].Text.ToString();

                    rowcount++;
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
                string despath = str_saved + @"\BM132-" + str_savef + ".xlsx";
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
                dtworkRow = null; //Dispose datarow
                dtworkRow02 = null; //Dispose datarow
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
