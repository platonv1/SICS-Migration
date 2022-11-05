using System;
using System.Data;


namespace Bordereaux_SICS_Mapping.BAL
{
    class BM058
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

                System.Data.DataRow dtworkRow01;
                System.Data.DataRow dtworkRow02;

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
                int edatacol = rawrange.Columns.Count;

                int prawrow = 1;

                string polnum = wsraw.Cells[2][5].Text.ToString();
                string reins = wsraw.Cells[prawrow, 6].Text.ToString();

                string policyd = wsraw.Cells[3][6].Text.ToString();
                string origd = wsraw.Cells[prawrow, 7].Text.ToString();
                string origd1 = wsraw.Cells[prawrow, 8].Text.ToString();
                string sum = wsraw.Cells[prawrow, 9].Text.ToString();
                string sum1 = wsraw.Cells[prawrow, 10].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                string gender = wsraw.Cells[prawrow, 5].Text.ToString();
                string dob = wsraw.Cells[prawrow, 3].Text.ToString();
                string expiry = wsraw.Cells[3][9].Text.ToString();
                string prem = wsraw.Cells[prawrow, 18].Text.ToString();
                string prem1 = wsraw.Cells[prawrow, 19].Text.ToString();
                string holder = wsraw.Cells[2][4].Text.ToString();
                string risk = wsraw.Cells[prawrow, 15].Text.ToString();
                string cert = wsraw.Cells[prawrow, 1].Text.ToString();
                string age = wsraw.Cells[prawrow, 4].Text.ToString();

                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

               
                int storee;
                bool chck;
                #region Data Processing
                while (rowcount != erawrow + 1) //loop
                {
                    chck = int.TryParse(cert, out storee);

                    if (cert != string.Empty && chck == false)
                    {
                        findboo = false;

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

                        comparestring = new string[] { "EXPIRY", "EXPIRED" };
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

                        comparestring = new string[] { "FULL RECAPTURE", "RECAPTURED", " RECAP", "PARTIAL RECAP", "FULL RECAP", "PARTIAL RECAPTURED" };
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
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();

                        dtworkRow01[0] = polnum;
                        dtworkRow01[3] = "DEATH";
                        dtworkRow01[4] = "TERMLIFE-GRP";
                        dtworkRow01[8] = "SURPLUS";
                        dtworkRow01[9] = "PAFM";
                        dtworkRow01[10] = "S";
                        dtworkRow01[13] = "GRP";
                        dtworkRow01[14] = "T";
                        dtworkRow01[24] = "MLY";
                        dtworkRow01[26] = "1.00";
                        dtworkRow01[28] = "1.00";
                        dtworkRow01[29] = "NATREID";
                        dtworkRow01[38] = "NONE";
                        dtworkRow01[19] = reins;
                        dtworkRow01[22] = reins;
                        dtworkRow01[20] = policyd;
                        dtworkRow01[25] = origd;
                        dtworkRow01[27] = sum;
                        dtworkRow01[31] = fullname;
                        dtworkRow01[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR4AL" : dtworkRow01[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow01[37] = dob.ToString();
                        dtworkRow01[40] = expiry;
                        dtworkRow01[77] = risk;
                        dtworkRow01[82] = holder;
                        dtworkRow01[79] = age;

                        DateTime reinsDate = Convert.ToDateTime(reins);

                        DateTime policydoDate = Convert.ToDateTime(policyd);

                        if (policydoDate.Year > reinsDate.Year)
                        {
                            dtworkRow01[21] = "TNEWBUS";
                            dtworkRow01[56] = "4000";
                            dtworkRow01[57] = prem;
                        }
                        else if (policydoDate.Year <= reinsDate.Year)
                        {
                            dtworkRow01[21] = "TRENEW";
                            dtworkRow01[58] = "4001";
                            dtworkRow01[59] = prem;
                        }

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR6AF" : dtworkRow01[76].ToString() + "|BR6AF";
                        }

                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out string str_outfname, out string str_outlname, out string str_outlifeid, "000");

                        string str_MI = objHlpr.fn_getMI(str_outfname);
                        dtworkRow01[34] = str_MI;

                        dtworkRow01[31] = objHlpr.fn_stringcleanup(fullname);
                        dtworkRow01[32] = str_outlname;

                        dtworkRow01[33] = str_outfname.Replace(" " + str_MI, string.Empty);

                        dtworkRow01[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow01[36] = objHlpr.fn_getgender(str_gender, str_outfname, dtworkRow01.Table.Columns[36].ColumnName);
                        }
                        dtworkRow02.ItemArray = dtworkRow01.ItemArray;

                        dtworkRow02[4] = "AD&D-GRP";
                        dtworkRow02[25] = origd1;
                        dtworkRow02[27] = sum1;

                        if (policydoDate.Year > reinsDate.Year)
                        {
                            dtworkRow02[21] = "TNEWBUS";
                            dtworkRow02[56] = "4000";
                            dtworkRow02[57] = prem;
                        }
                        else if (policydoDate.Year <= reinsDate.Year)
                        {
                            dtworkRow02[21] = "TRENEW";
                            dtworkRow02[58] = "4001";
                            dtworkRow02[59] = prem;
                        }
                        #region "New Requirements"
                        dtworkRow01[26] = string.Empty;

                        if (!String.IsNullOrEmpty(dtworkRow01[27].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[77].ToString()))
                        {
                            dtworkRow01[77] = dtworkRow01[27];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR1-1BZ" : dtworkRow01[76].ToString() + "|BR1-1BZ";
                        }
                        else if (!String.IsNullOrEmpty(dtworkRow01[25].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[77].ToString()))
                        {
                            dtworkRow01[75] = dtworkRow01[25];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR1-2BZ" : dtworkRow01[76].ToString() + "|BR1-2BZ";
                        }

                        if (!String.IsNullOrEmpty(dtworkRow01[77].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[27].ToString()))
                        {
                            dtworkRow01[27] = dtworkRow01[77];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR2-1AB" : dtworkRow01[76].ToString() + "|BR2-1AB";
                        }
                        else if (!String.IsNullOrEmpty(dtworkRow01[25].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[27].ToString()))
                        {
                            dtworkRow01[27] = dtworkRow01[25];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR2-2AB" : dtworkRow01[76].ToString() + "|BR2-2AB";
                        }

                        if (!String.IsNullOrEmpty(dtworkRow01[27].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[25].ToString()))
                        {
                            dtworkRow01[25] = dtworkRow01[27];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR3-1Z" : dtworkRow01[76].ToString() + "|BR3-1Z";
                        }
                        else if (!String.IsNullOrEmpty(dtworkRow01[77].ToString())
                            &&
                            String.IsNullOrEmpty(dtworkRow01[25].ToString()))
                        {
                            dtworkRow01[25] = dtworkRow01[77];
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR3-2Z" : dtworkRow01[76].ToString() + "|BR3-2Z";
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

                        if (dtworkRow01[13].ToString() == "GRP" || dtworkRow01[13].ToString() == "GCL" || dtworkRow01[13].ToString() == "GEB")
                        {
                            if (dtworkRow01[0].ToString().Length >= 7)
                            {
                                dtworkRow01[0] = dtworkRow01[0].ToString().Substring(dtworkRow01[0].ToString().Length - 7, 7) +
                                    initialNR +
                                    parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                            }
                            else
                            {
                                dtworkRow01[0] = dtworkRow01[0].ToString() +
                                    initialNR +
                                    parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                            }
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR5-1A" : dtworkRow01[76].ToString() + "|BR5-1A";

                            dtworkRow01[1] = polnum.ToString() + gender.Substring(0, 1);
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR5-2B" : dtworkRow01[76].ToString() + "|BR5-2B";

                            dtworkRow01[7] = polnum.ToString();
                        }
                        else
                        {
                            dtworkRow01[1] = string.Empty;
                            dtworkRow01[7] = string.Empty;
                        }

                        dtworkRow01[19] = dtworkRow01[20];

                        #endregion
                        dbl_BF += decimal.Parse(
                           String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                           );
                        dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                            );
                        dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                            );
                        dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                            );
                        dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                            );

                        objdt_template.Rows.Add(dtworkRow01);

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
                    polnum = wsraw.Cells[2][5].Text.ToString();
                    reins = wsraw.Cells[prawrow, 6].Text.ToString();
                    policyd = wsraw.Cells[3][6].Text.ToString();
                    origd = wsraw.Cells[prawrow, 7].Text.ToString();
                    origd1 = wsraw.Cells[prawrow, 8].Text.ToString();
                    sum = wsraw.Cells[prawrow, 9].Text.ToString();
                    sum1 = wsraw.Cells[prawrow, 10].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 2].Text.ToString();
                    gender = wsraw.Cells[prawrow, 5].Text.ToString();
                    dob = wsraw.Cells[prawrow, 3].Text.ToString();
                    expiry = wsraw.Cells[3][9].Text.ToString();
                    prem = wsraw.Cells[prawrow, 18].Text.ToString();
                    prem1 = wsraw.Cells[prawrow, 19].Text.ToString();
                    holder = wsraw.Cells[2][4].Text.ToString();
                    cert = wsraw.Cells[prawrow, 1].Text.ToString();
                    age = wsraw.Cells[prawrow, 4].Text.ToString();
                    rowcount++;
                }
                #endregion

                string despath = str_saved + @"\BM058-" + str_savef + ".xlsx";
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
                dtworkRow01 = null;
                dtworkRow02 = null;
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
