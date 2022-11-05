using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM050
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
                System.Data.DataRow dtworkRow01;
                System.Data.DataRow dtworkRow02;
                System.Data.DataRow dtworkRow03;
                System.Data.DataRow dtworkRow04;
                System.Data.DataRow dtworkRow05;
               
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
                string orig = wsraw.Cells[prawrow, 12].Text.ToString();
                string ceded = wsraw.Cells[prawrow, 16].Text.ToString();
                string gender = wsraw.Cells[prawrow, 4].Text.ToString();
                string dob = wsraw.Cells[prawrow, 5].Text.ToString();
                string premium = wsraw.Cells[prawrow, 20].Text.ToString();
                string age = wsraw.Cells[prawrow, 6].Text.ToString();
                string full = wsraw.Cells[prawrow, 2].Text.ToString();
                string name = wsraw.Cells[prawrow, 3].Text.ToString();
                string reins = wsraw.Cells[3][3].Text.ToString();
                string policy = wsraw.Cells[3][2].Text.ToString();
                string holder = wsraw.Cells[3][1].Text.ToString();
                string count = wsraw.Cells[prawrow, 1].Text.ToString();
                string orig1 = wsraw.Cells[prawrow, 10].Text.ToString();
                string cedent1= wsraw.Cells[prawrow, 11].Text.ToString();
                string premium1= wsraw.Cells[prawrow, 14].Text.ToString();
                string effdate = wsraw.Cells[prawrow, 20].Text.ToString();
                string prem = wsraw.Cells[prawrow, 15].Text.ToString();
                string orig2 = wsraw.Cells[prawrow, 8].Text.ToString();
                string surplus = wsraw.Cells[prawrow, 13].Text.ToString();
                string prem1 = wsraw.Cells[prawrow, 17].Text.ToString();
                string prem2 = wsraw.Cells[prawrow, 18].Text.ToString();
                string prem3 = wsraw.Cells[prawrow, 19].Text.ToString();
                string prem4 = wsraw.Cells[prawrow, 21].Text.ToString();
                dob = string.Empty;
                string currency = string.Empty;
                string year12 = string.Empty;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;
                string polnum = string.Empty;
                int storee;
                bool chck;
           

                #region Data Processing

                while (rowcount != erawrow + 2)
                {
                    chck = int.TryParse(count, out storee);
                  
                    dtworkRow = objdt_template.NewRow();

                    if (count != string.Empty && chck == false)
                    {
                        findboo = false;

                        comparestring = new string[] { "FIRST", "NEW" };
                        foreach (string s in comparestring)
                        {
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                            switch (count.Contains(s))
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
                        else
                        {
                            dtworkRow[21] = TRANCODE;
                        }
                    }
                    else if ((count != string.Empty && chck == true) && (str_sheet == ("AEIGHT PACIFIC CORP")) || (str_sheet == ("CSA TRUE BLUE")) || (str_sheet == ("FLOOR FINISH-FLOOR LEVEL(FF-FL)"))
                        || (str_sheet == ("RCHITECTS INC")) || (str_sheet == ("SPRAYCRETE CORP")))
                    {
                        dtworkRow[5] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = reins;
                        dtworkRow[22] = reins;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[25] = orig;
                        dtworkRow[26] = ceded;
                        dtworkRow[27] = ceded;
                        dtworkRow[77] = ceded;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;
                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = premium.ToString();
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = premium.ToString();
                        }
                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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
                    else if ((count != string.Empty && chck == true) && (str_sheet == ("AG&P (ANGOLA)")) || (str_sheet == ("GININTUAN AGRO-INDUSTRIAL CORP")) || (str_sheet == ("GREAT SWISS MARITIME SERVICES")) || (str_sheet == ("MARBELLA MANILA ASSOC.")))
                    {
                        dtworkRow = objdt_template.NewRow();

                        dtworkRow[5] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = reins;
                        dtworkRow[22] = reins;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[25] = orig1;
                        dtworkRow[26] = orig;
                        dtworkRow[27] = orig;
                        dtworkRow[77] = orig;
                        dtworkRow[28] = cedent1;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = premium1.ToString();
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = premium1.ToString();
                        }

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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



                    else if ((count != string.Empty && chck == true) && (str_sheet == ("AMERICAN STAR APPAREL")) || (str_sheet == ("BRIDGES TRAVEL & TOURS")) || (str_sheet == ("FIL - NIPPON TECH SUPPLY INC")) || (str_sheet == ("CANON MARKETING")) || (str_sheet == ("FIL-NIPPON TECH SUPPLY INC"))
                             || (str_sheet == ("FLEXIBLE AUTOMATION SYSTEM CORP")) || (str_sheet == ("FRABELLE FISHING CORPP")) || (str_sheet == ("G TRAVEL PHIL")) || (str_sheet == ("GAKKEN PHIL")) || (str_sheet == ("GOOD SHEPHERD INSURANCE")) || (str_sheet == ("GRAND PLAZA HOTEL"))
                                || (str_sheet == ("GREENOLOGY INNOVATION & SE LPG")))

                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow03 = objdt_template.NewRow();
                        dtworkRow04 = objdt_template.NewRow();
                        dtworkRow05 = objdt_template.NewRow();
                        dtworkRow[5] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = effdate;
                        dtworkRow[22] = effdate;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + " , " + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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

                        if (surplus == ("0.00 "))
                        {
                            dtworkRow[8] = "SURPLUS";
                            dtworkRow[25] = orig1;
                            dtworkRow[26] = orig;
                            dtworkRow[27] = orig;
                            dtworkRow[77] = orig;
                            dtworkRow[28] = cedent1;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow[21] = "TRENEW";
                                dtworkRow[58] = "4001";
                                dtworkRow[59] = premium1;
                            }
                            else
                            {
                                dtworkRow[21] = "TNEWBUS";
                                dtworkRow[56] = "4000";
                                dtworkRow[57] = premium1;
                            }
                            dtworkRow01.ItemArray = dtworkRow.ItemArray;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow01[59] = prem;
                            }
                            else
                            {
                                dtworkRow01[57] = prem;
                            }
                            dtworkRow02.ItemArray = dtworkRow01.ItemArray;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow02[59] = ceded;
                            }
                            else
                            {
                                dtworkRow02[57] = ceded;
                            }
                            dtworkRow01[25] = "1.00";
                            dtworkRow02[25] = "1.00";
                            dtworkRow01[26] = "1.00";
                            dtworkRow02[26] = "1.00";
                            dtworkRow01[27] = "1.00";
                            dtworkRow02[27] = "1.00";
                            dtworkRow01[77] = "1.00";
                            dtworkRow02[77] = "1.00";
                            dtworkRow01[28] = "1.00";
                            dtworkRow02[28] = "1.00";
                        }

                        else if ((surplus != ("0.00 ")) && (prem1 != ("0.00 ")) && (prem2 != ("0.00 ")) && (prem3 != ("0.00 ")))
                        {
                            dtworkRow[8] = "QA";
                            dtworkRow[25] = orig1;
                            dtworkRow[26] = orig;
                            dtworkRow[27] = orig;
                            dtworkRow[77] = orig;
                            dtworkRow[28] = cedent1;
                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow[21] = "TRENEW";
                                dtworkRow[58] = "4001";
                                dtworkRow[59] = premium1;
                            }
                            else
                            {
                                dtworkRow[21] = "TNEWBUS";
                                dtworkRow[56] = "4000";
                                dtworkRow[57] = premium1;
                            }
                            dtworkRow01.ItemArray = dtworkRow.ItemArray;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow01[59] = prem;
                            }
                            else
                            {
                                dtworkRow01[57] = prem;
                            }

                            dtworkRow01[25] = "1.00";
                            dtworkRow01[26] = "1.00";
                            dtworkRow01[27] = "1.00";
                            dtworkRow01[77] = "1.00";
                            dtworkRow01[28] = "1.00";

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow02[59] = ceded;
                            }
                            else
                            {
                                dtworkRow02[57] = ceded;
                            }

                            dtworkRow02[25] = "1.00";
                            dtworkRow02[26] = "1.00";
                            dtworkRow02[27] = "1.00";
                            dtworkRow02[77] = "1.00";
                            dtworkRow02[28] = "1.00";


                            dtworkRow04[25] = "1.00";
                            dtworkRow05[25] = "1.00";
                            dtworkRow03[25] = "1.00";
                            dtworkRow02.ItemArray = dtworkRow01.ItemArray;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow03[59] = prem1;

                            }
                            else
                            {
                                dtworkRow03[57] = prem1;
                            }

                            dtworkRow03[8] = "SURPLUS";
                            dtworkRow03[26] = surplus;
                            dtworkRow03[27] = surplus;
                            dtworkRow03[28] = "1.00";

                            dtworkRow03.ItemArray = dtworkRow02.ItemArray;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow04[59] = prem2;
                            }
                            else
                            {
                                dtworkRow04[57] = prem2;
                            }
                            dtworkRow04.ItemArray = dtworkRow03.ItemArray;

                            dtworkRow04[26] = "1.00";
                            dtworkRow04[27] = "1.00";
                            dtworkRow04[77] = "1.00";
                            dtworkRow04[28] = "1.00";

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow05[59] = prem3;

                            }
                            else
                            {
                                dtworkRow05[57] = prem3;
                            }
                            dtworkRow05[26] = "1.00";
                            dtworkRow05[27] = "1.00";
                            dtworkRow05[77] = "1.00";
                            dtworkRow05[28] = "1.00";

                            dtworkRow05.ItemArray = dtworkRow04.ItemArray;

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
                            if (dtworkRow04 != null)
                            {
                                dbl_BF += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow04[57].ToString()) ? "0" : dtworkRow04[57].ToString()
                                );
                                dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow04[59].ToString()) ? "0" : dtworkRow04[59].ToString()
                                    );
                                dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow04[61].ToString()) ? "0" : dtworkRow04[61].ToString()
                                    );
                                dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow04[63].ToString()) ? "0" : dtworkRow04[63].ToString()
                                    );
                                dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow04[77].ToString()) ? "0" : dtworkRow04[77].ToString()
                                    );

                                objdt_template.Rows.Add(dtworkRow04);
                            }
                            if (dtworkRow05 != null)
                            {
                                dbl_BF += decimal.Parse(
                                String.IsNullOrEmpty(dtworkRow05[57].ToString()) ? "0" : dtworkRow05[57].ToString()
                                );
                                dbl_BH += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow05[59].ToString()) ? "0" : dtworkRow05[59].ToString()
                                    );
                                dbl_BJ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow05[61].ToString()) ? "0" : dtworkRow05[61].ToString()
                                    );
                                dbl_BL += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow05[63].ToString()) ? "0" : dtworkRow05[63].ToString()
                                    );
                                dbl_BZ += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow05[77].ToString()) ? "0" : dtworkRow05[77].ToString()
                                    );

                                objdt_template.Rows.Add(dtworkRow05);
                            }
                        }

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

                        if (dtworkRow01 != null)
                        {
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
                        }
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

                    else if ((count != string.Empty && chck == true) && ((str_sheet == ("OMYA CHEMICAL MERCHANTS")) || (str_sheet == ("OMYA MINERALS PHIL INC")) || (str_sheet == ("YAMAZEN MACHINERY & TOOLS"))
                        || (str_sheet == ("HGL DEVT. CORP")) || (str_sheet == ("JAIME V. ONGPIN FOUNDATION")) || (str_sheet == ("KOPPEL INC")) || (str_sheet == ("LIFE GEN DISTRIBUTION INC."))
                        || (str_sheet == ("LANTRO PHILS.")) || (str_sheet == ("MEINHARDT PHILS., INC.")) || (str_sheet == ("NORWEGIAN MARITIME FOUNDATION")) || (str_sheet == ("PHIL GLOBAL COMMUNICATION"))
                        || (str_sheet == ("TRANS-PHIL GROUP")) || (str_sheet == ("WESTMINSTER SEAFARER"))))
                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow03 = objdt_template.NewRow();
                        dtworkRow04 = objdt_template.NewRow();
                        dtworkRow05 = objdt_template.NewRow();
                        dtworkRow[8] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = policy;
                        dtworkRow[22] = policy;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[25] = orig1;
                        dtworkRow[26] = orig;
                        dtworkRow[27] = orig;
                        dtworkRow[77] = orig;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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

                        if (surplus == ("0.00 "))
                        {
                            dtworkRow[25] = orig1;
                            dtworkRow[26] = orig;
                            dtworkRow[27] = orig;
                            dtworkRow[77] = orig;

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow[21] = "TRENEW";
                                dtworkRow[58] = "4001";
                                dtworkRow[59] = premium1;
                            }
                            else
                            {
                                dtworkRow[21] = "TNEWBUS";
                                dtworkRow[56] = "4000";
                                dtworkRow[57] = premium1;
                            }

                            dtworkRow01.ItemArray = dtworkRow.ItemArray;
                            dtworkRow01[8] = "GYRT";
                            dtworkRow01[25] = "1.00";
                            dtworkRow01[26] = "1.00";
                            dtworkRow01[27] = "1.00";
                            dtworkRow01[77] = "1.00";
                            dtworkRow01[28] = "1.00";

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow01[59] = prem;
                            }
                            else
                            {
                                dtworkRow01[57] = prem;
                            }
                            dtworkRow02.ItemArray = dtworkRow01.ItemArray;
                            dtworkRow02[8] = "TPD";
                            dtworkRow02[25] = "1.00";
                            dtworkRow02[26] = "1.00";
                            dtworkRow02[27] = "1.00";
                            dtworkRow02[77] = "1.00";
                            dtworkRow02[28] = "1.00";

                            if (TRANCODE == ("TRENEW"))
                            {
                                dtworkRow02[59] = ceded;
                            }
                            else
                            {
                                dtworkRow02[57] = ceded;
                            }

                        }
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

                        if (dtworkRow01 != null)
                        {
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
                        }
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
                    else if ((count != string.Empty && chck == true) && ((str_sheet == ("FIL-FACIFIC APPAREL")) || (str_sheet == ("FIRST GLIDER OPTIONS PH")) || (str_sheet == ("FLEET MARITIME SERVICES"))
                           || (str_sheet == ("MRM PHILIPPINES")) || (str_sheet == ("NORTH SEA MARINE SERVICES")) || (str_sheet == ("PHIL GLOBAL COMMUNICATION"))))
                    {
                        dtworkRow = objdt_template.NewRow();
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow03 = objdt_template.NewRow();
                        dtworkRow04 = objdt_template.NewRow();
                        dtworkRow05 = objdt_template.NewRow();
                        dtworkRow[8] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = policy;
                        dtworkRow[22] = policy;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[25] = orig;
                        dtworkRow[26] = ceded;
                        dtworkRow[27] = ceded;
                        dtworkRow[28] = premium1;
                        dtworkRow[77] = ceded;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;
                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = effdate;
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = effdate;
                        }

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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

                        dtworkRow01.ItemArray = dtworkRow.ItemArray;
                        dtworkRow01[8] = "ADDDI";
                        dtworkRow01[25] = surplus;
                        dtworkRow01[26] = prem;
                        dtworkRow01[27] = prem;
                        dtworkRow01[77] = prem;
                        dtworkRow01[28] = premium;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow01[59] = prem3;
                        }
                        else
                        {
                            dtworkRow01[57] = prem3;
                        }

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

                        if (dtworkRow01 != null)
                        {
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
                        }

                    }
                    else if ((count != string.Empty && chck == true) && (str_sheet == ("GRACE MARINE & SHIPPING CORP")))

                    {
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();

                        dtworkRow[8] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";
                        dtworkRow[19] = policy;
                        dtworkRow[22] = policy;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";
                        dtworkRow[25] = orig;
                        dtworkRow[26] = ceded;
                        dtworkRow[27] = ceded;
                        dtworkRow[28] = premium1;
                        dtworkRow[77] = ceded;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = effdate;
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = effdate;
                        }

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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

                        dtworkRow01.ItemArray = dtworkRow.ItemArray;
                        dtworkRow01[8] = "ADDDI";
                        dtworkRow01[25] = surplus;
                        dtworkRow01[26] = prem1;
                        dtworkRow01[27] = prem1;
                        dtworkRow01[77] = prem1;
                        dtworkRow01[28] = prem;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow01[59] = prem4;
                        }
                        else
                        {
                            dtworkRow01[57] = prem4;
                        }

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

                        if (dtworkRow01 != null)
                        {
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
                        }

                    }
                    else if ((count != string.Empty && chck == true) && ((str_sheet == ("HYATT INDUSTRIAL MANUFACTURING")) || (str_sheet == ("KINGWANLY ELECTRICAL&EQUIPMENT"))))

                    {
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();

                        dtworkRow[8] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";

                        dtworkRow[19] = policy;
                        dtworkRow[22] = policy;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";

                        dtworkRow[25] = orig1;
                        dtworkRow[26] = orig;
                        dtworkRow[27] = orig;
                        dtworkRow[28] = cedent1;
                        dtworkRow[77] = orig;

                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = premium1;
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = premium1;
                        }

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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
                        dtworkRow01.ItemArray = dtworkRow.ItemArray;
                        dtworkRow01[8] = "ADDDI";
                        dtworkRow01[25] = orig1;
                        dtworkRow01[26] = orig;
                        dtworkRow01[27] = orig;
                        dtworkRow01[77] = orig;
                        dtworkRow01[28] = cedent1;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow01[21] = "TRENEW";
                            dtworkRow01[58] = "4001";
                            dtworkRow01[59] = prem;
                        }
                        else
                        {
                            dtworkRow01[57] = prem;
                        }

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

                        if (dtworkRow01 != null)
                        {
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
                        }

                    }

                    else if ((count != string.Empty && chck == true) && ((str_sheet == ("SWAN INSURANCE AGENCY"))))

                    {
                        dtworkRow = objdt_template.NewRow();
                        dtworkRow[8] = "GYRT";
                        dtworkRow[7] = "GYRT";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "GRP";
                        dtworkRow[14] = "T";

                        dtworkRow[19] = policy;
                        dtworkRow[22] = policy;
                        dtworkRow[20] = policy;
                        dtworkRow[24] = "YLY";
                        dtworkRow[24] = "PHP";

                        dtworkRow[25] = orig1;
                        dtworkRow[26] = orig;
                        dtworkRow[27] = orig;
                        dtworkRow[28] = cedent1;
                        dtworkRow[77] = orig;

                        dtworkRow[29] = "NATREID";
                        dtworkRow[31] = full + "," + name;
                        dtworkRow[36] = gender;
                        dtworkRow[37] = dob;
                        dtworkRow[38] = "NONE";
                        dtworkRow[39] = "NONE";
                        dtworkRow[78] = age;
                        dtworkRow[82] = holder;

                        string fullname;
                        fullname = full + " , " + name;

                        if (TRANCODE == ("TRENEW"))
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = premium1;
                        }
                        else
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = premium1;
                        }

                        string birth;
                        int age1;
                        age1 = Convert.ToInt32(age);
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();

                        DateTime oDate = Convert.ToDateTime(dob);
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

                        dtworkRow[0] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        dtworkRow[1] = "GYRT" + str_outfname.Substring(0, 1) + str_outlname.Substring(0, 1) + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year.ToString();
                        

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
                     //   objdt_template.Rows.Add(dtworkRow01);// inpu8trow+++

                    }
                    prawrow++;
                    orig = wsraw.Cells[prawrow, 12].Text.ToString();
                    ceded = wsraw.Cells[prawrow, 16].Text.ToString();
                    gender = wsraw.Cells[prawrow, 4].Text.ToString();
                    dob = wsraw.Cells[prawrow, 5].Text.ToString();
                    premium = wsraw.Cells[prawrow, 20].Text.ToString();
                    age = wsraw.Cells[prawrow, 6].Text.ToString();
                    full = wsraw.Cells[prawrow, 2].Text.ToString();
                    name = wsraw.Cells[prawrow, 3].Text.ToString();
                    count = wsraw.Cells[prawrow, 1].Text.ToString();
                    orig1 = wsraw.Cells[prawrow, 10].Text.ToString();
                    cedent1 = wsraw.Cells[prawrow, 11].Text.ToString();
                    premium1 = wsraw.Cells[prawrow, 14].Text.ToString();
                    effdate = wsraw.Cells[prawrow, 20].Text.ToString();
                    prem = wsraw.Cells[prawrow, 15].Text.ToString();
                    orig2 = wsraw.Cells[prawrow, 8].Text.ToString();
                    surplus = wsraw.Cells[prawrow, 13].Text.ToString();
                    prem1 = wsraw.Cells[prawrow, 17].Text.ToString();
                    prem2 = wsraw.Cells[prawrow, 18].Text.ToString();
                    prem3 = wsraw.Cells[prawrow, 19].Text.ToString();
                    prem4 = wsraw.Cells[prawrow, 21].Text.ToString();
                    polnum = string.Empty;
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
                string despath = str_saved + @"\BM050-" + str_savef + ".xlsx";
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
                dtworkRow01 = null; //Dispose datarow
                dtworkRow02 = null; //Dispose datarow
                dtworkRow03= null; //Dispose datarow
                dtworkRow04 = null; //Dispose datarow
                dtworkRow05 = null; //Dispose datarow
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