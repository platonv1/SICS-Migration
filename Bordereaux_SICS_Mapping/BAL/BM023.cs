using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM023
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
                decimal dbl_BF01 = 0, dbl_BF02 = 0, dbl_BF03 = 0, dbl_BF04 = 0, dbl_BF05 = 0, dbl_BF06 = 0, dbl_BF07 = 0, dbl_BF08 = 0, dbl_BF09 = 0, dbl_BF10 = 0, dbl_BF11 = 0, dbl_BF12 = 0,
                    dbl_BH01 = 0, dbl_BH02 = 0, dbl_BH03 = 0, dbl_BH04 = 0, dbl_BH05 = 0, dbl_BH06 = 0, dbl_BH07 = 0, dbl_BH08 = 0, dbl_BH09 = 0, dbl_BH10 = 0, dbl_BH11 = 0, dbl_BH12 = 0,
                    dbl_BJ01 = 0, dbl_BJ02 = 0, dbl_BJ03 = 0, dbl_BJ04 = 0, dbl_BJ05 = 0, dbl_BJ06 = 0, dbl_BJ07 = 0, dbl_BJ08 = 0, dbl_BJ09 = 0, dbl_BJ10 = 0, dbl_BJ11 = 0, dbl_BJ12 = 0,
                    dbl_BL01 = 0, dbl_BL02 = 0, dbl_BL03 = 0, dbl_BL04 = 0, dbl_BL05 = 0, dbl_BL06 = 0, dbl_BL07 = 0, dbl_BL08 = 0, dbl_BL09 = 0, dbl_BL10 = 0, dbl_BL11 = 0, dbl_BL12 = 0,
                    dbl_BZ01 = 0, dbl_BZ02 = 0, dbl_BZ03 = 0, dbl_BZ04 = 0, dbl_BZ05 = 0, dbl_BZ06 = 0, dbl_BZ07 = 0, dbl_BZ08 = 0, dbl_BZ09 = 0, dbl_BZ10 = 0, dbl_BZ11 = 0, dbl_BZ12 = 0;
                #endregion

                System.Data.DataRow dtworkRow01;
                System.Data.DataRow dtworkRow02;
                Helper objHlpr = new Helper();
                DataTable objdt_template01 = new DataTable();
                DataTable objdt_template02 = new DataTable();
                DataTable objdt_template03 = new DataTable();
                DataTable objdt_template04 = new DataTable();
                DataTable objdt_template05 = new DataTable();
                DataTable objdt_template06 = new DataTable();
                DataTable objdt_template07 = new DataTable();
                DataTable objdt_template08 = new DataTable();
                DataTable objdt_template09 = new DataTable();
                DataTable objdt_template10 = new DataTable();
                DataTable objdt_template11 = new DataTable();
                DataTable objdt_template12 = new DataTable();

                objdt_template01 = objHlpr.dt_formtemplate("JAN" + str_sheet);
                objdt_template02 = objHlpr.dt_formtemplate("FEB" + str_sheet);
                objdt_template03 = objHlpr.dt_formtemplate("MAR" + str_sheet);
                objdt_template04 = objHlpr.dt_formtemplate("APR" + str_sheet);
                objdt_template05 = objHlpr.dt_formtemplate("MAY" + str_sheet);
                objdt_template06 = objHlpr.dt_formtemplate("JUN" + str_sheet);
                objdt_template07 = objHlpr.dt_formtemplate("JUL" + str_sheet);
                objdt_template08 = objHlpr.dt_formtemplate("AUG" + str_sheet);
                objdt_template09 = objHlpr.dt_formtemplate("SEP" + str_sheet);
                objdt_template10 = objHlpr.dt_formtemplate("OCT" + str_sheet);
                objdt_template11 = objHlpr.dt_formtemplate("NOV" + str_sheet);
                objdt_template12 = objHlpr.dt_formtemplate("DEC" + str_sheet);

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

                int year1;
                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string branded = wsraw.Cells[prawrow, 2].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 6].Text.ToString();
                string sex = wsraw.Cells[prawrow, 5].Text.ToString();
                string dob = wsraw.Cells[prawrow, 7].Text.ToString();
                string age = wsraw.Cells[prawrow, 8].Text.ToString();
                string month = wsraw.Cells[prawrow, 1].Text.ToString();
                string effd = wsraw.Cells[prawrow, 3].Text.ToString();
                string remarks = wsraw.Cells[prawrow, 4].Text.ToString();
                string pref = wsraw.Cells[prawrow, 9].Text.ToString();
                string ai = wsraw.Cells[prawrow, 35].Text.ToString();
                string year = wsraw.Cells[3][1].Text.ToString();
                string gender = string.Empty;
                year = year.Substring(year.Length - 4, 4);
                year1 = Convert.ToInt32(year);

                string tempmonth = string.Empty;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                int storee;
                bool chck;
                int reind1;
                double mul = .15;
                double dbl_xOne = 1;

                double osa1 = 0, osa2 = 0, isar1 = 0, isar2 = 0, prem1 = 0, prem2 = 0, sat1 = 0, sat2 = 0;
         
                #region Data Processing
                while (rowcount != erawrow + 2)
                {
                    chck = false;
                    string[] polnum1 = polnum.Split('-');

                    if (polnum1.Length == 3)
                    {
                        if (int.TryParse(polnum1[1], out storee) && int.TryParse(polnum1[2], out storee))
                        {
                            chck = true;
                        }
                    }
                    else if (int.TryParse(polnum, out storee))
                    {
                        chck = true;
                    }
                    if (month.ToUpper().Contains("JAN") || (month.ToUpper().Contains("FEB")) || (month.ToUpper().Contains("MAR")) || (month.ToUpper().Contains("APR"))
                        || (month.ToUpper().Contains("MAY")) || (month.ToUpper().Contains("JUN")) || (month.ToUpper().Contains("JUL")) || (month.ToUpper().Contains("AUG"))
                        || (month.ToUpper().Contains("SEP")) || (month.ToUpper().Contains("OCT")) || (month.ToUpper().Contains("NOV")) || (month.ToUpper().Contains("DEC")))

                    {
                        tempmonth = month.Substring(0, 3);
                    }

                    
                    if (chck == true)
                    {
                        dtworkRow01 = null;
                        dtworkRow02 = null;

                        switch (tempmonth.ToUpper())
                        {
                            case "JAN":
                                dtworkRow01 = objdt_template01.NewRow();
                                dtworkRow02 = objdt_template01.NewRow();
                                break;
                            case "FEB":
                                dtworkRow01 = objdt_template02.NewRow();
                                dtworkRow02 = objdt_template02.NewRow();
                                break;
                            case "MAR":
                                dtworkRow01 = objdt_template03.NewRow();
                                dtworkRow02 = objdt_template03.NewRow();
                                break;
                            case "APR":
                                dtworkRow01 = objdt_template04.NewRow();
                                dtworkRow02 = objdt_template04.NewRow();
                                break;
                            case "MAY":
                                dtworkRow01 = objdt_template05.NewRow();
                                dtworkRow02 = objdt_template05.NewRow();
                                break;
                            case "JUN":
                                dtworkRow01 = objdt_template06.NewRow();
                                dtworkRow02 = objdt_template06.NewRow();
                                break;
                            case "JUL":
                                dtworkRow01 = objdt_template07.NewRow();
                                dtworkRow02 = objdt_template07.NewRow();
                                break;
                            case "AUG":
                                dtworkRow01 = objdt_template08.NewRow();
                                dtworkRow02 = objdt_template08.NewRow();
                                break;
                            case "SEP":
                                dtworkRow01 = objdt_template09.NewRow();
                                dtworkRow02 = objdt_template09.NewRow();
                                break;
                            case "OCT":
                                dtworkRow01 = objdt_template10.NewRow();
                                dtworkRow02 = objdt_template10.NewRow();
                                break;
                            case "NOV":
                                dtworkRow01 = objdt_template11.NewRow();
                                dtworkRow02 = objdt_template11.NewRow();
                                break;
                            case "DEC":
                                dtworkRow01 = objdt_template12.NewRow();
                                dtworkRow02 = objdt_template12.NewRow();
                                break;
                        }
                        dtworkRow01[0] = polnum;
                        dtworkRow01[5] = branded;
                        dtworkRow01[8] = "SURPLUS";
                        dtworkRow01[9] = "PAFM";
                        dtworkRow01[10] = "S";
                        dtworkRow01[13] = "IND";
                        dtworkRow01[14] = "T";
                        dtworkRow01[23] = "PHP";
                        dtworkRow01[29] = "NATREID";
                        dtworkRow01[24] = "YLY";
                        dtworkRow01[38] = "NONE";
                        dtworkRow01[26] = "1.00";
                        dtworkRow01[28] = "1.00";
                        dtworkRow01[31] = fullname;
                        dtworkRow01[36] = sex;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR4AL" : dtworkRow01[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow01[37] = dob.ToString();
                        dtworkRow01[79] = age;
                        dtworkRow01[20] = effd;
                        dtworkRow01[22] = effd.Substring(0,effd.Length -4) + year1;
                        dtworkRow01[41] = year1;
                        dtworkRow01[76] = remarks;

                        dtworkRow01[39] = objHlpr.fn_getmortality(pref);
                        if (objHlpr.fn_isDMort(dtworkRow01[39].ToString()))
                        {
                            dtworkRow01[39] = "STANDARD";
                            dtworkRow01[76] = String.IsNullOrEmpty(dtworkRow01[76].ToString()) ? "BR8AN" : dtworkRow01[76].ToString() + "|BR8AN";
                        }

                        dtworkRow01[25] = osa1.ToString();
                        dtworkRow01[27] = isar1.ToString();
                        dtworkRow01[77] = sat1.ToString();

                        int effd2;
                        effd2 = Convert.ToInt32(effd.Substring(effd.Length - 4, 4));

                        if (remarks.ToUpper().Contains("TERMINATED"))
                        {
                            dtworkRow01[21] = "TCONTER";
                            dtworkRow01[62] = "4004";
                            dtworkRow01[56] = "";
                            dtworkRow01[57] = "";
                            dtworkRow01[58] = "";
                            dtworkRow01[59] = "";
                            dtworkRow01[63] = prem1;
                        }
                        else
                        {
                            if (effd2 >= year1)
                            {
                                dtworkRow01[21] = "TNEWBUS";
                                dtworkRow01[56] = "4000";
                                dtworkRow01[57] = prem1;
                            }
                            if (effd2 < year1)
                            {
                                dtworkRow01[21] = "TRENEW";
                                dtworkRow01[58] = "4001";
                                dtworkRow01[59] = prem1;
                            }
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

                        if (str_outfname.Substring(str_outfname.Length - 1) == str_MI.Trim())
                        {
                            dtworkRow01[33] = str_outfname.Substring(0, str_outfname.Length - 1);
                            //str_outfname.Replace(" " + str_MI, string.Empty);
                        }
                        else
                        {
                            dtworkRow01[33] = str_outfname;
                        }
                        

                        dtworkRow01[30] = str_outlifeid;

                        if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            dtworkRow01[36] = objHlpr.fn_getgender(str_gender, dtworkRow01[33].ToString());
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

                        

                        //ISSUE#010-Start---------
                        if (String.IsNullOrEmpty(dtworkRow01[19].ToString()))
                        {
                            if (dtworkRow01[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                dtworkRow01[19] = dtworkRow01[20];
                            }
                            else
                            {
                                dtworkRow01[19] = dtworkRow01[22];
                            }
                        }
                        //ISSUE#010-End-----------
                        #endregion


                        dtworkRow02.ItemArray = dtworkRow01.ItemArray;
                        dtworkRow02[25] = osa2.ToString();
                        dtworkRow02[27] = isar2.ToString();
                        dtworkRow02[77] = sat2.ToString();
                        dtworkRow02[5] = "ADB";

                        if (effd2 >= year1)
                        {
                            dtworkRow02[21] = "TNEWBUS";
                            dtworkRow02[56] = "4000";
                            dtworkRow02[57] = prem2;
                        }
                        if (effd2 < year1)
                        {
                            dtworkRow02[21] = "TRENEW";
                            dtworkRow02[58] = "4001";
                            dtworkRow02[59] = prem2;
                        }
                        if (remarks.ToUpper().Contains("TERMINATED"))
                        {
                            dtworkRow02[21] = "TCONTER";
                            dtworkRow02[62] = "4004";
                            dtworkRow02[56] = "";
                            dtworkRow02[57] = "";
                            dtworkRow02[58] = "";
                            dtworkRow02[59] = "";
                            dtworkRow02[63] = prem2;
                        }


                        //ISSUE#017-Start---------
                        if (dtworkRow01[25].ToString() == "0")
                        {
                            dtworkRow01[25] = "1";
                        }
                        if (dtworkRow01[26].ToString() == "0")
                        {
                            dtworkRow01[26] = "1";
                        }
                        if (dtworkRow01[27].ToString() == "0")
                        {
                            dtworkRow01[27] = "1";
                        }
                        if (dtworkRow01[28].ToString() == "0")
                        {
                            dtworkRow01[28] = "1";
                        }
                        if (dtworkRow01[77].ToString() == "0")
                        {
                            dtworkRow01[77] = "1";
                        }


                        if (dtworkRow02[25].ToString() == "0")
                        {
                            dtworkRow02[25] = "1";
                        }
                        if (dtworkRow02[26].ToString() == "0")
                        {
                            dtworkRow02[26] = "1";
                        }
                        if (dtworkRow02[27].ToString() == "0")
                        {
                            dtworkRow02[27] = "1";
                        }
                        if (dtworkRow02[28].ToString() == "0")
                        {
                            dtworkRow02[28] = "1";
                        }
                        if (dtworkRow02[77].ToString() == "0")
                        {
                            dtworkRow02[77] = "1";
                        }
                        //ISSUE#017-End-----------


                        switch (tempmonth.ToUpper())
                        {
                            case "JAN":
                                objdt_template01.Rows.Add(dtworkRow01);

                                dbl_BF01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );

                                if (dtworkRow02 != null)
                                {
                                    dbl_BF01 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH01 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ01 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL01 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ01 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template01.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "FEB":
                                objdt_template02.Rows.Add(dtworkRow01);

                                dbl_BF02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );

                                if (dtworkRow02 != null)
                                {
                                    dbl_BF02 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH02 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ02 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL02 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ02 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template02.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "MAR":
                                objdt_template03.Rows.Add(dtworkRow01);

                                dbl_BF03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF03 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH03 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ03 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL03 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ03 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template03.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "APR":
                                objdt_template04.Rows.Add(dtworkRow01);

                                dbl_BF04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF04 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH04 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ04 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL04 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ04 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template04.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "MAY":
                                objdt_template05.Rows.Add(dtworkRow01);

                                dbl_BF05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF05 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH05 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ05 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL05 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ05 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template05.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "JUN":
                                objdt_template06.Rows.Add(dtworkRow01);

                                dbl_BF06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF06 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH06 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ06 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL06 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ06 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template06.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "JUL":
                                objdt_template07.Rows.Add(dtworkRow01);

                                dbl_BF07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF07 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH07 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ07 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL07 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ07 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template07.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "AUG":
                                objdt_template08.Rows.Add(dtworkRow01);

                                dbl_BF08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF08 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH08 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ08 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL08 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ08 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template08.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "SEP":
                                objdt_template09.Rows.Add(dtworkRow01);

                                dbl_BF09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF09 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH09 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ09 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL09 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ09 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template09.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "OCT":
                                objdt_template10.Rows.Add(dtworkRow01);

                                dbl_BF10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF10 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH10 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ10 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL10 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ10 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template10.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "NOV":
                                objdt_template11.Rows.Add(dtworkRow01);

                                dbl_BF11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF11 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH11 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ11 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL11 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ11 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template11.Rows.Add(dtworkRow02);
                                }
                                break;
                            case "DEC":
                                objdt_template12.Rows.Add(dtworkRow01);

                                dbl_BF12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[57].ToString()) ? "0" : dtworkRow01[57].ToString()
                                    );
                                dbl_BH12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[59].ToString()) ? "0" : dtworkRow01[59].ToString()
                                    );
                                dbl_BJ12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[61].ToString()) ? "0" : dtworkRow01[61].ToString()
                                    );
                                dbl_BL12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[63].ToString()) ? "0" : dtworkRow01[63].ToString()
                                    );
                                dbl_BZ12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow01[77].ToString()) ? "0" : dtworkRow01[77].ToString()
                                    );
                                if (dtworkRow02 != null)
                                {
                                    dbl_BF12 += decimal.Parse(
                                    String.IsNullOrEmpty(dtworkRow02[57].ToString()) ? "0" : dtworkRow02[57].ToString()
                                    );
                                    dbl_BH12 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[59].ToString()) ? "0" : dtworkRow02[59].ToString()
                                        );
                                    dbl_BJ12 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[61].ToString()) ? "0" : dtworkRow02[61].ToString()
                                        );
                                    dbl_BL12 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[63].ToString()) ? "0" : dtworkRow02[63].ToString()
                                        );
                                    dbl_BZ12 += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow02[77].ToString()) ? "0" : dtworkRow02[77].ToString()
                                        );

                                    objdt_template12.Rows.Add(dtworkRow02);
                                }
                                break;
                            default:
                                break;
                        }
                    }

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                    branded = wsraw.Cells[prawrow, 2].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 6].Text.ToString();
                    sex = wsraw.Cells[prawrow, 5].Text.ToString();
                    dob = wsraw.Cells[prawrow, 7].Text.ToString();
                    age = wsraw.Cells[prawrow, 8].Text.ToString();
                    month = wsraw.Cells[prawrow, 1].Text.ToString();
                    effd = wsraw.Cells[prawrow, 3].Text.ToString();
                    remarks = wsraw.Cells[prawrow, 4].Text.ToString();
                    pref = wsraw.Cells[prawrow, 9].Text.ToString();
                    year = wsraw.Cells[3][1].Text.ToString();
                    ai = wsraw.Cells[prawrow, 35].Text.ToString();

                    gender = string.Empty;
                    
                    osa1 = 0.00;
                    osa2 = 0.00;
                    isar1 = 0.00;
                    isar2 = 0.00;
                    prem1 = 0.00;
                    prem2 = 0.00;
                    sat1 = 0.00;
                    sat2 = 0.00;
                    
                    if (!remarks.ToUpper().Contains("TERMINATED"))
                    {
                        chck = false;
                        polnum1 = polnum.Split('-');

                        if (polnum1.Length == 3)
                        {
                            if (int.TryParse(polnum1[1], out storee) && int.TryParse(polnum1[2], out storee))
                            {

                                if (branded.ToUpper().Contains("MPP"))
                                {
                                    try
                                    {
                                        osa1 = double.Parse(wsraw.Cells[prawrow, 16].Text.ToString()) * dbl_xOne;
                                    }
                                    catch { }

                                    try
                                    {
                                        osa2 = double.Parse(wsraw.Cells[prawrow, 18].Text.ToString()) * dbl_xOne;
                                    }
                                    catch { }

                                }
                                else //if (branded.ToUpper().Contains("GSP"))
                                {
                                    try
                                    {
                                        osa1 = double.Parse(wsraw.Cells[prawrow, 11].Text.ToString()) * dbl_xOne;
                                    }
                                    catch { }

                                    try
                                    {
                                        osa2 = double.Parse(wsraw.Cells[prawrow, 14].Text.ToString()) * dbl_xOne;
                                    }
                                    catch { }

                                }

                                try
                                {
                                    isar1 = double.Parse(wsraw.Cells[prawrow, 20].Text.ToString()) * mul;
                                }
                                catch { }

                                try
                                {
                                    isar2 = double.Parse(wsraw.Cells[prawrow, 23].Text.ToString()) * mul;
                                }
                                catch { }

                                try
                                {
                                    prem1 = double.Parse(wsraw.Cells[prawrow, 34].Text.ToString()) * mul;
                                }
                                catch { }

                                try
                                {
                                    prem2 = double.Parse(wsraw.Cells[prawrow, 35].Text.ToString()) * mul;
                                }
                                catch { }

                                try
                                {
                                    sat1 = double.Parse(wsraw.Cells[prawrow, 30].Text.ToString()) * mul;
                                }
                                catch { }

                                try
                                {
                                    sat2 = double.Parse(wsraw.Cells[prawrow, 31].Text.ToString()) * mul;
                                }
                                catch { }

                            }
                        }
                        else if (polnum1.Length == 1) 
                        {
                            try
                            {
                                osa1 = double.Parse(wsraw.Cells[prawrow, 11].Text.ToString()) * mul;
                            }
                            catch { }

                            try
                            {
                                osa2 = double.Parse(wsraw.Cells[prawrow, 14].Text.ToString()) * mul;
                            }
                            catch { }
                        }
                    }

                    

                    rowcount++;
                }
                #endregion
                string despath = string.Empty;
                if (objdt_template01.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template01.NewRow();
                    objdt_template01.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template01.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF01 + dbl_BH01 + dbl_BJ01 + dbl_BL01;
                    objdt_template01.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template01.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ01;
                    objdt_template01.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template01, boo_open, str_saved + @"\BM023-JAN-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template02.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template02.NewRow();
                    objdt_template02.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template02.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF02 + dbl_BH02 + dbl_BJ02 + dbl_BL02;
                    objdt_template02.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template02.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ02;
                    objdt_template02.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template02, boo_open, str_saved + @"\BM023-FEB-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template03.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template03.NewRow();
                    objdt_template03.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template03.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF03 + dbl_BH03 + dbl_BJ03 + dbl_BL03;
                    objdt_template03.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template03.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ03;
                    objdt_template03.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template03, boo_open, str_saved + @"\BM023-MAR-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template04.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template04.NewRow();
                    objdt_template04.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template04.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF04 + dbl_BH04 + dbl_BJ04 + dbl_BL04;
                    objdt_template04.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template04.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ04;
                    objdt_template04.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template04, boo_open, str_saved + @"\BM023-APR-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template05.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template05.NewRow();
                    objdt_template05.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template05.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF05 + dbl_BH05 + dbl_BJ05 + dbl_BL05;
                    objdt_template05.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template05.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ05;
                    objdt_template05.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template05, boo_open, str_saved + @"\BM023-MAY-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template06.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template06.NewRow();
                    objdt_template06.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template06.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF06 + dbl_BH06 + dbl_BJ06 + dbl_BL06;
                    objdt_template06.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template06.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ06;
                    objdt_template06.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template06, boo_open, str_saved + @"\BM023-JUN-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template07.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template07.NewRow();
                    objdt_template07.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template07.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF07 + dbl_BH07 + dbl_BJ07 + dbl_BL07;
                    objdt_template07.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template07.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ07;
                    objdt_template07.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template07, boo_open, str_saved + @"\BM023-JUL-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template08.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template08.NewRow();
                    objdt_template08.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template08.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF08 + dbl_BH08 + dbl_BJ08 + dbl_BL08;
                    objdt_template08.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template08.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ08;
                    objdt_template08.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template08, boo_open, str_saved + @"\BM023-AUG-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template09.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template09.NewRow();
                    objdt_template09.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template09.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF09 + dbl_BH09 + dbl_BJ09 + dbl_BL09;
                    objdt_template09.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template09.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ09;
                    objdt_template09.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template09, boo_open, str_saved + @"\BM023-SEP-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template10.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template10.NewRow();
                    objdt_template10.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template10.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF10 + dbl_BH10 + dbl_BJ10 + dbl_BL10;
                    objdt_template10.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template10.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ10;
                    objdt_template10.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template10, boo_open, str_saved + @"\BM023-OCT-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template11.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template11.NewRow();
                    objdt_template11.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template11.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF11 + dbl_BH11 + dbl_BJ11 + dbl_BL11;
                    objdt_template11.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template11.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ11;
                    objdt_template11.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template11, boo_open, str_saved + @"\BM023-NOV-" + str_savef + ".xlsx");
                }
                //---------------------------------------------------------------------------------------------------------------
                if (objdt_template12.Rows.Count > 0)
                {
                    #region "Compute Hash Total"
                    dtworkRow01 = objdt_template12.NewRow();
                    objdt_template12.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template12.NewRow();
                    dtworkRow01[0] = "Total Premium:";
                    dtworkRow01[1] = dbl_BF12 + dbl_BH12 + dbl_BJ12 + dbl_BL12;
                    objdt_template12.Rows.Add(dtworkRow01);

                    dtworkRow01 = objdt_template12.NewRow();
                    dtworkRow01[0] = "Total Sum at Risk:";
                    dtworkRow01[1] = dbl_BZ12;
                    objdt_template12.Rows.Add(dtworkRow01);
                    #endregion
                    objHlpr.fn_savemultiple(objdt_template12, boo_open, str_saved + @"\BM023-DEC-" + str_savef + ".xlsx");
                }
                eapp.DisplayAlerts = false;
                wsraw = null;
                wbraw.SaveAs(str_raw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
                wbraw.Close();
                wbraw = null;
                eapp = null;
                ////
                dtworkRow01 = null; //Dispose datarow
                dtworkRow02 = null;
                objdt_template01.Dispose();
                objdt_template01 = null;
                objdt_template02.Dispose();
                objdt_template02 = null;
                objdt_template03.Dispose();
                objdt_template03 = null;
                objdt_template04.Dispose();
                objdt_template04 = null;
                objdt_template05.Dispose();
                objdt_template05 = null;
                objdt_template06.Dispose();
                objdt_template06 = null;
                objdt_template07.Dispose();
                objdt_template07 = null;
                objdt_template08.Dispose();
                objdt_template08 = null;
                objdt_template09.Dispose();
                objdt_template09 = null;
                objdt_template10.Dispose();
                objdt_template10 = null;
                objdt_template11.Dispose();
                objdt_template11 = null;
                objdt_template12.Dispose();
                objdt_template12 = null;

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
