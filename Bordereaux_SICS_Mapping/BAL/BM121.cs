﻿using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM121
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
                string plan = wsraw.Cells[prawrow, 6].Text.ToString();
                string reins = wsraw.Cells[prawrow, 17].Text.ToString();
                string orig= wsraw.Cells[prawrow, 20].Text.ToString();
                string ceded = wsraw.Cells[prawrow, 25].Text.ToString();
                string ret = wsraw.Cells[prawrow, 22].Text.ToString();
                string fullname = wsraw.Cells[prawrow,3].Text.ToString();
                string gender = wsraw.Cells[prawrow, 10].Text.ToString();
                string dob = wsraw.Cells[prawrow, 9].Text.ToString();
                string pref= wsraw.Cells[prawrow, 13].Text.ToString();
                string premium= wsraw.Cells[prawrow, 28].Text.ToString();

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
                    else if (polnum != string.Empty && chck == true)
                    {

                        dtworkRow = objdt_template.NewRow();

                        dtworkRow[0] = polnum;
                        dtworkRow[5] = plan;
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                
                        dtworkRow[29] = "NATREID";
                        dtworkRow[25] = orig;
                        dtworkRow[26] = ceded;
                        dtworkRow[77] = ceded;
                        dtworkRow[27] = ret;
                        dtworkRow[28] = "1.00";
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
                        dtworkRow[38] = "NONE";


                        //int bmy;
                        //string bmyear1;
                        //int effective1;
                        //eff = eff.Substring(eff.Length - 4, 4);
                        //effective1 = Convert.ToInt32(eff);
                        //oDate = Convert.ToDateTime(bmyear);
                        //bmyear1 = oDate.Year.ToString();
                        //bmy = Convert.ToInt32(bmyear1);

                        //if (bmy >= effective1)
                        //{
                        //    dtworkRow[21] = "TRENEW";
                        //    dtworkRow[58] = "4001";
                        //    dtworkRow[59] = premium.ToString();
                        //}

                        //else if (bmy < effective1)
                        //{
                        //    dtworkRow[21] = "TNEWBUS";
                        //    dtworkRow[56] = "4000";
                        //    dtworkRow[57] = premium.ToString();
                        //}

                        decimal classific;
                        classific = Convert.ToDecimal(pref);
                        if (classific == 1)
                        {
                            dtworkRow[39] = "STANDARD";
                        }
                        else if (classific == 0)
                        {
                            dtworkRow[39] = "STANDARD";
                        }
                        else if (classific == Decimal.Parse("1.25"))
                        {
                            dtworkRow[39] = "CLASSA";
                        }
                        else if (classific == Decimal.Parse("1.375"))
                        {
                            dtworkRow[39] = "CLASSAA";
                        }
                        else if (classific == Decimal.Parse("1.50"))
                        {
                            dtworkRow[39] = "CLASSB";
                        }
                        else if (classific == Decimal.Parse("1.75"))
                        {
                            dtworkRow[39] = "CLASSC";
                        }
                        else if (classific == 2)
                        {
                            dtworkRow[39] = "CLASSD";
                        }
                        else if (classific == Decimal.Parse("2.25"))
                        {
                            dtworkRow[39] = "CLASSE";
                        }
                        else if (classific == Decimal.Parse("2.5"))
                        {
                            dtworkRow[39] = "CLASSF";
                        }
                        else if (classific == Decimal.Parse("2.75"))
                        {
                            dtworkRow[39] = "CLASSG";
                        }
                        else if (classific == 3)
                        {
                            dtworkRow[39] = "CLASSH";
                        }
                        else if (classific == Decimal.Parse("3.25"))
                        {
                            dtworkRow[39] = "CLASSI";
                        }
                        else if (classific == Decimal.Parse("3.5"))
                        {
                            dtworkRow[39] = "CLASSJ";
                        }
                        else if (classific == Decimal.Parse("3.75"))
                        {
                            dtworkRow[39] = "CLASSK";
                        }
                        else if (classific == 4)
                        {
                            dtworkRow[39] = "CLASSL";
                        }
                        else if (classific == Decimal.Parse("4.25"))
                        {
                            dtworkRow[39] = "CLASSM";
                        }
                        else if (classific == Decimal.Parse("4.5"))
                        {
                            dtworkRow[39] = "CLASSN";
                        }
                        else if (classific == Decimal.Parse("4.75"))
                        {
                            dtworkRow[39] = "CLASSO";
                        }
                        else if (classific == 5)
                        {
                            dtworkRow[39] = "CLASSP";
                        }

                        if (reins.Length == 9)
                        {
                            dtworkRow[19] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 9, 4);
                            dtworkRow[20] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 9, 4);
                            dtworkRow[22] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 9, 4);
                        }
                        if (reins.Length == 10)
                        {
                            dtworkRow[19] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 10, 4);
                            dtworkRow[20] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 10, 4);
                            dtworkRow[22] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 10, 4);
                        }
                        if (reins.Length == 8)
                        {
                            dtworkRow[19] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 8, 4);
                            dtworkRow[20] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 8, 4);
                            dtworkRow[22] = reins.Substring(reins.Length - 4, 2) + "/" + reins.Substring(reins.Length - 2, 2) + "/" + reins.Substring(reins.Length - 8, 4);
                        }
                        dtworkRow[37] = dob;

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

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 4].Text.ToString();
                    plan = wsraw.Cells[prawrow, 6].Text.ToString();
                    reins = wsraw.Cells[prawrow, 17].Text.ToString();
                    orig = wsraw.Cells[prawrow, 20].Text.ToString();
                    ceded = wsraw.Cells[prawrow, 25].Text.ToString();
                    ret = wsraw.Cells[prawrow, 22].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                    gender = wsraw.Cells[prawrow, 10].Text.ToString();
                    dob = wsraw.Cells[prawrow, 9].Text.ToString();
                    pref = wsraw.Cells[prawrow, 13].Text.ToString();
                    premium = wsraw.Cells[prawrow, 28].Text.ToString();
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
                string despath = str_saved + @"\BM121-" + str_savef + ".xlsx";
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