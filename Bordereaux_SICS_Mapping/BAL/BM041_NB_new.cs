using System;
using System.Data;
using System.Linq;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM041_NB_new
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

                decimal dbl_BF_PHP = 0, dbl_BH_PHP = 0, dbl_BJ_PHP = 0, dbl_BL_PHP = 0, dbl_BZ_PHP = 0,
                       dbl_BF_USD = 0, dbl_BH_USD = 0, dbl_BJ_USD = 0, dbl_BL_USD = 0, dbl_BZ_USD = 0;

                if (boo_clean)
                {
                    wsraw = objHlpr.fn_extendwidth(wsraw);
                }

                int erawrow = rawrange.Rows.Count;
                int erawcol = rawrange.Columns.Count;

                int prawrow = 1;

                string busmean = "", btype = "", remarks;
                string adjamt = wsraw.Cells[prawrow, 31].Text.ToString();
                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                string fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string gender = wsraw.Cells[prawrow, 5].Text.ToString();
                string smoker = wsraw.Cells[prawrow, 6].Text.ToString();
                string bustype = wsraw.Cells[prawrow, 8].Text.ToString();
                string age = wsraw.Cells[prawrow, 7].Text.ToString();
                string paid = wsraw.Cells[prawrow, 14].Text.ToString();
                string age1 = wsraw.Cells[prawrow, 9].Text.ToString();
                string rating = wsraw.Cells[prawrow, 12].Text.ToString();
                string status = wsraw.Cells[prawrow, 4].Text.ToString();
                string premyr = wsraw.Cells[prawrow, 10].Text.ToString();
                string polnum2 = wsraw.Cells[prawrow, 1].Text.ToString();
                string branded = wsraw.Cells[prawrow, 2].Text.ToString();
                string curr = wsraw.Cells[prawrow, 16].Text.ToString();
                string cededsum = wsraw.Cells[prawrow, 4].Text.ToString();
                string inisum = wsraw.Cells[prawrow, 8].Text.ToString();
                string inisum2 = wsraw.Cells[prawrow, 8].Text.ToString();
                string effdt = wsraw.Cells[prawrow, 3].Text.ToString();
                string prem = wsraw.Cells[prawrow, 12].Text.ToString();
                string classpref = wsraw.Cells[prawrow, 9].Text.ToString();
                string premyr1 = wsraw.Cells[prawrow, 12].Text.ToString();
                string total = wsraw.Cells[prawrow, 15].Text.ToString();
                string code = wsraw.Cells[prawrow, 2].Text.ToString();
                string risk = wsraw.Cells[prawrow, 5].Text.ToString();
                string transa = wsraw.Cells[prawrow, 6].Text.ToString();
                string adjprem = wsraw.Cells[prawrow, 17].Text.ToString();
                string busty = wsraw.Cells[prawrow, 20].Text.ToString();
                string adjcurr = wsraw.Cells[prawrow, 21].Text.ToString();
                string adjyr = wsraw.Cells[1][1].Text.ToString();
                string pref = wsraw.Cells[prawrow, 15].Text.ToString();
                string adjs = wsraw.Cells[prawrow, 10].Text.ToString();
                string adjtc = wsraw.Cells[prawrow, 12].Text.ToString();
                string adjbt = wsraw.Cells[prawrow, 13].Text.ToString();
                string adjcur = wsraw.Cells[prawrow, 14].Text.ToString();
                string adjc = wsraw.Cells[prawrow, 19].Text.ToString();
                string fyry = wsraw.Cells[prawrow, 6].Text.ToString();
                string polstore;
                string bcstore;
                string ipstore;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                int storee;
                double prefstore;
                bool chck;
                bool cla;

                #region Data Processing
               
                    while (rowcount != erawrow + 2)
                    {
                        chck = int.TryParse(polnum, out storee);

                        if (polnum != string.Empty && chck == false)
                        {
                            findboo = false;
                            if (str_sheet.Contains("ADJ"))
                            {
                                TRANCODE = "ADJUST";
                            }
                            else if (str_sheet.Contains("REN"))
                            {
                                TRANCODE = "TRENEW";
                            }
                            else if (str_sheet.Contains("NB"))
                            {
                                TRANCODE = "TNEWBUS";
                            }
                            else
                            {
                                TRANCODE = "TRENEW";
                            }
                        }
                        else if (chck == true)
                        {
                            polnum2 = wsraw.Cells[prawrow, 26].Text.ToString();
                            branded = wsraw.Cells[prawrow, 3].Text.ToString();
                            remarks = wsraw.Cells[prawrow, 24].Text.ToString();
                            btype = wsraw.Cells[prawrow, 4].Text.ToString();
                            effdt = wsraw.Cells[prawrow, 5].Text.ToString();
                            cededsum = wsraw.Cells[prawrow, 8].Text.ToString();
                            inisum = wsraw.Cells[prawrow, 13].Text.ToString();
                            inisum2 = wsraw.Cells[prawrow, 13].Text.ToString();
                            classpref = wsraw.Cells[prawrow, 15].Text.ToString();
                            curr = wsraw.Cells[prawrow, 25].Text.ToString();
                            prem = wsraw.Cells[prawrow, 20].Text.ToString();
                            age = wsraw.Cells[prawrow, 11].Text.ToString();
                            age1 = wsraw.Cells[prawrow,14].Text.ToString();
                            code = wsraw.Cells[prawrow, 2].Text.ToString();
                            risk = wsraw.Cells[prawrow, 9].Text.ToString();
                            transa = wsraw.Cells[prawrow, 10].Text.ToString();
                            adjprem = wsraw.Cells[prawrow, 18].Text.ToString();
                            busty = wsraw.Cells[prawrow, 20].Text.ToString();
                            adjcurr = wsraw.Cells[prawrow, 21].Text.ToString();
                            bustype = wsraw.Cells[prawrow, 12].Text.ToString();
                            rating = wsraw.Cells[prawrow, 19].Text.ToString();
                            total = wsraw.Cells[prawrow, 23].Text.ToString();
                            
                                dtworkRow = objdt_template.NewRow();
                                string fac = "T";
                                string fac1 ="F";

                                if (bustype == "F")
                                {
                                    dtworkRow[14] = fac1;
                                    dtworkRow[83] = objHlpr.fn_getrefcode(fac1);
                                }
                                else if (bustype == "A")
                                {
                                    dtworkRow[14] = fac;
                                    dtworkRow[83] = objHlpr.fn_getrefcode(fac);
                                }
                                //dtworkRow[83] = objHlpr.fn_getrefcode(busmean);
                                if (curr.ToString() == "PESO")
                                {
                                    dtworkRow[23] = "PHP";
                                }
                                else
                                {
                                    dtworkRow[23] = "USD";
                                }

                                if (smoker.ToUpper().Contains("N"))
                                {
                                    dtworkRow[38] = "NSMOK";
                                }
                                else if (smoker.ToUpper().Contains("S"))
                                {
                                    dtworkRow[38] = "SMOK";
                                }

                                if (polnum.StartsWith("08"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "VARLIFE-GU";
                                }
                                else if (!polnum.StartsWith("08") && !(btype == "ADB" || btype == "TDB"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "TRADITIONALLIFE";
                                }
                                else if (!polnum.StartsWith("08") && (btype == "ADB"))
                                {
                                    dtworkRow[3] = "DEATH";
                                    dtworkRow[4] = "ADB-IND";
                                }
                                else if (!polnum.StartsWith("08") && (btype == "TDB"))
                                {
                                    dtworkRow[3] = "DISAB";
                                    dtworkRow[4] = "WOPDIIND";
                                }

                                dtworkRow[0] = "'" + polnum.ToString().Trim(new char[0]);
                                dtworkRow[1] = "'" + polnum.ToString().Trim(new char[0]);
                                dtworkRow[5] = branded.ToString();
                                dtworkRow[8] = "SURPLUS";
                                dtworkRow[10] = "S";
                                dtworkRow[9] = "PAFM";
                                dtworkRow[13] = "IND";
                                dtworkRow[41] = rating.ToString();
                                dtworkRow[20] = paid.ToString();//
                                dtworkRow[24] = "YLY";
                                dtworkRow[25] = cededsum.ToString();
                                dtworkRow[27] = inisum2.ToString();
                                dtworkRow[77] = inisum2.ToString();
                                dtworkRow[29] = "NATREID";
                                dtworkRow[33] = fullnames.ToString();
                                dtworkRow[36] = gender.ToString();
                                dtworkRow[76] = remarks;
                                dtworkRow[21] = TRANCODE;
                                #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                                {
                                    dob = "7/1/1900";
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                                }
                                #endregion
                                dtworkRow[37] = dob.ToString();

                                    dtworkRow[22] = effdt;
                              

                                dtworkRow[79] = age.ToString();


                        if (str_sheet.Contains("ADJ"))
                        {
                            if (fyry == "FY")
                            {
                                dtworkRow[60] = "4002";
                                dtworkRow[61] = adjamt.ToString();
                            }
                            else
                            {
                                dtworkRow[62] = "4004";
                                dtworkRow[63] = adjamt.ToString();
                            }
                        }
                        else 
                        {
                            if (TRANCODE.Contains("TNEWBUS"))
                            {
                                dtworkRow[56] = "4000";
                                dtworkRow[57] = total.ToString();
                            }
                            else if (TRANCODE.Contains("TRENEW"))
                            {
                                dtworkRow[58] = "4001";
                                dtworkRow[59] = total.ToString();
                            }
                            else if (TRANCODE.Contains("ADJUST"))
                            {
                                dtworkRow[60] = "4002";
                                dtworkRow[61] = adjamt.ToString();
                            }
                        }
                            


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

                                #region "New Requirements - No Name"
                                if (String.IsNullOrEmpty(fullnames))
                                {
                                    fullnames = polnum.ToString();
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR6AF" : dtworkRow[76].ToString() + "|BR6AF";
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
                        //    dtworkRow[34] = str_MI;
                        //}

                        //Updated logic for names 05 / 19 / 2022
                        dtworkRow [31] = fullnames; /*objHlpr.fn_stringcleanup(fullnames);*/

                        objHlpr2.fn_separateLastNameFirstNameV4(fullnames, out fullnames, out string strLastName, out string strFirstName, out string strMiddleInitial);

                        dtworkRow [32] = objHlpr2.fn_removeCharacters(strLastName);/*str_outlname;*/

                        dtworkRow [33] = objHlpr2.fn_removeCharacters(strFirstName);/*str_outfname.Replace(" " + str_MI, string.Empty);*/

                        dtworkRow [30] = objHlpr.fn_LifeID(strFirstName, strLastName, dob);/*str_outlifeid;*/
                        dtworkRow [34] = strMiddleInitial;

                        if(String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                                {
                                    dtworkRow[36] = objHlpr.fn_getgender(str_gender, dtworkRow[33].ToString());
                        }

                                dtworkRow[39] = objHlpr.fn_getmortality(pref);
                                if (objHlpr.fn_isDMort(dtworkRow[39].ToString()))
                                {
                                    dtworkRow[39] = "STANDARD";
                                    dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR8AN" : dtworkRow[76].ToString() + "|BR8AN";
                                }
                        //if (str_sheet.Contains("NB"))////////////////////////
                        //{
                        //    dtworkRow[22] = "FY";

                        //}
                        //else if (str_sheet.Contains("Ren"))
                        //{
                        //    dtworkRow[23] = "RY";
                        //}

                      

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
                                if (!String.IsNullOrEmpty(strFirstName))
                                {
                                    initialNR = strFirstName.Substring(0, 1);
                                }
                                if (!String.IsNullOrEmpty(strLastName))
                                {
                                    initialNR += strLastName.Substring(0, 1);
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
                        
                                if (dtworkRow[23].ToString() == "PHP")
                                {
                                    dbl_BF_PHP += decimal.Parse(
                                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                                           );
                                    dbl_BH_PHP += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                                        );
                                    dbl_BJ_PHP += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                                        );
                                    dbl_BL_PHP += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                                        );
                                    dbl_BZ_PHP += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                                        );
                                }
                                else if (dtworkRow[23].ToString() == "USD")
                                {
                                    dbl_BF_USD += decimal.Parse(
                                           String.IsNullOrEmpty(dtworkRow[57].ToString()) ? "0" : dtworkRow[57].ToString()
                                           );
                                    dbl_BH_USD += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[59].ToString()) ? "0" : dtworkRow[59].ToString()
                                        );
                                    dbl_BJ_USD += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[61].ToString()) ? "0" : dtworkRow[61].ToString()
                                        );
                                    dbl_BL_USD += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[63].ToString()) ? "0" : dtworkRow[63].ToString()
                                        );
                                    dbl_BZ_USD += decimal.Parse(
                                        String.IsNullOrEmpty(dtworkRow[77].ToString()) ? "0" : dtworkRow[77].ToString()
                                        );
                                }
                           
                                #endregion

                                objdt_template.Rows.Add(dtworkRow);

                        }
                        prawrow++;
                        polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                        fullnames = wsraw.Cells[prawrow, 2].Text.ToString();
                        dob = wsraw.Cells[prawrow, 7].Text.ToString();
                        gender = wsraw.Cells[prawrow, 9].Text.ToString();
                        smoker = wsraw.Cells[prawrow, 10].Text.ToString();
                        bustype = wsraw.Cells[prawrow, 12].Text.ToString();
                        curr = wsraw.Cells[prawrow, 25].Text.ToString();
                        paid = wsraw.Cells[prawrow, 14].Text.ToString();
                        premyr = wsraw.Cells[prawrow, 10].Text.ToString();
                        premyr1 = wsraw.Cells[prawrow, 12].Text.ToString();
                        age1 = wsraw.Cells[prawrow, 14].Text.ToString();
                        code = wsraw.Cells[prawrow, 2].Text.ToString();
                        risk = wsraw.Cells[prawrow, 9].Text.ToString();
                        transa = wsraw.Cells[prawrow, 10].Text.ToString();
                        adjprem = wsraw.Cells[prawrow, 18].Text.ToString();
                        busty = wsraw.Cells[prawrow, 20].Text.ToString();
                        adjcurr = wsraw.Cells[prawrow, 21].Text.ToString();
                        bustype = wsraw.Cells[prawrow, 12].Text.ToString();
                        rating = wsraw.Cells[prawrow, 19].Text.ToString();
                        total = wsraw.Cells[prawrow, 23].Text.ToString();
                        adjamt = wsraw.Cells[prawrow, 31].Text.ToString();
                        pref = wsraw.Cells[prawrow, 15].Text.ToString();
                        fyry = wsraw.Cells[prawrow, 6].Text.ToString();
                    rowcount++;
                    }
                
               
                #endregion

                #region "Compute Hash Total"
                dtworkRow = objdt_template.NewRow();
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Premium PHP:";
                dtworkRow[1] = dbl_BF_PHP + dbl_BH_PHP + dbl_BJ_PHP + dbl_BL_PHP;
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Premium USD:";
                dtworkRow[1] = dbl_BF_USD + dbl_BH_USD + dbl_BJ_USD + dbl_BL_USD;
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Sum at Risk PHP:";
                dtworkRow[1] = dbl_BZ_PHP;
                objdt_template.Rows.Add(dtworkRow);

                dtworkRow = objdt_template.NewRow();
                dtworkRow[0] = "Total Sum at Risk USD:";
                dtworkRow[1] = dbl_BZ_USD;
                objdt_template.Rows.Add(dtworkRow);
                #endregion

                string despath = str_saved + @"\BM041-" + str_sheet + "-" + str_savef + ".xlsx";
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
