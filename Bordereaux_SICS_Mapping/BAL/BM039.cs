using System;
using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM039
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

                string group = wsraw.Cells[prawrow, 1].Text.ToString();
                string poldate = wsraw.Cells[prawrow, 2].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                string gender = wsraw.Cells[prawrow, 6].Text.ToString();
                string age = wsraw.Cells[prawrow, 5].Text.ToString();
                string orig = wsraw.Cells[prawrow, 7].Text.ToString();
                string orig1 = wsraw.Cells[prawrow, 8].Text.ToString();
                string orig2 = wsraw.Cells[prawrow, 9].Text.ToString();
                string ini = wsraw.Cells[prawrow, 10].Text.ToString();
                string ini1 = wsraw.Cells[prawrow, 11].Text.ToString();
                string ini2 = wsraw.Cells[prawrow, 12].Text.ToString();
                string dob = wsraw.Cells[prawrow, 4].Text.ToString();
                string premium = wsraw.Cells[prawrow, 16].Text.ToString();
                string premium1 = wsraw.Cells[prawrow, 17].Text.ToString();
                string premium2 = wsraw.Cells[prawrow, 18].Text.ToString();
                string polnum = string.Empty;
                string year  = wsraw.Cells[3, 2].Text.ToString();
                string currency = string.Empty;
                string year12 = string.Empty;
                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                int storee;
                bool chck;
         
                #region Data Processing


                while (rowcount != erawrow + 2)
                {
                    chck = int.TryParse(fullname, out storee);
                    fullname = objHlpr.fn_stringcleanup(fullname);


                    if (fullname == string.Empty && chck == false)
                    {
                        findboo = false;
                    }
                    else if (fullname != string.Empty && chck == true)
                    {

                        dtworkRow = objdt_template.NewRow();
                        dtworkRow01 = objdt_template.NewRow();
                        dtworkRow02 = objdt_template.NewRow();
                        dtworkRow[5] = "Life Grp";
                        dtworkRow[8] = "SURPLUS";
                        dtworkRow[9] = "PAFM";
                        dtworkRow[10] = "S";
                        dtworkRow[13] = "IND";
                        dtworkRow[14] = "T";
                        dtworkRow[7] = group;
                        dtworkRow[20] = poldate;
                        dtworkRow[19] = poldate;
                        dtworkRow[22] = poldate;
                        dtworkRow[23] = "PHP";
                        dtworkRow[24] = "YLY";
                        dtworkRow[27] = ini;
                        dtworkRow[25] = orig;
                        dtworkRow[29] = "NATREID";
                        dtworkRow[79] = age;
                        dtworkRow[39]= "STANDARD";
                        dtworkRow[36] = gender;
                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            dob = "7/1/1900";
                            dtworkRow[76] = String.IsNullOrEmpty(dtworkRow[76].ToString()) ? "BR4AL" : dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion

                        dtworkRow[37] = dob.ToString();
                        dtworkRow[83] = "R";
                        dtworkRow[82] = group;
                        dtworkRow[31] = fullname;
                        dtworkRow[32] = fullname;
                        dtworkRow[33] = fullname;
                        dob = dob.TrimStart(' ').TrimEnd(' ').Replace("/", String.Empty);
                        dtworkRow[30] = fullname + dob;
                        dtworkRow[0] = "RLife" + dob;

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

                        int year1;
                        string poldate1;
                        int poldate2;

                    
                        year1 = Convert.ToInt32(year);
                        DateTime oDate = Convert.ToDateTime(poldate);
                        poldate1 = oDate.Year.ToString();
                        poldate2 = Convert.ToInt32(poldate1);

                        if (year1 >= poldate2)
                        {
                            dtworkRow[21] = "TRENEW";
                            dtworkRow[58] = "4001";
                            dtworkRow[59] = premium;
                        }

                        else if (year1 < poldate2)
                        {
                            dtworkRow[21] = "TNEWBUS";
                            dtworkRow[56] = "4000";
                            dtworkRow[57] = premium;
                        }

                        dtworkRow01.ItemArray = dtworkRow.ItemArray;
                        dtworkRow01[5] = "AD&D";
                        dtworkRow01[25] = orig1;
                        dtworkRow01[27] = ini1;

                        if (year1 >= poldate2)
                        {
                            dtworkRow01[59] = premium1;
                        }
                        else if (year1 < poldate2)
                        {
                            dtworkRow01[57] = premium1;
                        }
                        dtworkRow02.ItemArray = dtworkRow01.ItemArray;
                        dtworkRow02[5] = "TPD";
                        dtworkRow02[25] = orig1;
                        dtworkRow02[27] = ini1;

                        if (year1 >= poldate2)
                        {
                            dtworkRow02[59] = premium2;
                        }
                        else if (year1 < poldate2)
                        {
                            dtworkRow02[57] = premium2;
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

                    prawrow++;
                    group = wsraw.Cells[prawrow, 1].Text.ToString();
                    poldate = wsraw.Cells[prawrow, 2].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                    gender = wsraw.Cells[prawrow, 6].Text.ToString();
                    age = wsraw.Cells[prawrow, 5].Text.ToString();
                    orig = wsraw.Cells[prawrow, 7].Text.ToString();
                    orig1 = wsraw.Cells[prawrow, 8].Text.ToString();
                    orig2 = wsraw.Cells[prawrow, 9].Text.ToString();
                    ini = wsraw.Cells[prawrow, 10].Text.ToString();
                    ini1 = wsraw.Cells[prawrow, 11].Text.ToString();
                    ini2 = wsraw.Cells[prawrow, 12].Text.ToString();
                    dob = wsraw.Cells[prawrow, 4].Text.ToString();
                    premium = wsraw.Cells[prawrow, 16].Text.ToString();
                    premium1 = wsraw.Cells[prawrow, 17].Text.ToString();
                    premium2 = wsraw.Cells[prawrow, 18].Text.ToString();
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
                string despath = str_saved + @"\BM039-" + str_savef + ".xlsx";
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
