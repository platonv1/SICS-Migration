using System;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Windows.Forms;
namespace Bordereaux_SICS_Mapping.BAL
{
    public static class My_DataTable_Extensions
    {
        public static void ExportToExcel(this System.Data.DataTable DataTable, string ExcelFilePath = null)
        {
            try
            {
                int ColumnsCount;

                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbooks.Add();

                Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];

                for (int i = 0; i < ColumnsCount; i++)
                    Header[i] = DataTable.Columns[i].ColumnName;

                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header;
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                HeaderRange.Font.Bold = true;

                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                int lastUsedRow, lastUsedColumn = 0;

                if ((ExcelFilePath != null && ExcelFilePath != "") && (System.IO.File.Exists(ExcelFilePath)))
                {

                    Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook wbinput = eapp.Workbooks.Open(ExcelFilePath);
                    Microsoft.Office.Interop.Excel.Worksheet wsinput = wbinput.Sheets[1];

                    lastUsedRow = wsinput.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                               Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                    lastUsedColumn = wsinput.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns,
                               Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                    Cells = new object[RowsCount + lastUsedRow, ColumnsCount];

                    for (int x = 2; x <= lastUsedRow; x++)
                        for (int y = 1; y <= ColumnsCount; y++)
                            Cells[x - 2, y - 1] = wsinput.Cells[x, y].Text.ToString();

                    wsinput = null;
                    wbinput.Close();
                    wbinput = null;
                    eapp = null;

                    lastUsedRow = lastUsedRow - 1;
                }
                else { lastUsedRow = 0; }

                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j + lastUsedRow, i] = DataTable.Rows[j][i];

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + lastUsedRow + 1, ColumnsCount])).Value = Cells;

                //Format number to prevent scientific notation
                int loopctr = 2;
                Microsoft.Office.Interop.Excel.Range formatRange;
                for (loopctr = 2; loopctr <= RowsCount + lastUsedRow + 1; loopctr++)
                {
                    formatRange = Worksheet.get_Range("H" + loopctr.ToString());
                    formatRange.NumberFormat = "####################";
                    formatRange.ColumnWidth = 20;

                    formatRange = Worksheet.get_Range("A" + loopctr.ToString());
                    formatRange.NumberFormat = "####################";
                    formatRange.ColumnWidth = 20;
                }

                Helper objHlpr = new Helper();

                Worksheet = objHlpr.fn_extendwidth(Worksheet);
                objHlpr = null;

                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(ExcelFilePath);
                        Excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else
                {
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }
    }

    class Helper : _Global
    {

        _Global _var = new _Global();

        public Microsoft.Office.Interop.Excel.Worksheet fn_extendwidth(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int lastUsedColumn = ws.Cells.Find("*", System.Reflection.Missing.Value,
                              System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                              Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns,
                              Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                              false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            for (int i = 1; i <= lastUsedColumn; i++)
            {
                ws.Columns[i].ColumnWidth = 20;
            }
            return fn_rawcleanup(ws);
        }

        public Microsoft.Office.Interop.Excel.Worksheet fn_rawcleanup(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int lastUsedRow = ws.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                               Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            int lastUsedColumn = ws.Cells.Find("*", System.Reflection.Missing.Value,
                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                       Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns,
                       Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            for (int x = 1; x <= lastUsedRow; x++)
            {
                for (int y = 1; y <= lastUsedColumn; y++)
                {
                    ws.Cells[x, y] = fn_stringcleanup(ws.Cells[x, y].Text.ToString());
                }
            }

            return ws;
        }

        public Microsoft.Office.Interop.Excel._Worksheet fn_extendwidth(Microsoft.Office.Interop.Excel._Worksheet ws)
        {
            int lastUsedColumn = ws.Cells.Find("*", System.Reflection.Missing.Value,
                              System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                              Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns,
                              Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                              false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            for (int i = 1; i <= lastUsedColumn; i++)
            {
                ws.Columns[i].ColumnWidth = 20;
            }
            return ws;
        }
        public DataTable dt_formtemplate(string str_sheet)
        {
            DataTable objDT = new DataTable(str_sheet);

            objDT.Columns.Add("POLICY_NUMBER", typeof(String));
            objDT.Columns.Add("CEDENT_CESSION_NUMBER", typeof(String));
            objDT.Columns.Add("PROPOSAL_NUMBER", typeof(String));
            objDT.Columns.Add("BENEFIT_COVERED", typeof(String));
            objDT.Columns.Add("INSURANCE_PRODUCT", typeof(String));
            objDT.Columns.Add("BRANDED_PRODUCT_CEDENT_CODE", typeof(String));
            objDT.Columns.Add("BRANDED_PRODUCT_SICS_CODE", typeof(String));
            objDT.Columns.Add("GROUP_SCHEME_ID", typeof(String));
            objDT.Columns.Add("REINSURANCE_PRODUCT", typeof(String));
            objDT.Columns.Add("TYPE_OF_BUSINESS", typeof(String));
            objDT.Columns.Add("REINSURANCE_METHODS", typeof(String));
            objDT.Columns.Add("CONTRACTUAL_RELATIONSHIP", typeof(String));
            objDT.Columns.Add("BUSINESS_REGION", typeof(String));
            objDT.Columns.Add("CLASS_OF_BUSINESS", typeof(String));
            objDT.Columns.Add("BUSINESS_TYPE", typeof(String));
            objDT.Columns.Add("MAIN_BENEFIT/RIDER", typeof(String));
            objDT.Columns.Add("RENEWAL_FREQUENCY", typeof(String));
            objDT.Columns.Add("EXPERIENCE_REFUND_INDICATOR", typeof(String));
            objDT.Columns.Add("UNDERWRITING_METHOD", typeof(String));
            objDT.Columns.Add("REINSURANCE_START_DATE", typeof(String));
            objDT.Columns.Add("POLICY_START_DATE", typeof(String));
            objDT.Columns.Add("TRANS_CODE", typeof(String));
            objDT.Columns.Add("TRANS_EFFECTIVE_DATE", typeof(String));
            objDT.Columns.Add("CESSION_CURRENCY", typeof(String));
            objDT.Columns.Add("PREMIUM_FREQUENCY", typeof(String));
            objDT.Columns.Add("ORIGINAL_SUM_ASSURED", typeof(String));
            objDT.Columns.Add("CEDED_SUM_ASSURED", typeof(String));
            objDT.Columns.Add("INITIAL_SUM_AT_RISK", typeof(String));
            objDT.Columns.Add("CEDENT_RETENTION", typeof(String));
            objDT.Columns.Add("LIFE1_ID_TYPE", typeof(String));
            objDT.Columns.Add("LIFE1_ID", typeof(String));
            objDT.Columns.Add("LIFE1_FULL_NAME", typeof(String));
            objDT.Columns.Add("LIFE1_LAST_NAME", typeof(String));
            objDT.Columns.Add("LIFE1_FIRST_NAME", typeof(String));
            objDT.Columns.Add("LIFE1_MIDDLE_NAME", typeof(String));
            objDT.Columns.Add("LIFE1_TITLE", typeof(String));
            objDT.Columns.Add("LIFE1_GENDER", typeof(String));
            objDT.Columns.Add("LIFE1_DATE_OF_BIRTH", typeof(String));
            objDT.Columns.Add("LIFE1_SMOKER_STATUS", typeof(String));
            objDT.Columns.Add("PREFERRED_CLASSIFIC", typeof(String));
            objDT.Columns.Add("RISK_EXPIRY_DATE", typeof(String));
            objDT.Columns.Add("POLICY_YEAR", typeof(String));
            objDT.Columns.Add("BENEFIT_TERM", typeof(String));
            objDT.Columns.Add("BENEFIT_TERM_UNITS", typeof(String));
            objDT.Columns.Add("OCCUPATION_CLASS", typeof(String));
            objDT.Columns.Add("OCCUPATION_CODE", typeof(String));
            objDT.Columns.Add("LOADING_FACTOR", typeof(String));
            objDT.Columns.Add("LOADING_DURATION_TYPE", typeof(String));
            objDT.Columns.Add("LOADING_DURATION", typeof(String));
            objDT.Columns.Add("LOADING_DURATION_UNIT", typeof(String));
            objDT.Columns.Add("LOADING_REASON", typeof(String));
            objDT.Columns.Add("LOADING_AMOUNT", typeof(String));
            objDT.Columns.Add("LOADING_PERCENT", typeof(String));
            objDT.Columns.Add("LOADING_AGE", typeof(String));
            objDT.Columns.Add("LOADING_START_AFTER", typeof(String));
            objDT.Columns.Add("LOADING_START_AFTER_U1", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_1", typeof(String));
            objDT.Columns.Add("First_Year_Premium", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_2", typeof(String));
            objDT.Columns.Add("Renewal_Premiums", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_3", typeof(String));
            objDT.Columns.Add("FY_Refunds_and_Adjustments", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_4", typeof(String));
            objDT.Columns.Add("RY_Refunds_and_Adjustments", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_5", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-5", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_6", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-6", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_7", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-7", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_8", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-8", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_9", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-9", typeof(String));
            objDT.Columns.Add("ENTRY_CODE_10", typeof(String));
            objDT.Columns.Add("PREMIUM_AMOUNT-10", typeof(String));
            objDT.Columns.Add("REMARKS", typeof(String));
            objDT.Columns.Add("SUM_AT_RISK", typeof(String));
            objDT.Columns.Add("LIFE1_ATTAINED_AGE", typeof(String));
            objDT.Columns.Add("LIFE1_ISSUE_AGE", typeof(String));
            objDT.Columns.Add("REINSURANCE_COMMISSION", typeof(String));
            objDT.Columns.Add("  NUMBER_OF_LIVES", typeof(String));
            objDT.Columns.Add("GROUP POLICYHOLDER", typeof(String));
            objDT.Columns.Add("REFUNDING CODE", typeof(String));


            return objDT;
        }
        public void fn_killexcel()
        {
            Process [] excelProcesses = Process.GetProcessesByName("Excel");
            foreach(Process p in excelProcesses)
            {
                if(string.IsNullOrEmpty(p.MainWindowTitle))
                {
                    //p.Dispose();
                    p.Kill();
                }
            }
            //var processes = from p in Process.GetProcessesByName("EXCEL")
            //                select p;

            //foreach(var process in processes)
            //{
            //    if(process.MainWindowTitle == "Microsoft Excel")
            //    process.Kill();
            //}
        }
        public void fn_savefile(DataTable objDT, string str_path)
        {
            objDT.ExportToExcel(str_path);
        }
        public void fn_openfile(string str_path)
        {
            Process.Start(str_path);
        }

        string[] str_suffix = {
            "JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "VI"
        };

        #region NOTES
        //trailing spaces needed for any comparison to avoid searching text mixed with other letters
        #endregion
        string[] str_lnprefix = {
            " DE ", " DEL ", " DELA ", " DELOS ", " DELAS ",
            " LA ", " LAS ", " LOS ",
            " SAN ", " STA ", " STA. ", " STO ", " STO. ", " SANTO ", " SANTA "
        };

        public string fn_getMI(string str_fname)
        {
            string[] arr_fname;
            arr_fname = str_fname.Split(' ');
            for (int i = 0; i <= arr_fname.Length - 1; i++)
            {
                if (arr_fname[i].Length == 1)
                {
                    return arr_fname[i];
                }

            }
            return " ";
        }
        public void fn_getnamesandlifeID(string str_fullname, string dob, out string str_fname, out string str_lname, out string str_lifeID, string str_customBM = "000")
        {
            HelperV21 objHlpr2 = new HelperV21();
            string str_uncleanfull = str_fullname;
            str_fullname = fn_stringcleanup(str_fullname);
            str_fname = string.Empty;
            str_lname = string.Empty;
            str_lifeID = string.Empty;

            objHlpr2.fn_checkFullNameIsDummy(str_fullname, out bool boldummy);

            string [] arr_fullname;
            arr_fullname = str_fullname.Split(' ');
            
            if (boldummy == true)
            {
                str_uncleanfull = str_fullname;
                str_fname = "DummyFirstName";
                str_lname = "DummyLastName";
            }
            else if (arr_fullname.Length == 1 && str_customBM == "000")
            {
                #region NOTES
                //Will set fname to fullname if array is only 1
                //Lastname will be left blank
                #endregion

                str_fname = arr_fullname[0];
            }
            else if (str_fullname.Contains(",") && str_customBM == "000")
            {
                #region NOTES
                //if comma is grater than 1, 3 array item will be moved to firstname
                #endregion

                arr_fullname = str_fullname.Split(',');
                str_lname = arr_fullname[0];
                str_fname = arr_fullname[1];

                for (int i = 2; i <= arr_fullname.Length - 1; i++)
                {
                    str_fname += " " + arr_fullname[i];
                }

                str_lname = fn_stringcleanup(str_lname);
                str_fname = fn_stringcleanup(str_fname);

                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname;
                str_lname = _var.str_final_lname;

                #region NOTES
                //will move leading suffix to end of fname
                #endregion
                arr_fullname = fn_stringcleanup(str_fname).Split(' ');
                int y = fn_getleadsuffix(arr_fullname, out _var.str_leadsuffix);
                str_fname = string.Empty;

                for (int yy = y; yy <= arr_fullname.Length - 1; yy++)
                {
                    str_fname += " " + arr_fullname[yy];
                }

                str_fname += " " + _var.str_leadsuffix;
            }
            else if (str_lnprefix.Any(str_fullname.Contains) && str_customBM == "000")
            {
                #region NOTES
                //Will arrange the lastname according to the str_lnprefix arrangement
                #endregion
                int int_fullnamectr = 0;
                bool boo_cleanup = true;

                while (boo_cleanup)
                {
                    arr_fullname = fn_stringcleanup(str_fullname).Split(' ');
                    bool boo_skip = false;
                    foreach (string i in arr_fullname)
                    {
                        if (i == string.Empty)
                        {
                            int_fullnamectr++;
                            continue;
                        }
                        boo_skip = false;

                        foreach (string ii in str_lnprefix)
                        {
                            if (i.Trim() == ii.Trim())
                            {
                                string str_templname = string.Empty;
                                if (int_fullnamectr != arr_fullname.Length - 1)
                                {
                                    if (((i == "DE" && arr_fullname[int_fullnamectr + 1] == "LA") ||
                                        (i == "DE" && arr_fullname[int_fullnamectr + 1] == "LOS") ||
                                        (i == "DE" && arr_fullname[int_fullnamectr + 1] == "LAS"))
                                        && (int_fullnamectr != arr_fullname.Length - 2))
                                    {
                                        str_templname = " " + arr_fullname[int_fullnamectr] + " " + arr_fullname[int_fullnamectr + 1] + " " + arr_fullname[int_fullnamectr + 2];

                                        arr_fullname[int_fullnamectr] = string.Empty;
                                        arr_fullname[int_fullnamectr + 1] = string.Empty;
                                        arr_fullname[int_fullnamectr + 2] = string.Empty;

                                        boo_skip = true;
                                    }
                                    else
                                    {
                                        str_templname = " " + arr_fullname[int_fullnamectr] + " " + arr_fullname[int_fullnamectr + 1];

                                        arr_fullname[int_fullnamectr] = string.Empty;
                                        arr_fullname[int_fullnamectr + 1] = string.Empty;

                                        boo_skip = true;
                                    }
                                }
                                else
                                {
                                    str_templname = " " + arr_fullname[int_fullnamectr];

                                    arr_fullname[int_fullnamectr] = string.Empty;
                                }
                                str_lname += str_templname;
                            }

                            if (boo_skip)
                            {
                                break;
                            }
                        }
                        int_fullnamectr++;
                    }

                    foreach (string i in arr_fullname)
                    {
                        str_fname += " " + i;
                    }

                    boo_cleanup = false;

                    arr_fullname = str_fname.Split(' ');
                    foreach (string i in arr_fullname)
                    {
                        foreach (string ii in str_lnprefix)
                        {
                            if (i.Trim() == ii.Trim())
                            {
                                boo_cleanup = true;
                            }
                        }
                    }
                }

                #region NOTES
                //will move leading suffix to end of fname
                #endregion
                arr_fullname = fn_stringcleanup(str_fname).Split(' ');
                int y = fn_getleadsuffix(arr_fullname, out _var.str_leadsuffix);
                str_fname = string.Empty;

                for (int yy = y; yy <= arr_fullname.Length - 1; yy++)
                {
                    str_fname += " " + arr_fullname[yy];
                }

                str_fname += " " + _var.str_leadsuffix;

                //Detect DELA CRUZ as middle name
                arr_fullname = str_fullname.Trim().Split(new string[] { str_lname }, StringSplitOptions.None);
                if (arr_fullname.Length > 1 && arr_fullname[1] != string.Empty)
                {
                    str_fname = arr_fullname[0] + " " + str_lname;

                    arr_fullname = fn_stringcleanup(str_fullname.Replace(str_lname, string.Empty).Replace(arr_fullname[0], string.Empty)).Split(' ');

                    str_lname = string.Empty;

                    for (int yyy = 0; yyy <= arr_fullname.Length - 1; yyy++)
                    {
                        bool findboo = false;
                        string append_suffix = string.Empty;

                        foreach (string xx in str_suffix)
                        {
                            if (xx == arr_fullname[yyy].Trim())
                            {
                                findboo = true;
                                append_suffix = xx;
                            }
                        }

                        if (!findboo)
                        {
                            str_lname += " " + arr_fullname[yyy];
                        }
                        else
                        {
                            str_fname += " " + append_suffix;
                        }

                    }
                }
            }
            else if (str_customBM == "025")
            {
                str_lname = arr_fullname[arr_fullname.Length - 1];

                for (int ii = 0; ii <= arr_fullname.Length - 2; ii++)
                {
                    str_fname += " " + arr_fullname[ii];
                }

                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname;
                str_lname = _var.str_final_lname;

                #region NOTES
                //will move leading suffix to end of fname
                #endregion
                arr_fullname = fn_stringcleanup(str_fname).Split(' ');
                int y = fn_getleadsuffix(arr_fullname, out _var.str_leadsuffix);
                str_fname = string.Empty;

                for (int yy = y; yy <= arr_fullname.Length - 1; yy++)
                {
                    str_fname += " " + arr_fullname[yy];
                }

                str_fname += " " + _var.str_leadsuffix;
            }
            else if (str_customBM == "021" || str_customBM == "051-A")
            {
                str_lname = arr_fullname[0];

                for (int ii = 1; ii <= arr_fullname.Length - 1; ii++)
                {
                    str_fname += " " + arr_fullname[ii];
                }

                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname;
                str_lname = _var.str_final_lname;

                #region NOTES
                //will move leading suffix to end of fname
                #endregion
                arr_fullname = fn_stringcleanup(str_fname).Split(' ');
                int y = fn_getleadsuffix(arr_fullname, out _var.str_leadsuffix);
                str_fname = string.Empty;

                for (int yy = y; yy <= arr_fullname.Length - 1; yy++)
                {
                    str_fname += " " + arr_fullname[yy];
                }

                str_fname += " " + _var.str_leadsuffix;
            }
            else if (str_customBM == "019")
            {
                int s_ctr = 0;
                if (str_fullname.ToUpper().Contains("JR") || str_fullname.ToUpper().Contains("SR"))
                {
                    s_ctr = 1;
                }

                if (arr_fullname.Length == 2)
                {
                    str_fname = arr_fullname[0];
                    str_lname = arr_fullname[1];
                }
                else if (arr_fullname.Length == 3)
                {
                    str_fname = arr_fullname[0];
                    str_lname = arr_fullname[2];
                }
                else if (arr_fullname.Length > 3)
                {
                    for (int ii = 0; ii <= arr_fullname.Length - (3 + s_ctr); ii++)
                    {
                        str_fname += " " + arr_fullname[ii];
                    }
                    str_lname = arr_fullname[arr_fullname.Length - 1];
                }

                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname.Trim();
                str_lname = _var.str_final_lname;

                if (str_lname.Trim() == string.Empty)
                {
                    str_lname = arr_fullname[arr_fullname.Length - 2];

                }
            }
            else if (str_customBM == "098")
            {
                arr_fullname = fn_stringcleanup(str_fullname.ToUpper().Replace(".", "")).Split(' ');
                str_lname = arr_fullname[arr_fullname.Length - 1];
                int x = 0;
                if ((str_lname == "JR") || (str_lname == "SR") || (str_lname == "I") || (str_lname == "II") || (str_lname == "III") || (str_lname == "IV") || (str_lname == "V"))
                {
                    str_lname = arr_fullname[arr_fullname.Length - 2];
                    x = 3;

                }
                else
                {
                    x = 2;
                }

                for (int i = 0; i <= arr_fullname.Length - x; i++)
                {
                    str_fname += " " + arr_fullname[i];
                }
                if (x == 3)
                {
                    str_fname = str_fname + " " + arr_fullname[arr_fullname.Length - 1];
                }
            }
            else if (str_uncleanfull.Contains("     "))
            {
                #region NOTES
                //if comma is grater than 1, 3 array item will be moved to firstname
                #endregion

                arr_fullname = str_uncleanfull.Replace("     ", "|").Split('|');
                str_lname = arr_fullname[0];
                str_fname = arr_fullname[1];

                for (int i = 2; i <= arr_fullname.Length - 1; i++)
                {
                    str_fname += " " + arr_fullname[i];
                }

                str_lname = fn_stringcleanup(str_lname);
                str_fname = fn_stringcleanup(str_fname);

                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname;
                str_lname = _var.str_final_lname;
            }
            else
            {
                #region NOTES
                //Cannot identify double names execpt for MA and MA.
                //will move leading suffix to end of fname
                #endregion

                int i = fn_getleadsuffix(arr_fullname, out _var.str_leadsuffix);

                str_fname = arr_fullname[i];

                if (str_fname.Trim() == "MA" || str_fname.Trim() == "MA.")
                {
                    str_fname += " " + arr_fullname[i + 1];
                    i += 2;
                }
                else
                {
                    i += 1;
                }

                str_fname += _var.str_leadsuffix;

                for (int ii = i; ii <= arr_fullname.Length - 1; ii++)
                {
                    str_lname += " " + arr_fullname[ii];
                }


                fn_removesuffix(str_fname, str_lname, out _var.str_final_fname, out _var.str_final_lname);
                str_fname = _var.str_final_fname;
                str_lname = _var.str_final_lname;
            }

            str_lname = fn_stringcleanup(str_lname).Replace(".", "");
            str_fname = fn_stringcleanup(str_fname).Replace(".", "");


            #region NOTES
            //process SICSID
            #endregion
            //DateTime oDate = fn_convertStringtoDateV3(dob);
            DateTime oDate = DateTime.ParseExact(dob, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            string str_formatfname = string.Empty;
            string str_formatlname = string.Empty;

            for (int x = 0; x <= str_fname.Length - 1; x++)
            {
                if (str_fname.Substring(x, 1) == " ")
                {
                    continue;
                }
                str_formatfname += str_fname.Substring(x, 1);
                if (str_formatfname.Length == 2)
                {
                    break;
                }
            }


            if (!str_fullname.Contains(","))
            {
                string str_MI = fn_getMI(str_lname);
                str_lname = str_lname.Replace(str_MI + " ", string.Empty);
            }


            for (int x = 0; x <= str_lname.Length - 1; x++)
            {
                if (str_lname.Substring(x, 1) == " ")
                {
                    continue;
                }
                str_formatlname += str_lname.Substring(x, 1);
                if (str_formatlname.Length == 5)
                {
                    break;
                }
            }

            if (boldummy == false)
            {
                str_lifeID = str_formatlname + str_formatfname + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year;
                str_lifeID = fn_stringcleanup(str_lifeID);
            }
            
        }
        public void fn_removesuffix(string str_fname, string str_lname, out string str_final_fname, out string str_final_lname)
        {
            #region NOTES
            //Will move all suffix to fname
            //Middle initial V will be treated as suffix
            #endregion
            str_final_fname = str_fname;
            str_final_lname = str_lname;

            foreach (string i in str_suffix)
            {
                if (str_lname.Contains(i))
                {
                    string[] arr_tempname = str_lname.Split(' ');

                    foreach (string ii in arr_tempname)
                    {
                        if (i == ii.Trim())
                        {
                            str_final_fname += " " + i;
                            str_final_lname = " " + str_final_lname + " ";
                            str_final_lname = str_final_lname.Replace(" " + ii + " ", " ");
                        }
                    }
                }
            }
        }
        public int fn_getleadsuffix(string [] arr_fullname, out string str_leadsuffix)
        {
            int i = 0;
            int i_check = 0;

            str_leadsuffix = string.Empty;
            foreach(string x in arr_fullname)
            {
                foreach(string xx in str_suffix)
                {
                    if(xx == x.Trim())
                    {
                        i++;
                        str_leadsuffix += " " + xx;
                    }
                }

                if(i == i_check)
                {
                    break;
                }

                i_check += 1;
            }

            return i;
        }


        public string fn_getrefcode(string str_busstype)
        {
            if (str_busstype.ToUpper() == "T")
            {
                return "R";
            }
            else
            {
                return "NR";
            }
        }
        public string fn_stringcleanupRemoveSpace(string str_toclean)
        {
            str_toclean = str_toclean.ToUpper();
            str_toclean = str_toclean.Replace(" ", string.Empty);
            str_toclean = str_toclean.Trim();
            return str_toclean;
        }
        public string fn_stringcleanup(string str_toclean)
        {
            str_toclean = str_toclean.ToUpper();
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("'", " ");
            str_toclean = str_toclean.Trim();
            return str_toclean;
        }
        public string fn_numbercleanup(string str_toclean)
        {
            str_toclean = str_toclean.ToUpper();
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Trim();
            str_toclean = str_toclean.Replace("-", "").Replace("(", String.Empty).Replace(")", String.Empty);
            return str_toclean;
        }
        public string fn_numbercleanup_negative(string str_toclean)
        {
            str_toclean = str_toclean.ToUpper();
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("  ", " ");
            str_toclean = str_toclean.Replace("'", " ");
            str_toclean = str_toclean.Trim();

            if (str_toclean == "-")
            {
                return "0";
            }
            str_toclean = str_toclean.Replace("-", "").Replace("(", "-").Replace(")", String.Empty).Replace(",", String.Empty);

            return str_toclean;
        }

        public string fn_parenthesistoNegative(string str_toclean)
        {
            if (str_toclean.StartsWith("(") && str_toclean.EndsWith(")"))
            {
                return "-" + str_toclean.Replace("(", "").Replace(")", "");
            }
            else
            {
                return str_toclean;
            }
        }
        public string fn_getgender(string str_genderexcel, string str_rawname)
        {
            if (objdt_GenderDB.Rows.Count == 0)
            {
                Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wbdata = eapp.Workbooks.Open(str_genderexcel);
                Microsoft.Office.Interop.Excel.Worksheet wsdata = wbdata.Sheets["SUMMARY"];
                Microsoft.Office.Interop.Excel.Range datarange = wsdata.UsedRange;

                int edatarow = datarange.Rows.Count;

                objdt_GenderDB.Columns.Add("FNAME", typeof(String));
                objdt_GenderDB.Columns.Add("SEX", typeof(String));


                for (int intLoop = 1; intLoop <= edatarow + 1; intLoop++)
                {
                    dtworkRow = objdt_GenderDB.NewRow();
                    dtworkRow[0] = wsdata.Cells[intLoop, 1].Text.ToString().ToUpper();
                    dtworkRow[1] = wsdata.Cells[intLoop, 2].Text.ToString();


                    objdt_GenderDB.Rows.Add(dtworkRow);
                }

                wsdata = null;
                wbdata.Close();
                wbdata = null;

                eapp = null;

            }

            DataRow[] foundRows = objdt_GenderDB.Select("FNAME = '" + str_rawname.ToUpper() + "'");
            if (foundRows.Length != 0)
            { return foundRows[0][1].ToString(); }
            else { return string.Empty; }

        }
        public string fn_getgenderv2(string strFirstname)
        {
            string strSex = string.Empty;
            //bool genderFail = false;
            
            try
            {
                fn_Getfirstname(strFirstname, out string strFirstName);
                strSex = null;
                string query = "SELECT * FROM dbo_gender WHERE firstname=" + "'" + strFirstName + "'";
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                //OdbcConnection cnDB = new OdbcConnection(szConnect);

                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                OdbcDataReader DbReader = DbCommand.ExecuteReader();

                if (DbReader.Read())
                {
                    strSex = DbReader.GetValue(1).ToString();
                    return strSex;

                }
                else
                {
                    strSex = "M";
                    Variables.boogenderfail = true;
                    return strSex;

                }

                DbReader.Close();
                cnDB.Dispose();
                cnDB.Close();
            }
            catch (Exception ex)
            {
                strSex = "M";
                Variables.boogenderfail = true;
                return strSex;
                
                //boo_genderfail = true;
                //if (boo_genderfail = true)
                //{
                    

                //    strSex = "M";
                //    var Result = MessageBox.Show("Please connect to VPN to fetch gender in the database", "Not connected to VPN " + ex.Message, MessageBoxButtons.OK);
                //    if(Result == DialogResult.OK)
                //    {
                //        System.Windows.Forms.Application.Exit();
                //    }
                //}
                //return strSex;
            }
        }

        public string fn_getgenderv3(string gender)
        {
            string strSex = string.Empty;
            //bool genderFail = false;

            try
            {
                if(string.IsNullOrEmpty(gender))
                {
                    return "MALE";
                }
                return gender;
            }
            catch(Exception ex)
            {
                strSex = "M";
                Variables.boogenderfail = true;
                return strSex;

                //boo_genderfail = true;
                //if (boo_genderfail = true)
                //{


                //    strSex = "M";
                //    var Result = MessageBox.Show("Please connect to VPN to fetch gender in the database", "Not connected to VPN " + ex.Message, MessageBoxButtons.OK);
                //    if(Result == DialogResult.OK)
                //    {
                //        System.Windows.Forms.Application.Exit();
                //    }
                //}
                //return strSex;
            }
        }

        public Boolean fn_boolGender(string strSex)
        {
            if (string.IsNullOrEmpty(strSex))
            {
                return Variables.boogenderfail = true;
            }
            else
            {
                return Variables.boogenderfail = false;
            }

        }


        public string fn_gettranseffectivedate(string valueIssueDate, string valueBmYear)
        {
            valueIssueDate = valueIssueDate.Substring(0, 6);
            string strTransEffectiveDate = valueIssueDate + valueBmYear;
            return strTransEffectiveDate;
            //01 / 01 / 2004

        }
        public void fn_macrobenlifebm010(string strValueCN, string strValueURC, string strValuePolno, string strPremDueDate, out string strCessionCode, out string strPolno,
        out string strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial, out string strSex, out string strDOB, out string strLifeID,
        out string strIssueDate, out string strMortality, out string strRefunding, out string strCededRetention,
        out string strOSA, out string strISA, out string strRemarksCode)
        {

            strCessionCode = ""; strPolno = "";
            strFullName = ""; strLastName = ""; strFirstName = ""; strMiddleInitial = ""; strSex = ""; strDOB = ""; strLifeID = "";
            strIssueDate = ""; strMortality = ""; strRefunding = ""; strFullName = "";
            strOSA = ""; strISA = ""; strRemarksCode = ""; strCededRetention = "";



            try
            {
                HelperV21 objHlpr2 = new HelperV21();

                if (!string.IsNullOrEmpty(strValueCN) && !string.IsNullOrEmpty(strValueURC)) //DB connect for BM010
                {

                    string query = "SELECT * FROM dbo_macro WHERE cession_no=" + "'" + strValueCN + "'"
                    + " " + "AND" + " " + "urc_cntl_no=" + "'" + strValueURC + "'" + " AND " + "company_name LIKE " + "'BENEFI%'";

                    //Console.WriteLine(query);
                    string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                    OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();

                    if (DbReader.Read())
                    {
                        if(DbReader.HasRows == true)
                        {
                        strPolno = DbReader.GetValue(3).ToString();
                        strCessionCode = DbReader.GetValue(4).ToString();
                        //strIssueAge = DbReader.GetValue(6).ToString();
                        strIssueDate = DbReader.GetValue(7).ToString();
                        strMortality = DbReader.GetValue(8).ToString();
                        strRefunding = DbReader.GetValue(9).ToString();
                        strFullName = DbReader.GetValue(10).ToString();
                        strDOB = DbReader.GetValue(11).ToString();
                        strSex = DbReader.GetValue(12).ToString();
                        strOSA = DbReader.GetValue(15).ToString();
                        strISA = DbReader.GetValue(16).ToString();
                        strCededRetention = DbReader.GetValue(17).ToString();

                        objHlpr2.fn_separateLastNameFirstName(strFullName, out strFirstName, out strLastName, out strMiddleInitial);
                        strLifeID = fn_LifeID(strFirstName, strLastName, strDOB);

                        DbReader.Close();
                        cnDB.Dispose();
                        cnDB.Close();

                        }
                    }
                    else if (DbReader.HasRows == false)
                    {
                        strPolno = strValuePolno;
                        strRemarksCode = "BR6";
                        strLastName = "DummyLastName";
                        strFirstName = "DummyFirstName";
                        strMiddleInitial = "DummyMiddleName";
                        strFullName = "DummyFullName";
                        strLifeID = strValuePolno;
                        strSex = "M";
                        strIssueDate = strPremDueDate;
                        strDOB = "07/01/1900";
                        strRefunding = "N";
                    }
                }
                else
                {
                    strPolno = strValuePolno;
                    strRemarksCode = "BR6";
                    strLastName = "DummyLastName";
                    strFirstName = "DummyFirstName";
                    strMiddleInitial = "DummyMiddleName";
                    strFullName = "DummyFullName";
                    strLifeID = strValuePolno;
                    strSex = "M";
                    strIssueDate = strPremDueDate;
                    strDOB = "07/01/1900";
                    strRefunding = "N";
                }
                

            }
            catch (Exception e)
            {
                strPolno = strValuePolno;
                strRemarksCode = "BR6";
                strLastName = "DummyLastName";
                strFirstName = "DummyFirstName";
                strMiddleInitial = "DummyMiddleName";
                strFullName = "DummyFullName";
                strLifeID = strValuePolno;
                strSex = "M";
                strIssueDate = strPremDueDate;
                strDOB = "07/01/1900";
                strRefunding = "N";

            }
        }



        public void fn_macrobenlifebm123(string strValueCN, string strValuePolNo, out string strPolNo, out string strIssueAge, out string strIssueDate,
        out string strMortality, out string strRefunding, out string strFullName, out string strFirstName, out string strLastName,
        out string strMI, out string strTitle, out string strDOB, out string strSex, out string strLifeID,
        out string str_LE_OSA, out string str_LE_ISR, out string str_LE_Ret, out string strRcDummyName, out string strCessionCode)
        {

            strPolNo = ""; string strCover7C; strIssueAge = ""; strIssueDate = ""; strMortality = ""; strRefunding = "";
            strFullName = ""; strLastName = ""; strFirstName = ""; strMI = ""; strTitle = ""; strDOB = ""; strSex = "";
            strLifeID = "";
            str_LE_OSA = ""; str_LE_ISR = ""; str_LE_Ret = ""; strRcDummyName = ""; strCessionCode = "";

            try
            {
                HelperV21 objHlpr2 = new HelperV21();
                if (!string.IsNullOrEmpty(strValueCN))
                {

                    string query = "SELECT * FROM dbo_macro WHERE cession_no=" + "'" + strValueCN + "'" + " AND " + "company_name= " + "'BENEFICIAL-PNB LIFE INSURANCE CO'";
                    string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                    OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();
                    DbReader.Read();

                    strPolNo = DbReader.GetValue(3).ToString();
                    strCessionCode = DbReader.GetValue(4).ToString();
                    strIssueAge = DbReader.GetValue(6).ToString();
                    strIssueDate = DbReader.GetValue(7).ToString();
                    strMortality = DbReader.GetValue(8).ToString();
                    strRefunding = DbReader.GetValue(9).ToString();
                    strFullName = DbReader.GetValue(10).ToString();
                    strDOB = DbReader.GetValue(11).ToString();
                    strSex = DbReader.GetValue(12).ToString();
                    strCover7C = DbReader.GetValue(13).ToString();
                    boo_genderfail = false;


                    objHlpr2.fn_separateLastNameFirstName(strFullName, out strFirstName, out strLastName, out strMI);
                    //fn_separatefullnamev4(strFullName, out strFirstName, out strLastName, out strMI);
                    Console.WriteLine(strFirstName + "" + strLastName + "" + strMI + "" + strTitle);
                    strLifeID = fn_LifeID(strFirstName, strLastName, strDOB);



                    if (strCover7C == "1")
                    {
                        str_LE_OSA = DbReader.GetValue(15).ToString();
                        str_LE_ISR = DbReader.GetValue(16).ToString();
                        str_LE_Ret = DbReader.GetValue(17).ToString();
                    }

                }
                else
                {
                    strPolNo = strValuePolNo;
                    strRcDummyName = "BR6";
                    strLastName = "DummyLastName";
                    strFirstName = "DummyFirstName";
                    strMI = "DummyMiddleName";
                    strFullName = strValuePolNo;
                    strLifeID = strValuePolNo;
                    strSex = "M";
                    strDOB = "07/01/1900";
                    strRefunding = "N";
                    str_LE_OSA = "1";
                    str_LE_ISR = "1";
                    str_LE_Ret = "1";
                    boo_genderfail = true;
                }
            }
            catch (Exception e)
            {
                strPolNo = strValuePolNo;
                strRcDummyName = "BR6";
                strLastName = "DummyLastName";
                strFirstName = "DummyFirstName";
                strMI = "DummyMiddleName";
                strFullName = strValuePolNo;
                strLifeID = strValuePolNo;
                strSex = "M";
                strDOB = "07/01/1900";
                strRefunding = "N";
                str_LE_OSA = "1";
                str_LE_ISR = "1";
                str_LE_Ret = "1";
                boo_genderfail = true;
            }

        }

        public string fn_getDOB(string strbirth, out string strDOB)
        {

            if (string.IsNullOrEmpty(strbirth))
            {
                strDOB = "07/01/1900";
                return strDOB;
            }
            else
            {
                strDOB = strbirth;
                return strDOB;
            }
        }




        public void fn_getbusinessTypeRefundingCode(string valueBusinessType, out string strBusinessType, out string strRefundingCode)
        {

            if (valueBusinessType.ToUpper().Contains("Z"))
            {
                strBusinessType = "F";
                strRefundingCode = "N";
            }
            else
            {
                strBusinessType = "T";
                strRefundingCode = "R";
            }
        }

        public void fn_seperateNames(string strFullname, out string strFirstName, out string strLastName)
        {

            var names = strFullname.Split(' ');
            strLastName = names[0];
            strFirstName = names[1];
            string strTitle;
            char strMI;

            if (strFullname.Contains("III"))
            {
                //strTitle = names[2];
                strLastName = names[0] + " " + names[1] + " " + names[2] + " " + names[3];
                strMI = strFullname[strFullname.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(4).ToArray();
                strFirstName = String.Join(" ", names);
            }
        }
        public void fn_savemultiple(DataTable objdt, bool boo_toopen, string str_savedir)
        {
            if (objdt.Rows.Count > 0)
            {
                fn_savefile(objdt, str_savedir);
                if (boo_toopen)
                {
                    fn_openfile(str_savedir);
                }
            }
        }


        public DateTime fn_setdefdob(int int_age = 0)
        {
            string birth = "July" + " " + "1" + " " + (DateTime.Now.Year - int_age).ToString();

            return Convert.ToDateTime(birth);
        }
        public string fn_getMonthNumber(string str_month)
        {

            switch (str_month.ToUpper())
            {
                case "JAN":
                    return "01";
                case "FEB":
                    return "02";
                case "MAR":
                    return "03";
                case "APR":
                    return "04";
                case "MAY":
                    return "05";
                case "JUN":
                    return "06";
                case "JUL":
                    return "07";
                case "AUG":
                    return "08";
                case "SEP":
                    return "09";
                case "OCT":
                    return "10";
                case "NOV":
                    return "11";
                case "DEC":
                    return "12";
                case "JANUARY":
                    return "01";
                case "FEBRUARY":
                    return "02";
                case "MARCH":
                    return "03";
                case "APRIL":
                    return "04";
                //case "MAY":
                //    return "05";
                case "JUNE":
                    return "06";
                case "JULY":
                    return "07";
                case "AUGUST":
                    return "08";
                case "SEPTEMBER":
                    return "09";
                case "OCTOBER":
                    return "10";
                case "NOVEMBER":
                    return "11";
                case "DECEMBER":
                    return "12";
                default:
                    return "0";
            }
        }

        public string fn_getmortality(string mort)
        {
            mort = mort.Trim();
            if (string.IsNullOrEmpty(mort))
            {
                return "STANDARD";
            }

            if (mort.ToUpper().Trim().Replace(" ", string.Empty).In("CLASSA", "CLASSAA", "CLASSB", "CLASSC", "CLASSD", "CLASSE", "CLASSF", "CLASSG"
                                , "CLASSH", "CLASSI", "CLASSJ", "CLASSK", "CLASSL", "CLASSM", "CLASSN", "CLASSO", "CLASSP"
                                , "CLASSI", "CLASSII", "CLASSIII", "CLASSIV", "CLASSV"
                                , "TABLEA", "TABLEAA", "TABLEB", "TABLEC", "TABLED", "TABLEE", "TABLEF", "TABLEG"
                               , "TABLEH", "TABLEI", "TABLEJ", "TABLEK", "TABLEL", "TABLEM", "TABLEN", "TABLEO", "TABLEP"))
            {
                return mort.ToUpper().Trim().Replace(" ", string.Empty);
            }

            if (mort.ToUpper().Trim().Replace(" ", string.Empty).IndexOf("SUBSTANDARD-") > -1)
            {
                return "CLASS" + mort.ToUpper().Trim().Replace(" ", string.Empty).Replace("SUBSTANDARD-", string.Empty);
            }

            switch (mort.ToUpper().Trim())
            {
                case "STANDARD":
                    return "STANDARD";
                case "1.00":
                    return "STANDARD";
                case "1.25":
                    return "CLASSA";
                case "1.375":
                    return "CLASSAA";
                case "1.5":
                    return "CLASSB";
                case "1.50":
                    return "CLASSB";
                case "1.75":
                    return "CLASSC";
                case "2.00":
                    return "CLASSD";
                case "2.25":
                    return "CLASSE";
                case "2.5":
                    return "CLASSF";
                case "2.50":
                    return "CLASSF";
                case "2.75":
                    return "CLASSG";
                case "3.00":
                    return "CLASSH";
                case "3.25":
                    return "CLASSI";
                case "3.5":
                    return "CLASSJ";
                case "3.50":
                    return "CLASSJ";
                case "3.75":
                    return "CLASSK";
                case "4.00":
                    return "CLASSL";
                case "4.25":
                    return "CLASSM";
                case "4.5":
                    return "CLASSN";
                case "4.50":
                    return "CLASSN";
                case "4.75":
                    return "CLASSO";
                case "5.00":
                    return "CLASSP";
                case "1":
                    return "STANDARD";
                case "2":
                    return "CLASSD";
                case "3":
                    return "CLASSH";
                case "4":
                    return "CLASSL";
                case "5":
                    return "CLASSP";

                case "STD":
                    return "STANDARD";
                case "A":
                    return "CLASSA";
                case "AA":
                    return "CLASSAA";
                case "B":
                    return "CLASSB";
                case "C":
                    return "CLASSC";
                case "D":
                    return "CLASSD";
                case "E":
                    return "CLASSE";
                case "F":
                    return "CLASSF";
                case "G":
                    return "CLASSG";
                case "H":
                    return "CLASSH";
                case "I":
                    return "CLASSI";
                case "J":
                    return "CLASSJ";
                case "K":
                    return "CLASSK";
                case "L":
                    return "CLASSL";
                case "M":
                    return "CLASSM";
                case "N":
                    return "CLASSN";
                case "O":
                    return "CLASSO";
                case "P":
                    return "CLASSP";
                case "S":
                    return "STANDARD";
                case "SUB":
                    return "SUBSTANDARD";

                case "II":
                    return "CLASSII";
                case "III":
                    return "CLASSIII";
                case "IV":
                    return "CLASSIV";
                case "V":
                    return "CLASSV";


                case "100":
                    return "STANDARD";
                case "125":
                    return "CLASSA";
                case "150":
                    return "CLASSB";
                case "175":
                    return "CLASSC";
                case "200":
                    return "CLASSD";
                case "225":
                    return "CLASSE";
                case "250":
                    return "CLASSF";
                case "275":
                    return "CLASSG";
                case "300":
                    return "CLASSH";
                case "325":
                    return "CLASSI";
                case "350":
                    return "CLASSJ";
                case "375":
                    return "CLASSK";
                case "400":
                    return "CLASSL";
                case "425":
                    return "CLASSM";
                case "450":
                    return "CLASSN";
                case "475":
                    return "CLASSO";
                case "500":
                    return "CLASSP";

                default:
                    return "STANDARD";
            }
        }

        public bool fn_isDMort(string mort)
        {

            if (mort.IndexOf("D-") >= 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string fn_dashtozero(string str_var)
        {

            if (str_var.Replace(" ", string.Empty) == "-")
            {
                return "0";
            }
            else
            {
                return str_var;
            }
        }
     
        public bool fn_policyNumChecker(string polnum, string Column2, string Column3, string Column4)
        {
            Console.WriteLine(polnum);
            if(polnum.Equals(string.Empty))
            {
                return false;
            }

            Boolean boo = false;
            Regex rx = new Regex(@"/\d|\d+[a-zA-Z0-9-_]|[a-zA-Z0-9-_]+\d+$|^UB\d{7,7}|\d{13,13}[/]\d{10,10}|\d{15,15}[/]\d{12,12}|\d{11,11}\W\d{2,2}\D|\d{3,3}\W\d{7,7}\D|\d{6,6}|\d{8,8}|\d{5,5}|^D\d{9,9}"); //put additional 

            if (rx.IsMatch(polnum) || rx.IsMatch(Column2))
            {
                boo = true;
            }
            else
            {
                boo = false;
            }

            if ((!Column2.Equals(string.Empty) || !Column3.Equals(string.Empty) || !Column4.Equals(string.Empty)) && boo && !polnum.Contains(" "))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool fn_isinQTR(int qtr, int month)
        {
            if ((qtr == 1 && month == 1) || (qtr == 1 && month == 2) || (qtr == 1 && month == 3))
            {
                return true;
            }
            if ((qtr == 2 && month == 4) || (qtr == 2 && month == 5) || (qtr == 2 && month == 6))
            {
                return true;
            }
            if ((qtr == 3 && month == 7) || (qtr == 3 && month == 8) || (qtr == 3 && month == 9))
            {
                return true;
            }
            if ((qtr == 4 && month == 10) || (qtr == 4 && month == 11) || (qtr == 4 && month == 12))
            {
                return true;
            }
            return false;
        }

        public int MonthDiff(DateTime d1, DateTime d2)
        {
            int m1;
            int m2;
            if (d1 < d2)
            {
                m1 = (d2.Month - d1.Month);//for years
                m2 = (d2.Year - d1.Year) * 12; //for months
            }
            else
            {
                m1 = (d1.Month - d2.Month);//for years
                m2 = (d1.Year - d2.Year) * 12; //for months
            }

            return m1 + m2;
        }

        public string fn_convertStringtoDateV2(string strDate)
        {

            try
            {
                string year = strDate.Substring(6, 4);
                string month = strDate.Substring(0, 2);
                string day = strDate.Substring(3, 2);
                string date = month + "/" + day + "/" + year;
                //DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy",CultureInfo.InvariantCulture);
                return date;
            }
            catch (Exception ex)
            {

                var date = DateTime.Now;
                return date.ToString("MM/dd/yyyy");

            }
        }

        public DateTime fn_convertStringtoDateV3(string strDate)
        {

            try
            {
                string month = strDate.Substring(0, 2);
                string day = strDate.Substring(3, 2);
                string year = strDate.Substring(6, 4);
                DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                result = result.Date;
                return result;
            }
            catch (Exception ex)
            {

                return DateTime.Now;

            }
        }

        public DateTime fn_convertStringtoDateV4(string strDate)
        {

            try
            {
                string month = strDate.Substring(5, 2);
                string day = strDate.Substring(7, 2);
                string year = strDate.Substring(0, 4);
                DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                result = result.Date;
                return result;
            }
            catch (Exception ex)
            {

                return DateTime.Now;

            }
        }

        public string fn_convertStringtoDateV5(string strDate)
        {

            try
            {
                strDate = Convert.ToDateTime(strDate).ToString("MM/dd/yyyy");
                return strDate;
            }
            catch(Exception ex)
            {
                strDate = DateTime.Now.ToString("MM/dd/yyyy");
                return strDate;
              

            }
        }



        public string fn_bpSicsCode(string valueSicsCode)
        {
            if (valueSicsCode == "LIFE")
            {
                return "BP289";
            }
            else if (valueSicsCode == "EXTRA")
            {
                return "BP293";
            }
            else if (valueSicsCode == "ADB")
            {
                return "BP291";
            }
            else if (valueSicsCode == "WPD")
            {
                return "BP292";
            }
            else
            {
                return "BP290";
            }
        }

        public void fn_Getfirstname(string strValueFN, out string strGetFirstName)
        {
            strGetFirstName = "";
            strValueFN = strValueFN.Trim().ToUpper();
            string [] strBlankSpace = { " " };
            string[] result = strValueFN.Split(strBlankSpace, StringSplitOptions.None);

            for (int i = 0; i < result.Length; i++)
            {
                if (i == 0)
                {
                    strGetFirstName = result[0];
                    break;
                }
            }
        }


        public DateTime fn_reformatDate(string strDate)
        {
            try
            {

                string month = strDate.Substring(0, 2);
                string day = strDate.Substring(3, 2);
                string year = strDate.Substring(6, 4);
                DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                result = result.Date;
                return result;
            }
            catch (Exception ex)
            {

                return DateTime.Now;

            }
        }

        public string fn_getDOB(string strDOB)
        {
            if (string.IsNullOrEmpty(strDOB))
            {
                return "07/01/1900";
            }
            else
            {
                string month = strDOB.Substring(0, 2);
                string day = strDOB.Substring(3, 2);
                string year = strDOB.Substring(6, 4);
                strDOB = month + "/" + day + "/" + year;
                return strDOB;
            }
        }

        public Double fn_getAttainAge(string valueBY, string valueDOB)
        {
            try
            {
                string year = valueDOB.Substring(6, 4);
                double dblAttainAge = Convert.ToDouble(valueBY) - Convert.ToDouble(year);
                return dblAttainAge;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public Double fn_getIssueAge(string valueDOB, string valuePSD)
        {
            try
            {
                string dobYear = valueDOB.Substring(6, 4);
                string psdYear = valuePSD.Substring(6, 4);
                double dblAttainAge = Convert.ToDouble(psdYear) - Convert.ToDouble(dobYear);
                return dblAttainAge;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        public void fn_isNumeric(string valuePrem,out  decimal dclPremium)
        {
            if (decimal.TryParse(valuePrem, out dclPremium))
            {
                dclPremium = Convert.ToDecimal(valuePrem);
            }
            else if (string.IsNullOrEmpty(valuePrem))
            {
                dclPremium = 0;
            }
            else
            {
                dclPremium = 0;
            }
        }



        //public string RemarksCode(string strRemarkscode)
        //{

        //}
        public string RemoveExtraZeros(string str, List<char> charsToRemove)
        {
            // return String.Join(String.Empty, str.Split(charsToRemove.ToArray()));

            return String.Concat(str.Split(charsToRemove.ToArray()));
        }

        public void fn_getBirthday(string Birthday, out string strBirthday)
        {
            if (!string.IsNullOrEmpty(Birthday))
            {
                strBirthday = Birthday;
            }
            else
            {
                strBirthday = "01/07/1900";
            }
        }

        public string fn_SmokerCode(string smoke)
        {

            try
            {
                if (string.IsNullOrEmpty(smoke))
                {
                    return "NONE";
                }
                else if (smoke.ToUpper() == "N" || (smoke.ToUpper() == "NS"))
                {
                    return "NSMOK";
                }
                else if (smoke.ToUpper() == "Blended")
                {
                    return "BLENDED";
                }
                else if (smoke.ToUpper() == "smoker")
                {
                    return "SMOK";
                }
                else
                {
                    return "SMOK";
                }
            }
            catch
            {
                return "NONE";
            }

        }

        public string fn_SmokerCodeV2(string smoke)
        {

            try
            {
                if(string.IsNullOrEmpty(smoke) || smoke.ToUpper() == "N")
                {
                    return "NSMOK";
                }
                else if(smoke.ToUpper() == "Y")
                {
                    return "SMOK";
                }
                else
                {
                    return "NSMOK";
                }
            }
            catch
            {
                return "NSMOK";
            }

        }

        public string fn_smokercode(string isSmoker, string str_bm)
        {
            if (str_bm == "019" || str_bm == "043")
            {
                if (isSmoker == "Non-Smoker")
                {
                    return "NSMOK";
                }
                else
                {
                    return "SMOK";
                }
            }
            else if (str_bm == "053")
            {
                if (isSmoker == "S")
                {
                    return "SMOK";
                }
                else
                {
                    return "NSMOK";
                }
            }
            else
            {
                return "";
            }
        }

        public void fn_getbusinesstype(string cedantValue, out string strCessionNo)
        {
            if (cedantValue.ToUpper().Trim().Contains("NR"))
            {
                strCessionNo = "F";
            }
            else
            {
                strCessionNo = "T";
            }
        }

        public string fn_checkBusinessTypeV1(string valueCessionCode)
        {
            if (string.IsNullOrEmpty(valueCessionCode))
            {
                return "T";
            }
            else
            {
                return valueCessionCode;
            }
        }



        public string fn_checkBusinessTypeV2(string businessType)
        {
            businessType.Trim().ToUpper();
            if ( businessType == "AUTOMATIC" || string.IsNullOrEmpty(businessType))
            {
                return "T";
            }
            else
            {
                return "F";
            }
        }

        public void fn_GetRemarksCodeBenlife(string valueBirthday, string valueName, string valuePolicyNo, string valueNoGender, string valueMortality,
            double dblvalueISR, double dblvalueNAAR, double dblvalueOSA, out string strRemarksCode)

        {
            string strRemarksDOB = "";
            string strRemarksPolicyno = "";
            string strRemarksNoGender = "";
            string strRemarksValue1 = "";
            string strRemarksValue3 = "";
            string strRemarksValue4 = "";
            string strRemarksValue5 = "";
            string strRemarksValue6 = "";
            string strMortality = "";


            //strRemarksSAR = ""; strRemarksOSR1 = "";


            if (valueBirthday == "07/01/1900")
            {
                strRemarksDOB = "BR4";
            }
            else
            {
                strRemarksDOB = "";
            }


            if (valueMortality == "STANDARD")
            {
                strMortality = "BR8AN";
            }
            else
            {
                strMortality = "";
            }

            if (valueName == valuePolicyNo)
            {
                strRemarksPolicyno = "BR6";

            }
            else
            {
                strRemarksPolicyno = "";
            }
            if (string.IsNullOrEmpty(valueNoGender))
            {
                strRemarksNoGender = "BR61";

            }
            else
            {
                strRemarksNoGender = "";
            }


            if (dblvalueISR != 1 && dblvalueNAAR == 1)
            {
                strRemarksValue1 = "BR1-1BZ";
            }

            else if (dblvalueOSA != 1 && dblvalueNAAR == 1)
            {
                strRemarksValue1 = "BR1-1BZ";
            }
            else if (dblvalueNAAR != 1 && dblvalueISR == 1)
            {
                strRemarksValue3 = "BR2-1AB";

            }
            else if (dblvalueISR != 1 && (dblvalueISR == 1))
            {
                strRemarksValue4 = "BR2-2AB";
            }
            else if (dblvalueISR != 1 && dblvalueOSA == 1)
            {
                strRemarksValue5 = "BR3-1Z";
            }
            else if (dblvalueNAAR != 1 && dblvalueOSA == 1)
            {
                strRemarksValue6 = "BR3-2Z";
            }

            strRemarksCode = strRemarksDOB + "|" + strRemarksPolicyno + "|" + strRemarksNoGender + "|" + strMortality + "|" + strRemarksValue1 + "|" + strRemarksValue3 + "|" + strRemarksValue4
                + "|" + strRemarksValue5 + "|" + strRemarksValue6;
            Console.WriteLine(strRemarksCode);

        }


        public void fn_GetRemarksCodeBenlife123(string valueBirthday, string valueName, string valuePolicyNo, string valueNoGender, string valueMortality,
     double dblvalueISR, double dblvalueNAAR, double dblvalueOSA, double dbl_ValuePremiumAmnt, out string strRemarksCode)

        {
            string strRemarksDOB = "";
            string strRemarksPolicyno = "";
            string strRemarksNoGender = "";
            string strRemarksValue1 = "";
            string strRemarksValue3 = "";
            string strRemarksValue4 = "";
            string strRemarksValue5 = "";
            string strRemarksValue6 = "";
            string strMortality = "";
            strRemarksCode = "";


            //strRemarksSAR = ""; strRemarksOSR1 = "";


            if (valueBirthday == "07/01/1900")
            {
                strRemarksDOB = "BR4";
            }
            else
            {
                strRemarksDOB = "";
            }


            if (valueMortality == "STANDARD")
            {
                strMortality = "BR8AN";
            }
            else
            {
                strMortality = "";
            }

            if (valueName == valuePolicyNo)
            {
                strRemarksPolicyno = "BR6";

            }
            else
            {
                strRemarksPolicyno = "";
            }
            if (string.IsNullOrEmpty(valueNoGender))
            {
                strRemarksNoGender = "BR61";

            }
            else
            {
                strRemarksNoGender = "";
            }


            if (dblvalueNAAR == dblvalueISR)
            {
                strRemarksValue1 = "BR1-1BZ";

            }
            if (dbl_ValuePremiumAmnt == dblvalueOSA)
            {
                strRemarksValue1 = "BR1-2BZ";
            }

            else if (dblvalueISR == dblvalueNAAR)
            {
                strRemarksValue3 = "BR2-1AB";

            }

            else if (dblvalueISR == dblvalueOSA)
            {
                strRemarksValue5 = "BR2-2AB";
            }
            else if (dblvalueISR == dblvalueOSA)
            {
                strRemarksValue5 = "BR2-2AB";
            }
            else if (dblvalueOSA == dblvalueISR)
            {
                strRemarksValue5 = "BR3-1Z";
            }
            else if (dblvalueOSA == dblvalueNAAR)
            {
                strRemarksValue5 = "BR3-2Z";
            }

            strRemarksCode = strRemarksDOB + "|" + strRemarksPolicyno + "|" + strRemarksNoGender + "|" + strMortality + "|" + strRemarksValue1 + "|" + strRemarksValue3 + "|" + strRemarksValue4
            + "|" + strRemarksValue5 + "|" + strRemarksValue6;
            Console.WriteLine(strRemarksCode);
        }

        public void fn_GetRemarksCode(string valueBirthday, string valueName, string valueNoGender, out string strRemarksCode)

        {
            string strRemarksDOB = "";
            string strRemarksDummyName = "";
            string strRemarksNoGender = "";

            //strRemarksSAR = ""; strRemarksOSR1 = "";


            if (valueBirthday == "07/01/1900")
            {
                strRemarksDOB = "BR4";
            }


            //if (valueName == "DummyFullName")
            //{
            //    strRemarksDummyName = "BR6";

            //}

            if (string.IsNullOrEmpty(valueNoGender))
            {
                strRemarksNoGender = "BR7";

            }

            strRemarksCode = strRemarksDOB + "|" + strRemarksDummyName + "|" + strRemarksNoGender;

        }




        public void fn_osabreakdown(Double value1SAR, Double value2RIS, Double value3OSA, out string strOrignalSum)
        {


            //Convert.ToDecimal(value1SAR);//COL AG
            //Convert.ToDecimal(value2RIS); //COL AF
            //Convert.ToDecimal(value3OSA); //COL X
            double x;
            //Convert.ToDecimal(strOrignalSum);
            //=(AG/(AF+AG))*X

            //Convert.ToInt32(value2RIS);
            x = value2RIS + value1SAR; //AF + AG
            if (x == 0)
            {
                x = 1;
            }
            x = value1SAR / x;
            x = x * value3OSA;
            if (x == 0)
            {
                x = 1;
            }
            strOrignalSum = Convert.ToString(x);


        }

        public void fn_checksheetname(string str_sheet, out int NARColNo)
        {

            if (str_sheet == "Peso Renewals" || (str_sheet == "Renewals-Peso") || (str_sheet == "Dollar Renewals"))
            {

                NARColNo = 12;
            }
            else if (str_sheet == "Renewals-Dollar")
            {
                NARColNo = 11; //CURRENCY
            }
            else
            {
                NARColNo = 21;
            }
        }


        public void fn_checksheetnameV2(string strSheet, string valueFullName, out string strFullName, out string strLastName, out string strFirstName, out string strMiddleInitial)
        {
            strSheet.TrimEnd();
            strFirstName = ""; strLastName = ""; strMiddleInitial = ""; strFullName = "";
            HelperV21 objHlpr2 = new HelperV21();
            if (strSheet == "CWSLAI" || strSheet == "AGRICOM TPC" || strSheet == "ARGEM LOANS TPC" || strSheet == "AVIDA & SUBSIDIARIES" || strSheet == "BDO PELAC" || strSheet == "BDO SM EMPLOYEES" || strSheet == "BDO COMMBANK"
            || strSheet == "BDO REMEDIAL" || strSheet == "CITYSTATE SAVINGS BANK" || strSheet == "CRC REALTY")
            {
                objHlpr2.fn_separateLastNameFirstNameV2(valueFullName, out strFullName, out strLastName, out strFirstName, out strMiddleInitial); //Name has no MiddileInitial
            }

            else
            {
                objHlpr2.fn_separateLastNameFirstNameV3(strFullName, out strLastName, out strFirstName, out strMiddleInitial);
            }
        }


        public decimal fn_getcedentretention(decimal valueSumRetro, decimal valueYourShare)
        {
            decimal dclCedentRentention = valueSumRetro - valueYourShare;
            if (dclCedentRentention == 0 || dclCedentRentention < 0)
            {
                return 1;
            }
            else
            {
                return dclCedentRentention;
            }



        }
        public void fn_getTotalPremiumV1(decimal valueT, decimal valueV)
        {

            Variables.TotalPremium += valueT + valueV;
        }

        public void fn_getTotalPremiumV2(string valueCurrency, string valueCessionCode, decimal valueLife, decimal valueADB, decimal valueWPD, decimal valuePDD)
        {

            if (valueCessionCode == "F" && valueCurrency == "PHP" || valueCessionCode == "F" && valueCurrency == "USD")
            {
                Variables.TotalFaculPremium += valueLife + valueADB + valueWPD + valuePDD;
            }

            else if (valueCessionCode == "T" && valueCurrency == "PHP" || valueCessionCode == "T" && valueCurrency == "USD")
            {
                Variables.TotalTreatyPremium += valueLife + valueADB + valueWPD + valuePDD;
            }

        }
        public void fn_getTotalPremiumV3(decimal valueLIFE, decimal valueEXTRA, decimal valueADB, decimal valueSAR, decimal valueSARDI, decimal valueSumAtRisk, out decimal TotalPremium, out decimal TotalSAR)
        {

            _var.dbl_BF += valueLIFE;
            _var.dbl_BH += valueEXTRA;
            _var.dbl_BJ += valueSAR;
            _var.dbl_BL += valueSARDI;
            _var.dbl_FBH += valueADB;
            _var.dbl_BZ += valueSumAtRisk;


            TotalPremium = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL + _var.dbl_FBH;
            TotalSAR = _var.dbl_BZ;

        }
       
        //public void fn_getTotalPremiumV5(decimal valuePremium)
        //{
        //    Variables.TotalPremium += valuePremium;
        //}



        public void fn_getTotalPremiumV6(decimal valuePremAdd, decimal dblPrem_AddBasic, decimal valuePremAddRider, decimal valuePremLife, decimal valuePremTpdi, decimal valuePremUma)
        {

            if (valuePremAdd != 0 || valuePremAdd != 1 || valuePremAddRider != 0 || valuePremAddRider != 1 || valuePremLife != 0 || valuePremLife != 1 || valuePremTpdi != 0 || valuePremTpdi != 1 || valuePremUma != 0 || valuePremUma != 1)
            {
                Variables.TotalPremium += valuePremAdd + valuePremAddRider + valuePremLife + valuePremTpdi + valuePremUma;

            }
        }

     

        public void fn_getTotalSumAtRiskV2(decimal valueNarAdd, decimal dblNarBasic, decimal valueNarRider, decimal valueNarLife, decimal valueNarTpdi, decimal valueNarUma)
        {

            if (valueNarAdd != 0 || valueNarAdd != 1 || dblNarBasic != 0 || dblNarBasic != 1 || valueNarRider != 0 || valueNarRider != 1 || valueNarLife != 0 || valueNarLife != 1 || valueNarTpdi != 0 || valueNarTpdi != 1 || valueNarUma != 0 || valueNarUma != 1)
            {
                Variables.TotalSumAtRisk += valueNarAdd + dblNarBasic + valueNarRider + valueNarLife + valueNarTpdi + valueNarUma;

            }
        }

        public void fn_getTotalSumAtRiskV3(string valueCessionCode, string valueCurrency, decimal valueSAR)
        {
            if (valueCessionCode == "F" && valueCurrency == "PHP" || valueCessionCode == "F" && valueCurrency == "USD")
            {
                Variables.TotalFaculSAR += valueSAR;
            }
            else if (valueCessionCode == "T" && valueCurrency == "PHP" || valueCessionCode == "T" && valueCurrency == "USD")
            {
                Variables.TotalTreatySAR += valueSAR;
            }

        }
  




        public void fn_getTotalCommission(decimal valueX, decimal valueZ, decimal valueAB)
        {
            Variables.TotalCommission += valueX + valueZ + valueAB;

        }


        public void fn_getBusinessType(string valueBusinessType, out string valueCessionCode)
        {
            if (valueBusinessType == "T")
            {
                valueCessionCode = "T";
            }
            else if (string.IsNullOrEmpty(valueBusinessType))
            {
                valueCessionCode = "T";
            }
            else
            {
                valueCessionCode = "F";
            }

        }
        
      
        public string fn_gettranscode(string valueColumnL, string valueSheetName)
        {
            string strTranscode = valueColumnL.Substring(0, 1);

            if (valueSheetName.ToUpper().Contains("R&A_RECAP"))
            {
                if (strTranscode.ToUpper() == "S")
                {

                    return "TFULLSUR";
                }
                else if (strTranscode.ToUpper() == "M")
                {
                    return "TFULLMAT";
                }
                else if (strTranscode.ToUpper() == "D")
                {
                    return "TREINS";
                }
                else if (valueColumnL.ToUpper().Contains("EXPIRED") || strTranscode.ToUpper() == "E")
                {
                    return "TEXPIRY";
                }
                else if (valueColumnL.ToUpper().Contains("PCV") || valueColumnL.ToUpper().Contains("FPU"))
                {
                    return "ADJUST";
                }
                else
                {
                    return "TFULLREC";
                }
                    

            }
            else
            {
                return "TRENEW";
            }

        }

        public string fn_gettranscodev2(int valueOYear, int valueEY, out bool bolEntry)
        {
            if(valueOYear > valueEY)
            {
                bolEntry = true;
                return "TRENEW";
            }
            else if(valueOYear == valueEY)
            {
                bolEntry = false;
                return "TNEWBUS";
            }
            else
            {
                bolEntry = true;
                return "TRENEW";
            }
        }


        

        public string fn_computecededretention(string ValueOSA, string ValueCSA)
        {
            string strCR = "";

            decimal dcmCR = Convert.ToDecimal(ValueOSA) - Convert.ToDecimal(ValueCSA);
            strCR = Convert.ToString(dcmCR);
            return strCR;

        }
        #region Aljohn
        public DateTime fn_reformatDatev2( string strDate )
        {
            string day = strDate.Substring(0, 2);
            string strmonth = strDate.Substring(3, 3);
            string month = fn_getMonthNumber(strmonth);
            string year = strDate.Substring(7, 4);
            DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            result = result.Date;
            return result;
        }


        //LASTNAME SUFFIX FIRSTNAME MIDDLEINITIAL
        public void fn_separatefullname( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            if ( names.Length <= 2 )
            {
                strLastName = names [0];
                strMiddleInitial = "";
                strFirstName = names [1];
            }
            else if ( strFullname.ToUpper().Contains("JR ") || strFullname.ToUpper().Contains("SR ") || strFullname.ToUpper().Contains("III ") )
            {
                if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
                {
                    if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                    {
                        strLastName = names [0] + " " + names [1] + " " + names [2];
                        strMiddleInitial = names[names.Length - 1];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(3).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names [0] + " " + names [1];
                        strMiddleInitial = names[names.Length - 1];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(2).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    strLastName = names [0] + " " + names [1];
                    strMiddleInitial = names[names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(2).ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                {
                    strLastName = names [0] + " " + names [1];
                    strMiddleInitial = names[names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(2).ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [0];
                    strMiddleInitial = names[names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(1).ToArray();
                    strFirstName = String.Join(" ", names);
                }

            }
            else if ( strFullname.ToUpper().Contains("DE LOS ") || strFullname.ToUpper().Contains("DE LA ") )
            {
                if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
                {
                    strLastName = names [0] + " " + names [1] + " " + names [2] + " " + names [3];
                    strMiddleInitial = names[names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(4).ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [0] + " " + names [1] + " " + names [2];
                    strMiddleInitial = names[names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(3).ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else
            {
                strLastName = names [0];
                strMiddleInitial = names[names.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(1).ToArray();
                strFirstName = String.Join(" ", names);
            }
        }
        //FIRSTNAME MIDDLENAME LASTNAME SUFFIX
        public void fn_separatefullnamev2( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');

            if ( names.Length <= 2 )
            {
                strLastName = names [1];
                strMiddleInitial = "";
                strFirstName = names [0];
            }
            else if (names.Contains("JR") || names.Contains("SR") || names.Contains("III") || names.Contains("II"))
            {
                if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
                {
                    if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                    {
                        strLastName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 4];
                        names = names.Reverse().Skip(4).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 3];
                        names = names.Reverse().Skip(3).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    if (names.Length <= 3)
                    {
                        strLastName = names[names.Length - 2] + " " + names[names.Length - 1];
                        strMiddleInitial = "";
                        names = names.Reverse().Skip(2).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names[names.Length - 2] + " " + names[names.Length - 1];
                        strMiddleInitial = names[names.Length - 3];
                        names = names.Reverse().Skip(3).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                {
                    if (names.Length <= 3)
                    {
                        strLastName = names[names.Length - 2] + " " + names[names.Length - 1];
                        strMiddleInitial = "";
                        names = names.Reverse().Skip(2).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names[names.Length - 2] + " " + names[names.Length - 1];
                        strMiddleInitial = names[names.Length - 3];
                        names = names.Reverse().Skip(3).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    strLastName = names.Last();
                    strMiddleInitial = strMiddleInitial = names[names.Length - 2];
                    names = names.Reverse().Skip(2).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE LOS ") || strFullname.ToUpper().Contains("DE LA ") )
            {
                if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
                {
                    strLastName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = names [names.Length - 5];
                    names = names.Reverse().Skip(5).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = names [names.Length - 4];
                    names = names.Reverse().Skip(4).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else
            {
                strLastName = names.Last();
                strMiddleInitial = strMiddleInitial = names [names.Length - 2];
                names = names.Reverse().Skip(2).Reverse().ToArray();
                strFirstName = String.Join(" ", names);
            }


        }
        //LASTNAME, FIRSTNAME MIDDLENAME
        public void fn_separatefullnamev3( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(',');
            strFirstName = "";
            strMiddleInitial = "";
            strLastName = names [0];
            string out_suffix = String.Empty;
            string [] Suffixes = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            string [] MISuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA" };
            if(names.Contains("JR") || names.Contains("SR") || names.Contains("II") || names.Contains("III"))
            {
                var fname = names[0].Split(' ');
                strFirstName = fname.Last();
                fname = fname.Take(fname.Count() - 1).ToArray();
                strLastName = String.Join(" ", fname);
            }
            names = names [1].TrimStart().Split(' ');
            if ( names.Length < 2 )
            {
                strMiddleInitial = "";
                strFirstName = names[0] + strFirstName;
            }
            else
            { // ACLA, CECILIO JR. ATUEL
               //Dueñas, Zaldy jr. Del Prado


                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || strFullname.ToUpper().Contains("VII"))
                {
                    foreach(var suffix in Suffixes)
                    {
                        foreach(var name in names)
                        {
                            if(name == suffix)
                            {
                                out_suffix = name;
                                strFirstName = names [0] + " " + out_suffix;
                                strMiddleInitial = names.Last();
                                break;
                            }
                            else; continue;
                        }

                    }
                }
                else if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL"))
                {
                    if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || strFullname.ToUpper().Contains("VII"))
                    {

                        foreach(var MI in MISuffix)
                        {
                            foreach(var mi_ in names)
                            {
                                if(mi_ == MI)
                                {
                                    strFirstName = names [0] + " " + names[1];
                                    names = names.Skip(2).ToArray();
                                    strMiddleInitial = names.ToString();
                                    break;
                                }
                                
                            }
                        }

                    }
                    
                }
                else
                {
                    strMiddleInitial = names.Last();
                    names = names.Take(names.Count() - 1).ToArray();
                    strFirstName = String.Join(" ", names);
                }
               
            }

        }
        // LASTNAME FIRSTNAME MIDDLEININTIAL SUFFIX
        public void fn_separatefullnamev4( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            if ( names.Length <= 2 )
            {
                strLastName = names [0];
                strMiddleInitial = "";
                strFirstName = names [1];
            }
            else if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
            {
                if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
                {
                    if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                    {
                        strLastName = names [0] + " " + names [1] + " " + names.AsQueryable().Last();
                        strMiddleInitial = names [names.Length - 2];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(3).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names [0] + " " + names.AsQueryable().Last();
                        strMiddleInitial = names [names.Length - 2];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(1).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    strLastName = names [0] + " " + names.AsQueryable().Last();
                    strMiddleInitial = names [names.Length - 2];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(1).ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                {
                    strLastName = names [0] + " " + names [1];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(2).ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [0];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(1).ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE LOS ") || strFullname.ToUpper().Contains("DE LA ") )
            {
                if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
                {
                    strLastName = names [0] + " " + names [1] + " " + names [2] + " " + names.AsQueryable().Last();
                    strMiddleInitial = names [names.Length - 2];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(3).ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [0] + " " + names [1] + " " + names [2];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(3).ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else
            {
                strLastName = names [0];
                strMiddleInitial = names [names.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(1).ToArray();
                strFirstName = String.Join(" ", names);
            }
        }
        // FIRSTNAME LASTNAME SUFFIX
        public void fn_separatefullnamev5( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            if ( names.Length <= 2 )
            {
                strLastName = names [1];
                strMiddleInitial = "";
                strFirstName = names [0];
            }
            else if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
            {
                if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
                {
                    if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                    {
                        strLastName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = "";
                        names = names.Reverse().Skip(3).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = "";
                        names = names.Reverse().Skip(2).Reverse().ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = "";
                    names = names.Reverse().Skip(2).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                {
                    strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = "";
                    names = names.Reverse().Skip(2).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names.Last();
                    strMiddleInitial = "";
                    names = names.Reverse().Skip(1).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else if ( strFullname.ToUpper().Contains("DE LOS ") || strFullname.ToUpper().Contains("DE LA ") )
            {
                if ( strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("III") )
                {
                    strLastName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = "";
                    names = names.Reverse().Skip(4).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    strMiddleInitial = "";
                    names = names.Reverse().Skip(3).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
            }
            else
            {
                strLastName = names.Last();
                strMiddleInitial = "";
                names = names.Reverse().Skip(1).Reverse().ToArray();
                strFirstName = String.Join(" ", names);
            }
        }
        //LASTNAME, FIRSTNAME MIDDLE INITIAL (some name have no middle initial)
        public void fn_separatefullnamev6( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(',');
            strFirstName = "";
            strLastName = names [0];
            if (names[0].ToUpper().Contains("JR") || names[0].ToUpper().Contains("SR") || names[0].ToUpper().Contains("II") || names[0].ToUpper().Contains("III"))
            {
                var fname = names[0].Split(' ');
                strFirstName = fname.Last();
                fname = fname.Take(fname.Count() - 1).ToArray();
                strLastName = String.Join(" ", fname);
            }
            names = names [1].Split(' ');
            if ( strFullname.Contains(".") )
            {
                strMiddleInitial = names.Last();
                names = names.Take(names.Count() - 1).ToArray();
            }
            else
            {
                strMiddleInitial = "";
            }
            strFirstName = String.Join(" ", names) + " " + strFirstName;

        }
        //FIRSTNAME MIDDLEINITIAL(some name have no middle initial) LASTNAME
        public void fn_separatefullnamev7( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            strLastName = names [0];
            names = names [1].Split(' ');
            if ( strFullname.Contains(".") )
            {
                strMiddleInitial = names.Last();
            }
            else
            {
                strMiddleInitial = "";
            }
            names = names.Take(names.Count() - 1).ToArray();
            strFirstName = String.Join(" ", names);

        }
        //LASTNAME FIRSTNAME SUFFIX MIDDLENAME
        public void fn_separatefullnamev8( string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial )
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            if ( names.Length <= 2 )
            {
                strLastName = names [0];
                strMiddleInitial = "";
                strFirstName = names [1];
            }
            else if ( strFullname.ToUpper().Contains("JR ") || strFullname.ToUpper().Contains("SR ") || strFullname.ToUpper().Contains("III ") )
            {
                if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
                {
                    if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                    {
                        strLastName = names [0] + " " + names [1];
                        strMiddleInitial = names [names.Length - 1];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(2).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                    else
                    {
                        strLastName = names [0];
                        strMiddleInitial = names [names.Length - 1];
                        names = names.Take(names.Length - 1).ToArray();
                        names = names.Skip(1).ToArray();
                        strFirstName = String.Join(" ", names);
                    }
                }
                else
                {
                    strLastName = names [0];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(1).ToArray();
                    strFirstName = String.Join(" ", names);
                }

            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("DE") || names.Contains("DEL") || names.Contains("DELA") || names.Contains("DELOS") || names.Contains("SAN") )
                {
                    strLastName = names [0] + " " + names [1];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(2).ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [0];
                    strMiddleInitial = names [names.Length - 1];
                    names = names.Take(names.Length - 1).ToArray();
                    names = names.Skip(1).ToArray();
                    strFirstName = String.Join(" ", names);
                }

            }
            else if ( strFullname.ToUpper().Contains("DE LOS ") || strFullname.ToUpper().Contains("DE LA ") )
            {

                strLastName = names [0] + " " + names [1] + " " + names [2];
                strMiddleInitial = names [names.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(3).ToArray();
                strFirstName = String.Join(" ", names);

            }
            else
            {
                strLastName = names [0];
                strMiddleInitial = names [names.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(1).ToArray();
                strFirstName = String.Join(" ", names);
            }
        }

        public void fn_separatefullnamev9(string strFullname, out string strFirstName, out string strLastName)
        {
            strFullname = strFullname.TrimEnd();
            strFullname = strFullname.ToUpper();
            var names = strFullname.Split(' ');
            if ( names.Length <= 2 )
            {
                strLastName = names [1];
                strFirstName = names [0];
            }
            else if ( strFullname.ToUpper().Contains("DE ") || strFullname.ToUpper().Contains("DELA ") || strFullname.ToUpper().Contains("DEL ") || strFullname.ToUpper().Contains("DELOS ") || strFullname.ToUpper().Contains("SAN ") )
            {
                if ( names.Contains("JR") || names.Contains("SR") || names.Contains("III") )
                {
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    names = names.Reverse().Skip(3).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }
                else
                {
                    strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                    names = names.Reverse().Skip(2).Reverse().ToArray();
                    strFirstName = String.Join(" ", names);
                }

            }
            else if ( names.Contains("JR.") || names.Contains("SR.") || names.Contains("III") )
            {
                strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                names = names.Reverse().Skip(2).Reverse().ToArray();
                strFirstName = String.Join(" ", names);
            }
            else
            {
                strLastName = names [names.Length - 1];
                names = names.Reverse().Skip(1).Reverse().ToArray();
                strFirstName = String.Join(" ", names);
            }
        }
        #endregion

     
        public string fn_GetMortality(double MortRating)
        {
            double[] array = new double[18] { 1, 1.25, 1.375, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 3.25, 3.5, 3.75, 4, 4.25, 4.5, 4.75, 5 };


            double TargetNumber = MortRating;

            var nearest = array.OrderBy(x => Math.Abs((double)x - TargetNumber)).First();

            return fn_getmortality(Convert.ToString(nearest));

        }


        public void fn_CheckTransCode(string businesstype, out string transcode)
        {
            if (businesstype.ToUpper() == "NEW BUSINESS" || businesstype.ToUpper() == "NEW")
            {
                transcode = "TNEWBUS";
            }
            else if (businesstype.ToUpper() == "RENEWAL")
            {
                transcode = "TRENEW";
            }
            else
            {
                transcode = " ";
            }

        }


        public string fn_CheckTransCodeV2(string valueSheetName)
        {
            if (valueSheetName.ToUpper().Contains("REN"))
            {
                return "TRENEW";
            }
            else if (valueSheetName.ToUpper().Contains("ADJUST"))
            {
                return "ADJUST";
            }
            else
            {
                return "TNEWBUS";
            }

        }

        public string fn_CheckTransCodeV3(string businesstype)
        {
            if(businesstype.ToUpper() == "RENEWAL")
            {
                return "TRENEW";
            }
            else if (businesstype.ToUpper() == "FIRST" || businesstype.ToUpper() == "NEW")
            {
                return "TNEWBUS";
            }
            else if (businesstype.ToUpper() == "TERMINATION" || businesstype.ToUpper() == "TERMINATED")
            {
                return "TCONTER";
            }
            else if (businesstype.ToUpper() == "REINSTATED" || businesstype.ToUpper() == "REINSTATEMENT")
            {
                return "TREINS";
            }
            else if (businesstype.ToUpper() == "CANCELLED")
            {
                return "TCANCINC";
            }
            else if (businesstype.ToUpper() == "EXPIRY" || businesstype.ToUpper() == "EXPIRED")
            {
                return "TEXPIRY";
            }
            else if (businesstype.ToUpper() == "SURRENDERED" || businesstype.ToUpper() == "SURRENDER" || businesstype.ToUpper() == "FULL SURENDERED")
            {
                return "TFULLSUR";
            }
            else if (businesstype.ToUpper() == "EXTENDED" || businesstype.ToUpper() == "TERM" || businesstype.ToUpper() == "ETI")
            {
                return "TEXTTER";
            }
            else if (businesstype.ToUpper() == "MATURITY" || businesstype.ToUpper() == "MATURED")
            {
                return "TFULLMAT";
            }
            else if (businesstype.ToUpper().Contains("PAID"))
            {
                return "TFULLPU";
            }
            else if (businesstype.ToUpper().Contains("CAP"))
            {
                return "TFULLREC";
            }
            else if (businesstype.ToUpper().Contains("LAPSE"))
            {
                return "TLAPSE";
            }
            else
            {
                return "ADJUST";
            }
                    
        }




        public string fn_LifeID(string strFirstName, string strLastName, string strBirthday)
        {
            strFirstName = strFirstName.ToUpper().Trim();
            strLastName = strLastName.ToUpper().Replace(" ", "");
            strBirthday = fn_convertStringtoDateV5(strBirthday);
            strBirthday = strBirthday.Replace("/", "");


            //if (strFirstName == null)
            //{
            //    strFirstName = "DummyFirstName";
            //}
            //if (strLastName == null)
            //{
            //    strLastName = "DummyLastName";
            //}
            //if (strBirthday == null)
            //{
            //    strBirthday = "07/01/1900";
            //}

            if (strLastName.Length < 5 && strFirstName.Length < 2)
            {
                return strLastName + strFirstName + strBirthday;
            }
            else if (strLastName.Length >= 5 && strFirstName.Length < 2)
            {
                return strLastName.Substring(0, 5) + strFirstName + strBirthday;
            }
            else if (strLastName.Length >= 5)
            {
                return strLastName.Substring(0, 5) + strFirstName.Substring(0, 2) + strBirthday;
            }
            else
            {
                return strLastName + strFirstName.Substring(0, 2) + strBirthday;
            }

        }

        public DateTime fn_convertStringtoDate(string strDate)
        {

            try
            {
                string year = strDate.Substring(0, 4);
                string month = strDate.Substring(4, 2);
                string day = strDate.Substring(strDate.Length - 2);
                DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                result = result.Date;
                return result;
            }
            catch (Exception ex)
            {

                return DateTime.Now;

            }
        }

        public void fn_searchpolicydb(string strPolicyNo, out string strAge, out string strFullName, out string strBirthdate, out string strGender)
        {
            try
            {
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                string query = "SELECT * FROM dbo_macro WHERE policy_no=" + "'" + strPolicyNo + "'";

                OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                OdbcDataReader DbReader = DbCommand.ExecuteReader();

                if (DbReader.Read())
                {
                    strAge = DbReader.GetValue(6).ToString();
                    strFullName = DbReader.GetValue(10).ToString();
                    strBirthdate = DbReader.GetValue(11).ToString();
                    strGender = DbReader.GetValue(12).ToString();
                }
                else
                {
                    strAge = "Dummy";
                    strFullName = "Dummy";
                    strBirthdate = "Dummy";
                    strGender = "Dummy";
                }

                DbReader.Close();
                cnDB.Dispose();
                cnDB.Close();
            }
            catch (Exception e)
            {
                strAge = "Policy number doesn't exist";
                strFullName = "Policy number doesn't exist";
                strBirthdate = "Policy number doesn't exist";
                strGender = "Policy number doesn't exist";
            }
        }

        public void fn_CheckingforA_AB_BZColumn(string valueZ, string valueAB, string valueBZ, out string strOriginalSum, out string strInitialSum, out string strSumAtRisk, out string strRemarksCode)
        {

            //BZ = SAR
            //AB = ISR
            //Z = OSR
            string ValueRC;

            if (valueBZ == null)
            {
                if (valueAB == null)
                {
                    strOriginalSum = valueZ;
                    strInitialSum = valueZ;
                    strSumAtRisk = valueZ;
                    ValueRC = "BR1-1|BR2-1";
                    strRemarksCode = ValueRC;

                }
                else if (valueZ == null)
                {
                    strOriginalSum = valueAB;
                    strInitialSum = valueAB;
                    strSumAtRisk = valueAB;
                    ValueRC = "BR3-1";
                    strRemarksCode = ValueRC;
                }
                else
                {
                    strOriginalSum = valueZ;
                    strInitialSum = valueAB;
                    strSumAtRisk = valueAB;
                    ValueRC = "BR3-1";
                    strRemarksCode = ValueRC;
                }

            }
            else if (valueAB == null)
            {
                if (valueZ == null)
                {
                    strOriginalSum = valueBZ;
                    strInitialSum = valueBZ;
                    strSumAtRisk = valueBZ;
                    ValueRC = "BR3-1";
                    strRemarksCode = ValueRC;
                }
                else
                {
                    strOriginalSum = valueZ;
                    strInitialSum = valueBZ;
                    strSumAtRisk = valueBZ;
                    ValueRC = "BR2-1 | BR2-2";
                    strRemarksCode = ValueRC;
                }
            }
            else if (valueZ == null)
            {
                strOriginalSum = valueAB;
                strInitialSum = valueAB;
                strSumAtRisk = valueBZ;
                ValueRC = "BR3-1";
                strRemarksCode = ValueRC;
            }
            else
            {
                strOriginalSum = valueZ;
                strInitialSum = valueAB;
                strSumAtRisk = valueBZ;
                ValueRC = "";
                strRemarksCode = ValueRC;
            }

        }
        public decimal fn_CheckingValueZeroOrEmpty(string Value)
        {

            if (Value == null || string.IsNullOrEmpty(Value))
            {
                return 1;
            }
            else if (Decimal.Parse(Value) == 0)
            {

                return 1;
            }
            else
            {
                return Decimal.Parse(Value);
            }
        }



        public decimal fn_multiplier(decimal Value, decimal multiplier)
        {

            if (Value != 0)
            {
                return Convert.ToDecimal(Value) * multiplier;

            }
            else
            {
                return 0;
            }
        }

        public string fn_CheckingValueZeroOrEmptyV2(string Value, out string strValueOne)
        {
            decimal decValueone;
            if (Value == null)
            {
                strValueOne = "1";
                return strValueOne;
            }
            else if (Decimal.Parse(Value) == 0)
            {
                strValueOne = "1";
                return strValueOne;
            }
            else
            {
                decValueone = Decimal.Parse(Value) / 100;
                strValueOne = Convert.ToString(decValueone);
                return strValueOne;
            }
        }

        public void fn_GetEntryCode(string Tcode, bool Fyear, out string str_entryCode, out bool boo_FY)
        {
            str_entryCode = null;
            boo_FY = false;

            if (Tcode == "ADJUST" && (boo_FY == true))
            {
                str_entryCode = "4002";
            }

        }




        public DataTable fn_Loadmacro(string str_macro)
        {
            _Global _var = new _Global();

            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbdata = eapp.Workbooks.Open(str_macro);
            Microsoft.Office.Interop.Excel.Worksheet wsdata = wbdata.Sheets[1];
            Microsoft.Office.Interop.Excel.Range datarange = wsdata.UsedRange;

            int edatarow = datarange.Rows.Count;

            _var.objdt_MACRO.Columns.Add("SPN", typeof(String));
            _var.objdt_MACRO.Columns.Add("URC", typeof(String));
            _var.objdt_MACRO.Columns.Add("POLICY NO.", typeof(String));
            _var.objdt_MACRO.Columns.Add("CESSION TYPE CODE", typeof(String));
            _var.objdt_MACRO.Columns.Add("CURRENCY CODE", typeof(String));
            _var.objdt_MACRO.Columns.Add("ISSUE AGE 1", typeof(String));
            _var.objdt_MACRO.Columns.Add("ISSUE DATE", typeof(String));
            _var.objdt_MACRO.Columns.Add("ISSUE DATE 1", typeof(String));
            _var.objdt_MACRO.Columns.Add("MORT RATING CODE", typeof(String));
            _var.objdt_MACRO.Columns.Add("REFUNDING CODE", typeof(String));
            _var.objdt_MACRO.Columns.Add("INSRD7M", typeof(String));
            _var.objdt_MACRO.Columns.Add("BIRTH7D", typeof(String));
            _var.objdt_MACRO.Columns.Add("DATE OF BIRTH", typeof(String));
            _var.objdt_MACRO.Columns.Add("SEX7C", typeof(String));
            _var.objdt_MACRO.Columns.Add("COVER7C", typeof(String));
            _var.objdt_MACRO.Columns.Add("BENEFIT", typeof(String));
            _var.objdt_MACRO.Columns.Add("AMT7INSRD7A", typeof(String));
            _var.objdt_MACRO.Columns.Add("AMT7REINSRD7A", typeof(String));
            _var.objdt_MACRO.Columns.Add("CED7RETN7A", typeof(String));
            _var.objdt_MACRO.Columns.Add("COMPANY NAME", typeof(String));
            _var.objdt_MACRO.Columns.Add("SOURCE", typeof(String));
            _var.objdt_MACRO.Columns.Add("CN", typeof(String));

            for (int intLoop = 3; intLoop <= edatarow + 1; intLoop++)
            {
                _var.dtworkRow = _var.objdt_MACRO.NewRow();
                _var.dtworkRow[0] = wsdata.Cells[intLoop, 1].Text.ToString();
                _var.dtworkRow[1] = wsdata.Cells[intLoop, 2].Text.ToString();
                _var.dtworkRow[2] = wsdata.Cells[intLoop, 3].Text.ToString();
                _var.dtworkRow[3] = wsdata.Cells[intLoop, 4].Text.ToString();
                _var.dtworkRow[4] = wsdata.Cells[intLoop, 5].Text.ToString();
                _var.dtworkRow[5] = wsdata.Cells[intLoop, 6].Text.ToString();
                _var.dtworkRow[6] = wsdata.Cells[intLoop, 7].Text.ToString();
                _var.dtworkRow[7] = wsdata.Cells[intLoop, 8].Text.ToString();
                _var.dtworkRow[8] = wsdata.Cells[intLoop, 9].Text.ToString();
                _var.dtworkRow[9] = wsdata.Cells[intLoop, 10].Text.ToString();

                _var.dtworkRow[10] = wsdata.Cells[intLoop, 11].Text.ToString();
                _var.dtworkRow[11] = wsdata.Cells[intLoop, 12].Text.ToString();
                _var.dtworkRow[12] = wsdata.Cells[intLoop, 13].Text.ToString();
                _var.dtworkRow[13] = wsdata.Cells[intLoop, 14].Text.ToString();
                _var.dtworkRow[14] = wsdata.Cells[intLoop, 15].Text.ToString();
                _var.dtworkRow[15] = wsdata.Cells[intLoop, 16].Text.ToString();
                _var.dtworkRow[16] = wsdata.Cells[intLoop, 17].Text.ToString();
                _var.dtworkRow[17] = wsdata.Cells[intLoop, 18].Text.ToString();
                _var.dtworkRow[18] = wsdata.Cells[intLoop, 19].Text.ToString();
                _var.dtworkRow[19] = wsdata.Cells[intLoop, 20].Text.ToString();
                _var.dtworkRow[20] = wsdata.Cells[intLoop, 21].Text.ToString();
                _var.dtworkRow[21] = wsdata.Cells[intLoop, 22].Text.ToString();

                _var.objdt_MACRO.Rows.Add(_var.dtworkRow);
            }

            wsdata = null;
            wbdata.Close();
            wbdata = null;

            eapp = null;

            return _var.objdt_MACRO;
        }

        public DataTable fn_LoadOCCCode()
        {
            _Global _var = new _Global();

            _var.objdt_OCCCODE.Columns.Add("CODE", typeof(String));
            _var.objdt_OCCCODE.Columns.Add("_NAME", typeof(String));

            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MANGR"; _var.dtworkRow[1] = "MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ABTRW"; _var.dtworkRow[1] = "ABATTOIR WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACCNT"; _var.dtworkRow[1] = "ACCOUNTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACRBT"; _var.dtworkRow[1] = "ACROBAT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACTOR"; _var.dtworkRow[1] = "ACTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACTRS"; _var.dtworkRow[1] = "ACTRESS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACTRY"; _var.dtworkRow[1] = "ACTUARY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACPNT"; _var.dtworkRow[1] = "ACUPUNCTURIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADJST"; _var.dtworkRow[1] = "ADJUSTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADMNT"; _var.dtworkRow[1] = "ADMINISTRATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADLTD"; _var.dtworkRow[1] = "ADULT EDUCATION TUTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADVRT"; _var.dtworkRow[1] = "ADVERTISER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AESTH"; _var.dtworkRow[1] = "AESTHETICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AFPNU"; _var.dtworkRow[1] = "AFP NON-UNIFORM PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AFPOF"; _var.dtworkRow[1] = "AFP OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTBK"; _var.dtworkRow[1] = "AGENT - BOOKING"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTNS"; _var.dtworkRow[1] = "AGENT - INSURANCE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTRN"; _var.dtworkRow[1] = "AGENT - RENTAL EQUIPMENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTRP"; _var.dtworkRow[1] = "AGENT - REPOSSESSION"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTSC"; _var.dtworkRow[1] = "AGENT - SECRET SERVICE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTSP"; _var.dtworkRow[1] = "AGENT - SPORTS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTTH"; _var.dtworkRow[1] = "AGENT - THEATRICAL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGTTC"; _var.dtworkRow[1] = "AGENT - TICKET"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AGRCL"; _var.dtworkRow[1] = "AGRICULTURIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARTRF"; _var.dtworkRow[1] = "AIR TRAFFIC CONTROL CLERK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARCRF"; _var.dtworkRow[1] = "AIRCRAFT WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ALTRN"; _var.dtworkRow[1] = "ALTERNATIVE MEDICINE PRACTITIONER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AMBSD"; _var.dtworkRow[1] = "AMBASSADOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AMSMT"; _var.dtworkRow[1] = "AMUSEMENT ARCADE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANLYS"; _var.dtworkRow[1] = "ANALYST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANMBR"; _var.dtworkRow[1] = "ANIMAL BREEDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANMSH"; _var.dtworkRow[1] = "ANIMAL SHELTER WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANMTR"; _var.dtworkRow[1] = "ANIMAL TRAINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANMTP"; _var.dtworkRow[1] = "ANIMAL TRAPPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANNCR"; _var.dtworkRow[1] = "ANNOUNCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ANQSR"; _var.dtworkRow[1] = "ANTIQUES RESTORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "AQFRM"; _var.dtworkRow[1] = "AQUATIC FARMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARCHT"; _var.dtworkRow[1] = "ARCHITECT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARMRY"; _var.dtworkRow[1] = "ARMORY KEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARMOF"; _var.dtworkRow[1] = "ARMY OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARTPR"; _var.dtworkRow[1] = "ART APPRAISER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARTRS"; _var.dtworkRow[1] = "ART RESTORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARTSN"; _var.dtworkRow[1] = "ARTISAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ARTST"; _var.dtworkRow[1] = "ARTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASBST"; _var.dtworkRow[1] = "ASBESTOS STRIPPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASPHL"; _var.dtworkRow[1] = "ASPHALT WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASSYR"; _var.dtworkRow[1] = "ASSAYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASSMB"; _var.dtworkRow[1] = "ASSEMBLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASTCH"; _var.dtworkRow[1] = "ASSEMBLY TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASSOR"; _var.dtworkRow[1] = "ASSESSOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASSTN"; _var.dtworkRow[1] = "ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ASTRL"; _var.dtworkRow[1] = "ASTROLOGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ATHLT"; _var.dtworkRow[1] = "ATHLETE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ATTND"; _var.dtworkRow[1] = "ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ACTNR"; _var.dtworkRow[1] = "AUCTIONEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADGST"; _var.dtworkRow[1] = "AUDIOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ADITR"; _var.dtworkRow[1] = "AUDITOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BCKFF"; _var.dtworkRow[1] = "BACK OFFICE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BCTRL"; _var.dtworkRow[1] = "BACTERIOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BAGGR"; _var.dtworkRow[1] = "BAGGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BALFF"; _var.dtworkRow[1] = "BAILLIFF"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BAKER"; _var.dtworkRow[1] = "BAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BANKR"; _var.dtworkRow[1] = "BANKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRSTF"; _var.dtworkRow[1] = "BAR STAFF"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRGYT"; _var.dtworkRow[1] = "BARANGAY TANOD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BARBR"; _var.dtworkRow[1] = "BARBER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BARKR"; _var.dtworkRow[1] = "BARKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRTND"; _var.dtworkRow[1] = "BARTENDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTICN"; _var.dtworkRow[1] = "BEAUTICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BOKPR"; _var.dtworkRow[1] = "BEEKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BELLBY"; _var.dtworkRow[1] = "BELLBOY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BCHMS"; _var.dtworkRow[1] = "BIOCHEMIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLGST"; _var.dtworkRow[1] = "BIOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLCKS"; _var.dtworkRow[1] = "BLACKSMITH"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLSTR"; _var.dtworkRow[1] = "BLASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLCHR"; _var.dtworkRow[1] = "BLEACHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLNDR"; _var.dtworkRow[1] = "BLENDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLCKR"; _var.dtworkRow[1] = "BLOCKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLOWR"; _var.dtworkRow[1] = "BLOWER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTBLD"; _var.dtworkRow[1] = "BOAT BUILDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTSWN"; _var.dtworkRow[1] = "BOATSWAIN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BDYGR"; _var.dtworkRow[1] = "BODYGUARD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLRPR"; _var.dtworkRow[1] = "BOILER OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BLRMK"; _var.dtworkRow[1] = "BOILERMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BOLTR"; _var.dtworkRow[1] = "BOLTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BMBDS"; _var.dtworkRow[1] = "BOMB DISPOSAL UNIT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BNSTT"; _var.dtworkRow[1] = "BONE SETTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BKBND"; _var.dtworkRow[1] = "BOOKBINDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BKPR"; _var.dtworkRow[1] = "BOOKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BMMN"; _var.dtworkRow[1] = "BOOMMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BORER"; _var.dtworkRow[1] = "BORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTTLR"; _var.dtworkRow[1] = "BOTTLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BONCR"; _var.dtworkRow[1] = "BOUNCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRKMN"; _var.dtworkRow[1] = "BRAKEMAN MOTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRTTC"; _var.dtworkRow[1] = "BRATTICEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BREKR"; _var.dtworkRow[1] = "BREAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BREWR"; _var.dtworkRow[1] = "BREWER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRCKL"; _var.dtworkRow[1] = "BRICKLAYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRCKM"; _var.dtworkRow[1] = "BRICKMASON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BROKR"; _var.dtworkRow[1] = "BROKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BILDR"; _var.dtworkRow[1] = "BUILDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRNSH"; _var.dtworkRow[1] = "BURNISHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BRSRD"; _var.dtworkRow[1] = "BURSAR - EDUCATION"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BSCND"; _var.dtworkRow[1] = "BUS CONDUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BSNAS"; _var.dtworkRow[1] = "BUSINESS PROCESS ASSOCIATE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BSNSS"; _var.dtworkRow[1] = "BUSINESSMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTCHR"; _var.dtworkRow[1] = "BUTCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BUTLR"; _var.dtworkRow[1] = "BUTLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "BTTRM"; _var.dtworkRow[1] = "BUTTERMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CBLMN"; _var.dtworkRow[1] = "CABLEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CADET"; _var.dtworkRow[1] = "CADET"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CAGER"; _var.dtworkRow[1] = "CAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLBRT"; _var.dtworkRow[1] = "CALIBRATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLLCN"; _var.dtworkRow[1] = "CALL CENTER AGENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMRMN"; _var.dtworkRow[1] = "CAMERAMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNTNS"; _var.dtworkRow[1] = "CANTEEN ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNVSS"; _var.dtworkRow[1] = "CANVASSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CAPTN"; _var.dtworkRow[1] = "CAPTAIN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRDDL"; _var.dtworkRow[1] = "CARD DEALER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRGVR"; _var.dtworkRow[1] = "CAREGIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRTKR"; _var.dtworkRow[1] = "CARETAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRPNT"; _var.dtworkRow[1] = "CARPENTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CARIR"; _var.dtworkRow[1] = "CARRIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CARVR"; _var.dtworkRow[1] = "CARVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRWSH"; _var.dtworkRow[1] = "CARWASH ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CASHR"; _var.dtworkRow[1] = "CASHIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CSNCR"; _var.dtworkRow[1] = "CASINO CROUPIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CSNWR"; _var.dtworkRow[1] = "CASINO WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CASTR"; _var.dtworkRow[1] = "CASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CATRR"; _var.dtworkRow[1] = "CATERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CATHP"; _var.dtworkRow[1] = "CATERING HELPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CTTLH"; _var.dtworkRow[1] = "CATTLE HERDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLNGF"; _var.dtworkRow[1] = "CEILING FIXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLLPH"; _var.dtworkRow[1] = "CELL PHONE TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMTRY"; _var.dtworkRow[1] = "CEMETERY WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHMBR"; _var.dtworkRow[1] = "CHAMBERMAID"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHFFR"; _var.dtworkRow[1] = "CHAUFFER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHCKR"; _var.dtworkRow[1] = "CHECKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHCAT"; _var.dtworkRow[1] = "CHECKROOM ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHSMK"; _var.dtworkRow[1] = "CHEESEMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHEF"; _var.dtworkRow[1] = "CHEF"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHMEN"; _var.dtworkRow[1] = "CHEMICAL ENGINEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHMWK"; _var.dtworkRow[1] = "CHEMICAL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHMST"; _var.dtworkRow[1] = "CHEMIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHFXC"; _var.dtworkRow[1] = "CHIEF EXECUTIVE OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHFPR"; _var.dtworkRow[1] = "CHIEF OPERATING OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHLDC"; _var.dtworkRow[1] = "CHILD CARE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHRPD"; _var.dtworkRow[1] = "CHIROPODIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHRPR"; _var.dtworkRow[1] = "CHIROPRACTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CHRNM"; _var.dtworkRow[1] = "CHURNMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNDRM"; _var.dtworkRow[1] = "CINDERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRCSP"; _var.dtworkRow[1] = "CIRCUS PERFORMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRCSW"; _var.dtworkRow[1] = "CIRCUS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLMSD"; _var.dtworkRow[1] = "CLAIMS ADJUSTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLMSS"; _var.dtworkRow[1] = "CLAIMS ASSESSOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLSSR"; _var.dtworkRow[1] = "CLASSROOM ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLENR"; _var.dtworkRow[1] = "CLEANER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLERK"; _var.dtworkRow[1] = "CLERK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLMBR"; _var.dtworkRow[1] = "CLIMBER TOPMEN-HIGH"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COACH"; _var.dtworkRow[1] = "COACH"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CSTGR"; _var.dtworkRow[1] = "COASTGUARD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CBBLR"; _var.dtworkRow[1] = "COBBLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CCKPT"; _var.dtworkRow[1] = "COCKPIT OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CFFNM"; _var.dtworkRow[1] = "COFFIN MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLLCT"; _var.dtworkRow[1] = "COLLECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CLRMX"; _var.dtworkRow[1] = "COLORMIXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COMDN"; _var.dtworkRow[1] = "COMEDIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMMNT"; _var.dtworkRow[1] = "COMMENTATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMPSR"; _var.dtworkRow[1] = "COMPOSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMPND"; _var.dtworkRow[1] = "COMPOUNDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CMPTR"; _var.dtworkRow[1] = "COMPTROLLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CPTRE"; _var.dtworkRow[1] = "COMPUTER ENGINEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CPTRP"; _var.dtworkRow[1] = "COMPUTER PROGRAMMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CPTRT"; _var.dtworkRow[1] = "COMPUTER TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNCRG"; _var.dtworkRow[1] = "CONCIERGE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNCRT"; _var.dtworkRow[1] = "CONCRETE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNDNS"; _var.dtworkRow[1] = "CONDENSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CONDC"; _var.dtworkRow[1] = "CONDUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CONER"; _var.dtworkRow[1] = "CONER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CONFC"; _var.dtworkRow[1] = "CONFECTIONER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CONSR"; _var.dtworkRow[1] = "CONSERVATIONIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNSTB"; _var.dtworkRow[1] = "CONSTABLE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNSTR"; _var.dtworkRow[1] = "CONSTRUCTION WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNSLT"; _var.dtworkRow[1] = "CONSULTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNTRC"; _var.dtworkRow[1] = "CONTRACTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNTRL"; _var.dtworkRow[1] = "CONTROLLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNVRT"; _var.dtworkRow[1] = "CONVERTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COOK"; _var.dtworkRow[1] = "COOK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COOPR"; _var.dtworkRow[1] = "COOPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COPDR"; _var.dtworkRow[1] = "COPRA DRIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CORKR"; _var.dtworkRow[1] = "CORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRNPC"; _var.dtworkRow[1] = "CORN PICKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRNR"; _var.dtworkRow[1] = "CORONER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNSLR"; _var.dtworkRow[1] = "COUNSELOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CNTRM"; _var.dtworkRow[1] = "COUNTERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "COURR"; _var.dtworkRow[1] = "COURIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRBFS"; _var.dtworkRow[1] = "CRAB FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRFTP"; _var.dtworkRow[1] = "CRAFTPERSON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRNDR"; _var.dtworkRow[1] = "CRANE DRIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRNMN"; _var.dtworkRow[1] = "CRANEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRDTC"; _var.dtworkRow[1] = "CREDIT CONTROLLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRMTR"; _var.dtworkRow[1] = "CREMATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRMNL"; _var.dtworkRow[1] = "CRIMINOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRCDL"; _var.dtworkRow[1] = "CROCODILE FARMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRSHR"; _var.dtworkRow[1] = "CRUSHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CRBLY"; _var.dtworkRow[1] = "CURB LAYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CURER"; _var.dtworkRow[1] = "CURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CUSTM"; _var.dtworkRow[1] = "CUSTOMS OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "CUTTR"; _var.dtworkRow[1] = "CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DNCNS"; _var.dtworkRow[1] = "DANCE INSTRUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DANCR"; _var.dtworkRow[1] = "DANCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DTPRC"; _var.dtworkRow[1] = "DATA PROCESSOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DEALR"; _var.dtworkRow[1] = "DEALER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DBTCL"; _var.dtworkRow[1] = "DEBT COLLECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DCKHN"; _var.dtworkRow[1] = "DECKHAND"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DCRTR"; _var.dtworkRow[1] = "DECORATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DLVRY"; _var.dtworkRow[1] = "DELIVERYMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DMLTN"; _var.dtworkRow[1] = "DEMOLITION WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DNTLS"; _var.dtworkRow[1] = "DENTAL ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DNTLH"; _var.dtworkRow[1] = "DENTAL HYGIENIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DNTST"; _var.dtworkRow[1] = "DENTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRRCK"; _var.dtworkRow[1] = "DERRICKMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSGNR"; _var.dtworkRow[1] = "DESIGNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DTCTV"; _var.dtworkRow[1] = "DETECTIVE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DVLPR"; _var.dtworkRow[1] = "DEVELOPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DCTTR"; _var.dtworkRow[1] = "DIE CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DIETN"; _var.dtworkRow[1] = "DIETICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DPLMT"; _var.dtworkRow[1] = "DIPLOMAT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRCTR"; _var.dtworkRow[1] = "DIRECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSCJC"; _var.dtworkRow[1] = "DISC JOCKEY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSHWH"; _var.dtworkRow[1] = "DISHWAHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSHWS"; _var.dtworkRow[1] = "DISHWASHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSPTC"; _var.dtworkRow[1] = "DISPATCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSPSL"; _var.dtworkRow[1] = "DISPOSAL CREW"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSTLL"; _var.dtworkRow[1] = "DISTILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DSTRB"; _var.dtworkRow[1] = "DISTRIBUTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DTCHD"; _var.dtworkRow[1] = "DITCHDIGGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DIVER"; _var.dtworkRow[1] = "DIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DCKMS"; _var.dtworkRow[1] = "DOCK MASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DOCKR"; _var.dtworkRow[1] = "DOCKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DGCTC"; _var.dtworkRow[1] = "DOG CATCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DMSTC"; _var.dtworkRow[1] = "DOMESTIC HELPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DORMN"; _var.dtworkRow[1] = "DOORMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRFTS"; _var.dtworkRow[1] = "DRAFTSMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRSSR"; _var.dtworkRow[1] = "DRESSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRSSM"; _var.dtworkRow[1] = "DRESSMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRLLR"; _var.dtworkRow[1] = "DRILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "DRIVR"; _var.dtworkRow[1] = "DRIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ECNMT"; _var.dtworkRow[1] = "ECONOMIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EDITR"; _var.dtworkRow[1] = "EDITOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ELCTR"; _var.dtworkRow[1] = "ELECTRICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ELCPL"; _var.dtworkRow[1] = "ELECTROPLATER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ELVOP"; _var.dtworkRow[1] = "ELEVATOR OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EMBMR"; _var.dtworkRow[1] = "EMBALMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ENCDR"; _var.dtworkRow[1] = "ENCODER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ENGNR"; _var.dtworkRow[1] = "ENGINEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ENGTC"; _var.dtworkRow[1] = "ENGINEERING TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ENGRV"; _var.dtworkRow[1] = "ENGRAVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ENTRT"; _var.dtworkRow[1] = "ENTERTAINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EQSTR"; _var.dtworkRow[1] = "EQUESTRIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ERCTR"; _var.dtworkRow[1] = "ERECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ESTMT"; _var.dtworkRow[1] = "ESTIMATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ETCHR"; _var.dtworkRow[1] = "ETCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EVNTC"; _var.dtworkRow[1] = "EVENT COORDINATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXMNR"; _var.dtworkRow[1] = "EXAMINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXCTV"; _var.dtworkRow[1] = "EXECUTIVE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXCAS"; _var.dtworkRow[1] = "EXECUTIVE ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXHBT"; _var.dtworkRow[1] = "EXHIBITION ORGANISER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXPLV"; _var.dtworkRow[1] = "EXPLOSIVES MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXTRM"; _var.dtworkRow[1] = "EXTERMINATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "EXTRC"; _var.dtworkRow[1] = "EXTRACTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FCTRY"; _var.dtworkRow[1] = "FACTORY WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FTHHL"; _var.dtworkRow[1] = "FAITH HEALER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRMBL"; _var.dtworkRow[1] = "FARM BLASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRMFN"; _var.dtworkRow[1] = "FARM FENCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRMLB"; _var.dtworkRow[1] = "FARM LABORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FARMR"; _var.dtworkRow[1] = "FARMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHND"; _var.dtworkRow[1] = "FASHION DESIGNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSTFD"; _var.dtworkRow[1] = "FAST FOOD CREW MEMBER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FEEDR"; _var.dtworkRow[1] = "FEEDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRMNT"; _var.dtworkRow[1] = "FERMENTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FILLR"; _var.dtworkRow[1] = "FILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLMPR"; _var.dtworkRow[1] = "FILM PROJECTIONIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FNNCL"; _var.dtworkRow[1] = "FINANCIAL ADVISER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FNSHR"; _var.dtworkRow[1] = "FINISHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRFFC"; _var.dtworkRow[1] = "FIRE OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRFGH"; _var.dtworkRow[1] = "FIREFIGHTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRWRK"; _var.dtworkRow[1] = "FIREWORKS ASSEMBLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FISCL"; _var.dtworkRow[1] = "FISCAL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHCR"; _var.dtworkRow[1] = "FISH CURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHPR"; _var.dtworkRow[1] = "FISH PROCESSOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHRM"; _var.dtworkRow[1] = "FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHRY"; _var.dtworkRow[1] = "FISHERY OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHNG"; _var.dtworkRow[1] = "FISHING GUIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHMN"; _var.dtworkRow[1] = "FISHMONGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSHPN"; _var.dtworkRow[1] = "FISHPOND WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FTNSS"; _var.dtworkRow[1] = "FITNESS INSTRUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLGMN"; _var.dtworkRow[1] = "FLAGMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLGHT"; _var.dtworkRow[1] = "FLIGHT ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLRMN"; _var.dtworkRow[1] = "FLOOR MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLRST"; _var.dtworkRow[1] = "FLORIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FLRML"; _var.dtworkRow[1] = "FLOUR MILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FDPRC"; _var.dtworkRow[1] = "FOOD PROCESS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FDSRV"; _var.dtworkRow[1] = "FOOD SERVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FDSTL"; _var.dtworkRow[1] = "FOOD STALL HOLDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRMN"; _var.dtworkRow[1] = "FOREMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSTRG"; _var.dtworkRow[1] = "FOREST RANGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FSTER"; _var.dtworkRow[1] = "FORESTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRGMN"; _var.dtworkRow[1] = "FORGEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FORGR"; _var.dtworkRow[1] = "FORGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRKLF"; _var.dtworkRow[1] = "FORK LIFT DRIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRTNT"; _var.dtworkRow[1] = "FORTUNE TELLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRGHT"; _var.dtworkRow[1] = "FREIGHT MAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRTPC"; _var.dtworkRow[1] = "FRUIT PICKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FMGTR"; _var.dtworkRow[1] = "FUMIGATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FNRLD"; _var.dtworkRow[1] = "FUNERAL DIRECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRNCM"; _var.dtworkRow[1] = "FURNACEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRNDL"; _var.dtworkRow[1] = "FURNITURE DEALER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRNMK"; _var.dtworkRow[1] = "FURNITURE MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "FRNRS"; _var.dtworkRow[1] = "FURNITURE RESTORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GAFFR"; _var.dtworkRow[1] = "GAFFER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLLYH"; _var.dtworkRow[1] = "GALLEY HAND"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLVNZ"; _var.dtworkRow[1] = "GALVANIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRBGC"; _var.dtworkRow[1] = "GARBAGE COLLECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRDNR"; _var.dtworkRow[1] = "GARDENER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRMNT"; _var.dtworkRow[1] = "GARMENT WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GSFTT"; _var.dtworkRow[1] = "GAS FITTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GSSTT"; _var.dtworkRow[1] = "GAS STATION ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GSWRK"; _var.dtworkRow[1] = "GAS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GTKPR"; _var.dtworkRow[1] = "GATEKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GTHRR"; _var.dtworkRow[1] = "GATHERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GAUGR"; _var.dtworkRow[1] = "GAUGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GMCTT"; _var.dtworkRow[1] = "GEM CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GCHMS"; _var.dtworkRow[1] = "GEOCHEMIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GDTCN"; _var.dtworkRow[1] = "GEODETIC ENGINEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLGST"; _var.dtworkRow[1] = "GEOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GPHYS"; _var.dtworkRow[1] = "GEOPHYSICIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GILDR"; _var.dtworkRow[1] = "GILDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLSSW"; _var.dtworkRow[1] = "GLASS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLAZR"; _var.dtworkRow[1] = "GLAZIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLUER"; _var.dtworkRow[1] = "GLUER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLDSM"; _var.dtworkRow[1] = "GOLDSMITH"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GLFCD"; _var.dtworkRow[1] = "GOLF CADDIE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRADR"; _var.dtworkRow[1] = "GRADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRNLT"; _var.dtworkRow[1] = "GRANULATORMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRESR"; _var.dtworkRow[1] = "GREASER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRLLR"; _var.dtworkRow[1] = "GRILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRNDR"; _var.dtworkRow[1] = "GRINDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRNDC"; _var.dtworkRow[1] = "GROUNDCREW"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GRNDS"; _var.dtworkRow[1] = "GROUNDS KEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GROTR"; _var.dtworkRow[1] = "GROUTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "GSTRL"; _var.dtworkRow[1] = "GUEST RELATIONS OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRDRS"; _var.dtworkRow[1] = "HAIRDRESSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HMMRM"; _var.dtworkRow[1] = "HAMMERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HNDCR"; _var.dtworkRow[1] = "HANDICRAFTPERSON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HNDLR"; _var.dtworkRow[1] = "HANDLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRBMS"; _var.dtworkRow[1] = "HARBOR MASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRBWK"; _var.dtworkRow[1] = "HARBOR WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRPNF"; _var.dtworkRow[1] = "HARPOON FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRVST"; _var.dtworkRow[1] = "HARVESTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HTCHR"; _var.dtworkRow[1] = "HATCHERY WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HAULR"; _var.dtworkRow[1] = "HAULER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HLTHR"; _var.dtworkRow[1] = "HEALTH RECORD TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HCKLR"; _var.dtworkRow[1] = "HECKLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRDNS"; _var.dtworkRow[1] = "HERD INSPECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HDCLL"; _var.dtworkRow[1] = "HIDECELLAR WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSTPR"; _var.dtworkRow[1] = "HOIST OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRCWK"; _var.dtworkRow[1] = "HORSE and DOG RACING WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HRSSH"; _var.dtworkRow[1] = "HORSESHOER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSPAD"; _var.dtworkRow[1] = "HOSPITAL AIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSPOR"; _var.dtworkRow[1] = "HOSPITAL ORDERLY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HOSTS"; _var.dtworkRow[1] = "HOST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSTSS"; _var.dtworkRow[1] = "HOSTESS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSKPR"; _var.dtworkRow[1] = "HOUSEKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HSWFE"; _var.dtworkRow[1] = "HOUSEWIFE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HCKST"; _var.dtworkRow[1] = "HUCKSTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HMNRS"; _var.dtworkRow[1] = "HUMAN RESOURCES PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HNTNG"; _var.dtworkRow[1] = "HUNTING GUIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HUSKR"; _var.dtworkRow[1] = "HUSKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "HYDRT"; _var.dtworkRow[1] = "HYDROTHERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ILSTR"; _var.dtworkRow[1] = "ILLUSTRATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "IMGRT"; _var.dtworkRow[1] = "IMMIGRATION OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSPT"; _var.dtworkRow[1] = "INSPECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSTL"; _var.dtworkRow[1] = "INSTALLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSTR"; _var.dtworkRow[1] = "INSTRUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSLT"; _var.dtworkRow[1] = "INSULATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSDJ"; _var.dtworkRow[1] = "INSURANCE ADJUSTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSBK"; _var.dtworkRow[1] = "INSURANCE BROKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INSIV"; _var.dtworkRow[1] = "INSURANCE INVESTIGATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INTRD"; _var.dtworkRow[1] = "INTERIOR DESIGNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INTRP"; _var.dtworkRow[1] = "INTERPRETOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "INVSG"; _var.dtworkRow[1] = "INVESTIGATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "IRGTN"; _var.dtworkRow[1] = "IRRIGATION WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ITPRF"; _var.dtworkRow[1] = "IT PROFESSIONAL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JLSHR"; _var.dtworkRow[1] = "JAI ALAI USHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JLFFC"; _var.dtworkRow[1] = "JAIL OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JANTR"; _var.dtworkRow[1] = "JANITOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JNTRS"; _var.dtworkRow[1] = "JANITRESS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JWLLR"; _var.dtworkRow[1] = "JEWELLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JOCKY"; _var.dtworkRow[1] = "JOCKEY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JRNLS"; _var.dtworkRow[1] = "JOURNALIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JUDGE"; _var.dtworkRow[1] = "JUDGE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JMPRF"; _var.dtworkRow[1] = "JUMPER FIREFIGHTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "JUVNL"; _var.dtworkRow[1] = "JUVENILE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KYCTT"; _var.dtworkRow[1] = "KEY CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KLNTT"; _var.dtworkRow[1] = "KILN ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KLNMN"; _var.dtworkRow[1] = "KILN MAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KTCHN"; _var.dtworkRow[1] = "KITCHEN AIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KNEDR"; _var.dtworkRow[1] = "KNEADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "KNTTR"; _var.dtworkRow[1] = "KNITTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LBELR"; _var.dtworkRow[1] = "LABELER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LBRTR"; _var.dtworkRow[1] = "LABORATORY TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LBORR"; _var.dtworkRow[1] = "LABORER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LMNTR"; _var.dtworkRow[1] = "LAMINATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LNDSC"; _var.dtworkRow[1] = "LANDSCAPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LNDMN"; _var.dtworkRow[1] = "LAUNDRYMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LNDWM"; _var.dtworkRow[1] = "LAUNDRYWOMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LWYER"; _var.dtworkRow[1] = "LAWYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LTHRF"; _var.dtworkRow[1] = "LEATHER FINISHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LTHRT"; _var.dtworkRow[1] = "LEATHER TANNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LCTRR"; _var.dtworkRow[1] = "LECTURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LBRRN"; _var.dtworkRow[1] = "LIBRARIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LFGRD"; _var.dtworkRow[1] = "LIFEGUARD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LNMN"; _var.dtworkRow[1] = "LINEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LTRRY"; _var.dtworkRow[1] = "LITERARY AGENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LTHGR"; _var.dtworkRow[1] = "LITHOGRAPHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LVSTC"; _var.dtworkRow[1] = "LIVESTOCK INSEMINATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LOADR"; _var.dtworkRow[1] = "LOADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LNXMN"; _var.dtworkRow[1] = "LOAN EXAMINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LBSTR"; _var.dtworkRow[1] = "LOBSTER FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LCKSM"; _var.dtworkRow[1] = "LOCKSMITH"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LGGER"; _var.dtworkRow[1] = "LOGGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LMBRY"; _var.dtworkRow[1] = "LUMBER YARD WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "LMBRJ"; _var.dtworkRow[1] = "LUMBERJACK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MCHNP"; _var.dtworkRow[1] = "MACHINE OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MCHNT"; _var.dtworkRow[1] = "MACHINE TOOL SETTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MCHNS"; _var.dtworkRow[1] = "MACHINIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MAIDS"; _var.dtworkRow[1] = "MAID"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNTCH"; _var.dtworkRow[1] = "MAINTENANCE TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNTWK"; _var.dtworkRow[1] = "MAINTENANCE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MKPRT"; _var.dtworkRow[1] = "MAKE-UP ARTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MANGR"; _var.dtworkRow[1] = "MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNHLM"; _var.dtworkRow[1] = "MANHOLE MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNCRS"; _var.dtworkRow[1] = "MANICURIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNFCT"; _var.dtworkRow[1] = "MANUFACTURING - SKILLED WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MRBLS"; _var.dtworkRow[1] = "MARBLE SETTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MRNNG"; _var.dtworkRow[1] = "MARINE ENGINEER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MRSHL"; _var.dtworkRow[1] = "MARSHAL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MASON"; _var.dtworkRow[1] = "MASON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MASSR"; _var.dtworkRow[1] = "MASSEUR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MASSE"; _var.dtworkRow[1] = "MASSEUSE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTBNR"; _var.dtworkRow[1] = "MEAT BONER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTCRR"; _var.dtworkRow[1] = "MEAT CURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTCTT"; _var.dtworkRow[1] = "MEAT CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MCHNC"; _var.dtworkRow[1] = "MECHANIC"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MEDIA"; _var.dtworkRow[1] = "MEDIA"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MDCLS"; _var.dtworkRow[1] = "MEDICAL SECRETARY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MDCLT"; _var.dtworkRow[1] = "MEDICAL TECHNOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MELTR"; _var.dtworkRow[1] = "MELTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MENDR"; _var.dtworkRow[1] = "MENDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MRCRZ"; _var.dtworkRow[1] = "MERCERIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MRCHN"; _var.dtworkRow[1] = "MERCHANT BANKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MSSNG"; _var.dtworkRow[1] = "MESSENGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MSSMN"; _var.dtworkRow[1] = "MESSMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTLLR"; _var.dtworkRow[1] = "METALLURGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTRRD"; _var.dtworkRow[1] = "METER READER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MDWFE"; _var.dtworkRow[1] = "MIDWIFE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MLTDV"; _var.dtworkRow[1] = "MILITARY DIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MLTPL"; _var.dtworkRow[1] = "MILITARY PILOT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MLTRV"; _var.dtworkRow[1] = "MILITARY RESERVIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MLLPR"; _var.dtworkRow[1] = "MILL OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MLLWR"; _var.dtworkRow[1] = "MILL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MILLR"; _var.dtworkRow[1] = "MILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MINER"; _var.dtworkRow[1] = "MINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MSSNR"; _var.dtworkRow[1] = "MISSIONARY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MIXER"; _var.dtworkRow[1] = "MIXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MODEL"; _var.dtworkRow[1] = "MODEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MOLDR"; _var.dtworkRow[1] = "MOLDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNYCH"; _var.dtworkRow[1] = "MONEY CHANGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MNYLN"; _var.dtworkRow[1] = "MONEY LENDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MSCLY"; _var.dtworkRow[1] = "MOSAIC LAYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MTRMN"; _var.dtworkRow[1] = "MOTORMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MOVER"; _var.dtworkRow[1] = "MOVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "MUSCN"; _var.dtworkRow[1] = "MUSICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NANNY"; _var.dtworkRow[1] = "NANNY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NRCMF"; _var.dtworkRow[1] = "NARCOM OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NVYFF"; _var.dtworkRow[1] = "NAVY OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NBGNT"; _var.dtworkRow[1] = "NBI AGENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NWSRP"; _var.dtworkRow[1] = "NEWS REPORTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NUNNS"; _var.dtworkRow[1] = "NUN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NURSE"; _var.dtworkRow[1] = "NURSE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NRSAD"; _var.dtworkRow[1] = "NURSE AIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "NTGTH"; _var.dtworkRow[1] = "NUT GATHERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OBSRV"; _var.dtworkRow[1] = "OBSERVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OCCPN"; _var.dtworkRow[1] = "OCCUPATIONAL THERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OFFCL"; _var.dtworkRow[1] = "OFFICE CLERK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OFCMG"; _var.dtworkRow[1] = "OFFICE MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OFCPL"; _var.dtworkRow[1] = "OFFICE PLANTS MAINTENANCE OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OILER"; _var.dtworkRow[1] = "OILER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPREX"; _var.dtworkRow[1] = "OPERATOR - EXCAVATING MACHINE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPRGN"; _var.dtworkRow[1] = "OPERATOR - GENERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPRHE"; _var.dtworkRow[1] = "OPERATOR - HEAVY EQUIPMENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPRMC"; _var.dtworkRow[1] = "OPERATOR - MACHINE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPROT"; _var.dtworkRow[1] = "OPERATOR - OTHERS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPRPL"; _var.dtworkRow[1] = "OPERATOR - PLANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPTCH"; _var.dtworkRow[1] = "OPHTHALMIC TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPTHL"; _var.dtworkRow[1] = "OPHTHALMOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPTCN"; _var.dtworkRow[1] = "OPTICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OPTMT"; _var.dtworkRow[1] = "OPTOMETRIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ORNMT"; _var.dtworkRow[1] = "ORNAMENTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ORTHD"; _var.dtworkRow[1] = "ORTHONDONTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ORTHP"; _var.dtworkRow[1] = "ORTHOPTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OTHER"; _var.dtworkRow[1] = "OTHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "OYSFS"; _var.dtworkRow[1] = "OYSTER FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PACKR"; _var.dtworkRow[1] = "PACKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PADLR"; _var.dtworkRow[1] = "PADDLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PNTMX"; _var.dtworkRow[1] = "PAINT MIXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PANTR"; _var.dtworkRow[1] = "PAINTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLLTM"; _var.dtworkRow[1] = "PALLET MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRMDC"; _var.dtworkRow[1] = "PARAMEDIC"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PSTRZ"; _var.dtworkRow[1] = "PASTEURIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PASTR"; _var.dtworkRow[1] = "PASTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PTHLG"; _var.dtworkRow[1] = "PATHOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PTRLM"; _var.dtworkRow[1] = "PATROLMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PTTRN"; _var.dtworkRow[1] = "PATTERN MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PAVIR"; _var.dtworkRow[1] = "PAVIOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PWNSH"; _var.dtworkRow[1] = "PAWN SHOP OWNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PWNBR"; _var.dtworkRow[1] = "PAWNBROKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PYLDR"; _var.dtworkRow[1] = "PAYLOADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRLFS"; _var.dtworkRow[1] = "PEARL FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PEDLR"; _var.dtworkRow[1] = "PEDDLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PDTRC"; _var.dtworkRow[1] = "PEDIATRICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHMCT"; _var.dtworkRow[1] = "PHARMACIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHAST"; _var.dtworkRow[1] = "PHARMACY ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHTNG"; _var.dtworkRow[1] = "PHOTOENGRAVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHTGR"; _var.dtworkRow[1] = "PHOTOGRAPHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHTLT"; _var.dtworkRow[1] = "PHOTOLITHOGRAPHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHYSC"; _var.dtworkRow[1] = "PHYSICAL THERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHYMD"; _var.dtworkRow[1] = "PHYSICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PHYTP"; _var.dtworkRow[1] = "PHYSIOTHERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PICKR"; _var.dtworkRow[1] = "PICKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLDRV"; _var.dtworkRow[1] = "PILE DRIVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PILER"; _var.dtworkRow[1] = "PILER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PILOT"; _var.dtworkRow[1] = "PILOT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PPDRL"; _var.dtworkRow[1] = "PIPE DRILLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PPFTT"; _var.dtworkRow[1] = "PIPEFITTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLNTM"; _var.dtworkRow[1] = "PLANT MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLSTR"; _var.dtworkRow[1] = "PLASTERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTCT"; _var.dtworkRow[1] = "PLATE CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLATER"; _var.dtworkRow[1] = "PLATER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLCKR"; _var.dtworkRow[1] = "PLUCKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLMBR"; _var.dtworkRow[1] = "PLUMBER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PNPNN"; _var.dtworkRow[1] = "PNP NON-UNIFORM PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PNPFF"; _var.dtworkRow[1] = "PNP OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PONTR"; _var.dtworkRow[1] = "POINTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLSHR"; _var.dtworkRow[1] = "POLISHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTCL"; _var.dtworkRow[1] = "POLITICAL ADVISER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTCN"; _var.dtworkRow[1] = "POLITICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PNDMN"; _var.dtworkRow[1] = "PONDMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PORTR"; _var.dtworkRow[1] = "PORTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PTFSH"; _var.dtworkRow[1] = "POT FISHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "POTTR"; _var.dtworkRow[1] = "POTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTRY"; _var.dtworkRow[1] = "POULTRY DRESSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTSX"; _var.dtworkRow[1] = "POULTRY SEXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PLTMN"; _var.dtworkRow[1] = "POULTRYMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PWRSH"; _var.dtworkRow[1] = "POWER SHOVEL OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRASM"; _var.dtworkRow[1] = "PRECISION ASSEMBLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRMKR"; _var.dtworkRow[1] = "PRECISION INSTRUMENT MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRSDN"; _var.dtworkRow[1] = "PRESIDENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRIES"; _var.dtworkRow[1] = "PRIEST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRNCP"; _var.dtworkRow[1] = "PRINCIPAL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRNTR"; _var.dtworkRow[1] = "PRINTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRSNW"; _var.dtworkRow[1] = "PRISON WARDEN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRVTT"; _var.dtworkRow[1] = "PRIVATE TUTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRBTN"; _var.dtworkRow[1] = "PROBATION OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRCSS"; _var.dtworkRow[1] = "PROCESS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRDCR"; _var.dtworkRow[1] = "PRODUCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRCRW"; _var.dtworkRow[1] = "PRODUCTION CREW"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRDWK"; _var.dtworkRow[1] = "PRODUCTION LINE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRDMG"; _var.dtworkRow[1] = "PRODUCTION MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRFSS"; _var.dtworkRow[1] = "PROFESSOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRGRM"; _var.dtworkRow[1] = "PROGRAMMER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRMDZ"; _var.dtworkRow[1] = "PROMODIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRPRT"; _var.dtworkRow[1] = "PROPRIETOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRSCT"; _var.dtworkRow[1] = "PROSECUTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PRSTH"; _var.dtworkRow[1] = "PROSTHESIS MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PSYCH"; _var.dtworkRow[1] = "PSYCHIATRIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PSYLG"; _var.dtworkRow[1] = "PSYCHOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PBLCR"; _var.dtworkRow[1] = "PUBLIC RELATIONS OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PBLSH"; _var.dtworkRow[1] = "PUBLISHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PUDLR"; _var.dtworkRow[1] = "PUDDLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PMPMN"; _var.dtworkRow[1] = "PUMPMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "PYRTC"; _var.dtworkRow[1] = "PYROTECHNIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "QTINS"; _var.dtworkRow[1] = "QUALITY CONTROL INSPECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "QTMGR"; _var.dtworkRow[1] = "QUALITY CONTROL MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "QTCTS"; _var.dtworkRow[1] = "QUALITY CONTROL STAFF"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "QTCTR"; _var.dtworkRow[1] = "QUALITY CONTROLLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "QARRY"; _var.dtworkRow[1] = "QUARRY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RACER"; _var.dtworkRow[1] = "RACER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDNCR"; _var.dtworkRow[1] = "RADIO ANNOUNCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDPRT"; _var.dtworkRow[1] = "RADIO OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDLGS"; _var.dtworkRow[1] = "RADIOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDTCH"; _var.dtworkRow[1] = "RADIOLOGY TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDTHR"; _var.dtworkRow[1] = "RADIOTHERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RFTSM"; _var.dtworkRow[1] = "RAFTSMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RLRDW"; _var.dtworkRow[1] = "RAILROAD WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RNCHR"; _var.dtworkRow[1] = "RANCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RLTR"; _var.dtworkRow[1] = "REALTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RCPTN"; _var.dtworkRow[1] = "RECEPTIONIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RCRDR"; _var.dtworkRow[1] = "RECORDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RCRTN"; _var.dtworkRow[1] = "RECREATIONAL THERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RECTR"; _var.dtworkRow[1] = "RECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RFREE"; _var.dtworkRow[1] = "REFEREE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RFLXL"; _var.dtworkRow[1] = "REFLEXOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RGSTR"; _var.dtworkRow[1] = "REGISTRAR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RLGSW"; _var.dtworkRow[1] = "RELIGIOUS WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RPRMN"; _var.dtworkRow[1] = "REPAIRMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RPRTR"; _var.dtworkRow[1] = "REPORTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RSCWR"; _var.dtworkRow[1] = "RESCUE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RSRCH"; _var.dtworkRow[1] = "RESEARCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RSTRT"; _var.dtworkRow[1] = "RESTAURANTEUR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RETLR"; _var.dtworkRow[1] = "RETAILER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RETRD"; _var.dtworkRow[1] = "RETIRED"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RIGGR"; _var.dtworkRow[1] = "RIGGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RIVTR"; _var.dtworkRow[1] = "RIVETER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDMNT"; _var.dtworkRow[1] = "ROAD MAINTENANCE WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RDTST"; _var.dtworkRow[1] = "ROAD TESTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RODMN"; _var.dtworkRow[1] = "RODMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ROOFR"; _var.dtworkRow[1] = "ROOFER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RBCUR"; _var.dtworkRow[1] = "RUBBER CURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RBPRO"; _var.dtworkRow[1] = "RUBBER PROOFER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RBSTR"; _var.dtworkRow[1] = "RUBBER STRAINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "RBTCH"; _var.dtworkRow[1] = "RUBBER TECHNOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SDDLM"; _var.dtworkRow[1] = "SADDLE MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SFTYF"; _var.dtworkRow[1] = "SAFETY OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SAILR"; _var.dtworkRow[1] = "SAILOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SLAST"; _var.dtworkRow[1] = "SALES ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SLSRP"; _var.dtworkRow[1] = "SALES REPRESENTATIVE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SLSMN"; _var.dtworkRow[1] = "SALESMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SNDBL"; _var.dtworkRow[1] = "SANDBLASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SWYER"; _var.dtworkRow[1] = "SAWYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCFFL"; _var.dtworkRow[1] = "SCAFFOLDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCNTS"; _var.dtworkRow[1] = "SCIENTIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCRWR"; _var.dtworkRow[1] = "SCRIPTWRITER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCLPT"; _var.dtworkRow[1] = "SCULPTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SFRRN"; _var.dtworkRow[1] = "SEAFARER - NON-OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SFRRF"; _var.dtworkRow[1] = "SEAFARER - OFFICER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SFRRT"; _var.dtworkRow[1] = "SEAFARER - OTHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SMSTR"; _var.dtworkRow[1] = "SEAMSTRESS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCRTR"; _var.dtworkRow[1] = "SECRETARY"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCNSL"; _var.dtworkRow[1] = "SECURITY CONSULTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCRTY"; _var.dtworkRow[1] = "SECURITY GUARD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SMNRR"; _var.dtworkRow[1] = "SEMINAR ORGANIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SRVCC"; _var.dtworkRow[1] = "SERVICE CREW"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHPSH"; _var.dtworkRow[1] = "SHEEP SHEARER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHTFX"; _var.dtworkRow[1] = "SHEET FIXER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHTMT"; _var.dtworkRow[1] = "SHEET METAL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHPHR"; _var.dtworkRow[1] = "SHEPHERD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHPRP"; _var.dtworkRow[1] = "SHIP REPAIRER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHMKR"; _var.dtworkRow[1] = "SHOEMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHRPT"; _var.dtworkRow[1] = "SHORE PATROL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHTBL"; _var.dtworkRow[1] = "SHOTBLASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SHTFR"; _var.dtworkRow[1] = "SHOTFIRER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SGNLM"; _var.dtworkRow[1] = "SIGNALMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SINGR"; _var.dtworkRow[1] = "SINGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SKTNG"; _var.dtworkRow[1] = "SKATING RINK PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SKPPR"; _var.dtworkRow[1] = "SKIPPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SNKCT"; _var.dtworkRow[1] = "SNAKE CATCHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SCLWR"; _var.dtworkRow[1] = "SOCIAL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SLDRR"; _var.dtworkRow[1] = "SOLDERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SLCTR"; _var.dtworkRow[1] = "SOLICITOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SORTR"; _var.dtworkRow[1] = "SORTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SNDTC"; _var.dtworkRow[1] = "SOUND TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPCLC"; _var.dtworkRow[1] = "SPECIAL ACTION FORCES"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPCLS"; _var.dtworkRow[1] = "SPECIALIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPTHR"; _var.dtworkRow[1] = "SPEECH THERAPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPAST"; _var.dtworkRow[1] = "SPORTS ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPVTN"; _var.dtworkRow[1] = "SPORTS VENUE ATTENDANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPRYR"; _var.dtworkRow[1] = "SPRAYER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STBLM"; _var.dtworkRow[1] = "STABLEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STGHN"; _var.dtworkRow[1] = "STAGE HAND"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STGMN"; _var.dtworkRow[1] = "STAGE MANAGER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STNR"; _var.dtworkRow[1] = "STAINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STMPR"; _var.dtworkRow[1] = "STAMPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STTCN"; _var.dtworkRow[1] = "STATISTICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STMFT"; _var.dtworkRow[1] = "STEAMFITTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STLWR"; _var.dtworkRow[1] = "STEEL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STPLJ"; _var.dtworkRow[1] = "STEEPLEJACK"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STNCL"; _var.dtworkRow[1] = "STENCIL CUTTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STNGR"; _var.dtworkRow[1] = "STENOGRAPHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STRTY"; _var.dtworkRow[1] = "STEREOTYPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STVDR"; _var.dtworkRow[1] = "STEVEDORE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STCKB"; _var.dtworkRow[1] = "STOCKBROKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STRKP"; _var.dtworkRow[1] = "STOREKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STRTS"; _var.dtworkRow[1] = "STREET SWEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STDNT"; _var.dtworkRow[1] = "STUDENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STDNC"; _var.dtworkRow[1] = "STUDIO ANNOUNCER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "STNTM"; _var.dtworkRow[1] = "STUNTMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPRNT"; _var.dtworkRow[1] = "SUPERINTENDENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPRVS"; _var.dtworkRow[1] = "SUPERVISOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SPPLR"; _var.dtworkRow[1] = "SUPPLIER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SRGN"; _var.dtworkRow[1] = "SURGEON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SRVYR"; _var.dtworkRow[1] = "SURVEYOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SWPR"; _var.dtworkRow[1] = "SWEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "SWMMN"; _var.dtworkRow[1] = "SWIMMING POOL PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TAILR"; _var.dtworkRow[1] = "TAILOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLNTS"; _var.dtworkRow[1] = "TALENT SCOUT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLLYM"; _var.dtworkRow[1] = "TALLYMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TTTST"; _var.dtworkRow[1] = "TATTOOIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TXCNS"; _var.dtworkRow[1] = "TAX CONSULTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TXDRM"; _var.dtworkRow[1] = "TAXIDERMIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TECHR"; _var.dtworkRow[1] = "TEACHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TCHRD"; _var.dtworkRow[1] = "TEACHER AIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TMLDR"; _var.dtworkRow[1] = "TEAM LEADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TCAST"; _var.dtworkRow[1] = "TECHINICAL ASSISTANT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TCNTR"; _var.dtworkRow[1] = "TECHNICAL CONTROLLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TCHNC"; _var.dtworkRow[1] = "TECHNICIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TCHNL"; _var.dtworkRow[1] = "TECHNOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLPHN"; _var.dtworkRow[1] = "TELEPHONE OPERATOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TELLR"; _var.dtworkRow[1] = "TELLER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TELRR"; _var.dtworkRow[1] = "TELLER - ROVING"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRRZZ"; _var.dtworkRow[1] = "TERRAZZO WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TSTER"; _var.dtworkRow[1] = "TESTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "THMPR"; _var.dtworkRow[1] = "THEME PARK WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TILMK"; _var.dtworkRow[1] = "TILE MAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TILER"; _var.dtworkRow[1] = "TILER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLLMN"; _var.dtworkRow[1] = "TILLMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TMBRM"; _var.dtworkRow[1] = "TIMBERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TMKPR"; _var.dtworkRow[1] = "TIMEKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLLCL"; _var.dtworkRow[1] = "TOLL COLLECTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TLKPR"; _var.dtworkRow[1] = "TOOLKEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TOLMK"; _var.dtworkRow[1] = "TOOLMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TORGD"; _var.dtworkRow[1] = "TOUR GUIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TXCLG"; _var.dtworkRow[1] = "TOXICOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRCER"; _var.dtworkRow[1] = "TRACER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRADR"; _var.dtworkRow[1] = "TRADER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRFFC"; _var.dtworkRow[1] = "TRAFFIC AIDE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRNCN"; _var.dtworkRow[1] = "TRAIN CONDUCTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRANR"; _var.dtworkRow[1] = "TRAINER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRNSM"; _var.dtworkRow[1] = "TRANSMITTER WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRVLG"; _var.dtworkRow[1] = "TRAVEL AGENT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRSRR"; _var.dtworkRow[1] = "TREASURER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRSRG"; _var.dtworkRow[1] = "TREE SURGEON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRBLM"; _var.dtworkRow[1] = "TROUBLEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TRCKR"; _var.dtworkRow[1] = "TRUCKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TNNLW"; _var.dtworkRow[1] = "TUNNEL WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TUTOR"; _var.dtworkRow[1] = "TUTOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TYPST"; _var.dtworkRow[1] = "TYPIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TYRML"; _var.dtworkRow[1] = "TYRE MOULDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "TYRRP"; _var.dtworkRow[1] = "TYRE REPAIRER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "UNDWR"; _var.dtworkRow[1] = "UNDERWRITER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "UNMPL"; _var.dtworkRow[1] = "UNEMPLOYED"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "UNKNW"; _var.dtworkRow[1] = "UNKNOWN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "UPHLS"; _var.dtworkRow[1] = "UPHOLSTERER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "URBPL"; _var.dtworkRow[1] = "URBAN PLANNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "USHER"; _var.dtworkRow[1] = "USHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ULTYM"; _var.dtworkRow[1] = "UTILITY MAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VALET"; _var.dtworkRow[1] = "VALET"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VTCLN"; _var.dtworkRow[1] = "VAT CLEANER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VENDR"; _var.dtworkRow[1] = "VENDOR"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VTRNR"; _var.dtworkRow[1] = "VETERINARIAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VGNRN"; _var.dtworkRow[1] = "VIGNERON"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VVSCT"; _var.dtworkRow[1] = "VIVISECTIONIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VCLST"; _var.dtworkRow[1] = "VOCALIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VLCNL"; _var.dtworkRow[1] = "VOLCANOLOGIST"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "VLCNZ"; _var.dtworkRow[1] = "VULCANIZER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WAITR"; _var.dtworkRow[1] = "WAITER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WATRS"; _var.dtworkRow[1] = "WAITRESS"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WARDN"; _var.dtworkRow[1] = "WARDEN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRDRB"; _var.dtworkRow[1] = "WARDROBE PERSONNEL"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRHSM"; _var.dtworkRow[1] = "WAREHOUSEMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WASHR"; _var.dtworkRow[1] = "WASHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WTCMK"; _var.dtworkRow[1] = "WATCHMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WTCMN"; _var.dtworkRow[1] = "WATCHMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WTRMS"; _var.dtworkRow[1] = "WATERMASTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WTHRM"; _var.dtworkRow[1] = "WEATHERMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WEAVR"; _var.dtworkRow[1] = "WEAVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WGHER"; _var.dtworkRow[1] = "WEIGHER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WELDR"; _var.dtworkRow[1] = "WELDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WHLWR"; _var.dtworkRow[1] = "WHEELWRIGHT"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WNDWC"; _var.dtworkRow[1] = "WINDOW CLEANER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WNDWF"; _var.dtworkRow[1] = "WINDOW FITTER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WNMKR"; _var.dtworkRow[1] = "WINEMAKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRWND"; _var.dtworkRow[1] = "WIRE WINDER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRMNH"; _var.dtworkRow[1] = "WIREMAN - HOUSE"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRMNS"; _var.dtworkRow[1] = "WIREMAN - SWITCHBOARD"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WDCRV"; _var.dtworkRow[1] = "WOOD CARVER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WDWRK"; _var.dtworkRow[1] = "WOOD WORKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WLCLS"; _var.dtworkRow[1] = "WOOL CLASSER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRCKR"; _var.dtworkRow[1] = "WRECKER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "WRTER"; _var.dtworkRow[1] = "WRITER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "YRDMN"; _var.dtworkRow[1] = "YARDMAN"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "YRNSP"; _var.dtworkRow[1] = "YARN SPINNER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);
            _var.dtworkRow = _var.objdt_OCCCODE.NewRow(); _var.dtworkRow[0] = "ZOKPR"; _var.dtworkRow[1] = "ZOO KEEPER"; _var.objdt_OCCCODE.Rows.Add(_var.dtworkRow);

            return _var.objdt_OCCCODE;
        }
    }

    static class SqlStyleExtensions
    {
        public static bool In(this string me, params string[] set)
        {
            return set.Contains(me);
        }
    }

}
