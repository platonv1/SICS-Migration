using System;
using System.Data;
using System.Linq;
using System.Globalization;
namespace Bordereaux_SICS_Mapping.BAL
{
    class BM053
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false, string str_macro = "")
        {
            _Global _var = new _Global();
            Helper objHlpr = new Helper();
            HelperV21 objHlpr2 = new HelperV21();
            DataTable objdt_template = new DataTable();

            DataTable dt_OCC = new DataTable();
            dt_OCC = objHlpr.fn_LoadOCCCode();

            objdt_template = objHlpr.dt_formtemplate(str_sheet);

            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(str_raw);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets[str_sheet];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;

            int erawrow = rawrange.Rows.Count;
            int int_RowCnt = 0;

            double dbl_rate = 0.09;
            int intLoop = 1;
            try
            {
                for (intLoop = 1; intLoop <= erawrow + 1; intLoop++)
                {
                    string str_PolNum = wsraw.Cells[intLoop, 2].Text.ToString();

                    if (!objHlpr.fn_policyNumChecker(str_PolNum, wsraw.Cells[intLoop, 3].Text.ToString(), wsraw.Cells[intLoop, 4].Text.ToString(), wsraw.Cells[intLoop, 5].Text.ToString()))
                    {
                        continue;
                    }


                    string str_bmyear = wsraw.Cells[4, 2].Text.ToString().Substring(wsraw.Cells[4, 2].Text.ToString().Length - 4, 4);
                    if (!int.TryParse(str_bmyear, out int int_bmyear))
                    {
                        int_bmyear = 0;
                    }

                    string str_PlanCode = wsraw.Cells[intLoop, 4].Text.ToString();
                    string str_PolicyTerm = wsraw.Cells[intLoop, 6].Text.ToString();

                    string str_DOB = wsraw.Cells[intLoop, 7].Text.ToString();
                    string str_DOBYear = str_DOB.Substring(str_DOB.Length - 4, 4);
                    if (!int.TryParse(str_DOBYear, out int int_DOBYear))
                    {
                        int_DOBYear = 0;
                    }

                    string str_Sex = wsraw.Cells[intLoop, 8].Text.ToString();
                    string str_Smoker = wsraw.Cells[intLoop, 9].Text.ToString();
                    string str_Occupation_Code = wsraw.Cells[intLoop, 10].Text.ToString(); 
                    string str_Mortality = wsraw.Cells[intLoop, 11].Text.ToString();

                    string str_IssueDate = wsraw.Cells[intLoop, 18].Text.ToString();
                    string str_IssueYear = str_IssueDate.Substring(str_IssueDate.Length - 4, 4);

                    switch (str_IssueYear)
                    {
                        case "2016":
                            dbl_rate = 0.07;
                            break;
                        case "2017":
                            dbl_rate = 0.08;
                            break;
                        default:
                            dbl_rate = 0.09;
                            break;
                    }

                    if (!int.TryParse(str_IssueYear, out int int_IssueYear))
                    {
                        int_IssueYear = 0;
                    }

                    string str_PremDueDate = str_IssueDate.Substring(0, str_IssueDate.Length - 4) + str_bmyear;
                    

                    string str_OSA = wsraw.Cells[intLoop, 21].Text.ToString();
                    string str_CAV = wsraw.Cells[intLoop, 23].Text.ToString();
                    string str_ISAR = wsraw.Cells[intLoop, 31].Text.ToString();
                    string str_SAR = wsraw.Cells[intLoop, 31].Text.ToString();
                    string str_Prem = wsraw.Cells[intLoop, 33].Text.ToString();
                    string str_Comm = wsraw.Cells[intLoop, 34].Text.ToString();
                    string str_Status = wsraw.Cells[intLoop, 35].Text.ToString();

                    if (double.TryParse(str_OSA, out double dbl_OSA))
                    {
                        dbl_OSA = dbl_OSA * dbl_rate;
                    }
                    else
                    {
                        dbl_OSA = 1;
                    }

                    if (double.TryParse(str_CAV, out double dbl_CAV))
                    {
                        dbl_CAV = dbl_CAV * dbl_rate;
                    }
                    else
                    {
                        dbl_CAV = 0;
                    }

                    if (double.TryParse(str_ISAR, out double dbl_ISAR))
                    {
                        dbl_ISAR = dbl_ISAR * dbl_rate;
                    }
                    else
                    {
                        dbl_ISAR = 1;
                    }

                    if (double.TryParse(str_SAR, out double dbl_SAR))
                    {
                        dbl_SAR = dbl_SAR * dbl_rate;
                    }
                    else
                    {
                        dbl_SAR = 1;
                    }

                    if (!double.TryParse(str_Prem.Replace("(","-").Replace(")", ""), out double dbl_Prem))
                    {
                        dbl_Prem = 0;
                    }

                    if (!double.TryParse(str_Comm.Replace("(", "-").Replace(")", ""), out double dbl_Comm))
                    {
                        dbl_Comm = 0;
                    }

                    string str_Fullname = wsraw.Cells[intLoop, 36].Text.ToString();

                    if (str_PremDueDate.Contains("02/29"))
                    {
                        str_PremDueDate = str_PremDueDate.Replace("/29", "/28");
                    }

                    DateTime dt_PremiumDate = Convert.ToDateTime(str_PremDueDate);
                    DateTime dt_IssueDate = Convert.ToDateTime(str_IssueDate);

                    _var.dtworkRow01 = objdt_template.NewRow();

                    string str_tcode = "";
                    if (str_Status.ToUpper().Contains("SURRENDERED"))
                    {
                        str_tcode = "TFULLSUR";
                    }
                    else if (str_Status.ToUpper().Contains("LAPSE"))
                    {
                        str_tcode = "TLAPSE";
                    }
                    else if (str_Status.ToUpper().Contains("REINSTATEMENT"))
                    {
                        str_tcode = "TREINS";
                    }
                    else if (str_Status.ToUpper().Contains("FULLY PAID-UP"))
                    {
                        str_tcode = "TFULLPU";
                    }
                    else
                    {
                        str_tcode = "TRENEW";
                    }
                    
                    _var.dtworkRow01[0] = "'" + str_PolNum.ToString();
                    _var.dtworkRow01[5] = str_PlanCode;
                    _var.dtworkRow01[8] = "QA";
                    _var.dtworkRow01[9] = "PA";
                    _var.dtworkRow01[13] = "IND";
                    _var.dtworkRow01[10] = "Q";
                    _var.dtworkRow01[14] = "T";
                    _var.dtworkRow01[23] = "PHP";
                    _var.dtworkRow01[24] = "MLY";
                    _var.dtworkRow01[29] = "NATREID";

                    DateTime dt_DOB = Convert.ToDateTime(str_DOB);
                    
                    _var.dtworkRow01[78] = Math.Round((dt_IssueDate - dt_DOB).TotalDays / 365);
                    _var.dtworkRow01[79] = (int_IssueYear - int_DOBYear).ToString();

                    DataRow[] foundRows = dt_OCC.Select("_NAME = " + "'" + str_Occupation_Code + "'");
                    if (foundRows.Length != 0)
                    {
                        _var.dtworkRow01[46] = foundRows[0][0].ToString();
                    }
                    else { _var.dtworkRow01[46] = "NONE"; }

                    if (dbl_CAV != 0)
                    {
                        _var.dtworkRow01[76] = "CAV: " + dbl_CAV.ToString() + " ";
                    }

                    //DOB
                    if (String.IsNullOrEmpty(str_DOB))
                    {
                        str_DOB = "07/01/1900";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR4AL" : _var.dtworkRow01[76].ToString() + "|BR4AL";
                    }
                    _var.dtworkRow01[37] = str_DOB;

                    //Smoker
                    _var.dtworkRow01[38] = objHlpr.fn_smokercode(str_Smoker, "053");

                    //Mortality
                    _var.dtworkRow01[39] = objHlpr.fn_getmortality(str_Mortality);
                    if (objHlpr.fn_isDMort(_var.dtworkRow01[39].ToString()))
                    {
                        _var.dtworkRow01[39] = "STANDARD";
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR8AN" : _var.dtworkRow01[76].ToString() + "|BR8AN";
                    }

                    _var.dtworkRow01[41] = str_bmyear;

                    _var.dtworkRow01[20] = str_IssueDate;
                    _var.dtworkRow01[22] = str_PremDueDate;

                    if (str_tcode == "TRENEW")
                    {
                        _var.dtworkRow01[19] = _var.dtworkRow01[22];
                        _var.dtworkRow01[58] = "4001";
                        _var.dtworkRow01[59] = dbl_Prem;
                    }
                    else
                    {   if (str_tcode == "TFULLPU")
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[22];
                            _var.dtworkRow01[62] = "4004";
                            _var.dtworkRow01[63] = dbl_Prem;
                        }
                        else if ((dt_PremiumDate - dt_IssueDate).TotalDays <= 365)
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[20];
                            _var.dtworkRow01[60] = "4002";
                            _var.dtworkRow01[61] = dbl_Prem;
                        }
                        else
                        {
                            _var.dtworkRow01[19] = _var.dtworkRow01[22];
                            _var.dtworkRow01[62] = "4004";
                            _var.dtworkRow01[63] = dbl_Prem;
                        }
                    }

                    _var.dtworkRow01[21] = str_tcode;

                    _var.dtworkRow01[25] = dbl_OSA;
                    _var.dtworkRow01[27] = dbl_ISAR;
                    _var.dtworkRow01[77] = dbl_SAR;
                    _var.dtworkRow01[28] = "1";
                    if (str_Comm != string.Empty)
                    {
                        _var.dtworkRow01[66] = "5005";
                        _var.dtworkRow01[67] = dbl_Comm;
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[77] = _var.dtworkRow01[27];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-1BZ" : _var.dtworkRow01[76].ToString() + "|BR1-1BZ";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()))
                    {
                        _var.dtworkRow01[75] = _var.dtworkRow01[25];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR1-2BZ" : _var.dtworkRow01[76].ToString() + "|BR1-2BZ";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[77];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-1AB" : _var.dtworkRow01[76].ToString() + "|BR2-1AB";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[25].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[27].ToString()))
                    {
                        _var.dtworkRow01[27] = _var.dtworkRow01[25];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR2-2AB" : _var.dtworkRow01[76].ToString() + "|BR2-2AB";
                    }

                    if (!String.IsNullOrEmpty(_var.dtworkRow01[27].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[27];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-1Z" : _var.dtworkRow01[76].ToString() + "|BR3-1Z";
                    }
                    else if (!String.IsNullOrEmpty(_var.dtworkRow01[77].ToString())
                        &&
                        String.IsNullOrEmpty(_var.dtworkRow01[25].ToString()))
                    {
                        _var.dtworkRow01[25] = _var.dtworkRow01[77];
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR3-2Z" : _var.dtworkRow01[76].ToString() + "|BR3-2Z";
                    }

                    //Name
                    if (String.IsNullOrEmpty(str_Fullname))
                    {
                        str_Fullname = str_PolNum;
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR6AF" : _var.dtworkRow01[76].ToString() + "|BR6AF";
                    }

                    objHlpr.fn_getnamesandlifeID(str_Fullname, str_DOB, out string str_outfname, out string str_outlname, out string str_outlifeid);

                    string str_MI = string.Empty;
                    string [] arr_fullname;
                    arr_fullname = str_Fullname.Split(',');
                    str_outlname = arr_fullname [0];

                    if(arr_fullname.Count() > 1)
                    {
                        str_outfname = arr_fullname [1];
                    }

                    if(arr_fullname.Count() > 2)
                    {
                        str_MI = arr_fullname [2];
                        _var.dtworkRow01 [34] = str_MI;
                    }


                    _var.dtworkRow01 [30] = str_outlifeid;
                    _var.dtworkRow01 [31] = objHlpr.fn_stringcleanup(str_Fullname);
                    _var.dtworkRow01 [32] = str_outlname.Trim();
                    _var.dtworkRow01 [33] = str_outfname.Replace(" " + str_MI, string.Empty).Trim();

                    //objHlpr2.fn_separateLastNameFirstNameV5(str_Fullname,out string strLastname,out string strFirstname,out string  strMiddlename);
                    ////string str_MI = objHlpr.fn_getMI(strFirstname);
                    //_var.dtworkRow01 [34] = strMiddlename;
                    //_var.dtworkRow01 [31] = objHlpr.fn_stringcleanup(str_Fullname);
                    //_var.dtworkRow01 [32] = strLastname;
                    //_var.dtworkRow01 [33] = strFirstname;
                    //_var.dtworkRow01 [30] = objHlpr.fn_LifeID(str_outfname, str_outlname, str_DOB);


                    DateTime oDate = DateTime.ParseExact(str_DOB, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                    string str_formatLname = (_var.dtworkRow01[32].ToString().Length >= 5) ? _var.dtworkRow01[32].ToString().Substring(0,5) : _var.dtworkRow01[32].ToString();
                    string str_formatFname = (_var.dtworkRow01[33].ToString().Length >= 2) ? _var.dtworkRow01[33].ToString().Substring(0,2) : _var.dtworkRow01[33].ToString();
                    //str_outlifeid = str_formatLname
                    //    + str_formatFname
                    //    + oDate.Month.ToString().PadLeft(2, '0') + oDate.Day.ToString().PadLeft(2, '0') + oDate.Year;
                    //_var.dtworkRow01[30] = str_outlifeid;

                    //Gender
                    if (!String.IsNullOrEmpty(str_Sex))
                    {
                        _var.dtworkRow01[36] = (str_Sex.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && !String.IsNullOrEmpty(str_gender))
                    {
                        str_Sex = objHlpr.fn_getgender(str_gender, _var.dtworkRow01[33].ToString());
                        _var.dtworkRow01[36] = str_Sex;
                        _var.dtworkRow01[76] = String.IsNullOrEmpty(_var.dtworkRow01[76].ToString()) ? "BR7AK" : _var.dtworkRow01[76].ToString() + "|BR7AK";
                    }
                    else if (String.IsNullOrEmpty(str_Sex) && String.IsNullOrEmpty(str_gender))
                    {
                        _var.dtworkRow01[36] = string.Empty;
                    }

                    if (String.IsNullOrEmpty(_var.dtworkRow01[36].ToString()))
                    {
                        _var.boo_genderfail = true;
                    }

                    _var.dtworkRow01[7] = string.Empty;

                    _var.dbl_comm += decimal.Parse(dbl_Comm.ToString());
                    _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[57].ToString()) ? "0" : _var.dtworkRow01[57].ToString()
                            );
                    _var.dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[59].ToString()) ? "0" : _var.dtworkRow01[59].ToString()
                            );
                    _var.dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[61].ToString()) ? "0" : _var.dtworkRow01[61].ToString()
                            );
                    _var.dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[63].ToString()) ? "0" : _var.dtworkRow01[63].ToString()
                            );
                    _var.dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow01[77].ToString()) ? "0" : _var.dtworkRow01[77].ToString()
                            );

                    objdt_template.Rows.Add(_var.dtworkRow01);

                }

                _var.dtworkRow01 = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Premium:";
                _var.dtworkRow01[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Sum at Risk:";
                _var.dtworkRow01[1] = _var.dbl_BZ;
                objdt_template.Rows.Add(_var.dtworkRow01);

                _var.dtworkRow01 = objdt_template.NewRow();
                _var.dtworkRow01[0] = "Total Commission:";
                _var.dtworkRow01[1] = _var.dbl_comm;
                objdt_template.Rows.Add(_var.dtworkRow01);

                if (_var.boo_genderfail)
                {
                    _var.dtworkRow01 = objdt_template.NewRow();
                    _var.dtworkRow01[0] = "Please check for blank genders";
                    objdt_template.Rows.Add(_var.dtworkRow01);
                }

                string despath = str_saved + @"\BM053-" + str_savef + ".xlsx";
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
                _var.dtworkRow01 = null;
                _var.dtworkRow02 = null;
                _var.dtworkRow03 = null;
                _var.dtworkRow04 = null;
                objdt_template.Dispose();
                objdt_template = null;
                objHlpr.fn_killexcel();
                objHlpr = null;
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message + Environment.NewLine + " *****On excel row line: " + intLoop + " *****";
            }
        }


    }
}