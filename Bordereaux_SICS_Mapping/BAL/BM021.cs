using System;
using System.Data;
using System.Linq;
using System.Globalization;

namespace Bordereaux_SICS_Mapping.BAL
{
    class BM021 
    {
        public string fn_process(string str_raw, string str_sheet, string str_saved, string str_savef, string str_gender = "", bool boo_open = false, bool boo_clean = false)
        {

            int rowcount = 1;

            try
            {
                _Global _var = new _Global();
                Helper objHlpr = new Helper();
                DataTable objdt_template = new DataTable();

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

                string polnum = wsraw.Cells[prawrow, 1].Text.ToString();    
                string branded = wsraw.Cells[prawrow, 10].Text.ToString();
                string reins = wsraw.Cells[prawrow, 4].Text.ToString();
                string sum = wsraw.Cells[prawrow, 11].Text.ToString();
                string risk = wsraw.Cells[prawrow, 16].Text.ToString();
                string retention = wsraw.Cells[prawrow, 14].Text.ToString();
                string fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                string gender = wsraw.Cells[prawrow, 8].Text.ToString();
                string dob = wsraw.Cells[prawrow, 6].Text.ToString();
                string pref = wsraw.Cells[prawrow, 13].Text.ToString();
                string premium = wsraw.Cells[prawrow, 18].Text.ToString();
                string issue = wsraw.Cells[prawrow, 7].Text.ToString(); //Attained Age
                string fyprem = wsraw.Cells[prawrow, 27].Text.ToString();
                string ryprem = wsraw.Cells[prawrow, 35].Text.ToString();
                string comprem = wsraw.Cells[prawrow, 43].Text.ToString();
                string facul = wsraw.Cells[prawrow, 31].Text.ToString();
                string rfacul = wsraw.Cells[prawrow, 39].Text.ToString();
                string cfacul = wsraw.Cells[prawrow, 47].Text.ToString();
                string risk1 = wsraw.Cells[prawrow, 17].Text.ToString();
                string paid = wsraw.Cells[prawrow, 5].Text.ToString();

                string TRANCODE = string.Empty;
                string[] comparestring = new string[] { "" };
                bool findboo = false;

                int storee;
                bool chck;
                decimal classific;

                #region Data Processing
                while (rowcount != erawrow + 2)
                {
                    chck = int.TryParse(polnum, out storee);
                    polnum = objHlpr.fn_stringcleanup(polnum);
                    if (polnum != string.Empty && chck == false)
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
                        _var.dtworkRow = objdt_template.NewRow();
                        _var.dtworkRow[0] = polnum.ToString();
                        _var.dtworkRow[1] = polnum.ToString();
                        _var.dtworkRow[5] = branded.ToString();
                        //_var.dtworkRow[3] = "DEATH";
                        //_var.dtworkRow[4] = "TRADITIONALLIFE";
                        _var.dtworkRow[8] = "SURPLUS";
                        _var.dtworkRow[9] = "PAFM";
                        _var.dtworkRow[10] = "S";
                        _var.dtworkRow[13] = "IND";
                        _var.dtworkRow[23] = "PHP";
                        _var.dtworkRow[24] = "YLY";
                        _var.dtworkRow[19] = reins.ToString();
                        _var.dtworkRow[20] = reins.ToString();
                        _var.dtworkRow[22] = paid.ToString();
                        _var.dtworkRow[25] = sum.ToString();
                        _var.dtworkRow[77] = sum.ToString();
                
                        if ((risk.ToString() == "0") || (risk.Trim() == "-") ||  (risk == String.Empty))
                        {
                            _var.dtworkRow[77] = "1";
                        }
                        else
                        {
                            _var.dtworkRow[77] = risk.ToString();
                        }

                        if ((risk1.ToString() == "0") || (risk1.Trim() == "-") || (risk1 == String.Empty))
                        {
                            _var.dtworkRow[27] = "1";
                        }
                        else
                        {
                            _var.dtworkRow[27] = risk1.ToString();
                        }

                        double d_out = 0;
                        if (!double.TryParse(retention, out d_out))
                        {
                            retention = "0";
                        }

                        _var.dtworkRow[28] = retention.ToString();
                        _var.dtworkRow[29] = "NATREID";
                        _var.dtworkRow[36] = gender.ToString();

                        #region "New Requirements - No DOB"
                        if (String.IsNullOrEmpty(dob))
                        {
                            //ISSUE#009-Start---------
                            dob = "07/01/1900";
                            //ISSUE#009-End-----------

                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR4AL" : _var.dtworkRow[76].ToString() + "|BR4AL";
                        }
                        #endregion
                        _var.dtworkRow[37] = dob;

                        _var.dtworkRow[79] = issue.ToString();

                        string fac = "T";
                        string fac1 = "F";

                        if ((facul == String.Empty) && (rfacul == String.Empty) && (cfacul == String.Empty))
                        {
                            _var.dtworkRow[14] = fac.ToString();
                            _var.dtworkRow[83] = "NR";
                        }
                        else
                        {
                            _var.dtworkRow[14] = fac1.ToString();
                            _var.dtworkRow[83] = "NR";
                        }

                        //ISSUE# Bug on mortality-Start---------
                        _var.dtworkRow[39] = objHlpr.fn_getmortality(pref);
                        if (objHlpr.fn_isDMort(_var.dtworkRow[39].ToString()))
                        {
                            _var.dtworkRow[39] = "STANDARD";
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR8AN" : _var.dtworkRow[76].ToString() + "|BR8AN";
                        }
                        //ISSUE# Bug on mortality-End-----------

                        fyprem = objHlpr.fn_numbercleanup_negative(fyprem); 
                        ryprem = objHlpr.fn_numbercleanup_negative(ryprem); 
                        comprem = objHlpr.fn_numbercleanup_negative(comprem); 
                        facul = objHlpr.fn_numbercleanup_negative(facul); 
                        rfacul = objHlpr.fn_numbercleanup_negative(rfacul); 

                        if (polnum != string.Empty)
                        {
                            findboo = false;

                            comparestring = new string[] { "REINSTATEMENT", "REINSTATED" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Contains(s))
                                {
                                    case true:
                                        _var.dtworkRow[21] = "TREINS";
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
                                        _var.dtworkRow[21] = "TCONTER";
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
                                        _var.dtworkRow[21] = "TCANCINC";
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
                                        _var.dtworkRow[21] = "TEXPIRY";
                                        findboo = true;
                                        break;
                                }
                            }
                            comparestring = new string[] { "EXTENDED TERM", "ETI" };
                            foreach (string s in comparestring)
                            {
                                switch (polnum.Contains(s))
                                {
                                    case true:
                                        _var.dtworkRow[21] = "TEXTTER";
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
                                        _var.dtworkRow[21] = "TFULLMAT";
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
                                        _var.dtworkRow[21] = "TFULLPU";
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
                                        _var.dtworkRow[21] = "TFULLREC";
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
                                        _var.dtworkRow[21] = "TFULLSUR";
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
                                        _var.dtworkRow[21] = "TLAPSE";
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
                                        _var.dtworkRow[21] = "ADJUST";
                                        findboo = true;
                                        break;
                                }
                            }

                            if (!findboo)
                            {
                                _var.dtworkRow[21] = TRANCODE;
                            }
                        }
                        else
                        {
                            _var.dtworkRow[21] = TRANCODE;
                        }
                        _var.dtworkRow[80] = "0";
                        if ((comprem != String.Empty) && (fyprem != "0") && (TRANCODE != "ADJUST"))
                        { 
                            _var.dtworkRow[21] = "TNEWBUS";
                            _var.dtworkRow[57] = decimal.Parse(String.IsNullOrEmpty(fyprem.ToString()) ? "0" : fyprem.ToString()) -
                                 decimal.Parse(String.IsNullOrEmpty(comprem.ToString()) ? "0" : comprem.ToString());
                            _var.dtworkRow[80] = comprem.ToString();
                            _var.dtworkRow[56] = "4000";
                        }
                        else if (TRANCODE.ToUpper().Contains("TLAPSE"))
                        {
                            if (fyprem != String.Empty)
                            {
                                _var.dtworkRow[60] = "4002";
                                _var.dtworkRow[61] = decimal.Parse(String.IsNullOrEmpty(fyprem.ToString()) ? "0" : fyprem.ToString());
                            }
                            else if (ryprem != String.Empty)
                            {
                                _var.dtworkRow[62] = "4004";
                                _var.dtworkRow[63] = decimal.Parse(String.IsNullOrEmpty(ryprem.ToString()) ? "0" : ryprem.ToString());
                            }
                        }
                        else if (TRANCODE.ToUpper().Contains("ADJUST"))
                        {
                            if (fyprem != String.Empty)
                            {
                                _var.dtworkRow[60] = "4002";
                                _var.dtworkRow[61] = decimal.Parse(String.IsNullOrEmpty(fyprem.ToString()) ? "0" : fyprem.ToString());
                            }
                            else if (ryprem != String.Empty)
                            {
                                _var.dtworkRow[62] = "4004";
                                _var.dtworkRow[63] = decimal.Parse(String.IsNullOrEmpty(ryprem.ToString()) ? "0" : ryprem.ToString());
                            }
                        }
                        else if (ryprem != String.Empty)
                        {
                            _var.dtworkRow[21] = "TRENEW";
                            _var.dtworkRow[59] = decimal.Parse(String.IsNullOrEmpty(ryprem.ToString()) ? "0" : ryprem.ToString());
                            _var.dtworkRow[58] = "4001";
                        }
                        else if (fyprem != String.Empty)
                        {
                            _var.dtworkRow[21] = "TNEWBUS";
                            _var.dtworkRow[56] = "4000";
                            _var.dtworkRow[57] = decimal.Parse(String.IsNullOrEmpty(fyprem.ToString()) ? "0" : fyprem.ToString()) -
                                 decimal.Parse(String.IsNullOrEmpty(comprem.ToString()) ? "0" : comprem.ToString());
                        }
                        else if ((ryprem == String.Empty) && (fyprem == String.Empty))
                        {
                            if (TRANCODE.ToUpper().Contains("TNEWBUS"))
                            {
                                _var.dtworkRow[56] = "4000";
                                _var.dtworkRow[57] = "0";
                            }
                            else if ((TRANCODE.ToUpper().Contains("TRENEW")) || (TRANCODE.ToUpper().Contains("TREINS")))
                            {
                                _var.dtworkRow[58] = "4001";
                                _var.dtworkRow[59] = "0";
                            }
                            else if ((TRANCODE.ToUpper().Contains("ADJUST")) || (TRANCODE.ToUpper().Contains("TLAPSE")))
                            {
                                _var.dtworkRow[60] = "4004";
                                _var.dtworkRow[61] = "0";
                            }
                        }
                        if ((facul != String.Empty))
                        {
                          _var.dtworkRow[56] = "4000";
                            _var.dtworkRow[57] = decimal.Parse(String.IsNullOrEmpty(facul.ToString()) ? "0" : facul.ToString()) -
                                 decimal.Parse(String.IsNullOrEmpty(cfacul.ToString()) ? "0" : cfacul.ToString()); ;
                        }
                        else if (rfacul != String.Empty)
                        {
                           _var.dtworkRow[59] = decimal.Parse(String.IsNullOrEmpty(rfacul.ToString()) ? "0" : rfacul.ToString()); ;
                           _var.dtworkRow[58] = "4001";
                        }

                        #region "New Requirements - No Name"
                        if (String.IsNullOrEmpty(fullname))
                        {
                            fullname = polnum.ToString();
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR6AF" : _var.dtworkRow[76].ToString() + "|BR6AF";
                        }
                        #endregion

                        objHlpr.fn_getnamesandlifeID(fullname, dob, out _var.str_outfname, out _var.str_outlname, out _var.str_outlifeid, "021");

                        string str_MI = objHlpr.fn_getMI(_var.str_outfname);
                        _var.dtworkRow[34] = str_MI;

                        _var.dtworkRow[31] = objHlpr.fn_stringcleanup(fullname);
                        _var.dtworkRow[32] = _var.str_outlname;

                        //ISSUE#-Start022---------
                        string[] arr_fname;
                        arr_fname = _var.str_outfname.Split(' ');

                        if (!String.IsNullOrEmpty(str_MI.Trim()))
                        {
                            for (int i = 0; i <= arr_fname.Length - 1; i++)
                            {
                                if (arr_fname[i] != str_MI)
                                {
                                    _var.dtworkRow[33] = String.IsNullOrEmpty(_var.dtworkRow[33].ToString()) ? arr_fname[i] : _var.dtworkRow[33].ToString() + " " + arr_fname[i];
                                }
                            }
                        }
                        else
                        {
                            //NIGNES 20200818
                            //Correct First name and Middlename out
                            //Start
                            string[] arr_mname;
                            arr_mname = _var.str_outfname.Split(' ');
                            if (arr_mname.Length > 1)
                            {
                                string[] str_suffix = {
                                            "JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "VI"
                                        };

                                if (str_suffix.Any(arr_mname[arr_mname.Length - 1].Contains))
                                {
                                    _var.dtworkRow[34] = arr_mname[arr_mname.Length - 2];
                                }
                                else
                                {
                                    _var.dtworkRow[34] = arr_mname[arr_mname.Length - 1];
                                }
                                _var.dtworkRow[33] = _var.str_outfname.Replace(" " + _var.dtworkRow[34].ToString(), string.Empty);
                            }
                            else
                            {
                                _var.dtworkRow[33] = _var.str_outfname;
                            }
                            arr_mname = null;
                            //NIGNES 20200818
                        }
                        //ISSUE#-End022-----------

                        _var.dtworkRow[30] = _var.str_outlifeid;

                        //ISSUE#020-Start---------
                        if (!String.IsNullOrEmpty(gender))
                        {
                            _var.dtworkRow[36] = (gender.ToUpper().IndexOf("F") == 0) ? "F" : "M";
                        }
                        //ISSUE#020-End-----------
                        else if (String.IsNullOrEmpty(gender) && !String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow[36] = objHlpr.fn_getgender(str_gender, _var.dtworkRow[33].ToString());
                            //ISSUE#003-Start---------
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR7AK" : _var.dtworkRow[76].ToString() + "|BR7AK";
                            //ISSUE#003-End-----------
                        }
                        else if (String.IsNullOrEmpty(gender) && String.IsNullOrEmpty(str_gender))
                        {
                            _var.dtworkRow[36] = string.Empty;
                        }

                        //ISSUE#013-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow[36].ToString()))
                        {
                            _var.str_GFailLines = String.IsNullOrEmpty(_var.str_GFailLines) ? prawrow.ToString() : _var.str_GFailLines + "," + prawrow.ToString();
                        }
                        //ISSUE#013-End-----------

                        #region "New Requirements"
                        _var.dtworkRow[26] = string.Empty;

                        if (!String.IsNullOrEmpty(_var.dtworkRow[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()))
                        {
                            _var.dtworkRow[77] = _var.dtworkRow[27];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR1-1BZ" : _var.dtworkRow[76].ToString() + "|BR1-1BZ";
                        }
                        else if (!String.IsNullOrEmpty(_var.dtworkRow[25].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()))
                        {
                            _var.dtworkRow[75] = _var.dtworkRow[25];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR1-2BZ" : _var.dtworkRow[76].ToString() + "|BR1-2BZ";
                        }

                        if (!String.IsNullOrEmpty(_var.dtworkRow[77].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[27].ToString()))
                        {
                            _var.dtworkRow[27] = _var.dtworkRow[77];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR2-1AB" : _var.dtworkRow[76].ToString() + "|BR2-1AB";
                        }
                        else if (!String.IsNullOrEmpty(_var.dtworkRow[25].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[27].ToString()))
                        {
                            _var.dtworkRow[27] = _var.dtworkRow[25];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR2-2AB" : _var.dtworkRow[76].ToString() + "|BR2-2AB";
                        }

                        if (!String.IsNullOrEmpty(_var.dtworkRow[27].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[25].ToString()))
                        {
                            _var.dtworkRow[25] = _var.dtworkRow[27];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR3-1Z" : _var.dtworkRow[76].ToString() + "|BR3-1Z";
                        }
                        else if (!String.IsNullOrEmpty(_var.dtworkRow[77].ToString())
                            &&
                            String.IsNullOrEmpty(_var.dtworkRow[25].ToString()))
                        {
                            _var.dtworkRow[25] = _var.dtworkRow[77];
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR3-2Z" : _var.dtworkRow[76].ToString() + "|BR3-2Z";
                        }

                        //ISSUE#009-Start---------
                        var parsedDOB = DateTime.ParseExact(dob, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                        //ISSUE#009-End-----------

                        string initialNR = string.Empty;
                        if (!String.IsNullOrEmpty(_var.str_outfname))
                        {
                            initialNR = _var.str_outfname.Substring(0, 1);
                        }
                        if (!String.IsNullOrEmpty(_var.str_outlname))
                        {
                            initialNR += _var.str_outlname.Substring(0, 1);
                        }

                        if (_var.dtworkRow[13].ToString() == "GRP" || _var.dtworkRow[13].ToString() == "GCL" || _var.dtworkRow[13].ToString() == "GEB")
                        {
                            if (_var.dtworkRow[0].ToString().Length >= 7)
                            {
                                _var.dtworkRow[0] = _var.dtworkRow[0].ToString().Substring(_var.dtworkRow[0].ToString().Length - 7, 7) +
                                    initialNR +
                                    parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                            }
                            else
                            {
                                _var.dtworkRow[0] = _var.dtworkRow[0].ToString() +
                                    initialNR +
                                    parsedDOB.Month.ToString().PadLeft(2, '0') + parsedDOB.Day.ToString().PadLeft(2, '0') + parsedDOB.Year;
                            }
                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR5-1A" : _var.dtworkRow[76].ToString() + "|BR5-1A";

                            //ISSUE#019-Start---------
                            _var.dtworkRow[1] = _var.dtworkRow[0].ToString() + gender.Substring(0, 1);
                            //ISSUE#019-End-----------

                            _var.dtworkRow[76] = String.IsNullOrEmpty(_var.dtworkRow[76].ToString()) ? "BR5-2B" : _var.dtworkRow[76].ToString() + "|BR5-2B";

                            _var.dtworkRow[7] = polnum.ToString();
                        }
                        else
                        {
                            _var.dtworkRow[1] = string.Empty;
                            _var.dtworkRow[7] = string.Empty;
                        }

                        //ISSUE#010-Start---------
                        if (String.IsNullOrEmpty(_var.dtworkRow[19].ToString()))
                        {
                            if (_var.dtworkRow[21].ToString().ToUpper() == "TNEWBUS")
                            {
                                _var.dtworkRow[19] = _var.dtworkRow[20];
                            }
                            else
                            {
                                _var.dtworkRow[19] = _var.dtworkRow[22];
                            }
                        }
                        //ISSUE#010-End-----------

                        //ISSUE#017-Start---------
                        if (_var.dtworkRow[25].ToString() == "0")
                        {
                            _var.dtworkRow[25] = "1";
                        }
                        if (_var.dtworkRow[26].ToString() == "0")
                        {
                            _var.dtworkRow[26] = "1";
                        }
                        if (_var.dtworkRow[27].ToString() == "0")
                        {
                            _var.dtworkRow[27] = "1";
                        }
                        if (_var.dtworkRow[28].ToString() == "0")
                        {
                            _var.dtworkRow[28] = "1";
                        }
                        if (_var.dtworkRow[77].ToString() == "0")
                        {
                            _var.dtworkRow[77] = "1";
                        }
                        //ISSUE#017-End-----------

                        #endregion

                        _var.dbl_BF += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[57].ToString()) ? "0" : _var.dtworkRow[57].ToString()
                            );
                        _var.dbl_BH += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[59].ToString()) ? "0" : _var.dtworkRow[59].ToString()
                            );
                        _var.dbl_BJ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[61].ToString()) ? "0" : _var.dtworkRow[61].ToString()
                            );
                        _var.dbl_BL += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[63].ToString()) ? "0" : _var.dtworkRow[63].ToString()
                            );
                        _var.dbl_BZ += decimal.Parse(
                            String.IsNullOrEmpty(_var.dtworkRow[77].ToString()) ? "0" : _var.dtworkRow[77].ToString()
                            );

                        objdt_template.Rows.Add(_var.dtworkRow);
                    }

                    prawrow++;
                    polnum = wsraw.Cells[prawrow, 1].Text.ToString();
                    branded = wsraw.Cells[prawrow, 10].Text.ToString();
                    reins = wsraw.Cells[prawrow, 4].Text.ToString();
                    sum = wsraw.Cells[prawrow, 11].Text.ToString();
                    risk = wsraw.Cells[prawrow, 16].Text.ToString();
                    retention = wsraw.Cells[prawrow, 14].Text.ToString();
                    fullname = wsraw.Cells[prawrow, 3].Text.ToString();
                    gender = wsraw.Cells[prawrow, 8].Text.ToString();
                    dob = wsraw.Cells[prawrow, 6].Text.ToString();
                    pref = wsraw.Cells[prawrow, 13].Text.ToString();
                    premium = wsraw.Cells[prawrow, 18].Text.ToString();
                    issue = wsraw.Cells[prawrow, 7].Text.ToString();
                    fyprem = wsraw.Cells[prawrow, 27].Text.ToString();
                    ryprem = wsraw.Cells[prawrow, 35].Text.ToString();
                    comprem = wsraw.Cells[prawrow, 43].Text.ToString();
                    facul = wsraw.Cells[prawrow, 31].Text.ToString();
                    rfacul = wsraw.Cells[prawrow, 39].Text.ToString();
                    cfacul = wsraw.Cells[prawrow, 47].Text.ToString();
                    risk1 = wsraw.Cells[prawrow, 17].Text.ToString();
                    paid = wsraw.Cells[prawrow, 5].Text.ToString();

                    rowcount++;
                }
                #endregion

                #region "Compute Hash Total"
                _var.dtworkRow = objdt_template.NewRow();
                objdt_template.Rows.Add(_var.dtworkRow);

                _var.dtworkRow = objdt_template.NewRow();
                _var.dtworkRow[0] = "Total Premium:";
                _var.dtworkRow[1] = _var.dbl_BF + _var.dbl_BH + _var.dbl_BJ + _var.dbl_BL;
                objdt_template.Rows.Add(_var.dtworkRow);

                _var.dtworkRow = objdt_template.NewRow();
                _var.dtworkRow[0] = "Total Sum at Risk:";
                _var.dtworkRow[1] = _var.dbl_BZ;
                objdt_template.Rows.Add(_var.dtworkRow);
                #endregion

                //ISSUE#013-Start---------
                #region "List all failed genders"
                if (_var.str_GFailLines != string.Empty)
                {
                    _var.dtworkRow = objdt_template.NewRow();
                    _var.dtworkRow[0] = "Gender Fail Lines on RAW:";
                    _var.dtworkRow[1] = _var.str_GFailLines;
                    objdt_template.Rows.Add(_var.dtworkRow);
                }
                #endregion
                //ISSUE#013-End-----------

                string despath = str_saved + @"\BM021-" + str_sheet + "-" + str_savef + ".xlsx";
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
                _var.dtworkRow = null; //Dispose datarow
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
