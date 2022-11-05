using System;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;



namespace Bordereaux_SICS_Mapping.BAL
{

    class HelperV21 :_Global
    {
        _Global _var = new _Global();
        Helper objHlpr = new Helper();

        
        /**************************************CURRENCY********************************************************
        PROGRAM: SICS 
        COLUMN: _var.dtworkRow01[23]
        DESCRIPTION: Return the value of currency depending on the BM and value passed to the parameter

        FUNCTION NAME: fn_getcurrency
        BORDEREUX NO: BM061
        */


        public string fn_getcurrency(string valueCurrency)

        {
            if(valueCurrency == "1")
            {
                return "IDR";
            }
            else
            {
                return "USD";
            }
            
        }
        public string fn_getcurrencyV2(string value)

        {
            value = value.ToUpper().Trim();
            if(value.Contains("$") || value.ToUpper().Contains("DOLLAR"))
            {
                return "USD";
            }
            else
            {
                return "PHP";
            }
        }




        /**********************************SPLIT NAMES**************************************
        PROGRAM: SICS 
        COLUMN: 
        DESCRIPTION:Split the full name Return the value of firstname and lastname

        FUNCTION NAME:
        fn_seperateforeignames:
        fn_seperateforeignamesV2:BM061
        */

        string [] Suffix = {
            "JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "VI","VII"
            };

        string [] LNSuffix = {
            " DE ", " DEL ", " DELA ", " DELOS ", " DELAS ",
            " LA ", " LAS ", " LOS ",
            " SAN ", " STA ", " STA. ", " STO ", " STO. ", " SANTO ", " SANTA "
        };


        public void fn_seperateforeignames(string strFullname, out string strFirstName, out string strLastName)
        {

            var names = strFullname.Split(' ');
            strLastName = names [0];
            strFirstName = names [1];
            string strTitle;
            char strMI;

            if(strFullname.Contains("III"))
            {
                //strTitle = names[2];
                strLastName = names [0] + " " + names [1] + " " + names [2] + " " + names [3];
                strMI = strFullname [strFullname.Length - 1];
                names = names.Take(names.Length - 1).ToArray();
                names = names.Skip(4).ToArray();
                strFirstName = String.Join(" ", names);
            }
        }

        public void fn_seperateforeignamesV2(string strFullname, out string strFirstName, out string strLastName, out string strMI)
        {

            var names = strFullname.Split(' ');
            strFirstName = "";
            strLastName = "";
            strMI = "";

            for(int j = 1; names.Length >= j; j++)
            {

                if(names.Length <= 1)
                {
                    strLastName = names [0];
                    //strFirstName = names[0];
                    break;
                }
                else if(names.Length <= 2)
                {
                    strFirstName = names [0];
                    strLastName = names [1];
                    break;
                }
                else if(names.Length >= 3)
                {
                    if(names.Length == 3)
                    {
                        strFirstName = names [0] + " " + names [1];
                        strLastName = names [2];
                        break;
                    }
                    else if(names.Length == 4)
                    {
                        strFirstName = names [0] + " " + names [1] + " " + names [2];
                        strLastName = names [3];
                        break;
                    }
                    else if(names.Length == 5)
                    {
                        strFirstName = names [0] + " " + names [1] + " " + names [2] + " " + names [3];
                        strLastName = names [4];
                        break;
                    }
                    else if(names.Length == 6)
                    {
                        strFirstName = names [0] + " " + names [1] + " " + names [2] + " " + names [3] + " " + names [4];
                        strLastName = names [5];
                        break;
                    }
                    else
                    {
                        strFirstName = names [0] + " " + names [1];
                        strLastName = names [2];
                        break;

                    }
                }


            }

        }


        // FirstName , LastName
        public void fn_separateFirstNameLastNameV1(string strFullname, out string strFirstName, out string strLastName)
        {
            strFullname = strFullname.TrimEnd();
            var names = strFullname.Split(' ');
            strFirstName = "";
            strLastName = "";

            Console.WriteLine(strFullname);


            if(names.Length == 2)
            {
                strFirstName = names [0];
                strLastName = names [1];

            }
            else if(names.Length == 3)
            {

                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV") || strFullname.ToUpper().Contains("V") || strFullname.ToUpper().Contains("VI"))
                {
                    if(names [0].ToUpper() == "JR" || names [0].ToUpper() == "JR." || names [0].ToUpper() == "SR" || names [0].ToUpper() == "SR." || names [0].ToUpper() == "II" || names [0].ToUpper() == "III" || names [0].ToUpper() == "IV" || names [0].ToUpper() == "V" || names [0].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 3];
                        strLastName = names [names.Length - 1];
                    }
                    else if(names [1].ToUpper() == "JR" || names [1].ToUpper() == "JR." || names [1].ToUpper() == "SR" || names [1].ToUpper() == "SR." || names [1].ToUpper() == "II" || names [1].ToUpper() == "III" || names [1].ToUpper() == "IV" || names [1].ToUpper() == "V" || names [1].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strLastName = names [names.Length - 1];
                    }
                    else if(names [2].ToUpper() == "JR" || names [2].ToUpper() == "JR." || names [2].ToUpper() == "SR" || names [2].ToUpper() == "SR." || names [2].ToUpper() == "II" || names [2].ToUpper() == "III" || names [2].ToUpper() == "IV" || names [2].ToUpper() == "V" || names [2].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strLastName = names [names.Length - 3];
                    }

                    else
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strLastName = names [names.Length - 3];
                    }

                }
                else if(names [0].ToUpper() == "DE" || names [0].ToUpper() == "DEL." || names [0].ToUpper() == "DELA" || names [0].ToUpper() == "DELOS" || names [0].ToUpper() == "LA" || names [0].ToUpper() == "LAS" || names [0].ToUpper() == "LOS" || names [0].ToUpper() == "SAN" || names [0].ToUpper() == "STA" || names [0].ToUpper() == "STA." || names [0].ToUpper() == "STO." || names [0].ToUpper() == "SANTO" || (names [0].ToUpper() == "SANTA" || (names [0].ToUpper() == "MA." || (names [0].ToUpper() == "MA"))))
                {
                    strFirstName = names [names.Length - 1];
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2];
                }
                else if(names [1].ToUpper() == "DE" || names [1].ToUpper() == "DEL." || names [1].ToUpper() == "DELA" || names [1].ToUpper() == "DELOS" || names [1].ToUpper() == "LA" || names [1].ToUpper() == "LAS" || names [1].ToUpper() == "LOS" || names [1].ToUpper() == "SAN" || names [1].ToUpper() == "STA" || names [1].ToUpper() == "STA." || names [1].ToUpper() == "STO." || names [1].ToUpper() == "SANTO" || (names [1].ToUpper() == "SANTA" || (names [1].ToUpper() == "MA." || (names [1].ToUpper() == "MA"))))
                {
                    strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                    strLastName = names [names.Length - 1];
                }
                else
                //Ma Agnes Mirasol
                {
                    strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                    strLastName = names [names.Length - 3];
                }

            }

            else if(names.Length == 4)
            {

                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV") || strFullname.ToUpper().Contains("V") || strFullname.ToUpper().Contains("VI"))
                {
                    if(names [0].ToUpper() == "JR" || names [0].ToUpper() == "JR." || names [0].ToUpper() == "SR" || names [0].ToUpper() == "SR." || names [0].ToUpper() == "II" || names [0].ToUpper() == "III" || names [0].ToUpper() == "IV" || names [0].ToUpper() == "V" || names [0].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 3];
                        strLastName = names [names.Length - 1];
                    }
                    else if(names [1].ToUpper() == "JR" || names [1].ToUpper() == "JR." || names [1].ToUpper() == "SR" || names [1].ToUpper() == "SR." || names [1].ToUpper() == "II" || names [1].ToUpper() == "III" || names [1].ToUpper() == "IV" || names [1].ToUpper() == "V" || names [1].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strLastName = names [names.Length - 1];
                    }
                    else if(names [2].ToUpper() == "JR" || names [2].ToUpper() == "JR." || names [2].ToUpper() == "SR" || names [2].ToUpper() == "SR." || names [2].ToUpper() == "II" || names [2].ToUpper() == "III" || names [2].ToUpper() == "IV" || names [2].ToUpper() == "V" || names [2].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strLastName = names [names.Length - 1];
                    }
                    else if(names [3].ToUpper() == "JR" || names [3].ToUpper() == "JR." || names [3].ToUpper() == "SR" || names [3].ToUpper() == "SR." || names [3].ToUpper() == "II" || names [3].ToUpper() == "III" || names [3].ToUpper() == "IV" || names [3].ToUpper() == "V" || names [3].ToUpper() == "VI")
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strLastName = names [names.Length - 4];

                    }
                    else
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strLastName = names [names.Length - 1];
                    }

                }
                else if(names [0].ToUpper() == "DE" || names [0].ToUpper() == "DEL." || names [0].ToUpper() == "DELA" || names [0].ToUpper() == "DELOS" || names [0].ToUpper() == "LA" || names [0].ToUpper() == "LAS" || names [0].ToUpper() == "LOS" || names [0].ToUpper() == "SAN" || names [0].ToUpper() == "STA" || names [0].ToUpper() == "STA." || names [0].ToUpper() == "STO." || names [0].ToUpper() == "SANTO" || (names [0].ToUpper() == "SANTA" || (names [0].ToUpper() == "MA." || (names [0].ToUpper() == "MA"))))
                {
                    strFirstName = names [names.Length - 1];
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2];
                }
                else if(names [1].ToUpper() == "DE" || names [1].ToUpper() == "DEL." || names [1].ToUpper() == "DELA" || names [1].ToUpper() == "DELOS" || names [1].ToUpper() == "LA" || names [1].ToUpper() == "LAS" || names [1].ToUpper() == "LOS" || names [1].ToUpper() == "SAN" || names [1].ToUpper() == "STA" || names [1].ToUpper() == "STA." || names [1].ToUpper() == "STO." || names [1].ToUpper() == "SANTO" || (names [1].ToUpper() == "SANTA" || (names [1].ToUpper() == "MA." || (names [1].ToUpper() == "MA"))))
                {

                    strFirstName = names [names.Length - 4] + " " + names [names.Length - 1];
                    strLastName = names [names.Length - 3] + " " + names [names.Length - 2];
                }
                else if(names [2].ToUpper() == "DE" || names [2].ToUpper() == "DEL." || names [2].ToUpper() == "DELA" || names [2].ToUpper() == "DELOS" || names [2].ToUpper() == "LA" || names [2].ToUpper() == "LAS" || names [2].ToUpper() == "LOS" || names [2].ToUpper() == "SAN" || names [2].ToUpper() == "STA" || names [2].ToUpper() == "STA." || names [2].ToUpper() == "STO." || names [2].ToUpper() == "SANTO" || (names [2].ToUpper() == "SANTA" || (names [2].ToUpper() == "MA." || (names [2].ToUpper() == "MA"))))
                {
                    strFirstName = names [names.Length - 4] + " " + names [names.Length - 3];
                    strLastName = names [names.Length - 2] + " " + names [names.Length - 1];
                }
                else
                {
                    strLastName = names [names.Length - 4];
                    strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                }

            }

        }

        // FirstName , LastName,MiddleName
        public void fn_separateFirstNameLastNameV2(string valueFullName, out string strFullname, out string strFirstname, out string strLastname, out string strMiddlename)
        {
            strFirstname = string.Empty; strLastname = string.Empty; strMiddlename = string.Empty; strFullname = string.Empty;
            try
            {
                valueFullName = valueFullName.TrimEnd();
                var names = valueFullName.Split();
                strFullname = valueFullName;

                Console.WriteLine(valueFullName);
                if(names.Length == 1)
                {
                    strFirstname = names [0];
                }
                else if(names.Length == 2)
                {

                    strFirstname = names [0];
                    strLastname = names [1];
                }



                else if(names.Length >= 3)//DE GUZMAN JR MAMARIL  //ESPINOSA, VIVIAN B.
                {
                    if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                    {
                        if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                        {

                            strFirstname = names [0] + " " + names [1];
                            names = names.Skip(2).ToArray();
                            strLastname = String.Join(" ", names);
                        }
                        else
                        {
                            strFirstname = names [0] + " " + names [1];
                            names = names.Skip(2).ToArray();
                            strLastname = String.Join(" ", names);

                        }
                    }
                    else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                    {
                        strFirstname = names [0];
                        names = names.Skip(1).ToArray();
                        strLastname = String.Join(" ", names);

                    }
                    else
                    {
                        //JUAN KEVIN G BELMONTE
                        //MA. ANNA CRISTINA Y PEREZ
                        if(names [1].Length > 2)
                        {

                            if(names [1].Length > 2 && names [2].Length > 2)
                            {

                                strFirstname = names [0] + " " + names [1] + " " + names [2];
                            }
                            else if(names [1].Length > 2)
                            {
                                strFirstname = names [0] + " " + names [1];
                            }
                            strMiddlename = names [names.Length - 2];
                            int intMiddleName = strMiddlename.Length;
                            if(strMiddlename.Contains(".") || intMiddleName == 1)
                            {
                                strMiddlename = names [names.Length - 2];
                                names = names.Skip(2).ToArray();
                                strLastname = names [names.Length - 1];
                            }
                            else
                            {
                                names = names.Skip(1).ToArray();
                                strLastname = names [names.Length - 1];
                                strMiddlename = names [names.Length - 2];
                            }
                        }
                        else
                        {
                            strFirstname = names [0];
                            strMiddlename = names [names.Length - 2];
                            int intMiddleName = strMiddlename.Length;
                            if(strMiddlename.Contains(".") || intMiddleName == 1)
                            {
                                strMiddlename = names [names.Length - 2];
                                names = names.Skip(2).ToArray();
                                strLastname = names [names.Length - 1];
                            }
                            else
                            {
                                names = names.Skip(1).ToArray();
                                strLastname = names [names.Length - 1];
                                strMiddlename = names [names.Length - 2];
                            }
                        }
                    }

                }

            }
            catch(Exception ex)
            {
                strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }



        /*****************************************************************
         * DESCRIPTION: Middle Initial has . or period
         * e.g MATEO, MARILOU V.
         * FUNCTION NAME: fn_separateLastNameFirstName
         * LastName, FirstName, MiddleName
         */

        public void fn_separateLastNameFirstName(string strFullname, out string strFirstName, out string strLastName, out string strMiddleInitial)
        {
            strFullname = strFullname.TrimEnd();
            var names = strFullname.Split(' ');
            strMiddleInitial = "";
            //strTitle = "";
            strFirstName = "";
            //string strSecondName = "";
            strLastName = "";


            Console.WriteLine(strFullname);

            if(names.Length == 2)
            {
                strLastName = names [0];
                strFirstName = names [1];
            }

            else if(names.Length == 3)
            {
                strLastName = names [names.Length - 3];

                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV"))
                {
                    if(names [1].ToUpper().Contains("JR") || names [1].ToUpper().Contains("JR.") || names [1].ToUpper().Contains("SR") || names [1].ToUpper().Contains("SR.") || names [1].ToUpper().Contains("II") || names [1].ToUpper().Contains("III") || names [1].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 1] + " " + names [names.Length - 2];
                    }
                    else if(names [2].ToUpper().Contains("JR") || names [2].ToUpper().Contains("JR.") || names [2].ToUpper().Contains("SR") || names [2].ToUpper().Contains("SR.") || names [2].ToUpper().Contains("II") || names [2].ToUpper().Contains("III") || names [2].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                    else if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 2];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                }
                else if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || strFullname.ToUpper().Contains("STA") || strFullname.ToUpper().Contains("STA.") || strFullname.ToUpper().Contains("STO.") || strFullname.ToUpper().Contains("SANTO") || (strFullname.ToUpper().Contains("SANTA")))
                {
                    Console.WriteLine(strFullname);
                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 2];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                }

                else
                {

                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                }

            }

            else if(names.Length == 4)
            {
                strLastName = names [names.Length - 4];
                //"JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "VI"
                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV"))
                {
                    if(names [1].ToUpper().Contains("JR") || names [1].ToUpper().Contains("JR.") || names [1].ToUpper().Contains("SR") || names [1].ToUpper().Contains("SR.") || names [1].ToUpper().Contains("II") || names [1].ToUpper().Contains("III") || names [1].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 2] + " " + names [1];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [2].ToUpper().Contains("JR") || names [2].ToUpper().Contains("JR.") || names [2].ToUpper().Contains("SR") || names [2].ToUpper().Contains("SR.") || names [2].ToUpper().Contains("II") || names [2].ToUpper().Contains("III") || names [2].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 3] + " " + names [2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [3].ToUpper().Contains("JR") || names [3].ToUpper().Contains("JR.") || names [3].ToUpper().Contains("SR") || names [3].ToUpper().Contains("SR.") || names [3].ToUpper().Contains("II") || names [3].ToUpper().Contains("III") || names [3].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];

                    }
                    else
                    {

                        if(names [1].Length == 1)
                        {
                            strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                            strMiddleInitial = names [1];
                        }
                        else if(names [2].Length == 1)
                        {
                            strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                            strMiddleInitial = names [3];
                        }
                        else if(names [3].Length == 1)
                        {
                            strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                            strMiddleInitial = names [3];
                        }

                    }
                }
                //" DE ", " DEL ", " DELA ", " DELOS ", " DELAS ", " LA ", " LAS ", " LOS "," SAN ", " STA ", " STA. ", " STO ", " STO. ", " SANTO ", " SANTA "
                else if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || names.Contains("STA") || names.Contains("STA.") || names.Contains("STO") || names.Contains("STO.") || names.Contains("SANTO") || (names.Contains("SANTA")))
                {
                    if(names [1].Length == 1)
                    {
                        strLastName = names [names.Length - 4] + " " + names [1];
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strLastName = names [names.Length - 4] + " " + names [1];
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strLastName = names [names.Length - 4] + " " + names [1];
                        strFirstName = names [names.Length - 2];
                        strMiddleInitial = names [3];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                }
                else
                {

                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [3];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    }


                }

            }
            else if(names.Length == 5)
            {
                strLastName = names [names.Length - 5];
                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV"))
                {
                    if(names [1].ToUpper().Contains("JR") || names [1].ToUpper().Contains("JR.") || names [1].ToUpper().Contains("SR") || names [1].ToUpper().Contains("SR.") || names [1].ToUpper().Contains("II") || names [1].ToUpper().Contains("III") || names [1].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1] + " " + names [names.Length - 4];

                    }
                    else if(names [2].ToUpper().Contains("JR") || names [2].ToUpper().Contains("JR.") || names [2].ToUpper().Contains("SR") || names [2].ToUpper().Contains("SR.") || names [2].ToUpper().Contains("II") || names [2].ToUpper().Contains("III") || names [2].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 3];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [3].ToUpper().Contains("JR") || names [3].ToUpper().Contains("JR.") || names [3].ToUpper().Contains("SR") || names [3].ToUpper().Contains("SR.") || names [3].ToUpper().Contains("II") || names [3].ToUpper().Contains("III") || names [3].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [4].ToUpper().Contains("JR") || names [4].ToUpper().Contains("JR.") || names [4].ToUpper().Contains("SR") || names [4].ToUpper().Contains("SR.") || names [4].ToUpper().Contains("II") || names [4].ToUpper().Contains("III") || names [4].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 2];
                    }
                    else
                    {

                        if(names [1].Length == 1)
                        {
                            strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                            strMiddleInitial = names [1];
                        }
                        else if(names [2].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                            strMiddleInitial = names [3];
                        }
                        else if(names [3].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                            strMiddleInitial = names [3];
                        }
                        else if(names [4].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                            strMiddleInitial = names [4];
                        }
                    }
                }
                else if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || strFullname.ToUpper().Contains("STA") || strFullname.ToUpper().Contains("STA.") || strFullname.ToUpper().Contains("STO") || strFullname.ToUpper().Contains("STO.") || strFullname.ToUpper().Contains("SANTO") || (strFullname.ToUpper().Contains("SANTA")))
                {
                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [4].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [4];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    }

                }
                else
                {

                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [4].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [4];
                    }
                    else
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                    }
                }
            }
            else if(names.Length == 6)
            {
                strLastName = names [names.Length - 6];
                if(strFullname.ToUpper().Contains("JR") || strFullname.ToUpper().Contains("JR.") || strFullname.ToUpper().Contains("SR") || strFullname.ToUpper().Contains("SR.") || strFullname.ToUpper().Contains("II") || strFullname.ToUpper().Contains("III") || strFullname.ToUpper().Contains("IV"))
                {
                    if(names [1].ToUpper().Contains("JR") || names [1].ToUpper().Contains("JR.") || names [1].ToUpper().Contains("SR") || names [1].ToUpper().Contains("SR.") || names [1].ToUpper().Contains("II") || names [1].ToUpper().Contains("III") || names [1].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 2] + names [names.Length - 3];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [2].ToUpper().Contains("JR") || names [2].ToUpper().Contains("JR.") || names [2].ToUpper().Contains("SR") || names [2].ToUpper().Contains("SR.") || names [2].ToUpper().Contains("II") || names [2].ToUpper().Contains("III") || names [2].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 3];
                        strMiddleInitial = names [names.Length - 1];
                    }
                    else if(names [3].ToUpper().Contains("JR") || names [3].ToUpper().Contains("JR.") || names [3].ToUpper().Contains("SR") || names [3].ToUpper().Contains("SR.") || names [3].ToUpper().Contains("II") || names [3].ToUpper().Contains("III") || names [3].ToUpper().Contains("IV"))
                    {
                        if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || strFullname.ToUpper().Contains("STA") || strFullname.ToUpper().Contains("STA.") || strFullname.ToUpper().Contains("STO") || strFullname.ToUpper().Contains("STO.") || strFullname.ToUpper().Contains("SANTO") || (strFullname.ToUpper().Contains("SANTA")))
                        {
                            //DE LOS REYES JR REYNALDO F
                            strLastName = names [names.Length - 6] + " " + names [names.Length - 5] + " " + names [names.Length - 4];
                            strFirstName = names [names.Length - 2] + " " + names [names.Length - 3];
                            strMiddleInitial = names [names.Length - 1];
                        }
                        else
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                            strMiddleInitial = names [names.Length - 1];
                        }

                    }
                    else if(names [4].ToUpper().Contains("JR") || names [4].ToUpper().Contains("JR.") || names [4].ToUpper().Contains("SR") || names [4].ToUpper().Contains("SR.") || names [4].ToUpper().Contains("II") || names [4].ToUpper().Contains("III") || names [4].ToUpper().Contains("IV"))
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [names.Length - 2];
                    }
                    else if(names [5].ToUpper().Contains("JR") || names [5].ToUpper().Contains("JR.") || names [5].ToUpper().Contains("SR") || names [5].ToUpper().Contains("SR.") || names [5].ToUpper().Contains("II") || names [5].ToUpper().Contains("III") || names [5].ToUpper().Contains("IV"))
                    {
                        if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || strFullname.ToUpper().Contains("STA") || strFullname.ToUpper().Contains("STA.") || strFullname.ToUpper().Contains("STO") || strFullname.ToUpper().Contains("STO.") || strFullname.ToUpper().Contains("SANTO") || (strFullname.ToUpper().Contains("SANTA")))
                        {

                            strLastName = names [names.Length - 6] + " " + names [names.Length - 5] + " " + names [names.Length - 4];
                            if(names [3].ToUpper().Contains("JR") || names [3].ToUpper().Contains("JR.") || names [3].ToUpper().Contains("SR") || names [3].ToUpper().Contains("SR.") || names [3].ToUpper().Contains("II") || names [3].ToUpper().Contains("III") || names [3].ToUpper().Contains("IV"))
                            {
                                strFirstName = names [names.Length - 2] + " " + names [3];
                                strMiddleInitial = names [names.Length - 1];
                            }
                            else if(names [4].ToUpper().Contains("JR") || names [4].ToUpper().Contains("JR.") || names [4].ToUpper().Contains("SR") || names [4].ToUpper().Contains("SR.") || names [4].ToUpper().Contains("II") || names [4].ToUpper().Contains("III") || names [4].ToUpper().Contains("IV"))
                            {
                                strFirstName = names [names.Length - 3] + " " + names [4];
                                strMiddleInitial = names [names.Length - 1];
                            }
                            else if(names [5].ToUpper().Contains("JR") || names [5].ToUpper().Contains("JR.") || names [5].ToUpper().Contains("SR") || names [5].ToUpper().Contains("SR.") || names [5].ToUpper().Contains("II") || names [5].ToUpper().Contains("III") || names [5].ToUpper().Contains("IV"))
                            {
                                strFirstName = names [names.Length - 3] + " " + names [4] + " " + names [5];

                            }
                        }
                    }
                    else
                    {

                        if(names [1].Length == 1)
                        {
                            strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                            strMiddleInitial = names [1];
                        }
                        else if(names [2].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                            strMiddleInitial = names [3];
                        }
                        else if(names [3].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                            strMiddleInitial = names [3];
                        }
                        else if(names [4].Length == 1)
                        {
                            strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                            strMiddleInitial = names [4];
                        }
                    }
                }
                else if(strFullname.ToUpper().Contains("DE") || strFullname.ToUpper().Contains("DEL.") || strFullname.ToUpper().Contains("DELA") || strFullname.ToUpper().Contains("DELOS") || strFullname.ToUpper().Contains("LA") || strFullname.ToUpper().Contains("LAS") || strFullname.ToUpper().Contains("LOS") || strFullname.ToUpper().Contains("SAN") || strFullname.ToUpper().Contains("STA") || strFullname.ToUpper().Contains("STA.") || strFullname.ToUpper().Contains("STO") || strFullname.ToUpper().Contains("STO.") || strFullname.ToUpper().Contains("SANTO") || (strFullname.ToUpper().Contains("SANTA")))
                {
                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [4].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [4];
                    }
                    else if(names [5].Length == 1)
                    {

                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [5];
                    }
                }
                else
                {

                    if(names [1].Length == 1)
                    {
                        strFirstName = names [names.Length - 3] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [1];
                    }
                    else if(names [2].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 2] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [3].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 1];
                        strMiddleInitial = names [3];
                    }
                    else if(names [4].Length == 1)
                    {
                        strFirstName = names [names.Length - 4] + " " + names [names.Length - 3] + " " + names [names.Length - 2];
                        strMiddleInitial = names [4];
                    }
                }

            }

        }

        public void fn_separateLastNameFirstNameV2(string valueFullName, out string strFullname, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            #region NOTES
            //Middle Initial is in between LastName and First Name e.g MANUEL R LAROCO
            #endregion
            strFirstname = string.Empty; strLastname = string.Empty; strMiddlename = string.Empty; strFullname = string.Empty;
            try
            {
                string out_suffix = string.Empty;
                valueFullName = valueFullName.TrimEnd();
                var names = valueFullName.Split();
                strFullname = valueFullName;
                string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA" };
                string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
                Console.WriteLine(valueFullName);
                if(names.Length == 1)
                {
                    strLastname = names [0];
                }
                else if(names.Length == 2)
                {
                    strLastname = names [0];
                    strFirstname = names [1];
                }

                else if(names.Length >= 3)//CHAM, DENZELL L.
                {

                    int intMI = 0;
                    if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                    {
                        if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                        {
                            bool bolSuff = false;
                            foreach(string lnsuff in LNSuffix)
                            {
                                if(lnsuff != names [0])
                                {
                                    continue;
                                }
                                else if(lnsuff == names [0])
                                {

                                    strLastname = names [0] + " " + names [1];
                                    intMI = strMiddlename.Length;
                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        if(names.Length > 3)
                                        {
                                            bolSuff = true;
                                            strMiddlename = names [names.Length - 1];
                                            names = names.Skip(2).ToArray();
                                            strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                            break;
                                        }

                                    }
                                    else
                                    {
                                        bolSuff = true;
                                        strMiddlename = names [names.Length - 1];
                                        names = names.Skip(2).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                        break;

                                    }
                                }
                            }

                            if(bolSuff == false)
                            {

                                strLastname = names [0];
                                if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                                {
                                    out_suffix = names [names.Length - 1];
                                    strFirstname = names [1] + " " + out_suffix;
                                }
                                else
                                {

                                    //names = names.Skip(1).ToArray();
                                    strMiddlename = names [names.Length - 1];
                                    intMI = strMiddlename.Length;


                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        strMiddlename = names [names.Length - 1];
                                        //names = names.Skip(1).ToArray();
                                        strFirstname = names [2] + " " + names [1];

                                    }
                                    else
                                    {
                                        strMiddlename = names [names.Length - 2];
                                        intMI = strMiddlename.Length;
                                        if(strMiddlename.Contains(".") || intMI == 1)
                                        {
                                            strMiddlename = names [names.Length - 2];
                                            names = names.Skip(1).ToArray();
                                            strFirstname = names [names.Length - 1];

                                        }
                                        else
                                        {
                                            names = names.Skip(1).ToArray();
                                            strFirstname = String.Join(" ", names.ToArray());
                                            strMiddlename = string.Empty;
                                        }
                                    }
                                }
                            }
                        }
                        else if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                        {
                            //CHAM, DENZELL L.
                            bool bolSuffix = false;
                            foreach(string lnsuff in LNSuffix)
                            {
                                if(lnsuff != names [0])
                                {
                                    continue;
                                }
                                else if(lnsuff == names [0])
                                {

                                    strLastname = names [0] + " " + names [1];
                                    strMiddlename = names [names.Length - 1];
                                    //if(strMiddlename.Length == 2 || strMiddlename.Length == 1){
                                    //    strMiddlename = names [names.Length - 1];
                                    //}
                                    intMI = strMiddlename.Length;
                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        bolSuffix = true;
                                        strMiddlename = names [names.Length - 1];
                                        names = names.Skip(2).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                        break;

                                    }

                                    else
                                    {
                                        bolSuffix = true;
                                        names = names.Skip(2).ToArray();
                                        strMiddlename = "";
                                        //strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                        strFirstname = names [names.Length - 1];
                                        break;
                                    }
                                }

                            }
                            if(bolSuffix == false)
                            {
                                if(names.Length >= 3)
                                {
                                    strLastname = names [0];
                                    strMiddlename = names [names.Length - 1];
                                    intMI = strMiddlename.Length;
                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        strMiddlename = names [names.Length - 1];
                                        names = names.Skip(1).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));


                                    }
                                    else
                                    {
                                        strMiddlename = string.Empty;
                                        strLastname = names [0];
                                        names = names.Skip(1).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray());
                                    }

                                }
                            }
                        }
                        else
                        {
                            //CHAM, DENZELL L.
                            foreach(string lnsuff in LNSuffix)
                            {
                                if(lnsuff != names [0])
                                {
                                    continue;
                                }
                                else if(lnsuff == names [0])
                                {
                                    strLastname = names [0] + " " + names [1];
                                    intMI = strMiddlename.Length;
                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        strMiddlename = names [names.Length - 1];

                                    }
                                    else
                                    {
                                        strFirstname = names [2] + " " + names [1];

                                    }
                                }

                            }
                            strLastname = names [0];
                            //names = names.Skip(1).ToArray();
                            strMiddlename = names [names.Length - 1];
                            intMI = strMiddlename.Length;
                            if(strMiddlename.Contains(".") || intMI == 1)
                            {
                                strMiddlename = names [names.Length - 1];
                                //names = names.Skip(1).ToArray();
                                if(names.Length == 3)
                                {
                                    strFirstname = names [names.Length - 2];
                                }
                                else
                                {
                                    strFirstname = names [1] + " " + names [2];
                                }


                            }
                            else
                            {
                                strMiddlename = names [names.Length - 2];
                                intMI = strMiddlename.Length;
                                if(strMiddlename.Contains(".") || intMI == 1)
                                {
                                    strMiddlename = names [names.Length - 2];
                                    names = names.Skip(1).ToArray();
                                    strFirstname = names [names.Length - 1];

                                }
                                else
                                {
                                    names = names.Skip(1).ToArray();
                                    strFirstname = String.Join(" ", names.ToArray());
                                    strMiddlename = string.Empty;

                                }

                            }

                        }
                    }



                    else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                    {

                        int intCountNames = names.Length;

                        //intItem = intCount;
                        foreach(string suffix in Suffix)
                        {
                            int i = 0;
                            string strCheckSuffix = suffix;
                            //OLIVIA LYNN TAN UY
                            //SAMUEL ROBERT CHUASON JR
                            if(names.Contains(suffix))
                            {
                                strLastname = names [0];
                                strMiddlename = names [names.Length - 1];
                                intMI = strMiddlename.Length;
                                if(strMiddlename != suffix)
                                {

                                    if(strMiddlename.Contains(".") || intMI == 1)
                                    {
                                        names = names.Skip(1).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                        strMiddlename = names [names.Length - 1];
                                        break;
                                    }
                                    else if(names [1].Length == 1)
                                    {
                                        strMiddlename = names [1];
                                        names = names.Skip(2).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray());
                                        break;
                                    }
                                    else
                                    {

                                        strMiddlename = String.Empty;
                                        names = names.Skip(1).ToArray();
                                        strFirstname = String.Join(" ", names.ToArray());
                                        break;
                                    }

                                }
                                else
                                {
                                    strMiddlename = string.Empty;
                                    names = names.Skip(1).ToArray();
                                    strFirstname = String.Join(" ", names.ToArray());
                                    break;

                                }


                            }


                            foreach(string item in names)
                            {
                                i++;
                                if(item == suffix)
                                {
                                    strLastname = names [0];
                                    names = names.Skip(1).ToArray();
                                    strFirstname = String.Join(" ", names.ToArray());

                                }
                                else if(i == intCountNames)
                                {
                                    strLastname = names [0];
                                    if(strCheckSuffix == "VI")
                                    {

                                        strMiddlename = names [names.Length - 1];
                                        int intMiddleName = strMiddlename.Length;
                                        if(strMiddlename.Contains(".") || intMiddleName == 1)
                                        {
                                            strMiddlename = names [names.Length - 1];
                                            names = names.Skip(1).ToArray();
                                            strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                        }
                                        else
                                        {
                                            names = names.Skip(1).ToArray();
                                            strFirstname = String.Join(" ", names.ToArray());
                                        }

                                    }
                                    else if(item == suffix)
                                    {
                                        strMiddlename = names [names.Length - 1];
                                        if(strMiddlename.Contains(".") || intMI == 1)
                                        {
                                            strMiddlename = names [names.Length - 1];
                                            //names = names.Skip(1).ToArray();
                                            if(names.Length == 3)
                                            {
                                                strFirstname = names [names.Length - 2];
                                            }
                                            else
                                            {
                                                strFirstname = names [1] + " " + names [2];
                                            }


                                        }
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        strLastname = names [0];
                        strMiddlename = names [names.Length - 1];
                        int intMiddleName = strMiddlename.Length;
                        if(strMiddlename.Contains(".") || intMiddleName == 1)
                        {
                            if(intMiddleName >= 3)
                            {
                                strMiddlename = string.Empty;
                                names = names.Skip(1).ToArray();
                                strFirstname = String.Join(" ", names.ToArray());
                            }
                            else
                            {
                                strMiddlename = names [names.Length - 1];
                                names = names.Skip(1).ToArray();
                                strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                            }


                        }
                        else
                        {
                            strMiddlename = names [names.Length - 2];
                            intMiddleName = strMiddlename.Length;
                            if(strMiddlename.Contains(".") || intMiddleName == 1)
                            {
                                if(names.Length == 3)
                                {
                                    strMiddlename = names [names.Length - 2];
                                    names = names.Skip(1).ToArray();
                                    strFirstname = names [names.Length - 1];
                                }
                                else if(names.Length == 4)
                                {
                                    strMiddlename = names [names.Length - 2];
                                    names = names.Skip(0).ToArray();
                                    strFirstname = names [names.Length - 1];
                                }

                            }
                            else
                            {
                                names = names.Skip(1).ToArray();
                                strFirstname = String.Join(" ", names.ToArray());
                                strMiddlename = string.Empty;
                            }
                        }
                    }

                }

            }
            catch(Exception ex)
            {
                strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }

        //LastName FirstName MI
        public void fn_separateLastNameFirstNameV3(string valueFullName, out string strLastname, out string strFirstName, out string strMiddleName)
        {
            strFirstName = string.Empty; strLastname = string.Empty; strMiddleName = string.Empty;
            List<string> listFirstName = new List<string>();
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            string out_suffix = "";
            #region
            //Middle name is in the last part of the full name e.g LABUDAHON LOURDES S
            #endregion
            try
            {
                valueFullName = valueFullName.TrimEnd();
                var names = valueFullName.Split();
                Console.WriteLine(valueFullName);
                if(names.Length == 1)
                {
                    strLastname = names [0];
                }
                else if(names.Length == 2)
                {
                    strLastname = names [0];
                    strFirstName = names [1];
                }
                else if(names.Length >= 3)
                {

                    if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                    {
                        if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                        {
                            strLastname = names [0] + " " + names [1];
                            names = names.Skip(1).ToArray();
                            strFirstName = String.Join(" ", names.ToArray().Take(names.Length - 1));
                            strMiddleName = names [names.Length - 1];

                        }
                        else
                        {
                            //LABUDAHON LOURDES S"
                            strLastname = names [0] + " " + names [1];
                            names = names.Skip(1).ToArray();
                            strFirstName = String.Join(" ", names.ToArray().Take(names.Length - 1));
                            strMiddleName = names [names.Length - 1];

                        }
                    }
                    else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV"))
                    {

                        if(!names [names.Length - 1].ToUpper().Contains("JR") || !names [names.Length - 1].ToUpper().Contains("JR.") || !names [names.Length - 1].ToUpper().Contains("SR") || !names [names.Length - 1].ToUpper().Contains("SR.") || !names [names.Length - 1].ToUpper().Contains("II") || !names [names.Length - 1].ToUpper().Contains("III") || !names [names.Length - 1].ToUpper().Contains("IV"))
                        {
                            strLastname = names [0];
                            strMiddleName = names [names.Length - 1];
                            int intMiddleName = strMiddleName.Length;
                            if(strMiddleName.Contains(".") || intMiddleName == 1)
                            {
                                strMiddleName = names [names.Length - 1];
                                //REYES III BENJAMIN E
                                foreach(string suffix in Suffix)
                                {
                                    foreach(string name in names)
                                    {
                                        if(name == suffix)
                                        {
                                            out_suffix = suffix;
                                            break;
                                        }
                                        else
                                        {
                                            listFirstName.Add(name);
                                            continue;
                                        }
                                    }
                                }
                                names = names.Skip(2).ToArray();
                                strFirstName = string.Join(" ", listFirstName) + " " + out_suffix;

                                //strFirstName = String.Join(" ", names.ToArray().Take(names.Length - 1));
                            }
                            else
                            {
                                strMiddleName = null;
                                names = names.Skip(1).ToArray();
                                strFirstName = String.Join(" ", names);
                            }

                        }
                        else
                        {
                            strLastname = names [0];
                            names = names.Skip(1).ToArray();
                            strFirstName = String.Join(" ", names);
                        }


                    }
                    else
                    {


                        strLastname = names [0];
                        strMiddleName = names [names.Length - 1];
                        int intMiddleName = strMiddleName.Length;
                        if(strMiddleName.Contains(".") || intMiddleName == 1)
                        {
                            strMiddleName = names [names.Length - 1];
                            names = names.Skip(1).ToArray();
                            strFirstName = String.Join(" ", names.ToArray().Take(names.Length - 1));
                        }
                        else
                        {
                            strMiddleName = null;
                            names = names.Skip(1).ToArray();
                            strFirstName = String.Join(" ", names);
                        }
                    }

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + " " + Environment.NewLine + "Column has to no Firstname or Lastname");
            }

        }


        //LastName, Firstname, MI
        public void fn_separateLastNameFirstNameV4(string valueFullName, out string strFullname, out string strLastname, out string strFirstname, out string strMiddlename)
        {

            #region NOTES
            // Name is delimited by comma e.g BALONSONG, ALVIN, ADVINCULA
            #endregion

            strFirstname = string.Empty; strLastname = string.Empty; strMiddlename = string.Empty; strFullname = string.Empty;
            try
            {
                string strSuffix = string.Empty;
                bool bolSuffixFound = false;
                bool bolMIFound = false;
                strFullname = valueFullName;
                //string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA" , "LA" };
                //string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI","VII" };
                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");



                if(valueFullName.Contains(","))
                {
                    var names = valueFullName.Split(',');
                    var LN = names [0].Split(' ');
                    var FN = names [1].Trim().Split(' ');




                    Console.WriteLine(valueFullName);
                    if(names.Length == 1)
                    {
                        strLastname = names [0];
                    }
                    else if(names.Length == 2)
                    {
                        if(names [0].ToUpper().Contains("DE") || names [0].ToUpper().Contains("DEL.") || names [0].ToUpper().Contains("DELA") || names [0].ToUpper().Contains("DELOS") || names [0].ToUpper().Contains("LAS") || names [0].ToUpper().Contains("SAN") || names [0].ToUpper().Contains("STA") || names [0].ToUpper().Contains("STA.") || names [0].ToUpper().Contains("STO") || names [0].ToUpper().Contains("STO.") || names [0].ToUpper().Contains("SANTO") || (strLastname.ToUpper().Contains("SANTA") || strLastname.ToUpper().Contains("LA")))
                        {
                            foreach(string lnsuffix in LNSuffix)
                            {
                                if(bolSuffixFound == false)
                                {
                                    foreach(string ln in LN)
                                    {
                                        if(ln.ToUpper() == lnsuffix)
                                        {
                                            strLastname = names [0];
                                            strFirstname = names [1];
                                            bolSuffixFound = true;
                                            break;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            if(bolSuffixFound == false)
                            {
                                strLastname = names [0];
                                strFirstname = names [1];
                            }

                        }
                        else if(names [0].ToUpper().Contains("JR") || names [0].ToUpper().Contains("JR.") || names [0].ToUpper().Contains("SR") || names [0].ToUpper().Contains("SR.") || names [0].ToUpper().Contains("II") || names [0].ToUpper().Contains("III") || names [0].ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                        {

                            foreach(string suffix in Suffix)
                            {
                                if(bolSuffixFound == false)
                                {

                                    foreach(string ln in LN)
                                    {
                                        if(ln.ToUpper() == suffix)
                                        {
                                            strSuffix = suffix.ToString();
                                            strLastname = String.Join(" ", LN.ToArray().Take(LN.Length - 1));
                                            strFirstname = names [names.Length - 1] + " " + strSuffix;
                                            bolSuffixFound = true;
                                            break;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }
                        else
                        {
                            strLastname = names [0];
                            strFirstname = names [1];
                        }

                    }

                    else if(names.Length >= 3)//CHAM, DENZELL L.
                    {
                        if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA") || valueFullName.ToUpper().Contains("LA")))
                        {
                            if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                            {
                                //DE LARA JR, JAMES, TANAYAN
                                foreach(string lnsuffix in LNSuffix)
                                {
                                    if(bolSuffixFound == false)
                                    {
                                        if(names [0].ToUpper().Contains(lnsuffix))
                                        {
                                            var lastName = names [0].Split();

                                            foreach(string suffix in Suffix)
                                            {
                                                if(bolSuffixFound == false)
                                                {
                                                    foreach(string item in lastName)
                                                    {

                                                        if(item == suffix)
                                                        {
                                                            strLastname = String.Join(" ", lastName.ToArray().Take(lastName.Length - 1));
                                                            strSuffix = suffix;
                                                            strMiddlename = names [names.Length - 1];
                                                            strFirstname = names [1] + " " + strSuffix;
                                                            bolSuffixFound = true;
                                                            break;

                                                        }
                                                        else
                                                        {
                                                            continue;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            continue;

                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            foreach(string suffix in Suffix)
                            {
                                if(bolSuffixFound == false)
                                {

                                    if(names [0].ToUpper().Contains(suffix))
                                    {
                                        if(bolSuffixFound == false)
                                        {
                                            var lastname = names [0].Split();
                                            foreach(var item in lastname)
                                            {
                                                if(item == suffix)
                                                {
                                                    if(lastname.Length == 2)
                                                    {
                                                        strSuffix = suffix;
                                                        strLastname = lastname [0];
                                                        names = names.Skip(1).ToArray();
                                                        strFirstname = String.Join(" ", names.ToArray().Take(lastname.Length - 1)) + " " + strSuffix;
                                                        strMiddlename = names [names.Length - 1];
                                                        bolSuffixFound = true;
                                                        break;
                                                    }
                                                    else if(lastname.Length > 2)
                                                    {
                                                        strSuffix = lastname [lastname.Length - 1];
                                                        strLastname = String.Join(" ", lastname.ToArray().Take(lastname.Length - 1));
                                                        names = names.Skip(1).ToArray();
                                                        strFirstname = String.Join(" ", names.ToArray());
                                                        strMiddlename = names [names.Length - 1];
                                                        bolSuffixFound = true;
                                                        break;
                                                    }

                                                }
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }

                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            if(bolSuffixFound == false)
                            {
                                strLastname = names [0];
                                strFirstname = names [1];
                                strMiddlename = names [2];
                            }

                            else if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA") || valueFullName.ToUpper().Contains("LA")))
                            {
                                if(bolSuffixFound == false)
                                {
                                    //DE LA CRUZ, LOURDES, TIRADOR
                                    foreach(string lnsuffix in LNSuffix)
                                    {
                                        if(bolSuffixFound == false)
                                        {
                                            if(names [0].ToUpper().Contains(lnsuffix))
                                            {
                                                var lastName = names [0].Split();

                                                foreach(string item in lastName)
                                                {
                                                    if(item == lnsuffix)
                                                    {

                                                        strLastname = lastName [0];
                                                        bolSuffixFound = true;
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                continue;

                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    strLastname = names [0];
                                    strFirstname = names [1];
                                    strMiddlename = names [2];
                                }
                            }
                        }
                        else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                        {
                            //ENRICO JR, CELEDONIO, CASTILLO
                            //SEBASTIAN JR, ORLANDO, SABILE
                            var lastName = names [0].Split();
                            foreach(string suffix in Suffix)
                            {
                                if(bolSuffixFound == false)
                                {
                                    if(names [0].ToUpper().Contains(suffix))
                                    {


                                        foreach(string item in lastName)
                                        {
                                            if(item == suffix)
                                            {
                                                strSuffix = item;
                                                foreach(string lastname in lastName)
                                                {
                                                    if(bolSuffixFound == false)
                                                    {

                                                        foreach(string lnsuffix in LNSuffix)
                                                        {
                                                            if(lastname.ToUpper().Contains(lnsuffix))
                                                            {
                                                                strLastname = lastName [0] + " " + lastName [1];
                                                                bolSuffixFound = true;
                                                                strMiddlename = names [names.Length - 1];
                                                                names = names.Skip(1).ToArray();
                                                                strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1)) + " " + strSuffix;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                continue;

                                                            }
                                                        }
                                                    }
                                                    else
                                                    { break; }
                                                }
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        continue;

                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            if(bolSuffixFound == false)
                            {

                                foreach(string suffix in Suffix)
                                {
                                    if(bolSuffixFound == false)
                                    {
                                        foreach(string ln in lastName)
                                        {
                                            if(ln.ToUpper() == suffix)
                                            {
                                                strSuffix = suffix;
                                                strLastname = String.Join(" ", lastName.ToArray().Take(lastName.Length - 1));

                                                strFirstname = names [1] + " " + strSuffix;
                                                strMiddlename = names [names.Length - 1];
                                                bolSuffixFound = true;
                                                break;
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                            }
                            if(bolSuffixFound == false)
                            {
                                strLastname = names [0];
                                names = names.Skip(1).ToArray();
                                strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1)) + " " + strSuffix;
                                strMiddlename = names [names.Length - 1];
                            }
                        }

                        else
                        {
                            strLastname = names [0];
                            strMiddlename = names [names.Length - 1];
                            names = names.Skip(1).ToArray();
                            strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                        }


                    }
                }
                else
                {
                    //JOSHUA LEONARD N CRUZ
                    //JENNIFER K NGO
                    //ALBERTO P BATO JR

                    var names = valueFullName.Split();
                    int intMI = 1;
                    int intCountLength = names.Length;
                    int intcount = 0;

                    if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                    {
                        if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                        {
                            foreach(string name in names)
                            {
                                if(intMI == name.Length)
                                {
                                    strMiddlename = name;

                                    foreach(string item in names)
                                    {
                                        foreach(string suffix in Suffix)
                                        {
                                            if(bolSuffixFound == false)
                                            {

                                                if(item == suffix)
                                                {
                                                    //strSuffix = suffix;
                                                    List<string> listFirstName = new List<string>();

                                                    foreach(string charName in names)
                                                    {
                                                        intcount++;
                                                        if(charName.Length != 1)
                                                        {
                                                            listFirstName.Add(charName);
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            strFirstname = string.Join(" ", listFirstName) + " " + suffix;
                                                            names = names.Skip(intcount).ToArray();
                                                            strLastname = string.Join(" ", names.ToArray().Take(names.Length - 1));
                                                            bolSuffixFound = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                        }
                        else
                        {
                            foreach(string item in names)
                            {
                                if(bolSuffixFound == false)
                                {
                                    int intMILengthChecker = item.Length;
                                    if(intMI == intMILengthChecker || item.Contains("."))
                                    {
                                        strMiddlename = item;
                                        List<string> listFirstName = new List<string>();
                                        foreach(string charName in names)
                                        {
                                            if(charName.Length != 1)
                                            {
                                                listFirstName.Add(charName);

                                                continue;


                                            }
                                            else
                                            {
                                                strFirstname = string.Join(" ", listFirstName);
                                                strLastname = names [names.Length - 1];
                                                bolSuffixFound = true;
                                                break;
                                            }
                                        }
                                        //strFirstname = names [0] + " " + names [1];
                                        //strLastname = names [names.Length - 1];
                                        //bolSuffixFound = true;
                                        //break;
                                    }
                                    else { continue; }
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }
                        if(bolSuffixFound == false)
                        {
                            var firstname = names [0];

                            foreach(string suffix in Suffix)
                            {
                                if(firstname == suffix)
                                {
                                    strFirstname = firstname + " " + names [1];
                                    names = names.Skip(2).ToArray();
                                    strMiddlename = names [names.Length - 1];
                                    names = names.Skip(1).ToArray();
                                    strFirstname = string.Join(" ", names);
                                    bolSuffixFound = true;
                                    break;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        //SHAINA MAE DELOS SANTOS
                        foreach(string item in names)
                        {
                            if(bolSuffixFound == false)
                            {

                                if(intMI == item.Length || item.Contains("."))
                                {
                                    strMiddlename = item;
                                    List<string> listFirstName = new List<string>();
                                    foreach(string charName in names)
                                    {
                                        if(charName.Length != 1)
                                        {
                                            listFirstName.Add(charName);

                                            continue;


                                        }
                                        else
                                        {
                                            strFirstname = string.Join(" ", listFirstName);
                                            strLastname = names [names.Length - 1];
                                            bolSuffixFound = true;
                                            break;
                                        }
                                    }
                                }
                                else { continue; }
                            }
                            else
                            {
                                break;
                            }
                        }

                        foreach(string lnsuffix in LNSuffix)
                        {
                            if(bolSuffixFound == false)
                            {
                                foreach(string name in names)
                                {
                                    if(name.ToUpper() == lnsuffix)
                                    {
                                        strFirstname = names [0] + " " + names [1];
                                        names = names.Skip(2).ToArray();
                                        strLastname = String.Join(" ", names.ToArray());
                                        bolSuffixFound = true;
                                        break;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }


                    }
                    else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                    {
                        foreach(string name in names)
                        {
                            intCountLength = name.Length;
                            if(bolSuffixFound == false)
                            {
                                if(intMI == intCountLength)
                                {
                                    strMiddlename = name;
                                    foreach(string suffix in Suffix)
                                    {
                                        if(bolSuffixFound == false)
                                        {
                                            foreach(string item in names)
                                            {
                                                if(suffix == item)
                                                {
                                                    if(names.Length > 4)
                                                    {
                                                        strFirstname = names [0] + " " + names [1] + " " + suffix;
                                                        names = names.Skip(3).ToArray();
                                                        strLastname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                                        bolSuffixFound = true;
                                                        break;
                                                    }

                                                    else if(names.Length <= 4)
                                                    {
                                                        strFirstname = names [0] + " " + suffix;
                                                        names = names.Skip(2).ToArray();
                                                        strLastname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                                                        bolSuffixFound = true;
                                                        break;
                                                    }

                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    else if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA")))
                    {

                    }
                    else
                    {

                        if(intCountLength == 3)
                        {
                            foreach(string item in names)
                            {
                                if(bolSuffixFound == false)
                                {
                                    int intMILengthChecker = item.Length;
                                    if(intMI == intMILengthChecker || item.Contains("."))
                                    {
                                        strMiddlename = item;
                                        strFirstname = names [0];
                                        strLastname = names [names.Length - 1];
                                        bolSuffixFound = true;
                                        break;
                                    }
                                    else { continue; }
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        else if(intCountLength > 3)
                        {
                            foreach(string item in names)
                            {
                                if(bolSuffixFound == false)
                                {

                                    if(intMI == item.Length || item.Contains("."))
                                    {
                                        strMiddlename = item;
                                        List<string> listFirstName = new List<string>();
                                        foreach(string charName in names)
                                        {
                                            if(charName.Length != 1)
                                            {
                                                listFirstName.Add(charName);

                                                continue;


                                            }
                                            else
                                            {
                                                strFirstname = string.Join(" ", listFirstName);
                                                strLastname = names [names.Length - 1];
                                                bolSuffixFound = true;
                                                break;
                                            }
                                        }
                                    }
                                    else { continue; }
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }
                    }
                }
            }
            catch(Exception ex)
            {
                strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }


        //LastName, FirtName, Middle Name
        public void fn_separateLastNameFirstNameV5(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {

            #region NOTES
            // Name is delimited by comma, some suffix is in the last name , middle initial is in the last part of the fullname  e.g MICLAT JR,TITUS,R
            #endregion

            strFirstname = string.Empty; strLastname = string.Empty; strMiddlename = string.Empty;
            try
            {

                #region
                //full name is delimited by comma some with suffix, middle initial is in the last part of the name
                #endregion
                string strSuffix = string.Empty;
                bool bolSuffixFound = false;
                bool bolMIFound = false;
                //string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
                //string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");


                var names = valueFullName.Split(',');
                var LN = names [0].Trim().Split(' ');

                Console.WriteLine(valueFullName);
                if(names.Length == 1)
                {
                    strLastname = names [0];
                }
                else if(names.Length == 2)
                {
                    if(names [0].ToUpper().Contains("DE") || names [0].ToUpper().Contains("DEL.") || names [0].ToUpper().Contains("DELA") || names [0].ToUpper().Contains("DELOS") || names [0].ToUpper().Contains("LAS") || names [0].ToUpper().Contains("SAN") || names [0].ToUpper().Contains("STA") || names [0].ToUpper().Contains("STA.") || names [0].ToUpper().Contains("STO") || names [0].ToUpper().Contains("STO.") || names [0].ToUpper().Contains("SANTO") || (strLastname.ToUpper().Contains("SANTA") || strLastname.ToUpper().Contains("LA")))
                    {
                        foreach(string lnsuffix in LNSuffix)
                        {
                            if(bolSuffixFound == false)
                            {
                                foreach(string ln in LN)
                                {
                                    if(ln.ToUpper() == lnsuffix)
                                    {
                                        strLastname = names [0];
                                        strFirstname = names [1];
                                        bolSuffixFound = true;
                                        break;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        if(bolSuffixFound == false)
                        {
                            strLastname = names [0];
                            strFirstname = names [1];
                        }

                    }
                    else if(names [0].ToUpper().Contains("JR") || names [0].ToUpper().Contains("JR.") || names [0].ToUpper().Contains("SR") || names [0].ToUpper().Contains("SR.") || names [0].ToUpper().Contains("II") || names [0].ToUpper().Contains("III") || names [0].ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                    {

                        foreach(string suffix in Suffix)
                        {
                            if(bolSuffixFound == false)
                            {

                                foreach(string ln in LN)
                                {
                                    if(ln.ToUpper() == suffix)
                                    {
                                        strSuffix = suffix.ToString();
                                        strLastname = String.Join(" ", LN.ToArray().Take(LN.Length - 1));
                                        strFirstname = names [names.Length - 1] + " " + strSuffix;
                                        bolSuffixFound = true;
                                        break;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                break;
                            }
                        }

                    }
                    else
                    {
                        strLastname = names [0];
                        strFirstname = names [1];
                    }

                }

                else if(names.Length >= 3)
                {
                    if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA") || valueFullName.ToUpper().Contains("LA")))
                    {
                        if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || valueFullName.ToUpper().Contains("VII"))
                        {
                            //DE LARA JR, JAMES, TANAYAN
                            foreach(string lnsuffix in LNSuffix)
                            {
                                if(bolSuffixFound == false)
                                {
                                    if(names [0].ToUpper().Contains(lnsuffix))
                                    {
                                        var lastName = names [0].Split();

                                        foreach(string suffix in Suffix)
                                        {
                                            if(bolSuffixFound == false)
                                            {
                                                foreach(string item in lastName)
                                                {

                                                    if(item == suffix)
                                                    {
                                                        strLastname = String.Join(" ", lastName.ToArray().Take(lastName.Length - 1));
                                                        strSuffix = suffix;
                                                        strMiddlename = names [names.Length - 1];
                                                        strFirstname = names [1] + " " + strSuffix;
                                                        bolSuffixFound = true;
                                                        break;

                                                    }
                                                    else
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        continue;

                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        foreach(string suffix in Suffix)
                        {
                            if(bolSuffixFound == false)
                            {

                                if(names [0].ToUpper().Contains(suffix))
                                {
                                    if(bolSuffixFound == false)
                                    {
                                        var lastname = names [0].Split();
                                        foreach(var item in lastname)
                                        {
                                            if(item == suffix)
                                            {
                                                if(lastname.Length == 2)
                                                {
                                                    strSuffix = suffix;
                                                    strLastname = lastname [0];
                                                    names = names.Skip(1).ToArray();
                                                    strFirstname = String.Join(" ", names.ToArray().Take(lastname.Length - 1)) + " " + strSuffix;
                                                    strMiddlename = names [names.Length - 1];
                                                    bolSuffixFound = true;
                                                    break;
                                                }
                                                else if(lastname.Length > 2)
                                                {
                                                    strSuffix = lastname [lastname.Length - 1];
                                                    strLastname = String.Join(" ", lastname.ToArray().Take(lastname.Length - 1));
                                                    names = names.Skip(1).ToArray();
                                                    strFirstname = String.Join(" ", names.ToArray());
                                                    strMiddlename = names [names.Length - 1];
                                                    bolSuffixFound = true;
                                                    break;
                                                }

                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }

                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        if(bolSuffixFound == false)
                        {
                            strLastname = names [0];
                            strFirstname = names [1];
                            strMiddlename = names [2];
                        }

                        else if(valueFullName.ToUpper().Contains("DE") || valueFullName.ToUpper().Contains("DEL.") || valueFullName.ToUpper().Contains("DELA") || valueFullName.ToUpper().Contains("DELOS") || valueFullName.ToUpper().Contains("LAS") || valueFullName.ToUpper().Contains("SAN") || valueFullName.ToUpper().Contains("STA") || valueFullName.ToUpper().Contains("STA.") || valueFullName.ToUpper().Contains("STO") || valueFullName.ToUpper().Contains("STO.") || valueFullName.ToUpper().Contains("SANTO") || (valueFullName.ToUpper().Contains("SANTA") || valueFullName.ToUpper().Contains("LA")))
                        {
                            if(bolSuffixFound == false)
                            {
                                //DE LA CRUZ, LOURDES, TIRADOR
                                foreach(string lnsuffix in LNSuffix)
                                {
                                    if(bolSuffixFound == false)
                                    {
                                        if(names [0].ToUpper().Contains(lnsuffix))
                                        {
                                            var lastName = names [0].Split();

                                            foreach(string item in lastName)
                                            {
                                                if(item == lnsuffix)
                                                {

                                                    strLastname = lastName [0];
                                                    bolSuffixFound = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            continue;

                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                strLastname = names [0];
                                strFirstname = names [1];
                                strMiddlename = names [2];
                            }
                        }
                    }
                    else if(valueFullName.ToUpper().Contains("JR") || valueFullName.ToUpper().Contains("JR.") || valueFullName.ToUpper().Contains("SR") || valueFullName.ToUpper().Contains("SR.") || valueFullName.ToUpper().Contains("II") || valueFullName.ToUpper().Contains("III") || valueFullName.ToUpper().Contains("IV") || names [0].ToUpper().Contains("VI") || names [0].ToUpper().Contains("VII"))
                    {
                        //ENRICO JR, CELEDONIO, CASTILLO
                        //SEBASTIAN JR, ORLANDO, SABILE
                        var lastName = names [0].Split();
                        foreach(string suffix in Suffix)
                        {
                            if(bolSuffixFound == false)
                            {
                                if(names [0].ToUpper().Contains(suffix))
                                {


                                    foreach(string item in lastName)
                                    {
                                        if(item == suffix)
                                        {
                                            strSuffix = item;
                                            foreach(string lastname in lastName)
                                            {
                                                if(bolSuffixFound == false)
                                                {

                                                    foreach(string lnsuffix in LNSuffix)
                                                    {
                                                        if(lastname.ToUpper().Contains(lnsuffix))
                                                        {
                                                            strLastname = lastName [0] + " " + lastName [1];
                                                            bolSuffixFound = true;
                                                            strMiddlename = names [names.Length - 1];
                                                            names = names.Skip(1).ToArray();
                                                            strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1)) + " " + strSuffix;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            continue;

                                                        }
                                                    }
                                                }
                                                else
                                                { break; }
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }

                                }
                                else
                                {
                                    continue;

                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        if(bolSuffixFound == false)
                        {

                            foreach(string suffix in Suffix)
                            {
                                if(bolSuffixFound == false)
                                {
                                    foreach(string ln in lastName)
                                    {
                                        if(ln.ToUpper() == suffix)
                                        {
                                            strSuffix = suffix;
                                            strLastname = String.Join(" ", lastName.ToArray().Take(lastName.Length - 1));

                                            strFirstname = names [1] + " " + strSuffix;
                                            strMiddlename = names [names.Length - 1];
                                            bolSuffixFound = true;
                                            break;
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }

                        }
                        if(bolSuffixFound == false)
                        {
                            strLastname = names [0];
                            names = names.Skip(1).ToArray();
                            strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1)) + " " + strSuffix;
                            strMiddlename = names [names.Length - 1];
                        }
                    }

                    else
                    {
                        strLastname = names [0];
                        strMiddlename = names [names.Length - 1];
                        names = names.Skip(1).ToArray();
                        strFirstname = String.Join(" ", names.ToArray().Take(names.Length - 1));
                    }


                }


            }
            catch(Exception ex)
            {
                valueFullName = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }

        //FirstName MI LastName or  LastName, FirstName Suffix MI.
        public void fn_separateLastNameFirstNameV6(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {
                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");
                string out_suffix = string.Empty;
                List<string> listFirstName = new List<string>();

                //SY, JAN WYNTON C.
                if(valueFullName.Contains(",")) //If name has comma
                {
                    var names = valueFullName.Split(',');
                    strLastname = names [0].Trim();

                    if(names [1].Length == 2)
                    {
                        names = names.Skip(1).ToArray();
                        strFirstname = String.Join(" ", names);
                    }
                    else
                    {
                        names = names.Skip(1).ToArray();
                        names = names [0].Split(' ');
                        strMiddlename = fn_removeCharacters(valueFullName.Split().Last());
                        names = names.Take(names.Length - 1).ToArray();
                        strFirstname = String.Join(" ", names);
                    }


                }
                else
                {
                    var names = valueFullName.Split(' ');
                    strLastname = valueFullName.Split().Last();
                    strFirstname = valueFullName.Split().First();


                    foreach(string suffix in Suffix)
                    {
                        if(suffix == strLastname)
                        {
                            out_suffix = valueFullName.Split().Last();
                            break;
                        }
                        else
                        {
                            strLastname = valueFullName.Split().Last();
                            continue;
                        }
                    }

                    foreach(string name in names)
                    {
                        if(name.Length != 1 && !name.Contains("."))
                        {
                            if(name != strLastname)
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    strFirstname = string.Join(" ", listFirstName) + " " + out_suffix;
                    strMiddlename = fn_getmiddlename(valueFullName);

                    names = names.Skip(2).ToArray();
                    if(!string.IsNullOrEmpty(out_suffix))
                    {
                        strLastname = String.Join(" ", names).Replace(out_suffix, " ");
                    }



                    //REMIGIO B BONIFACIO JR

                }

            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                //strFirstname = fn_checkFirstname(strFirstname);
                //strLastname = fn_checkLastname(strLastname);
                //strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }

        //Lastname Suffix FirstName MI.
        public void fn_separateLastNameFirstNameV7(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {

                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");
                string out_suffix = string.Empty;
                string out_lsuffix = string.Empty;
                //string out_lsuffix_ = string.Empty;
                bool bolSuffix = false;
                List<string> listFirstName = new List<string>();

                var names = valueFullName.Split(' ');
                var out_lsuffix_ = valueFullName.Split(' ');
                strMiddlename = fn_getmiddlename(valueFullName);
                //SAN PEDRO BEAU D

                foreach(string name in names) //get last name  ahd check for suffix
                {
                    if (bolSuffix == true)
                    {
                        break;
                    }
                    foreach(string lsuffix in LNSuffix)
                    {
                        if(name == lsuffix)
                        {
                            out_lsuffix = name + " " + out_lsuffix_[1];
                            strLastname = out_lsuffix;
                            bolSuffix = true;
                            break;
                        }

                        else
                        {
                            strLastname = valueFullName.Split().First();
                            names = names.Skip(1).ToArray();
                            continue;
                        }
                    }
                }

                if (bolSuffix == true)
                {
                    names = valueFullName.Split(' ');
                    names = names.Skip(2).ToArray();

                    foreach(string name in names)
                    {
                        if(name.Length != 1 && !name.Contains("."))
                        {
                            if(name != strLastname)
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else
                {
                    names = valueFullName.Split(' ');
                    names = names.Skip(1).ToArray();

                    if (strMiddlename != null)
                    {
                        names = names.Take(names.Length - 1).ToArray();
                    }
                    //REYES III BENJAMIN E
                    foreach(string name in names)
                    {
                        foreach(string suffix in Suffix) //check for suffixes
                        {
                            if(name == suffix)  
                            {
                                out_suffix = suffix;
                                bolSuffix = true;
                                break;
                            }
                            else { continue; }
                        }
                        if(name.Length != 1 && !name.Contains("."))
                        {
                            if(name != strLastname && name != out_suffix) //if its not equal to lastname therefore its a firstname
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                           
                        }
                        else
                        {

                            break;
                        }
                    }
                }
                strFirstname = String.Join(" ", listFirstName) + " " + out_suffix;
 
            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }


        //Lastname suffix, FirstName MI.
        public void fn_separateLastNameFirstNameV8(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {
                if (valueFullName == null)
                {
                    strFirstname = fn_checkFirstname(strFirstname);
                    strLastname = fn_checkLastname(strLastname);
                    strMiddlename = fn_checkMiddleName(strMiddlename);
                }
                else
                {
                    valueFullName = valueFullName.Trim();
                    valueFullName = valueFullName.Replace("  ", " ");
                    string out_suffix = "";

                    bool bolSuffix = false;
                    string out_lsuffix = "";
                    List<string> listFirstName = new List<string>();
                    List<string> listFirstLastName = new List<string>();

                    var names = valueFullName.Split(',');
                    strMiddlename = fn_getmiddlename(valueFullName);
                    if (!string.IsNullOrEmpty(strMiddlename))
                    {
                        names = names.Take(names.Length - 1).ToArray();
                    }
               
                    foreach (string lsuffx in LNSuffix) //get the lastname
                    {
                        foreach(string name in names)
                        {
                            if(name == lsuffx)
                            {
                                strLastname = lsuffx + " " + names.Skip(1).ToArray();
                                break;

                            }
                            else
                            {
                                strLastname = names [0].Trim();
                                continue;
                            }
                        }
                    }

                    names = strLastname.Split(' '); //check if lastname has Suffix
                    foreach(string suffix in Suffix)
                    {
                        foreach(string lname in names)
                        {
                            if(lname == suffix)
                            {
                                out_suffix = suffix;
                                strLastname = strLastname.Replace(lname, ""); //remove suffix if found
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }


                    names = valueFullName.Split(',');
                    names = names [1].Split(' ');
                    if(!string.IsNullOrEmpty(strLastname))
                    {
                        names = names.Skip(1).ToArray();
                    }
                    foreach(string name in names)
                    {
                        foreach(string suffix in Suffix) //check for suffixes
                        {
                            if(name == suffix)
                            {
                                out_suffix = suffix;
                                bolSuffix = true;
                                break;
                            }
                            else { continue; }
                        }
                    }


                    names = valueFullName.Split(',');
                    names = names [1].Split(' ');
                    if(!string.IsNullOrEmpty(strLastname))
                    {
                        names = names.Skip(1).ToArray();
                    }
                
                    foreach(string name in names)
                    {
                   
                        if(name.Length != 1 && !name.Contains(".") || name.ToUpper() == "MA.")
                        {
                            if(name != strLastname && name != out_suffix) //if its not equal to lastname therefore its a firstname
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                        }
                        else
                        {

                            break;
                        }
                    }

                   strFirstname = String.Join(" ", listFirstName) + " " + out_suffix;
                   
                }

            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }

        //Lastname, Firstname, MI
        public void fn_separateLastNameFirstNameV9(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = ""; strFirstname = ""; strMiddlename = "";
            try
            {
                var names = valueFullName.Split(',');

                strLastname = names.First();
                strFirstname = names [1];
                strMiddlename = names.Last();


            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }


        //Lastname suffix, FirstName MI
        public void fn_separateLastNameFirstNameV10(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {
                if(valueFullName == null)
                {
                    strFirstname = fn_checkFirstname(strFirstname);
                    strLastname = fn_checkLastname(strLastname);
                    strMiddlename = fn_checkMiddleName(strMiddlename);
                }
                else
                {
                    valueFullName = valueFullName.Trim();
                    valueFullName = valueFullName.Replace("  ", " ");
                    string out_suffix = "";

                    bool bolSuffix = false;
                    string out_lsuffix = "";
                    List<string> listFirstName = new List<string>();
                    List<string> listFirstLastName = new List<string>();

                    var names = valueFullName.Split(',');
                    strMiddlename = fn_getmiddlenameV2(valueFullName);
                    if(!string.IsNullOrEmpty(strMiddlename))
                    {
                        names = names.Take(names.Length - 1).ToArray();
                    }

                    foreach(string lsuffx in LNSuffix) //get the lastname
                    {
                        foreach(string name in names)
                        {
                            if(name == lsuffx)
                            {
                                strLastname = lsuffx + " " + names.Skip(1).ToArray();
                                break;

                            }
                            else
                            {
                                strLastname = names [0].Trim();
                                continue;
                            }
                        }
                    }

                    names = strLastname.Split(' '); //check if lastname has Suffix
                    foreach(string suffix in Suffix)
                    {
                        foreach(string lname in names)
                        {
                            if(lname == suffix)
                            {
                                out_suffix = suffix;
                                strLastname = strLastname.Replace(lname, ""); //remove suffix if found
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }


                    names = valueFullName.Split(',');
                    names = names [1].Split(' ');
                    if(!string.IsNullOrEmpty(strLastname))
                    {
                        names = names.Skip(1).ToArray();
                    }
                    foreach(string name in names)
                    {
                        foreach(string suffix in Suffix) //check for suffixes
                        {
                            if(name == suffix)
                            {
                                out_suffix = suffix;
                                bolSuffix = true;
                                break;
                            }
                            else { continue; }
                        }
                    }


                    names = valueFullName.Split(',');
                    names = names [1].Split(' ');
                    if(!string.IsNullOrEmpty(strLastname))
                    {
                        names = names.Skip(1).ToArray();
                    }

                    foreach(string name in names)
                    {

                        if(name.Length != 1 && !name.Contains(".") || name.ToUpper() == "MA.")
                        {
                            if(name != strLastname && name != out_suffix) //if its not equal to lastname therefore its a firstname
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                        }
                        else
                        {

                            break;
                        }
                    }

                    strFirstname = String.Join(" ", listFirstName) + " " + out_suffix;

                }

            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }

        //Lastname FirstName MI Suffix
        public void fn_separateLastNameFirstNameV11(string valueFullName, out string strLastname, out string strFirstname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {

                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");
                string out_suffix = string.Empty;
                string out_lsuffix = string.Empty;
                //string out_lsuffix_ = string.Empty;
                bool bolSuffix = false;
                List<string> listFirstName = new List<string>();

                var names = valueFullName.Split(' ');
                var out_lsuffix_ = valueFullName.Split(' ');
                strMiddlename = fn_getmiddlenameV2(valueFullName);
                //SAN PEDRO BEAU D

                foreach(string name in names) //get last name  ahd check for suffix
                {
                    if(bolSuffix == true)
                    {
                        break;
                    }
                    foreach(string lsuffix in LNSuffix)
                    {
                        if(name == lsuffix)
                        {
                            out_lsuffix = name + " " + out_lsuffix_ [1];
                            strLastname = out_lsuffix;
                            bolSuffix = true;
                            break;
                        }

                        else
                        {
                            strLastname = valueFullName.Split().First();
                            names = names.Skip(1).ToArray();
                            continue;
                        }
                    }
                }

                if(bolSuffix == true)
                {
                    names = valueFullName.Split(' ');
                    names = names.Skip(2).ToArray();

                    foreach(string name in names)
                    {
                        if(name.Length != 1 && !name.Contains("."))
                        {
                            if(name != strLastname)
                            {
                                listFirstName.Add(name);
                                continue;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else
                {
                    names = valueFullName.Split(' ');
                    names = names.Skip(1).ToArray();

                    //if(strMiddlename != null)
                    //{
                    //    names = names.Take(names.Length - 1).ToArray();
                    //}
                    //REYES III BENJAMIN E
                    foreach(string name in names)
                    {
                        foreach(string suffix in Suffix) //check for suffixes
                        {
                            if(name == suffix)
                            {
                                out_suffix = suffix;
                                bolSuffix = true;
                                break;
                            }
                            else { continue; }
                        }
                        if(name.Length != 1)
                        {
                            if(name != strLastname && name != out_suffix) //if its not equal to lastname therefore its a firstname
                            {
                                listFirstName.Add(name);
                                continue;
                            }

                        }
                        else
                        {

                            continue;
                        }
                    }
                }
                strFirstname = String.Join(" ", listFirstName) + " " + out_suffix;

            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }


        //FirstName  MI LastName Suffix
        public void fn_separateLastNameFirstNameV12(string valueFullName, out string strFirstname, out string strLastname, out string strMiddlename)
        {
            strLastname = string.Empty; strFirstname = string.Empty; strMiddlename = string.Empty;
            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA", "LA" };
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            try
            {
                bool bolnameSuffix = false;
                valueFullName = valueFullName.Trim();
                valueFullName = valueFullName.Replace("  ", " ");
                string out_suffix = string.Empty;
                string out_lsuffix = string.Empty;
                //string out_lsuffix_ = string.Empty;
                bool bolSuffix = false;
                List<string> listFirstName = new List<string>();
                int skipName = 0;

                var names = valueFullName.Split(' ');
                var out_lsuffix_ = valueFullName.Split(' ');
                strMiddlename = fn_getmiddlenameV2(valueFullName);

                 //Get SUFFIX
                foreach(string name in names)
                {
                    foreach(string suffix in Suffix) //check for suffixes
                    {
                        if(name == suffix)
                        {
                            out_suffix = suffix;
                            bolSuffix = true;
                            break;
                        }
                        else { continue; }
                    }
                }

                foreach(string name in names)//check if last has Lsuffix
                {
                    skipName++;

                    foreach(string lsuffix in LNSuffix)
                    {
                        if (name == lsuffix)
                        {
                            strLastname = String.Join(" ", name) + " " + name.Last();
                            break;
                        }
                        else
                        {
                            bolnameSuffix = false;
                        }
                        
                    }
                       
                }

                if (bolnameSuffix == false)
                {
                    strLastname = names.Last();
                    names = names.Take(names.Length - 1).ToArray();
                }
              
                foreach(string name in names)
                {
                  
                    if(name.Length > 2)
                    {
                        listFirstName.Add(name);
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
                strFirstname = String.Join(" ", listFirstName) + " " + out_suffix;

                //names = valueFullName.Split(' ');
                //names = names.Skip(skipName).ToArray();
                ////Get LASTNAME
                //foreach(string name in names)
                //{
                //    foreach(string Lsuffix in LNSuffix)
                //    {
                //        if(name == Lsuffix && name != out_suffix)
                //        {
                //            strLastname = String.Join(" ", Lsuffix) + " " + name.Last();
                //            bolnameSuffix = true;
                //            break;
                //        }

                //    }
                //}

                //if(bolnameSuffix != true)
                //{
                //    strLastname = names.Last();
                //}
            }
            catch(Exception ex)
            {
                //strFullname = fn_checkFullname(valueFullName);
                strFirstname = fn_checkFirstname(strFirstname);
                strLastname = fn_checkLastname(strLastname);
                strMiddlename = fn_checkMiddleName(strMiddlename);

            }
        }
        //Middle Name with .
        public string fn_getmiddlename(string Fullname)
        {
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            var names = Fullname.Split(' ');
            string MI = "";

            foreach(string name in names)
            {
                foreach(string suffix in Suffix)
                {
                    if(name != suffix && name.Length == 2 && name.Contains(".") || name.Length == 1)
                    {
                        if(name.Contains(".") && name != suffix)
                        {
                        MI = fn_removeCharacters(name);
                        return MI;
                        break;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return null;
        }

        //Middle Name with no .
        public string fn_getmiddlenameV2(string Fullname)
        {
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            var names = Fullname.Split(' ');
            string MI = "";

            foreach(string name in names)
            {
                foreach(string suffix in Suffix)
                {
                    if(name != suffix && name.Length == 2 && name.Contains(".") || name.Length == 1)
                    {
                       
                            MI = fn_removeCharacters(name);
                            return MI;
                            break;
                       
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return null;
        }

        public string fn_getsuffix(string Fullname)
        {
            string [] Suffix = { "IV", "JR", "JR.", "SR", "SR.", "II", "III", "V", "VI", "VII" };
            var names = Fullname.Split(' ');

            foreach(string name in names)
            {
                foreach(string suffix in Suffix)
                {
                    if(name == suffix && name.Length >= 2 && name.Contains("."))
                    {
                        return name.ToString();
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            return null;
        }


        /**************************************CLASS OF BUSINESS********************************************************
        PROGRAM: SICS 
        COLUMN: CLASS OF BUSINESS
        DESCRIPTION: Return the value of currency depending on the BM and value passed to the parameter

        FUNCTION NAME: fn_getcurrency
        BORDEREUX NO: BM061
        */

        public string fn_getcob(string valueCob)


        {
            switch(valueCob)
            {
                case "1":
                return "IND";

                default:
                return "GRP";
            }

            //return "GRP";
        }


        public void fn_getcobver2(string valueFileName, out string cob, out string transcode, out bool bolNegative)
        {
            transcode = string.Empty;
            cob = string.Empty;
            bolNegative = false;
            bool bolFile = true;
            string [] Filename = { "VOU201808LCR0246", "VOU201809LCR0175", "VOU201809LCR0169", "VOU201809LCR0181", "VOU201808LCR0253" };
            string [] Filename2 = { "VOU201809LCR0209", "VOU201808LCR0229", "VOU201807LCR0114" };
            string [] Filename3 = { "VOU201807LCR0123", "VOU201808LCR0239", "VOU201809LCR0210", "VOU201808LCR0243", "VOU201808LCR0228", "VOU201809LCR0208", "VOU201807LCR0113" };
            foreach(string file in Filename)
            {
                if(valueFileName.Contains(file))
                {
                    transcode = "TLAPSE";
                    cob = "GRP";
                    bolNegative = true;
                    break;
                }
                else
                {
                    bolFile = false;
                    continue;

                }
            }

            if(bolFile == false)
            {
                foreach(string item_ in Filename2)
                {
                    if(valueFileName.Contains(item_))
                    {
                        transcode = "TRENEW";
                        cob = "GRP";
                        bolNegative = false;
                        break;
                    }
                    else
                    {
                        bolFile = false;
                        continue;

                    }
                }
            }

            if(bolFile == false)
            {
                foreach(string item_ in Filename3)
                {
                    if(valueFileName.Contains(item_))
                    {
                        transcode = "TRENEW";
                        cob = "IND";
                        bolNegative = false;
                        break;
                    }
                    else
                    {
                        bolFile = false;
                        continue;
                    }
                }
            }

        }

        /**************************************CLASS OF BUSINESS********************************************************
        PROGRAM: SICS 
        COLUMN: INSURANCE_PRODUCT
        DESCRIPTION: Return the value of currency depending on the BM and value passed to the parameter

        FUNCTION NAME: fn_gettransactionproduct
        BORDEREUX NO: BM061
        */

        public string fn_gettransactionproduct(string Valuecob, string ValueCurrency, string ValueRisk)
        {
            string strInsuranceProduct = Valuecob + ValueCurrency + ValueRisk;
            if(strInsuranceProduct == "111")
            {
                return "TRADITIONALLIFE";
            }
            else if(strInsuranceProduct == "112")
            {
                return "CIRACIND";
            }
            else if(strInsuranceProduct == "113")
            {
                return "T&PD-LS-IND";
            }
            else if(strInsuranceProduct == "114")
            {
                return "RPAR";
            }
            else if(strInsuranceProduct == "115")
            {
                return "RPAR";
            }
            else if(strInsuranceProduct == "119" || strInsuranceProduct == "219")
            {
                return "WOP-DDI-IND";
            }
            else if(strInsuranceProduct == "121")
            {
                return "CREDITLIFE-GRP";
            }
            else if(strInsuranceProduct == "122")
            {
                return "CRITICALILLNESS";
            }
            else if(strInsuranceProduct == "123")
            {
                return "T&PD-LS-GRP";
            }
            else if(strInsuranceProduct == "124")
            {
                return "RENEWALPERSONAL";
            }
            else if(strInsuranceProduct == "125")
            {
                return "ADBR-GRP";
            }
            else if(strInsuranceProduct == "211")
            {
                return "TRADITIONALLIFE";
            }
            else if(strInsuranceProduct == "211")
            {
                return "TRADITIONALLIFE";
            }
            else if(strInsuranceProduct == "212")
            {
                return "CIRACIND";
            }
            else if(strInsuranceProduct == "213")
            {
                return "T&PD-LS-IND";
            }
            else if(strInsuranceProduct == "214")
            {
                return "RPAR";
            }
            else if(strInsuranceProduct == "215")
            {
                return "ADB-IND";
            }

            else if(strInsuranceProduct == "221")
            {
                return "CREDITLIFE-GRP";
            }
            else if(strInsuranceProduct == "222")
            {
                return "CRITICALILLNESS";
            }
            else if(strInsuranceProduct == "223")
            {
                return "T&PD-LS-GRP";
            }
            else if(strInsuranceProduct == "224")
            {
                return "RENEWALPERSONAL";
            }
            else if(strInsuranceProduct == "225")
            {
                return "ADBR-GRP";
            }
            else
            {
                return "TRADITIONALLIFE";
            }

        }

        public string fn_gettransactionproductV2(string valueRisk, string valueCOB)
        {
            if(valueRisk == "1" && valueCOB == "IND")
            {
                return "TRADITIONALLIFE";
            }
            else if(valueRisk == "2" && valueCOB == "IND")
            {
                return "CIRACIND";
            }
            else if(valueRisk == "3" && valueCOB == "IND")
            {
                return "T&PD-LS-GRP";
            }
            else if(valueRisk == "4" && valueCOB == "IND")
            {
                return "RPAR";
            }
            else if(valueRisk == "5" && valueCOB == "IND")
            {
                return "ADB-IND";
            }
            else if(valueRisk == "9" && valueCOB == "IND")
            {
                return "WOP-DDI-IND";
            }
            else if(valueRisk == "1" && valueCOB == "GRP")
            {
                return "CREDITLIFE-GRP";
            }
            else if(valueRisk == "2" && valueCOB == "GRP")
            {
                return "CRITICALILLNESS";
            }
            else if(valueRisk == "3" && valueCOB == "GRP")
            {
                return "T&PD-LS-GRP";
            }
            else if(valueRisk == "4" && valueCOB == "GRP")
            {
                return "RENEWALPERSONAL";
            }
            else if(valueRisk == "5" && valueCOB == "GRP")
            {
                return "ADBR-GRP";
            }
            else
            {
                return "TRADITIONALLIFE";
            }
        }

        /**************************************INSERT NAME and GENDER to Database ********************************************************
        PROGRAM: SICS 
        DESCRIPTION: Import data name and gender to Database dbo_gender
        FUNCTION NAME: fn_getcurrency
        */

        public void fn_searchnamesdb(string strFirstName, out string strGender, out string strAuthor)
        {
            strGender = ""; strAuthor = "";
            try
            {
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                string query = "SELECT * FROM dbo_gender WHERE firstname=" + "'" + strFirstName + "'";

                OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                OdbcDataReader DbReader = DbCommand.ExecuteReader();

                if(DbReader.Read())
                {
                    strGender = DbReader.GetValue(1).ToString();
                    strAuthor = DbReader.GetValue(3).ToString();
                }
                else
                {
                    strGender = "Name doesn't exist";
                    strAuthor = "";
                }

                DbReader.Close();
                cnDB.Dispose();
                cnDB.Close();
            }
            catch(Exception e)
            {
                strGender = "Name doesn't exist";
            }
        }

        public void fn_importmultiplenamesdb(string strMultiplenames, string User)
        {
            Helper objHlpr = new Helper();
            System.Data.DataTable objdt_template = new System.Data.DataTable();
            Microsoft.Office.Interop.Excel.Application eapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbraw = eapp.Workbooks.Open(strMultiplenames);
            Microsoft.Office.Interop.Excel.Worksheet wsraw = wbraw.Sheets ["Gender"];
            Microsoft.Office.Interop.Excel.Range rawrange = wsraw.UsedRange;


            //DataRow _var.dtworkRow01;
            int intLastRow = wsraw.Range ["A1"].End [XlDirection.xlDown].Row;


            for(int i = 2; i <= intLastRow; i++)
            {
                string FullName = Convert.ToString(wsraw.Range ["A" + i].Value);
            }

            objdt_template.Dispose();
            objdt_template = null;
            objHlpr.fn_killexcel();
            objHlpr = null;

        }
        public void fn_importnamegenderdb(string strFirstname, string strGender, string strAuthor)
        {
            string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1;" + "Encrypt=False";
            Console.WriteLine(strFirstname);
            try
            {
                string query = "INSERT INTO dbo_gender (firstname,gender,date_added,author) VALUES(" + "'" + strFirstname + "'" + "," + "'" + strGender + "'" + "," + "'" + DateTime.Now.ToString("dd-MMM-yy") + "'" + "," + "'" + strAuthor + "')";
                OdbcConnection cnDB = new OdbcConnection(Dbconnection);

                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                DbCommand.ExecuteNonQuery();

                cnDB.Dispose();
                cnDB.Close();
            }
            catch(Exception ex)
            {

                //var dialog = MessageBox.Show(strFirstname + " already exist in the database" + Environment.NewLine + Environment.NewLine + "Overwrite this name?", "Proceed Import", MessageBoxButtons.YesNo);
                //if(dialog == DialogResult.Yes)
                //{
                //string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                string query = "UPDATE dbo_gender SET firstname = " + "'" + strFirstname + "'" + "," + "gender =" + "'" + strGender + "'" + "," + "date_added =" + "'" + DateTime.Now.ToString("dd-MMM-yy") + "'" + "," + "author=" + "'" + strAuthor + "' " + "WHERE firstname = " + "'" + strFirstname + "'";

                OdbcConnection cnDB = new OdbcConnection(Dbconnection);

                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                DbCommand.ExecuteNonQuery();

                cnDB.Dispose();
                cnDB.Close();

                //}

            }
        }

        public void fn_getuserid(string username, out string strUserName)
        {
            strUserName = "";
            if(username == "avila.lr")
            {
                strUserName = "Beth";
            }
            else if(username == "cruz.bm")
            {
                strUserName = "Betsy";
            }
            else if(username == "platon.vm")
            {
                strUserName = "Vic";
            }
            else if(username == "robles.nm")
            {
                strUserName = "Neil";
            }
            else if(username == "yap.ma")
            {
                strUserName = "Michelle";
            }


        }

        /**************************************GET GENDER FROM Postresql Database********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Pass and return the value of from databse
        FUNCTION NAME: fn_getgenderv2
        FOR BORDEREAUX: 
        DATABASE: dbo_gender
        */
        public string fn_getgenderv2(string strFirstname, out string strSex)
        {
            strSex = "";
            try
            {
                objHlpr.fn_Getfirstname(strFirstname, out string strFirstName);
                string query = "SELECT * FROM dbo_gender WHERE firstname=" + "'" + strFirstName + "'";
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                //OdbcConnection cnDB = new OdbcConnection(szConnect);

                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                OdbcDataReader DbReader = DbCommand.ExecuteReader();

                if(DbReader.Read())
                {
                    strSex = DbReader.GetValue(1).ToString();
                    return strSex;

                }
                else
                {
                    Variables.boogenderfail = true;
                    strSex = "";
                    return strSex;
                }

                DbReader.Close();
                cnDB.Dispose();
                cnDB.Close();
            }
            catch(Exception ex)
            {

                return "M";
            }
        }

        /**************************************GET DATA FROM Postresql Database********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Pass and return the value  from database
        FUNCTION NAME: 
        FOR BORDEREAUX: BM048
        */
        public void fn_getFirstFinancialMacrodata(string valuePolNo, string valueFullName, out string strIssueAge,
        out string strMortality, out string strRefunding, out string strFullName, out string strFirstName, out string strLastName,
        out string strMiddlInitial, out string strTitle, out string strDOB, out string strSex, out string strLifeID,
        out string strRemarksCode, /*out string strCessionCode,*/ out string strBrandedProduct)

        {
            strIssueAge = ""; strMortality = ""; strRefunding = "";
            strFullName = ""; strLastName = ""; strFirstName = ""; strMiddlInitial = ""; strTitle = ""; strDOB = ""; strSex = "";
            strLifeID = "";
            strRemarksCode = ""; /*strCessionCode = "";*/ strBrandedProduct = "";


            try
            {

                if(!string.IsNullOrEmpty(valuePolNo))
                {
                    //FIRST LIFE FINANCIAL COMPANY, INC.
                    string query = "SELECT * FROM dbo_macro WHERE policy_no=" + "'" + valuePolNo + "'" + " AND " + "company_name= " + "'FIRST LIFE FINANCIAL COMPANY, INC.'";
                    string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                    OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();
                    DbReader.Read();

                    //strCessionCode = DbReader.GetValue(4).ToString();
                    strIssueAge = DbReader.GetValue(6).ToString();
                    strMortality = DbReader.GetValue(8).ToString();
                    strRefunding = DbReader.GetValue(9).ToString();
                    strDOB = DbReader.GetValue(11).ToString();
                    strSex = DbReader.GetValue(12).ToString();
                    strBrandedProduct = DbReader.GetValue(14).ToString();


                    strFullName = valueFullName;
                    fn_separateLastNameFirstNameV6(valueFullName, out strLastName, out strFirstName, out  strMiddlInitial);
                    //fn_separateLastNameFirstNameV2(valueFullName, out strFullName, out strFirstName, out strLastName, out strMiddlInitial);
                    //fn_separateFirstNameLastNameV2(valueFullName, out strFirstName, out strLastName, out strMiddleName);
                    strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                    strMortality = fn_getmortalityrating(strMortality);
                }
                else
                {
                    strRemarksCode = "BR6";
                    strLastName = "DummyLastName";
                    strFirstName = "DummyFirstName";
                    strMiddlInitial = "DummyMiddleName";
                    strFullName = "DummyFullName";
                    strLifeID = valuePolNo;
                    strSex = "M";
                    strDOB = "07/01/1900";
                    strMortality = "STANDARD";
                    strBrandedProduct = "LIFE";


                }
            }
            catch(Exception ex)
            {
                Variables.boomacrofail = true;
                strFullName = valueFullName;
                fn_separateLastNameFirstNameV6(valueFullName, out strLastName, out strFirstName, out strMiddlInitial);
                strDOB = "07/01/1900";
                strLifeID = objHlpr.fn_LifeID(strFirstName, strLastName, strDOB);
                strSex = objHlpr.fn_getgenderv2(strFirstName);
                strRemarksCode = "BR4";
                strMortality = "STANDARD";
                strBrandedProduct = "LIFE";
            }

        }



        /**************************************GET DATA FROM Postresql Database********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Pass and return the value  from database
        FUNCTION NAME: 
        FOR BORDEREAUX: BM033
        */



        public void fn_getmacro_prembord_umre(string dbName, string strValuePolNo, string strValueCertno, out string polno, out string Volume,
        out string ADB_Volume, out string SAR_Volume, out string SDI_Volume,
        out string life_ret, out string rid_ret, out string ADB_amt, out string SAR_amt, out string SARDI_amt,
        out string Firstname, out string MI, out string LastName, out string FullName, out string strFirstName,
        out string Sex,
        out string DOB,
        out string Mort, out string Attain_Age, out string Issue_Age, out string Lifeid,
        out string RemarksCode)

        {

            polno = "";
            Volume = ""; ADB_Volume = ""; SAR_Volume = ""; SDI_Volume = "";
            life_ret = ""; rid_ret = ""; ADB_amt = ""; SAR_amt = ""; SARDI_amt = "";
            Firstname = ""; LastName = ""; MI = ""; FullName = ""; strFirstName = ""; Sex = "";
            DOB = ""; string DOB_MO = ""; string DOB_DD = ""; string DOB_YR = "";
            Mort = ""; Attain_Age = ""; Issue_Age = ""; Lifeid = ""; RemarksCode = "";

            try
            {
                if(!string.IsNullOrEmpty(strValuePolNo))
                {

                    string query = "SELECT * FROM " + dbName + " WHERE policy_no=" + "'" + strValuePolNo + "'"
                    + " AND " + "certno =" + "'" + strValueCertno + "'";
                    string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                    OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();
                    DbReader.Read();

                    polno = DbReader.GetValue(0).ToString();
                    Volume = DbReader.GetValue(8).ToString();
                    ADB_Volume = DbReader.GetValue(9).ToString();
                    SAR_Volume = DbReader.GetValue(10).ToString();
                    SDI_Volume = DbReader.GetValue(11).ToString();

                    life_ret = DbReader.GetValue(12).ToString();
                    rid_ret = DbReader.GetValue(13).ToString();
                    ADB_amt = DbReader.GetValue(14).ToString();
                    SAR_amt = DbReader.GetValue(15).ToString();
                    SARDI_amt = DbReader.GetValue(16).ToString();

                    Firstname = DbReader.GetValue(28).ToString();
                    LastName = DbReader.GetValue(30).ToString();
                    FullName = LastName + " " + Firstname;
                    Sex = DbReader.GetValue(31).ToString();

                    DOB_MO = DbReader.GetValue(2).ToString();
                    DOB_DD = DbReader.GetValue(3).ToString();
                    DOB_YR = DbReader.GetValue(4).ToString();
                    DOB = DOB_MO + "/" + DOB_DD + "/" + DOB_YR;
                    objHlpr.fn_reformatDate(DOB);

                    Mort = DbReader.GetValue(7).ToString();
                    Mort = fn_getmortalityrating(Mort);
                    Attain_Age = DbReader.GetValue(32).ToString();
                    Issue_Age = DbReader.GetValue(5).ToString();
                    Lifeid = objHlpr.fn_LifeID(Firstname, LastName, DOB);
                    objHlpr.fn_Getfirstname(Firstname, out strFirstName);
                    DbReader.Close();
                    cnDB.Close();

                }
                else
                {
                    polno = strValuePolNo;
                    RemarksCode = "BR6";
                    LastName = "DummyLastName";
                    Firstname = "DummyFirstName";
                    MI = "DummyMiddleName";
                    FullName = "DummyFullName";
                    Lifeid = strValuePolNo;
                    Sex = "M";
                    DOB = "07/01/1900";

                    Volume = "0";
                    ADB_Volume = "0";
                    ADB_amt = "0";
                    SAR_Volume = "0";
                    SDI_Volume = "0";
                    life_ret = "0";
                    rid_ret = "0";
                    SARDI_amt = "0";
                    SAR_amt = "0";
                    Mort = "STANDARD";
                }

            }
            catch(Exception e)
            {

                try
                {
                    string query = "SELECT * FROM dbo_macro WHERE policy_no=" + "'" + strValuePolNo + "'"
                        + " AND cession_no=" + "'" + strValueCertno + "' " + "AND company_name LIKE 'INSULAR LIFE%'";
                    string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                    OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                    cnDB.Open();
                    OdbcCommand DbCommand = cnDB.CreateCommand();
                    DbCommand.CommandText = query;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();
                    DbReader.Read();

                    polno = DbReader.GetValue(3).ToString();
                    FullName = DbReader.GetValue(10).ToString();
                    Sex = DbReader.GetValue(12).ToString();
                    fn_separateLastNameFirstNameV2(FullName, out FullName, out LastName, out Firstname, out MI);
                    DOB = DbReader.GetValue(11).ToString();
                    Lifeid = objHlpr.fn_LifeID(Firstname, LastName, DOB);

                    Mort = DbReader.GetValue(8).ToString();
                    Mort = fn_getmortalityrating(Mort);
                    //Attain_Age = DbReader.GetValue(7).ToString();
                    Issue_Age = DbReader.GetValue(6).ToString();
                    life_ret = DbReader.GetValue(17).ToString();

                    Volume = "0";
                    ADB_Volume = "0";
                    SAR_Volume = "0";
                    SDI_Volume = "0";
                    life_ret = "0";
                    rid_ret = "0";
                    ADB_amt = "0";
                    SAR_amt = "0";
                    SARDI_amt = "0";

                    DbReader.Close();
                    cnDB.Close();
                }
                catch(Exception ex)
                {
                    polno = strValuePolNo;
                    FullName = "DummyFullName";
                    Firstname = "DummyFirstName";
                    LastName = "DummyLastName";
                    Sex = "M";
                    Mort = "STANDARD";
                    DOB = "07/01/1900";
                    Lifeid = polno;

                    Volume = "0";
                    ADB_Volume = "0";
                    ADB_amt = "0";
                    SAR_Volume = "0";
                    SDI_Volume = "0";
                    life_ret = "0";
                    rid_ret = "0";
                    SARDI_amt = "0";
                    SAR_amt = "0";

                }
            }
        }




        public void fn_searchInMacroDatabase(string valueSearch, string valueDbName)
        {

            string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
            string query = "SELECT * FROM " + valueDbName + " WHERE policy_no = '" + valueSearch + "'";
            OdbcConnection cnDB = new OdbcConnection(Dbconnection);

            cnDB.Open();
            OdbcCommand DbCommand = cnDB.CreateCommand();
            DbCommand.CommandText = query;
            OdbcDataReader DbReader = DbCommand.ExecuteReader();

            //macroDB.listView1.Columns.Clear();
            //macroDB.listView1.Items.Clear();
            //macroDB.listView1.GridLines = true;
            //macroDB.listView1.View = View.Details;
            //macroDB.listView1.Columns.Add("Cession No");
            //macroDB.listView1.Columns.Add("URC");
            //macroDB.listView1.Columns.Add("Policy No");
            //macroDB.listView1.Columns.Add("Cession Type Code");
            //macroDB.listView1.Columns.Add("Currency Code");
            //macroDB.listView1.Columns.Add("Issue Age");
            //macroDB.listView1.Columns.Add("Issue Date");
            //macroDB.listView1.Columns.Add("Mortality Rating Code");
            //macroDB.listView1.Columns.Add("Refunding Code");
            //macroDB.listView1.Columns.Add("NAME");
            //macroDB.listView1.Columns.Add("Date of Birth");
            //macroDB.listView1.Columns.Add("Gender");
            //macroDB.listView1.Columns.Add("Cover7c");
            //macroDB.listView1.Columns.Add("Benefit");
            //macroDB.listView1.Columns.Add("amt7insrd7a");
            //macroDB.listView1.Columns.Add("amt7reinsrd7a");
            //macroDB.listView1.Columns.Add("ced7retn7a");
            //macroDB.listView1.Columns.Add("Company Name");
            //macroDB.listView1.Columns.Add("Source Code");
            //macroDB.listView1.AutoResizeColumns((ColumnHeaderAutoResizeStyle.HeaderSize));

            while(DbReader.Read())
            {
                ListViewItem lv = new ListViewItem(DbReader [1].ToString());
                lv.SubItems.Add(DbReader [2].ToString());
                lv.SubItems.Add(DbReader [3].ToString());
                lv.SubItems.Add(DbReader [4].ToString());
                lv.SubItems.Add(DbReader [5].ToString());
                lv.SubItems.Add(DbReader [6].ToString());
                lv.SubItems.Add(DbReader [7].ToString());
                lv.SubItems.Add(DbReader [8].ToString());
                lv.SubItems.Add(DbReader [9].ToString());
                lv.SubItems.Add(DbReader [10].ToString());
                lv.SubItems.Add(DbReader [11].ToString());
                lv.SubItems.Add(DbReader [12].ToString());
                lv.SubItems.Add(DbReader [13].ToString());
                lv.SubItems.Add(DbReader [14].ToString());
                lv.SubItems.Add(DbReader [15].ToString());
                lv.SubItems.Add(DbReader [16].ToString());
                lv.SubItems.Add(DbReader [17].ToString());
                lv.SubItems.Add(DbReader [18].ToString());
                lv.SubItems.Add(DbReader [19].ToString());
                //macroDB.listView1.Items.Add(lv);
            }
            cnDB.Close();
            DbReader.Close();
        }


        /**************************************GET DATA FROM Postresql Database********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Returns the value of Mortality Rating
        FUNCTION NAME: fn_getmortalityrating
        FOR BORDEREAUX: BM061
        */
        public string fn_getmortalityrating(string Valuemortality)
        {
            if(Valuemortality == null)
            {
                return "STANDARD";
            }
            else if(Valuemortality.Contains("CLASS AA"))
            {
                return "CLASSAA";
            }
            else if(Valuemortality.Contains("Class A") || Valuemortality.Contains("CLASS A"))
            {
                return "CLASSA";
            }
            else if(Valuemortality.Contains("Class C") || Valuemortality.Contains("CLASS C"))
            {
                return "CLASSC";
            }
            else if(Valuemortality.Contains("Class E") || Valuemortality.Contains("CLASS E"))
            {
                return "CLASSE";
            }
            else if(Valuemortality.Contains("Class D") || Valuemortality.Contains("CLASS D"))
            {
                return "CLASSD";
            }
            else if(Valuemortality.Contains("Class B") || Valuemortality.Contains("CLASS B"))
            {
                return "CLASSB";
            }
            else if(Valuemortality.Contains("Class H") || Valuemortality.Contains("CLASS H"))
            {
                return "CLASSH";
            }
            else if(Valuemortality.Contains("Class I") || Valuemortality.Contains("CLASS I"))
            {
                return "CLASSI";
            }
            else if(Valuemortality.Contains("Class J") || Valuemortality.Contains("CLASS J"))
            {
                return "CLASSJ";
            }
            else if(Valuemortality.Contains("Class K") || Valuemortality.Contains("CLASS K"))
            {
                return "CLASSK";
            }
            else if(Valuemortality.Contains("Class L") || Valuemortality.Contains("CLASS L"))
            {
                return "CLASSL";
            }
            else if(Valuemortality.Contains("Class M") || Valuemortality.Contains("CLASS M"))
            {
                return "CLASSM";
            }
            else if(Valuemortality.Contains("Class N") || Valuemortality.Contains("CLASS N"))
            {
                return "CLASSN";
            }
            else if(Valuemortality.Contains("Class O") || Valuemortality.Contains("CLASS O"))
            {
                return "CLASSO";
            }
            else if(Valuemortality.Contains("Class P") || Valuemortality.Contains("CLASS P"))
            {
                return "CLASSP";
            }
            else
            {
                Valuemortality = Valuemortality.Trim().ToUpper();
                switch(Valuemortality)
                {
                    case "":
                    return "STANDARD";
                    case "Substandard":
                    return "SUBSTANDARD";
                    case "0.00":
                    return "STANDARD";
                    case "0":
                    return "STANDARD";
                    case "25.00":
                    return "CLASSA";
                    case "50.00":
                    return "CLASSB";
                    case "75.00":
                    return "CLASSC";
                    case "100.00":
                    return "STANDARD";
                    case "1.00":
                    return "STANDARD";
                    case "125.00":
                    return "CLASSA";
                    case "1.50":
                    return "CLASSB";
                    case "175.00":
                    return "CLASSC";
                    case "1.75":
                    return "CLASSC";
                    case "200.00":
                    return "CLASSD";
                    case "2.00":
                    return "CLASSD";
                    case "225.00":
                    return "CLASSE";
                    case "250.00":
                    return "CLASSF";
                    case "275.00":
                    return "CLASSG";
                    case "300.00":
                    return "CLASSH";
                    case "325.00":
                    return "CLASSI";
                    case "350.00":
                    return "CLASSJ";
                    case "375.00":
                    return "CLASSK";
                    case "400.00":
                    return "CLASSL";
                    case "425.00":
                    return "CLASSM";
                    case "450.00":
                    return "CLASSN";
                    case "475.00":
                    return "CLASSO";
                    case "500.00":
                    return "CLASSP";


                    case "125":
                    return "CLASSA";
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
                    case "150":
                    return "CLASSB";


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

                    case "TA":
                    return "CLASSA";
                    case "TB":
                    return "CLASSB";
                    case "TC":
                    return "CLASSC";
                    case "TD":
                    return "CLASSD";
                    case "TE":
                    return "CLASSE";
                    case "TF":
                    return "CLASSF";
                    case "TG":
                    return "CLASSG";
                    case "TH":
                    return "CLASSH";
                    case "TI":
                    return "CLASSI";
                    case "TJ":
                    return "CLASSJ";
                    case "TK":
                    return "CLASSK";
                    case "TL":
                    return "CLASSL";
                    case "TM":
                    return "CLASSM";
                    case "TO":
                    return "CLASSO";
                    case "TP":
                    return "CLASSP";
                    case "TR":
                    return "CLASSR";
                    case "TZ":
                    return "CLASSZ";

                    default:
                    return "STANDARD";

                }
            }
        }




        /**************************************GET TRANSEFFECTIVE DATE / REINSURANCE DATE********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Get TRANSEFFECTIVE DATE / REINSURANCE DATE
        FUNCTION NAME: fn_getTransReinsuranceDate
        FOR BORDEREAUX: BM061, BM048
        */
        public string fn_getTransReinsuranceDate(string valueTcode, string bmyear, string valueIssueDate, out string valueTRD)
        {
            valueIssueDate = Convert.ToDateTime(valueIssueDate).ToString("MM/dd/yyyy");
            if(string.IsNullOrEmpty(valueIssueDate))
            {
                valueTRD = String.Empty;
                return valueTRD;
            }
            else
            {
                valueTRD = "";
                var VID = valueIssueDate.Split(' ');
                valueIssueDate = VID [0];
                var RSD = valueIssueDate.Split('/');
                string month = RSD [0];
                string day = RSD [1];
                string year = RSD [2];
                string DateFormatted = month + "/" + day + "/" + year;
                DateTime valueDate = DateTime.ParseExact(DateFormatted, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                //string strTranseffective = valueDate.ToString("MM/dd/yyyy");

                DateTime dt_IssueDate = DateTime.Now;
                try
                {
                    dt_IssueDate = Convert.ToDateTime(valueIssueDate);

                }
                catch { Variables.boo_invalidIssueDate = true; }


                if(valueTcode == "TRENEW")
                {
                    if(string.IsNullOrEmpty(valueIssueDate))
                    {
                        return null;
                    }
                    else
                    {

                        string newDate = valueDate.ToString("MM/dd/yyyy");
                        newDate = newDate.Substring(0, 6);
                        newDate = newDate + bmyear;
                        valueTRD = newDate;
                        return newDate;
                    }
                }
                else
                {
                    if(string.IsNullOrEmpty(valueIssueDate))
                    {
                        return null;
                    }
                    else if(Variables.boo_invalidIssueDate)
                    {
                        valueTRD = dt_IssueDate.ToString("MM/dd/yyyy");
                        return dt_IssueDate.ToString("MM/dd/yyyy"); ;
                    }
                    else
                    {
                        if((valueDate - dt_IssueDate).TotalDays <= 365)
                        {
                            valueTRD = dt_IssueDate.ToString("MM/dd/yyyy");
                            return dt_IssueDate.ToString("MM/dd/yyyy");
                        }
                        else
                        {
                            valueTRD = valueDate.ToString("MM/dd/yyyy");
                            return valueDate.ToString("MM/dd/yyyy");


                        }
                    }
                }
            }

        }

        public void fn_getTransReinsuranceDateV2(string issueDate, string transcode, string bordereauYear, out string transEffectiveDate)
        {
            transEffectiveDate = ""; 
            string day, month;

            if(transcode == "TNEWBUS")
            {
                transEffectiveDate = issueDate;
            }
            else
            {
                day = issueDate.Substring(0, 2);
                month = issueDate.Substring(3, 2);
                transEffectiveDate = day + "/" + month + "/" + bordereauYear;
            }
        }

        public void fn_getTransReinsuranceDateV3(string issueDate, string bordereauYear, out string transEffectiveDate)
        {
            transEffectiveDate = "";
            string day, month;
            DateTime dt_PremiumDate = Convert.ToDateTime(issueDate);
            DateTime dt_IssueDate = DateTime.Today;
            if((dt_IssueDate - dt_PremiumDate).TotalDays <= 365)
            {
                
                transEffectiveDate = issueDate;
            }
            else
            {
                day = issueDate.Substring(0, 2);
                month = issueDate.Substring(3, 2);
                transEffectiveDate = day + "/" + month + "/" + bordereauYear;
            }
        }

        public void fn_getTransReinsuranceDateV4(string issueDate, string bordereauYear, out string transEffectiveDate, out string transCode)
        {
            transEffectiveDate = ""; transCode = "";  
            string day, month; string premDate = "";

            premDate = issueDate.Substring(0, 6) + bordereauYear;
            DateTime dt_PremiumDate = Convert.ToDateTime(premDate);
            DateTime dt_issueDate = Convert.ToDateTime(issueDate);

            if((dt_PremiumDate - dt_issueDate).TotalDays <= 365)
            {
                transCode = "TNEWBUS";
                transEffectiveDate = issueDate;
            }
            else
            {
                transCode = "TRENEW";
                day = issueDate.Substring(0, 2);
                month = issueDate.Substring(3, 2);
                transEffectiveDate = day + "/" + month + "/" + bordereauYear;
            }
        }


        public void fn_getTransReinsuranceDateV5(string issueDate, string bordereauYear, out string transEffectiveDate, out string transCode)
        {
            transEffectiveDate = ""; transCode = "";
            string day, month; string premDate = "";

            premDate = issueDate.Substring(0, 6) + bordereauYear;
            
                transCode = "TRENEW";
                day = issueDate.Substring(0, 2);
                month = issueDate.Substring(3, 2);
                transEffectiveDate = day + "/" + month + "/" + bordereauYear;
            
        }


        public void fn_getTransReinsuranceDateV6(string issueDate, string bordereauYear, out string transEffectiveDate)
        {
            transEffectiveDate = ""; 
            string day, month; string premDate = "";

            premDate = issueDate.Substring(0, 6) + bordereauYear;

            month = issueDate.Substring(3, 2);
            day = issueDate.Substring(0, 2);
            transEffectiveDate = day + "/" + month + "/" + bordereauYear;

        }

        public void fn_getTransReinsuranceDateV7(string transCode, string bordereauYear, string issueDate, out string transEffectiveDate, out string policyStartDate)
        {
            string day, month; policyStartDate = "";
            day = issueDate.Substring(3, 2);
            month = issueDate.Substring(0, 2);
            if (transCode == "TNEWBUS")
            {
                transEffectiveDate = month + "/" + day + "/" + bordereauYear;
               
            }
            else
            {
                transEffectiveDate = month + "/" + day + "/" + bordereauYear;
                policyStartDate = issueDate;

            }
           
        }


        /**************************************GET POLICY START DATE********************************************************
        PROGRAM: SICS 
        FOR BORDEREAUX: BM061, BM048
        */

        public string fn_getPolicyStartDate(string valueTcode, string valueIssueDate, string valueTransEffectiveDate)
        {

            if(valueTcode == "TRENEW")
            {
                return valueIssueDate;
            }
            else
            {
                if(string.IsNullOrEmpty(valueTcode))
                {
                    return null;
                }
                else
                {
                    return valueTransEffectiveDate;
                }
            }
        }


        /**************************************DateFormatting********************************************************
        PROGRAM: SICS 

        DESCRIPTION: Reformat a date March 10, 1954
        FOR BORDEREAUX: BM049
        */

        public DateTime fn_reformatDatev1(string valueDate)
        {
            try
            {   //September 25, 1953

                var actualDate = valueDate.Split(' ', ',');
                var formattedDate = actualDate [0];
                var newDate = formattedDate.Split('/');
                string month = newDate [0];
                string day = newDate [1];
                string year = newDate [2];

                DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                result = result.Date;

                return result;

            }
            catch(Exception ex)
            {
                return DateTime.ParseExact("07/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

        }

        public string fn_reformatDatev2(string valueDate)
        {
            string strDate;
            try
            {   //September 25, 1953
                if(!string.IsNullOrEmpty(valueDate))
                {
                    var actualDate = valueDate.Split(' ', ',');
                    var formattedDate = actualDate [0];
                    var newDate = formattedDate.Split('/');
                    string month = newDate [0];
                    string day = newDate [1];
                    string year = newDate [2];

                    DateTime result = DateTime.ParseExact(month + "/" + day + "/" + year, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                    result = result.Date;
                    strDate = Convert.ToString(result);
                    return strDate;
                }
                else
                {
                    return valueDate;
                }

            }
            catch(Exception ex)
            {
                DateTime result = DateTime.ParseExact("07/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture);
                strDate = Convert.ToString(result);
                return strDate;
            }

        }



        /************************************Check DATABASE TABLE*************************************************************************************
          FOR BORDEREAUX: BM033
        */
        public string fn_checkDatabaseTable(string valueSheetName)
        {
            if(valueSheetName.ToUpper().Contains("URC"))
            {
                return "dbo_umrc_prembord_umre";
            }
            else
            {
                return "dbo_prembord_nre";
            }
        }

        /************************************Generate Policy No*************************************************************************************
         * DESCRIPTION: Generate Policy no if no policy no is available
         FOR BORDEREAUX: BM049
        */
        public string fn_generatePolicyno(string valuePolicyNo, string valueFirstName, string valueMiddleInitial, string valueLastName, string valueDOB)
        {
            string PolicyNo; string FormattedDOB; string strFirstName; string strLastName; string strPolicyNo;

            valueFirstName = fn_checkFirstname(valueFirstName);
            valueLastName = fn_checkLastname(valueLastName);
            valueMiddleInitial = fn_checkMiddleName(valueMiddleInitial);

            try
            {
                if(string.IsNullOrEmpty(valuePolicyNo))
                {

                    PolicyNo = "";
                    valueMiddleInitial = fn_removeCharacters(valueMiddleInitial);
                    valueMiddleInitial = valueMiddleInitial.Substring(0, 1);
                    FormattedDOB = valueDOB.Replace("/", "");
                    strFirstName = valueFirstName.Substring(0, 1);
                    strLastName = valueLastName.Substring(0, 1);
                    strPolicyNo = PolicyNo + strFirstName + valueMiddleInitial + strLastName + FormattedDOB;
                    return strPolicyNo;
                }
                else
                {
                    //2-00-2540, 2-00-756
                    int CountNo = valuePolicyNo.Length;
                    if(CountNo >= 7)
                    {
                        PolicyNo = valuePolicyNo.Substring(0, 7);
                        valueMiddleInitial = fn_removeCharacters(valueMiddleInitial);
                        valueMiddleInitial = valueMiddleInitial.Substring(0, 1);
                        FormattedDOB = valueDOB.Replace("/", "");
                        strFirstName = valueFirstName.Substring(0, 1);
                        strLastName = valueLastName.Substring(0, 1);
                        strPolicyNo = PolicyNo + strFirstName + valueMiddleInitial + strLastName + FormattedDOB;
                        return strPolicyNo;
                    }
                    else
                    {
                        PolicyNo = valuePolicyNo;
                        valueMiddleInitial = fn_removeCharacters(valueMiddleInitial);
                        valueMiddleInitial = valueMiddleInitial.Substring(0, 1);
                        FormattedDOB = valueDOB.Replace("/", "");
                        strFirstName = valueFirstName.Substring(0, 1);
                        strLastName = valueLastName.Substring(0, 1);
                        strPolicyNo = PolicyNo + strFirstName + valueMiddleInitial + strLastName + FormattedDOB;
                        return strPolicyNo;
                    }

                }
            }
            catch(Exception ex)
            {
                PolicyNo = ""; strFirstName = ""; valueMiddleInitial = ""; strLastName = ""; FormattedDOB = "";
                return strPolicyNo = PolicyNo + strFirstName + valueMiddleInitial + strLastName + FormattedDOB;
            }


        }


        /************************************Check row if has value for Fullname*************************************************************************************
        * DESCRIPTION: Check Row
        FOR BORDEREAUX: BM049
        */

        public string fn_checkFullname(string valueFullname)
        {
            if(string.IsNullOrEmpty(valueFullname))
            {
                return "DummyFullName";
            }
            else
            {
                return valueFullname;
            }
        }


        /************************************Check row if has value for Lastname*************************************************************************************
        * DESCRIPTION: Check Row
        FOR BORDEREAUX: BM049
        */

        public string fn_checkLastname(string valueLastName)
        {
            if(string.IsNullOrEmpty(valueLastName))
            {
                return "DummyLastName";
            }
            else
            {
                valueLastName = fn_removeCharacters(valueLastName);
                return valueLastName;
            }
        }

        /************************************Check row if has value for FirstName *************************************************************************************
        * DESCRIPTION: Check Row for FirstName
        FOR BORDEREAUX: BM049
        */

       
        public string fn_checkFirstname(string valueFirstName)
        {
            if(string.IsNullOrEmpty(valueFirstName))
            {
                return "DummyFirstName";
            }
            else
            {
                return valueFirstName;
            }
        }

        /************************************Check row if has value for MiddleName *************************************************************************************
        * DESCRIPTION: Check Row for Middle Name
        FOR BORDEREAUX: BM049
        */

        public string fn_checkMiddleName(string valueMiddleName)
        {
            if(string.IsNullOrEmpty(valueMiddleName))
            {
                return "DummyMiddleName";
            }
            else
            {
                return valueMiddleName;
            }
        }

        /************************************Check Birthday if Dummy *************************************************************************************
        * DESCRIPTION: Check Row for Middle Name
        FOR BORDEREAUX: BM049
        */

        public string fn_checkDOB(string valueDOB)
        {
            if(string.IsNullOrEmpty(valueDOB) || valueDOB == "01/01/0001")
            {
                return "07/01/1900";
            }
            else
            {
                return valueDOB;
            }
        }

        public string fn_removeCharacters(string valueChar)
        {

            string [] LNSuffix = { "DE", "DEL", "DELA", "DELOS", "LAS", "SAN", "STA", "STO.", "SANTA" };
            if(!string.IsNullOrEmpty(valueChar))
            {
                var lnames = valueChar.Split();
                if(lnames.Length == 1)
                {

                    Regex pattern = new Regex("[.,:/s ]");
                    string result = pattern.Replace(valueChar, "");
                    return result;
                }

                else
                {

                    var lname = valueChar.Split();
                    foreach(string suffix in LNSuffix)
                    {
                        if(lname [0] == suffix)
                        {
                            Regex pattern2 = new Regex("[.,:/s]");
                            string result2 = pattern2.Replace(valueChar, "");
                            return result2;
                            break;
                        }
                        else
                        {
                            continue;

                        }
                    }
                    Regex pattern = new Regex("[.,:/s,]");
                    string result = pattern.Replace(valueChar, "");
                    result = result.Trim();
                    return result;
                }

            }
            return null;
        }



        public bool fn_cl100lookup(string valueChecker)
        {

            Regex c100 = new Regex("[CI 100]"); //put additional 

            if(c100.IsMatch(valueChecker))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public string fn_checkCessioncode(string valueString)
        {
            Regex pattern1 = new Regex("^FAC$");
            Regex pattern2 = new Regex("^AUTOMATIC$");
            if(pattern1.IsMatch(valueString))
            {

                return "F";
            }
            else if(pattern2.IsMatch(valueString))
            {
                return "T";
            }
            else
            {
                return "F";
            }

        }


        public void fn_checkFullNameIsDummy(string valuePolicy, out bool bolDummy)
        {
            bolDummy = false;

            if(valuePolicy.Contains("DUMMY"))
            {
                bolDummy = true;
            }
        }


        public string fn_getBusinessTypeTorF(decimal value1, decimal value2, decimal value3)
        {
            if(value1 != 0 || value2 != 0 || value3 != 0)
            {
                return "F";
            }
            else
            {
                return "T";
            }
        }


        public Boolean fn_getTranscode(string value, out string transCode) //BM021_A
        {
            transCode = string.Empty;


            if(value.ToUpper().Contains("NEW ISSUES") || value.ToUpper() == "FIRST YEAR")
            {
                transCode = "TNEWBUS";
                return true;

            }
            else if(value.ToUpper() == "RENEWAL")
            {
                transCode = "TRENEW";
                return true;
            }
            else if(value.ToUpper().Contains("TERM"))
            {
                transCode = "TCONTER";
                return true;
            }
            else if(value.ToUpper().Contains("LAPS"))
            {
                transCode = "TLAPSE";
                return true;
            }
            else if(value.ToUpper().Contains("SURR"))
            {
                transCode = "TFULLSUR";
                return true;
            }
            else if(value.ToUpper().Contains("REINS"))
            {
                transCode = "TFULLREC";
                return true;
            }
            else if(value.ToUpper() == "ADJUSTMENT FIRST YEAR")
            {
                transCode = "ADJUST";
                return true;
            }
            else if(value.ToUpper() == "ADJUSTMENT RENEWAL")
            {
                transCode = "ADJUST";
                return true;
            }

            else if(value.ToUpper().Contains("RECAP"))
            {
                transCode = "TFULLREC";
                return true;
            }
            else
            {
                transCode = "";
                return false;
            }
        }

        public string fn_getTranscodeV2(string value, out string transCode)
        {
            transCode = string.Empty;


            if(value.ToUpper().Contains("FY"))
            {
                transCode = "TNEWBUS";
                return transCode;

            }
            else 
            {
                transCode = "TRENEW";
                return transCode;
            }
           
         
        }

        public string fn_MaleOrFemale(string Gender)
        {
            if(Gender.ToUpper().Trim() == "FEMALE")
            {
                return "F";
            }
            else
            {
                return "M";
            }
        }

        public string fn_benefitcover(string value)
        {
            value = value.ToUpper().Trim();
            switch(value)
            {
                case "LIFE":
                return "DEATH";

                case "TDB":
                return "DISAB";

                case "ADB":
                return "DEATH";
            }
            return "DEATH";

        }

        public string fn_businessType(string value)
        {
            value = value.ToUpper().Trim();
            switch(value)
            {
                case "A":
                return "R";

                case "F":
                return "NR";

            }
            return "NR";
        }

        public string fn_businessTypeV2(string value)
        {
            value = value.ToUpper().Trim();
            switch(value)
            {
                case "A":
                return "T";
                case "P":
                return "T";
                case "F":
                return "F";
                

            }
            return "";

        }

        public string fn_refundingCode(string value)
        {
            value = value.ToUpper().Trim();
            switch(value)
            {
                case "A":
                return "T";
                case "P":
                return "T";
                case "F":
                return "F";

            }
            return "A";

        }


        public string fn_RemarksBusinessType(string value)
        {
            value = value.ToUpper().Trim();
            switch(value)
            {
                case "P":
                return "PTF";
            }
            return "";
        }

        public string fn_checkIssueDate(string issueDate)
        {
            string IssueDate = "";
            if(string.IsNullOrEmpty(issueDate))
            {
                return "";
            }
            else
            {
                IssueDate = Convert.ToDateTime(issueDate).ToString("MM/dd/yyyy");
                return issueDate;
            }
                

           
        }

        public string fn_getplanCode(string valuePolicyNo, string valueCessionNo)
        {
           string planCode = "";
            try
            {
               
                string query = "SELECT * FROM dbo_cocolife_plancode WHERE policy_no=" + "'" + valuePolicyNo + "'" + "AND " + "cession_no=" + "'" + valueCessionNo + "'";
                string Dbconnection = "DSN=SICS_Postgres_DB;" + "UID=sics;" + "PWD=sics_1";
                OdbcConnection cnDB = new OdbcConnection(Dbconnection);
                //OdbcConnection cnDB = new OdbcConnection(szConnect);

                cnDB.Open();
                OdbcCommand DbCommand = cnDB.CreateCommand();
                DbCommand.CommandText = query;
                OdbcDataReader DbReader = DbCommand.ExecuteReader();

                if(DbReader.Read())
                {
                    planCode = DbReader.GetValue(1).ToString();
                    return planCode;

                }
                else
                {
                    return "";
                }

                DbReader.Close();
                cnDB.Dispose();
                cnDB.Close();
            }
            catch(Exception ex)
            {

                return "";
            }
        }

        public string fn_getplanCodeV2(string plancode)
        {
            if(plancode.ToUpper().Contains("MRI"))
            {
                return "MRI";
            }
            else
            {
                return plancode;
            }
        }

        public void fn_RemarksCode(string policyNo, string benefitCover, string comments, out string insuredProd, out string remarksCode)
        {
            insuredProd = ""; remarksCode = "";
            policyNo = policyNo.Substring(0, 2);
            benefitCover = benefitCover.ToUpper().Trim();

            if(policyNo.Contains("08") && benefitCover == "LIFE")
            {
                insuredProd = "VARLIFE-GU";
                if(!string.IsNullOrEmpty(comments))
                {
                    remarksCode = "Var Life/" + comments;
                }
                else {
                    remarksCode = "Var Life";
                };
                
            }
            else if (benefitCover.ToUpper().Contains("TDB"))
            {
                insuredProd = "WOPDIIND";
                if(!string.IsNullOrEmpty(comments))
                {
                    remarksCode = "Trad Life/" + comments;
                }
                else { remarksCode = "Trad Life";
                };

            }
            else if(benefitCover.ToUpper().Contains("ADB"))
            {
                insuredProd = "ADB-IND";
                if(!string.IsNullOrEmpty(comments))
                {
                    remarksCode = "Trad Life/" + comments;
                }
                else
                {
                    remarksCode = "Trad Life";
                };
            }
            else
            {
                insuredProd = "TRADITIONALLIFE";
                if(!string.IsNullOrEmpty(comments))
                {
                    remarksCode = "Trad Life/" + comments;
                }
                else
                {
                    remarksCode = "Trad Life";
                };
                
            }
        }

        public string fn_PeriodCover(string periodcover, string issueDate)
        {
            string IssueDate = ""; string PeriodCover = ""; string transEffectiveDate = "";
            string day, month, year = "";
            var findYear = periodcover.Split(' ');

            foreach (var item in findYear)
            {
                if (Regex.IsMatch(item, @"\d"))
                {
                    year = item;
                    break;
                }
                else
                {
                    continue;
                }
            }

            month = issueDate.Substring(0, 2);
            day = issueDate.Substring(3, 2);
            transEffectiveDate = month + "/" + day + "/" + year;
            return transEffectiveDate;

        }

    }

}



