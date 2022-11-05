using System.Data;

namespace Bordereaux_SICS_Mapping.BAL
{
    class _Global
    {
        public string str_ver = "119.4";

        public string str_outfname; 
        public string str_outlname;
        public string str_outlifeid;
        //public string str_policyYear;

        //HASH Total
        public decimal dbl_BF = 0, dbl_BH = 0, dbl_BJ = 0, dbl_BL = 0, dbl_BZ = 0, dbl_comm = 0, dbl_NAR = 0;
        public decimal dbl_FBF = 0, dbl_FBH = 0, dbl_FBJ = 0, dbl_FBL = 0, dbl_FBZ = 0, dbl_FNAR = 0;
        public decimal dbl_EBF = 0, dbl_EBH = 0, dbl_EBJ = 0, dbl_EBL = 0, dbl_EBZ = 0, dbl_ENAR = 0;
        public decimal dbl_GBF = 0, dbl_GBH = 0, dbl_GBJ = 0, dbl_GBL = 0, dbl_GBZ = 0, dbl_GNAR = 0;
        public decimal dbl_SumatRisk = 0;
        public decimal dbl_LIFE = 0, dbl_ADB = 0, dbl_SAR = 0, dbl_SARDI = 0, dbl_EXTRA = 0, dbl_RIDRET = 0;
        public decimal dbl_Volume = 0, dbl_liferet = 0;
        public decimal dblPremium = 0, dblEMPremium = 0, dblCR = 0;


        public decimal dbl_BF_adj = 0, dbl_BH_adj = 0, dbl_BJ_adj = 0, dbl_BL_adj = 0, dbl_BZ_adj = 0;

        public decimal dbl_BF01 = 0, dbl_BF02 = 0, dbl_BF03 = 0, dbl_BF04 = 0, dbl_BF05 = 0, dbl_BF06 = 0, dbl_BF07 = 0, dbl_BF08 = 0, dbl_BF09 = 0, dbl_BF10 = 0, dbl_BF11 = 0, dbl_BF12 = 0,
                    dbl_BH01 = 0, dbl_BH02 = 0, dbl_BH03 = 0, dbl_BH04 = 0, dbl_BH05 = 0, dbl_BH06 = 0, dbl_BH07 = 0, dbl_BH08 = 0, dbl_BH09 = 0, dbl_BH10 = 0, dbl_BH11 = 0, dbl_BH12 = 0,
                    dbl_BJ01 = 0, dbl_BJ02 = 0, dbl_BJ03 = 0, dbl_BJ04 = 0, dbl_BJ05 = 0, dbl_BJ06 = 0, dbl_BJ07 = 0, dbl_BJ08 = 0, dbl_BJ09 = 0, dbl_BJ10 = 0, dbl_BJ11 = 0, dbl_BJ12 = 0,
                    dbl_BL01 = 0, dbl_BL02 = 0, dbl_BL03 = 0, dbl_BL04 = 0, dbl_BL05 = 0, dbl_BL06 = 0, dbl_BL07 = 0, dbl_BL08 = 0, dbl_BL09 = 0, dbl_BL10 = 0, dbl_BL11 = 0, dbl_BL12 = 0,
                    dbl_BZ01 = 0, dbl_BZ02 = 0, dbl_BZ03 = 0, dbl_BZ04 = 0, dbl_BZ05 = 0, dbl_BZ06 = 0, dbl_BZ07 = 0, dbl_BZ08 = 0, dbl_BZ09 = 0, dbl_BZ10 = 0, dbl_BZ11 = 0, dbl_BZ12 = 0;
        //ISSUE#013-Start---------

        public decimal dbl_BF_PHP = 0, dbl_BH_PHP = 0, dbl_BJ_PHP = 0, dbl_BL_PHP = 0, dbl_BZ_PHP = 0,
                       dbl_BF_USD = 0, dbl_BH_USD = 0, dbl_BJ_USD = 0, dbl_BL_USD = 0, dbl_BZ_USD = 0,
                       dbl_BF_PHP_UL = 0, dbl_BH_PHP_UL = 0, dbl_BJ_PHP_UL = 0, dbl_BL_PHP_UL = 0, dbl_BZ_PHP_UL = 0,
                       dbl_BF_USD_UL = 0, dbl_BH_USD_UL = 0, dbl_BJ_USD_UL = 0, dbl_BL_USD_UL = 0, dbl_BZ_USD_UL = 0;

        public string str_GFailLines = string.Empty;

        public string str_GFailLines01 = string.Empty;
        public string str_GFailLines02 = string.Empty;
        public string str_GFailLines03 = string.Empty;
        public string str_GFailLines04 = string.Empty;
        public string str_GFailLines05 = string.Empty;
        public string str_GFailLines06 = string.Empty;
        public string str_GFailLines07 = string.Empty;
        public string str_GFailLines08 = string.Empty;
        public string str_GFailLines09 = string.Empty;
        public string str_GFailLines10 = string.Empty;
        public string str_GFailLines11 = string.Empty;
        public string str_GFailLines12 = string.Empty;

        public string str_GFailLines_adj = string.Empty;
        //ISSUE#013-End-----------

        public System.Data.DataRow dtworkRow;

        public System.Data.DataRow dtworkRow01;
        public System.Data.DataRow dtworkRow02;
        public System.Data.DataRow dtworkRow03;
        public System.Data.DataRow dtworkRow04;
        public System.Data.DataRow dtworkRow05;
        public System.Data.DataRow dtworkRow06;
        public System.Data.DataRow dtworkRow07;
        public System.Data.DataRow dtworkRow08;
        public System.Data.DataRow dtworkRow09;
        public System.Data.DataRow dtworkRow10;

        public DataTable objdt_template01 = new DataTable();
        public DataTable objdt_template02 = new DataTable();
        public DataTable objdt_template03 = new DataTable();
        public DataTable objdt_template04 = new DataTable();
        public DataTable objdt_template05 = new DataTable();
        public DataTable objdt_template06 = new DataTable();
        public DataTable objdt_template07 = new DataTable();
        public DataTable objdt_template08 = new DataTable();
        public DataTable objdt_template09 = new DataTable();
        public DataTable objdt_template10 = new DataTable();
        public DataTable objdt_template11 = new DataTable();
        public DataTable objdt_template12 = new DataTable();

        public DataTable objdt_templateADJ = new DataTable();

        public DataTable objdt_MACRO = new DataTable();
        public DataTable objdt_OCCCODE = new DataTable();
        public DataTable objdt_GenderDB = new DataTable();

        //used on helper
        public string str_final_fname;
        public string str_final_lname;
        public string str_leadsuffix;
        public bool boo_genderfail = false;
        

        
    }

    public class Variables
    {
        public static string strBmYear;
        public static bool boomacrofail;
        public static bool boogenderfail = false;
        public static decimal TotalPremium;
        public static decimal TotalSumAtRisk;
        public static decimal TotalCommission;
        public static bool boo_invalidIssueDate = false;

        public static decimal TotalFaculPremium;
        public static decimal TotalTreatyPremium;
        public static decimal TotalTreatySAR;
        public static decimal TotalFaculSAR;
        public static decimal TotalQuotaPremium;
        public static decimal TotalSurplusPremium;
        public static decimal TotalQuotaSAR;
        public static decimal TotalSurplusSAR;



    }

    
}
