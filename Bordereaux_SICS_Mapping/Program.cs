using System;
using System.Windows.Forms;

namespace Bordereaux_SICS_Mapping
{
    static class Program
    {   
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new B2SM());
            //Application.Run(new Forms.ClassTester());
        }
    }
}
