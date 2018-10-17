using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Units_display
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static public Form1 f;

        [STAThread]
        
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            f = new Form1();
            //f.TopMost = true;
            Application.Run(f);
        }
    }
}
