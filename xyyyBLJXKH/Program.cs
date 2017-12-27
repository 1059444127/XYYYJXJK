using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace pathnetfsjk
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            IniFiles f = new IniFiles(Application.StartupPath + "\\sz.ini");
            string aa = f.ReadString("fsjk","fsms","1").Replace("\0","");
            
                Application.Run(new Frm_fsjk());
           
        }
    }
}