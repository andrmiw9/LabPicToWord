using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using LabPicturesToWord.Properties;

namespace LabPicturesToWord
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Settings Set = new Settings();
            Version AppVersion = Assembly.GetExecutingAssembly().GetName().Version;
            if(String.Equals(Set.AppVersion, AppVersion) == false)
            {
                Set.Upgrade();
                Set.AppVersion = AppVersion.ToString();
                Set.Save();
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
