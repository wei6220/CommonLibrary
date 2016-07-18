using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CommonLibrary
{
    class LogHelper
    {
        private static string LogPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Log";
        private static string sLogFullName = LogPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

        private static object lockMe = new object();
        public static void Write(string sLogStr)
        {
            if (!Directory.Exists(LogPath))
                Directory.CreateDirectory(LogPath);

            StreamWriter sw;
            lock (lockMe)
            {
                if (!File.Exists(sLogFullName))
                    sw = File.CreateText(sLogFullName);
                else
                    sw = new StreamWriter(sLogFullName, true);
                sw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " || " + sLogStr);
                sw.Close();
            }
        }

    }
}
