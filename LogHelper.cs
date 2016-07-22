using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CommonLibrary
{
    public class LogHelper
    {
        private static string LogPath = System.AppDomain.CurrentDomain.BaseDirectory + @"\Log";

        private static string LogFullName = LogPath + @"\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

        private static object lockMe = new object();
        public static void Write(string sLogStr)
        {
            if (!Directory.Exists(LogPath))
                Directory.CreateDirectory(LogPath);

            StreamWriter sw;
            lock (lockMe)
            {
                if (!File.Exists(LogFullName))
                    sw = File.CreateText(LogFullName);
                else
                    sw = new StreamWriter(LogFullName, true);
                sw.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " || " + sLogStr);
                sw.Close();
            }
        }

    }
}
