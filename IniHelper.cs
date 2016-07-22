using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;
using System.Configuration;
using System.Runtime.InteropServices;

namespace CommonLibrary
{
    public class IniHelper
    {
        private static string _iniPath = "IniPath";


        public static string getConfigSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        public static string getIniFilePath()
        {
            return getConfigSetting(_iniPath);
        }
        public static string getIniFilePath(string appKey)
        {
            return getConfigSetting(appKey);
        }

        #region #取.ini資料
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static string getIniValue(string section, string key, string path)
        {
            StringBuilder output = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", output, 255, path);
            return output.ToString();
        }
        #endregion

    }
}
