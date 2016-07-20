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
        public static string getConfigSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static string getIniFileFullName()
        {
            return getConfigSetting("IniPath");
        }

        public static string getIniValue(string path, string section, string key)
        {
            StringBuilder output = new StringBuilder();
            GetPrivateProfileString(section, key, "", output, 255, path);
            return output.ToString();
        }

        public static string getConnectionString(string path, string section)
        {
            try
            {
                // Data Source =.\SQLEXPRESS; Initial Catalog = InsydeBIOSConfigurationDB; Integrated Security = True
                string source = getIniValue(path, section, "servername");
                string iniCatalog = getIniValue(path, section, "datasource");
                return @"Data Source =" + source + "; Initial Catalog = " + iniCatalog + "; Integrated Security = True";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }
}
