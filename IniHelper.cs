using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Collections;

namespace CommonLibrary
{
    public class IniHelper
    {
        private static string _iniPathKey = "IniPath";
    

        public static string getConfigSetting(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        public static string getIniFilePath()
        {
            return getConfigSetting(_iniPathKey);
        }
        public static string getIniFilePath(string appKey)
        {
            return getConfigSetting(appKey);
        }

        public static string getConnectionString(string section, string iniPath)
        {
            try
            {
                string server, db, userid, password;
                if (string.Compare(Path.GetExtension(iniPath), ".xml", true) == 0)
                {
                    //by .xml
                    XmlHelper xmlHelper = new XmlHelper();
                    IList list = xmlHelper.Read(iniPath, "SETTINGS/" + section);
                    server = xmlHelper.Find(list, "servername").ToString();
                    db = xmlHelper.Find(list, "datasource").ToString();
                    userid = xmlHelper.Find(list, "userid").ToString();
                    password = xmlHelper.Find(list, "Password").ToString();
                }
                else
                {
                    //by .ini
                    server = getIniFileValue(section, "servername", iniPath);
                    db = getIniFileValue(section, "datasource", iniPath);
                    userid = getIniFileValue(section, "userid", iniPath);
                    password = getIniFileValue(section, "Password", iniPath);
                }

                return @"Data Source =" + server
                         + "; Initial Catalog = " + db
                         + "; Integrated Security = false;User ID = " + userid
                         + "; Password = '" + password + "';";
            }
            catch (Exception ex)
            {
                LogHelper.Write("[IniHelper] getConnectionString error!!");
                throw new Exception(ex.Message);
            }
        }

        #region #取.ini資料
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static string getIniFileValue(string section, string key, string path)
        {
            try
            {
                StringBuilder output = new StringBuilder(255);
                GetPrivateProfileString(section, key, "", output, 255, path);
                return output.ToString();
            }
            catch (Exception e)
            {
                LogHelper.Write("[IniHelper] getIniFileValue error!!");
                throw new Exception(e.Message);
            }
        }
        #endregion

    }
}
