﻿using System;
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
    /// <summary>
    /// 設定檔讀取元件
    /// </summary>
    public class IniHelper
    {
        private static string _defaultIniPathKey = "IniPath";
    
        /// <summary>
        /// 取得初始設定檔路徑(預設key:IniPath)
        /// </summary>
        /// <returns></returns>
        public static string GetFilePath()
        {
            return GetFilePath(_defaultIniPathKey);
        }

        /// <summary>
        /// 取得初始設定檔路徑
        /// </summary>
        /// <param name="appKey">自訂key</param>
        /// <returns></returns>
        public static string GetFilePath(string appKey)
        {
            return ConfigurationManager.AppSettings[appKey];
        }

        /// <summary>
        /// 取得設定值(預設路徑)
        /// </summary>
        /// <param name="section">設定檔section值</param>
        /// <param name="key">設定檔key值</param>
        /// <returns></returns>
        public static string GetValue(string section, string key)
        {
            return GetValue(section, key, GetFilePath());
        }

        /// <summary>
        /// 取得設定值
        /// </summary>
        /// <param name="section">設定檔section值</param>
        /// <param name="key">設定檔key值</param>
        /// <param name="path">自訂設定檔位置</param>
        /// <returns></returns>
        public static string GetValue(string section, string key, string path)
        {
            try
            {
                if (string.Compare(Path.GetExtension(path), ".xml", true) == 0)
                {
                    XmlHelper xmlHelper = new XmlHelper();
                    IList list = xmlHelper.Read(path, "SETTINGS/" + section);
                    return xmlHelper.Find(list, key).ToString();
                }
                else
                {
                    StringBuilder output = new StringBuilder(255);
                    GetPrivateProfileString(section, key, "", output, 255, path);
                    return output.ToString();
                }
            }
            catch (Exception e)
            {
                LogHelper.Write("[Error][CommonLibrary>> IniHelper>> GetIniFileValue]\n" + e.Message);
                throw new Exception("[CommonLibrary]IniHelper GetIniFileValue failed!!");
            }
        }

        /// <summary>
        /// 取得設定值(預設路徑)
        /// </summary>
        /// <param name="section">設定檔section值</param>
        /// <param name="key">設定檔key值</param>
        /// <returns></returns>
        public static IList GetValues(string section)
        {
            return GetValues(section, GetFilePath());
        }

        /// <summary>
        /// 取得設定值
        /// </summary>
        /// <param name="section">設定檔section值</param>
        /// <param name="key">設定檔key值</param>
        /// <param name="path">自訂設定檔位置</param>
        /// <returns></returns>
        public static IList GetValues(string section, string path)
        {
            try
            {
                if (string.Compare(Path.GetExtension(path), ".xml", true) == 0)
                {
                    XmlHelper xmlHelper = new XmlHelper();
                    return xmlHelper.Read(path, section);
                }
                return null;
            }
            catch (Exception e)
            {
                LogHelper.Write("[Error][CommonLibrary>> IniHelper>> GetIniFileValue]\n" + e.Message);
                throw new Exception("[CommonLibrary]IniHelper GetIniFileValue failed!!");
            }
        }


        #region #取.ini資料
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        #endregion

    }
}
