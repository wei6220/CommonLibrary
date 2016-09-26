using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CommonLibrary
{
    /// <summary>
    /// Xml讀寫元件
    /// </summary>
    public class XmlHelper
    {
        private XmlDocument xmlDoc;

        /// <summary>
        /// 讀取xml檔案，並轉出xml string
        /// </summary>
        /// <param name="Path">檔案路徑</param>
        /// <returns>xml string</returns>
        public string Load(string Path)
        {
            xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(Path);
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> XmlHelper>> Load]\n" + ex.Message);
                throw new Exception("[CommonLibrary]XmlHelper Load failed!!");
            }
            return xmlDoc.InnerXml;
        }
        /// <summary>
        /// 讀取xml檔案，並轉出list (根結點)
        /// </summary>
        /// <param name="Path">檔案路徑</param>
        /// <returns>list</returns>
        public IList Read(string Path)
        {
            return Read(Path, null);
        }
        /// <summary>
        /// 讀取xml檔案，並轉出list
        /// </summary>
        /// <param name="Path">檔案路徑</param>
        /// <param name="xPath">自訂節點</param>
        /// <returns>list</returns>
        public IList Read(string Path, string xPath)
        {
            if (string.IsNullOrEmpty(xPath))
                xPath = "/";
            xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(Path);
                XmlNodeList nodes = xmlDoc.SelectNodes(xPath);
                return (IList)ResolveNode(nodes);
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> XmlHelper>> Read]\n" + ex.Message);
                throw new Exception("[CommonLibrary]XmlHelper Read failed!!");
            }
        }
        /// <summary>
        /// 讀取xml字串，並轉出list
        /// </summary>
        /// <param name="XmlStr">xml字串</param>
        /// <param name="XPath">自訂節點</param>
        /// <returns>list</returns>
        public IList ReadXml(string XmlStr, string XPath)
        {
            if (string.IsNullOrEmpty(XPath))
                XPath = "/";
            xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.LoadXml(XmlStr);
                XmlNodeList nodes = xmlDoc.SelectNodes(XPath);
                return (IList)ResolveNode(nodes);
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> XmlHelper>> ReadXml]\n" + ex.Message);
                throw new Exception("[CommonLibrary]XmlHelper ReadXml failed!!");
            }
        }
        /// <summary>
        /// 自傳入的xml字串內，擷取子節點字串
        /// </summary>
        /// <param name="XmlStr">xml字串</param>
        /// <param name="XPath">欲擷取的節點</param>
        /// <returns>string</returns>
        public string ExtractXml(string XmlStr, string XPath)
        {
            if (string.IsNullOrEmpty(XPath))
                return XmlStr;

            xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.LoadXml(XmlStr);
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> XmlHelper>> ExtractXml]\n" + ex.Message);
            }
            XmlNodeList nodes = xmlDoc.SelectNodes(XPath);
            string xml = "";
            foreach (XmlNode node in nodes)
                xml += node.InnerXml;
            return xml;
        }
        /// <summary>
        /// 解析節點
        /// </summary>
        /// <param name="nodes">節點List</param>
        /// <returns>object(格式: List{keyvalue})</returns>
        private object ResolveNode(XmlNodeList nodes)
        {
            List<object> NodeList = new List<object>();

            foreach (XmlNode node in nodes)
            {
                //TODO:處理屬性
                //XmlElement element = (XmlElement)node; 
                string NodeName = node.Name;
                object NodeValue;
                if (node.HasChildNodes)
                    NodeValue = ResolveNode(node.ChildNodes);
                else
                {
                    if (node.Value == null)
                        NodeValue = "";
                    else
                        return node.Value;
                }
                NodeList.Add(new KeyValuePair<string, object>(NodeName, NodeValue));
            }
            return NodeList;
        }

        /// <summary>
        /// 找尋List中的節點 
        /// </summary>
        /// <param name="TargetList">target source</param>
        /// <param name="TargetValue">target key</param>
        /// <returns>target value(依第一次找的回傳)</returns>
        public string Find(IList TargetList, string TargetValue)
        {
            foreach (KeyValuePair<string, object> item in TargetList)
            {
                if (item.Value.GetType() == typeof(string))
                {
                    if (string.Compare(item.Key.ToString(), TargetValue, true) == 0)
                        return item.Value.ToString();
                }
                else
                {
                    return Find((IList)item.Value, TargetValue);
                }
            }
            return "";
        }

    }
}
