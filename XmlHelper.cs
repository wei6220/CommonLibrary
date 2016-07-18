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
    public class XmlHelper
    {
        XmlDocument doc;

        public string Load(string Path)
        {
            doc = new XmlDocument();
            try
            {
                doc.Load(Path);
            }
            catch (Exception)
            {
                LogHelper.Write("[MainForm]Xml file load error!!");
            }
            return doc.InnerXml;
        }

        public IList Read(string Path)
        {
            return Read(Path, null);
        }
        public IList Read(string Path, string RootNode)
        {
            if (string.IsNullOrEmpty(RootNode))
                RootNode = "/";
            doc = new XmlDocument();
            try
            {
                doc.Load(Path);
            }
            catch (Exception e)
            {
                LogHelper.Write("[MainForm]Xml file load error!!");
            }
            XmlNodeList nodes = doc.SelectNodes(RootNode);
            return (IList)ResolveNode(nodes);
        }

        public IList ReadXml(string XmlStr, string XPath)
        {
            if (string.IsNullOrEmpty(XPath))
                XPath = "/";
            doc = new XmlDocument();
            try
            {
                doc.LoadXml(XmlStr);
            }
            catch (Exception)
            {
                LogHelper.Write("[MainForm]Read XmlString error!!");
                return null;
            }
            XmlNodeList nodes = doc.SelectNodes(XPath);
            return (IList)ResolveNode(nodes);
        }

        public string ExtractXml(string XmlStr, string XPath)
        {
            if (string.IsNullOrEmpty(XPath))
                return XmlStr;

            doc = new XmlDocument();
            try
            {
                doc.LoadXml(XmlStr);
            }
            catch (Exception)
            {
                LogHelper.Write("[MainForm]Extract Xml String error!!");
            }
            XmlNodeList nodes = doc.SelectNodes(XPath);
            string xml = "";
            foreach (XmlNode node in nodes)
                xml += node.InnerXml;
            return xml;
        }

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
                    return node.Value;
                NodeList.Add(new KeyValuePair<string, object>(NodeName, NodeValue));
            }
            return NodeList;
        }

        public IList Find(IList TargetList, string TargetValue)
        {
            foreach (KeyValuePair<string, object> item in TargetList)
            {
                if (item.Value.GetType() == typeof(string))
                {
                    if (item.Value.ToString() == TargetValue)
                        return (IList)TargetList;
                }
                else
                {
                    return Find((IList)item.Value, TargetValue);
                }
            }
            return new List<object>();
        }

    }
}
