using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AutoPTD
{
    public class DocConfig
    {
        public string ConfigPath = "";
        public string PicturePath = "";
        public string DocName = "";

        public DocConfig()
        {
            try
            {
                ConfigPath = AppDomain.CurrentDomain.BaseDirectory + "DocConfig.xml";
                System.Xml.XmlDocument DOC = new System.Xml.XmlDocument();
                SetInitialConfig();
                LoadConfig();
            }
            catch
            { }
        }

        private void LoadConfig()
        {
            this.PicturePath = GetNodeValue("PicturePath").ToString();
            this.DocName = GetNodeValue("DocName").ToString();
        }

        private void SetInitialConfig()
        {
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.LoadXml("<?xml version=\"1.0\"?> <Configuration> </Configuration>");
            System.Xml.XmlNode root = doc.SelectSingleNode("Configuration");

            System.Xml.XmlNode tmpNode = root.SelectSingleNode("PicturePath");
            if (tmpNode == null)
            {
                tmpNode = (System.Xml.XmlNode)doc.CreateElement("PicturePath");
                this.PicturePath = tmpNode.InnerText = @"D:\TestPicture\";
                root.AppendChild(tmpNode);
            }

            tmpNode = root.SelectSingleNode("DocName");
            if (tmpNode == null)
            {
                tmpNode = (System.Xml.XmlNode)doc.CreateElement("DocName");
                this.DocName = tmpNode.InnerText = @"Test.docx";
                root.AppendChild(tmpNode);
            }


        }

        public string GetNodeValue(string NodeName)
        {
            try
            {

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("DocConfig.xml");
                XmlNode root = xmlDoc.SelectSingleNode("//Configuration");
                string value = "";

                if (root != null)
                {
                    value = (root.SelectSingleNode(NodeName)).InnerText;
                }
                return value;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Reading file error:" + ex,"Tips",MessageBoxButtons.OK);
                Console.WriteLine(ex);
                return null;

            }
        }

        public void SetNodeValue(string NodeName, string NodeValue)
        {
            try
            {
                System.Xml.XmlDocument DOC = new System.Xml.XmlDocument();
                DOC.Load(ConfigPath);
                System.Xml.XmlNode ROOT = DOC.SelectSingleNode("Configuration");
                System.Xml.XmlNode tmpNode = DOC.SelectSingleNode("//" + NodeName);

                if (tmpNode == null)
                {
                    tmpNode = (System.Xml.XmlNode)DOC.CreateElement(NodeName);
                    tmpNode.InnerText = NodeValue;
                    ROOT.AppendChild(tmpNode);
                }
                tmpNode.InnerText = NodeValue;
                DOC.Save(ConfigPath);
            }
            catch { }
        }

    }
}
