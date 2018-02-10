using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SGMLExtracter
{
    internal class SGMLHelper
    {
        Sgml.SgmlReader sgml = null;
        HashSet<String> topicset = new HashSet<string>();
        internal SGMLHelper()
        {
            sgml = new Sgml.SgmlReader();

        }
        public string StoreFolder { get; set; }
        internal void ReadSingleFile(TextReader str, string DocType = "SGML")
        {
            sgml.DocType = DocType;
            sgml.IgnoreDtd = true;
            sgml.InputStream = str;

            XmlDocument x = new XmlDocument();
            x.Load(sgml);
            
            XmlNodeList nodes = x.SelectNodes("root/REUTERS");

            foreach (XmlNode node in nodes)
            {
                var foo = node.InnerXml;
                var topic = node.SelectSingleNode("TOPICS");
                if (!String.IsNullOrEmpty(topic.InnerText))
                 topicset.Add(topic.InnerText);
                File.WriteAllText(Path.Combine(StoreFolder, Path.GetRandomFileName()), node.OuterXml);
            }

        }
        internal void ReadSingleFile(string filePath, string DocType = "SGML")
        {
            using (TextReader reader = File.OpenText(filePath))
            {
                //nasty Workaround 
                MemoryStream stream = new MemoryStream();
                StreamWriter writer = new StreamWriter(stream);
                writer.Write("<root>" + reader.ReadToEnd() + "</root>");
                stream.Flush();
                stream.Position = 0;
                
                ReadSingleFile(new StreamReader(stream));
            }
        }

        internal void ReadSingleFolder(string folder)
        {
            foreach (var file in Directory.GetFiles(folder, "*.sgm"))
            {
                ReadSingleFile(file);
            }

        }
    }
}