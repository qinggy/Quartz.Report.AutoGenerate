using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Esd.Report.AutoGenerate.Common
{
    public static class SerializeExtension
    {
        public static T XmlToObject<T>(this string xmlFileName) where T : class
        {
            var xmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, xmlFileName);
            string xml;
            using (var reander = File.OpenText(xmlPath))
            {
                xml = reander.ReadToEnd();
            }
            var doc = new XmlDocument();
            doc.LoadXml(xml);

            return JsonConvert.SerializeXmlNode(doc).FromJson<T>();
        }

        public static T FromJson<T>(this string jsonStr) where T : class
        {
            return string.IsNullOrWhiteSpace(jsonStr) ? null : JsonConvert.DeserializeObject<T>(jsonStr);
        }
    }
}
