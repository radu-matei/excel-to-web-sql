using System.Collections.Generic;
using System.IO;

namespace ExcelToWebSql.Generator
{
    public class XmlSerializer : ISerializer
    {
        public string Serialize(object sourceObject)
        {
            var xmlSerializer = new System.Xml.Serialization.XmlSerializer(sourceObject.GetType());
            using (StringWriter writer = new StringWriter())
            {
                xmlSerializer.Serialize(writer, sourceObject);
                return writer.ToString();
            }
        }
    }
}
