using Newtonsoft.Json;

namespace ExcelToWebSql.Generator
{
    public class JsonSerializer : ISerializer
    {
        public string Serialize(object sourceObject)
        {
            return JsonConvert.SerializeObject(sourceObject);
        }
    }
}
