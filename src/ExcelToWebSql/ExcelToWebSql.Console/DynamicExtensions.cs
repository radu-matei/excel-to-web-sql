using System;
using System.Collections.Generic;
using System.Dynamic;

namespace ExcelToWebSql.Generator
{
    public static class DynamicExtensions
    {
        public static void AddProperty(this ExpandoObject dynamicObject, string propertyName, object propertyValue)
        {
            var dynamicDictionary = dynamicObject as IDictionary<string, object>;

            if (dynamicDictionary.ContainsKey(propertyName))
                dynamicDictionary[propertyName] = propertyValue;

            else
                dynamicDictionary.Add(propertyName, propertyValue);
        }

        public static void Print(this ExpandoObject dynamicObject)
        {
            var dynamicDictionary = dynamicObject as IDictionary<string, object>;

            foreach(KeyValuePair<string, object> property in dynamicDictionary)
            {
                Console.WriteLine("{0}: {1}", property.Key, property.Value.ToString());
            }
            Console.WriteLine();
        }
    }
}
