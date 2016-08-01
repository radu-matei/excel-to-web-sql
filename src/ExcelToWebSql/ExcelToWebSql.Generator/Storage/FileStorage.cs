using System;
using System.IO;

namespace ExcelToWebSql.Generator.Storage
{
    public class FileStorage : IStorage
    {
        public void SaveDocument(string content, string outputFilePath)
        {
            try
            {
                using (var stream = File.Open(outputFilePath, FileMode.Create, FileAccess.Write))
                using (var sw = new StreamWriter(stream))
                {
                    sw.Write(content);
                    sw.Close();
                }
            }
            catch (Exception)
            {
                throw new AccessViolationException();
            }
        }
    }
}
