using System;

namespace ExcelToWebSql.Generator
{
    public interface IExcelConverter : IDisposable
    {
        void ConvertToJson(string sourceFilePath, string outputFilePath);
        void ConvertToXml(string sourceFilePath, string outputFilePath);
    }
}

