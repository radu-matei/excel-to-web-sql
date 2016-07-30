using System;

namespace ExcelToWebSql.Generator
{
    public interface IExcelConverter : IDisposable
    {
        void ConvertToJson(string outputPath);
        void ConvertToXml(string outputPath);
    }
}

