using System;

namespace ExcelToWebSql.Generator
{
    public interface IExcelSqlScriptGenerator : IDisposable
    {
        void GenerateSqlTableScipts(string sourceFilePath, string outputFilePath);
        void GenerateSqlInsertScripts(string sourceFilePath, string outputFilePath);
    }
}
