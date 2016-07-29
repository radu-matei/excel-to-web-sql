using System;

namespace ExcelToWebSql.Generator
{
    public interface IExcelSqlScriptGenerator : IDisposable
    {
        void GenerateSqlTableScipts();
        void GenerateSqlInsertScripts();
    }
}
