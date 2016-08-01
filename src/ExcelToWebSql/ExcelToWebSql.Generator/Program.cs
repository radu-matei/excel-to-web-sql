using System;
using System.IO;

namespace ExcelToWebSql.Generator
{
    class Program
    {
        static void Main(string[] args)
        {
            var sourceFilePath = Path.Combine(Environment.CurrentDirectory, "..\\..\\..\\..\\worksheet.xlsx");
            var outputFilePath = Path.Combine(Environment.CurrentDirectory, "..\\..\\..\\..\\Output\\");

            IExcelSqlScriptGenerator scriptGenerator = new ExcelSqlScriptGenerator();

            scriptGenerator.GenerateSqlTableScipts(sourceFilePath, outputFilePath);
            scriptGenerator.GenerateSqlInsertScripts(sourceFilePath, outputFilePath);

            scriptGenerator.Dispose();


            IExcelConverter excelConverter = new ExcelConverter();
            excelConverter.ConvertToJson(sourceFilePath,outputFilePath);

            excelConverter.Dispose();
        }
    }
}
