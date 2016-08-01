namespace ExcelToWebSql.Generator
{
    class Program
    {
        static void Main(string[] args)
        {
            var sourceFilePath = @"C:\Users\t-ramate\Documents\Visual Studio 2015\Projects\ConsoleExcelDataReader\ConsoleExcelDataReader\worksheet.xlsx";
            //IExcelSqlScriptGenerator scriptGenerator = new ExcelSqlScriptGenerator(sourceFilePath);

            //scriptGenerator.GenerateSqlTableScipts();
            //scriptGenerator.GenerateSqlInsertScripts();

            //scriptGenerator.Dispose();

            IExcelConverter excelConverter = new ExcelConverter(sourceFilePath);
            excelConverter.ConvertToJson(sourceFilePath);

            excelConverter.Dispose();
        }
    }
}
