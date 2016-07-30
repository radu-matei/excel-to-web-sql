namespace ExcelToWebSql.Generator
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Users\t-ramate\Documents\Visual Studio 2015\Projects\ConsoleExcelDataReader\ConsoleExcelDataReader\worksheet.xlsx";
            //IExcelSqlScriptGenerator scriptGenerator = new ExcelSqlScriptGenerator(path);

            //scriptGenerator.GenerateSqlTableScipts();
            //scriptGenerator.GenerateSqlInsertScripts();

            //scriptGenerator.Dispose();

            IExcelConverter excelConverter = new ExcelConverter(path);
            excelConverter.ConvertToJson(path);

            excelConverter.Dispose();
        }
    }
}
