namespace ExcelToWebSql.Generator.Storage
{
    public interface IStorage
    {
        void SaveDocument(string content, string outputFilePath);
    }
}
