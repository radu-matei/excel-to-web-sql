namespace ExcelToWebSql.Generator
{
    public interface ISerializer
    {
        string Serialize(object sourceObject);
    }
}
