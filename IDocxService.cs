namespace XDoc
{
    public interface IDocxService
    {
        string ExtractXml(string docxPath);
        StatusDocx SaveXml(string originalDocxPath, string modifiedXml, string outputDocxPath, string FileName);
    }
}
