using static XDoc.XPathProcessor;

namespace XDoc
{
    public interface IXMLParser
    {
        List<XmlSdtItem> GetXmlTree();
        void WriteXmlTree<T>(string Tag, T Value);
    }
}
