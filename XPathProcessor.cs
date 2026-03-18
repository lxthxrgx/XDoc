using Microsoft.Extensions.Configuration;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace XDoc
{
    public class PathSettings
    {
        public string PathToFolder { get; set; } = "";
    }
    public class XPathProcessor : IFile, IXMLParser, IDocxService
    {
        private readonly string _filePath;
        private XmlDocument doc;
        private XmlNamespaceManager nsManager;
        private readonly string _pathToTemplate;

        public XPathProcessor(){}

        public XPathProcessor(string pathToTemplate, string filepath)
        {
            _pathToTemplate = pathToTemplate;
            string dataXml = ExtractXml(_pathToTemplate);

            _filePath = filepath;

            doc = new XmlDocument();
            doc.LoadXml(dataXml);

            nsManager = new XmlNamespaceManager(doc.NameTable);
            nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        }

        public string ExtractXml(string pathToDocx)
        {
            try
            {
                using (var archive = ZipFile.OpenRead(pathToDocx))
                {
                    var entry = archive.GetEntry("word/document.xml");
                    if (entry == null)
                        throw new Exception($"word file is empty or broken: {pathToDocx}");

                    using (var reader = new StreamReader(entry.Open()))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (InvalidDataException)
            {
                throw new Exception($"word file is empty or broken: {pathToDocx}");
            }
        }

        public class XmlSdtItem
        {
            public string? Alias { get; set; }
            public string? Tag { get; set; }
            public string? Text { get; set; }

            public override string ToString()
            {
                return $"{Tag}: {Text}";
            }
        }
        public List<XmlSdtItem> GetXmlTree()
        {
            XPathProcessor xPath = new(_pathToTemplate, _filePath);
            string dataXml = xPath.ExtractXml(_pathToTemplate);

            XmlDocument doc = new XmlDocument();
            doc.Load(new StringReader(dataXml));

            XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
            nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            List<XmlSdtItem> result = new();

            var paragraphs = doc.SelectNodes("//w:sdt[w:*]", nsManager);
            if (paragraphs is not null)
            {
                foreach (XmlNode node in paragraphs)
                {
                    var item = new XmlSdtItem();

                    XmlNode? sdtPr = node.SelectSingleNode("w:sdtPr", nsManager);
                    if (sdtPr != null)
                    {
                        var alias = sdtPr.SelectSingleNode("w:alias", nsManager);
                        if (alias?.Attributes?["w:val"] != null)
                            item.Alias = alias.Attributes["w:val"].Value;

                        var tag = sdtPr.SelectSingleNode("w:tag", nsManager);
                        if (tag?.Attributes?["w:val"] != null)
                            item.Tag = tag.Attributes["w:val"].Value;
                    }

                    XmlNode? sdtContent = node.SelectSingleNode("w:sdtContent", nsManager);
                    if (sdtContent != null)
                    {
                        var textNode = sdtContent.SelectSingleNode(".//w:t", nsManager);
                        if (textNode != null)
                            item.Text = textNode.InnerText;
                    }

                    result.Add(item);
                }
            }

            return result;
        }

        public void WriteXmlTree<T>(string Tag, T Value)
        {
            var paragraphs = doc.SelectNodes("//w:sdt[w:*]", nsManager);

            if (paragraphs != null)
            {
                foreach (XmlNode node in paragraphs)
                {
                    XmlNode? sdtPr = node.SelectSingleNode("w:sdtPr", nsManager);

                    if (sdtPr != null)
                    {
                        var tag = sdtPr.SelectSingleNode("w:tag", nsManager);
                        if (tag?.Attributes?["w:val"]?.Value == Tag)
                        {
                            XmlNode? sdtContent = node.SelectSingleNode("w:sdtContent", nsManager);
                            if (sdtContent != null)
                            {
                                var textNodes = sdtContent.SelectNodes(".//w:t", nsManager);
                                if (textNodes != null && textNodes.Count > 0)
                                {
                                    textNodes[0].InnerText = Value?.ToString() ?? "";

                                    for (int i = 1; i < textNodes.Count; i++)
                                    {
                                        textNodes[i].InnerText = "";
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public class Utf8StringWriter : StringWriter
        {
            public override Encoding Encoding => new UTF8Encoding(false);
        }

        public void Save(string nameFile)
        {
            string updatedXml;
            using (var stringWriter = new Utf8StringWriter())
            using (var xmlWriter = XmlWriter.Create(stringWriter, new XmlWriterSettings
            {
                Indent = false,
                Encoding = new UTF8Encoding(false)
            }))
            {
                doc.Save(xmlWriter);
                updatedXml = stringWriter.ToString();
            }

            Console.WriteLine(_filePath + nameFile + ".docx");

            SaveAndValidate(_pathToTemplate, updatedXml, _filePath + nameFile + ".docx", nameFile);
        }

        public StatusDocx SaveAndValidate(string originalDocxPath, string modifiedXml, string outputDocxPath, string FileName)
        {
            SaveXml(originalDocxPath, modifiedXml, outputDocxPath, FileName);
            return isSave(outputDocxPath);
        }

        public StatusDocx isSave(string pathToCreatedDocx)
        {
            if (!File.Exists(pathToCreatedDocx))
            {
                return new StatusDocxError
                {
                    isCreated = false,
                    Exception = "Error while creating word file."
                };
            }

            return new StatusDocx
            {
                isCreated = true
            };
        }

        public StatusDocx SaveXml(string originalDocxPath, string modifiedXml, string outputDocxPath, string FileName)
        {
            try
            {
                using (FileStream fs = new FileStream(FileName, FileMode.Create))
                using (ZipArchive originalArchive = ZipFile.Open(originalDocxPath, ZipArchiveMode.Read))
                using (ZipArchive newArchive = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    foreach (var entry in originalArchive.Entries)
                    {
                        if (entry.FullName == "word/document.xml")
                        {
                            var newEntry = newArchive.CreateEntry("word/document.xml");
                            using (var entryStream = newEntry.Open())
                            using (var writer = new StreamWriter(entryStream))
                            {
                                writer.Write(modifiedXml);
                            }
                        }
                        else
                        {
                            var newEntry = newArchive.CreateEntry(entry.FullName);
                            using (var entryStream = newEntry.Open())
                            using (var originalStream = entry.Open())
                            {
                                originalStream.CopyTo(entryStream);
                            }
                        }
                    }

                }
                return new StatusDocx { isCreated = true };
            }
            catch (FileNotFoundException ex)
            {
                return new StatusDocxError { isCreated = false, Exception = $"Шаблон не знайдено: {ex.Message}" };
            }
            catch (UnauthorizedAccessException ex)
            {
                return new StatusDocxError { isCreated = false, Exception = $"Нема дозволе для запису: {ex.Message}" };
            }
            catch (InvalidDataException ex)
            {
                return new StatusDocxError { isCreated = false, Exception = $"Щось не так з файлом Docx: {ex.Message}" };
            }
            catch (IOException ex)
            {
                return new StatusDocxError { isCreated = false, Exception = $"Помилка при роботі з файлом (можливо він зайнятий іншим процесом): {ex.Message}" };
            }
            catch (Exception ex)
            {
                return new StatusDocxError {
                    isCreated = false,
                    Exception = $"Не передбачуванна помилка {ex.Message}"
                };
            }

        }
    }
}