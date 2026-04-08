# XDoc

A lightweight C# library for reading and processing `.docx` files via XML and XPath.

## Overview

DOCX files are ZIP archives containing XML. XDoc exposes a clean service-based API to open, parse, and query document content — without depending on Word or Office Interop.

## Project Structure

```
XDoc/
├── IFile.cs                  # Abstraction over a file resource
├── FileName.cs               # Value object representing a file name / path
├── Status.cs                 # Result / status type for operation outcomes
├── IXmlParser.cs             # Contract for XML parsing
├── XPathProcessor.cs         # XPath query execution against document XML
├── IDocxFileNameService.cs   # Service interface for resolving DOCX file names
├── IDocxService.cs           # Main service interface for DOCX operations
└── XDoc.csproj               # .NET project file
```

## Requirements

- .NET 8 or later
- C# 12+

## Installation

Clone the repository and reference the project directly:

```bash
git clone https://github.com/lxthxrgx/XDoc.git
```

Add to your solution:

```bash
dotnet sln add XDoc/XDoc.csproj
dotnet add YourProject reference ../XDoc/XDoc.csproj
```

## Usage

### Open and parse a DOCX file

```csharp
IDocxService docxService = // resolve via DI or instantiate
var result = docxService.Open(new FileName("document.docx"));

if (result.Status == Status.Ok)
{
    // work with parsed content
}
```

### Query content with XPath

```csharp
IXmlParser parser = // resolve
var processor = new XPathProcessor(parser);

var nodes = processor.Select("//w:t", document);
```

## Key Abstractions

| Interface | Responsibility |
|---|---|
| `IFile` | Represents a file resource (path, stream) |
| `IXmlParser` | Parses raw XML from a DOCX part |
| `IDocxFileNameService` | Resolves and validates DOCX file names |
| `IDocxService` | High-level DOCX open / read operations |

`Status` is used as a result discriminator — operations return a status value instead of throwing, keeping error handling explicit.

## License

MIT — see [LICENSE.txt](LICENSE.txt).
