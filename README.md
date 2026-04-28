# XDoc

A lightweight .NET library for filling **Word (`.docx`) templates** using **Content Controls** (structured document tags). Designed for server-side document generation — no Microsoft Word installation required.

---

## How it works

XDoc treats a `.docx` file as a ZIP archive containing `word/document.xml`. It parses the XML, finds Content Control elements (`<w:sdt>`) by their tag names, substitutes values, and saves the result as a new `.docx` file.

```
Template .docx  →  ExtractXml  →  WriteXmlTree (×N)  →  Save  →  Output .docx
```

---

## Requirements

- .NET 10.0+
- `Microsoft.Extensions.Configuration` (≥ 10.0.5)
- No Microsoft Word needed on the server

---

## Installation

Clone the repository and reference the project directly:

```bash
git clone https://github.com/lxthxrgx/XDoc.git
```

Add to your `.csproj`:

```xml
<ProjectReference Include="../XDoc/XDoc.csproj" />
```

---

## Quick Start

### 1. Prepare your Word template

In Microsoft Word, insert **Content Controls** via:  
`Developer → Insert → Plain Text Content Control`

For each field, open **Properties** and set a unique **Tag** (e.g. `ClientName`, `ContractDate`).

### 2. Use `XPathProcessor` in your code

```csharp
using XDoc;

// Create a processor: load template and set output folder
var processor = new XPathProcessor(
    pathToTemplate: "templates/contract.docx",
    filepath: "output/"
);

// Fill in the fields by tag name
processor.WriteXmlTree("ClientName", "test name");
processor.WriteXmlTree("ContractDate", DateTime.Now.ToString("dd/MM/yyyy"));
processor.WriteXmlTree("Address", "test Address");

// Save the resulting document
processor.Save("contract-001");
// → saves to: output/contract-001.docx
```

---

## API Reference

### `XPathProcessor`

The main class. Implements `IFile`, `IXMLParser`, and `IDocxService`.

#### Constructor

```csharp
XPathProcessor(string pathToTemplate, string filepath)
```

| Parameter        | Description                                      |
|------------------|--------------------------------------------------|
| `pathToTemplate` | Path to the `.docx` template file                |
| `filepath`       | Output folder path where generated files are saved |

---

#### `WriteXmlTree<T>(string tag, T value)`

Finds a Content Control by its **Tag** and sets its text value.

```csharp
processor.WriteXmlTree("ContractNumber", "2024-001");
processor.WriteXmlTree("Amount", 15000.00m);
```

| Parameter | Description                                      |
|-----------|--------------------------------------------------|
| `tag`     | The tag name set in Word Content Control properties |
| `value`   | Value to insert (any type, converted via `.ToString()`) |

> If the tag is not found in the document, the call is silently skipped.

---

#### `Save(string nameFile)`

Serializes the modified XML and saves the output as a `.docx` file.

```csharp
processor.Save("report-january");
// → output/report-january.docx
```

| Parameter  | Description                              |
|------------|------------------------------------------|
| `nameFile` | Output file name (without `.docx` extension) |

---

#### `GetXmlTree()` → `List<XmlSdtItem>`

Returns a list of all Content Controls found in the template. Useful for debugging or inspecting available tags.

```csharp
var fields = processor.GetXmlTree();
foreach (var item in fields)
{
    Console.WriteLine($"Tag: {item.Tag}, Alias: {item.Alias}, Current value: {item.Text}");
}
```

**`XmlSdtItem` properties:**

| Property | Description                              |
|----------|------------------------------------------|
| `Tag`    | Tag identifier of the Content Control    |
| `Alias`  | Human-readable label (set in Word)       |
| `Text`   | Current text value inside the control    |

---

#### `ExtractXml(string docxPath)` → `string`

Extracts raw `word/document.xml` content from a `.docx` file as a string.

```csharp
string xml = processor.ExtractXml("templates/contract.docx");
```

---

#### `SaveXml(...)` → `StatusDocx`

Low-level method. Replaces `word/document.xml` inside a `.docx` archive with modified XML and saves the result to a new file.

```csharp
StatusDocx status = processor.SaveXml(
    originalDocxPath: "templates/contract.docx",
    modifiedXml: xmlString,
    outputDocxPath: "output/contract-001.docx",
    FileName: "contract-001"
);
```

---

### `StatusDocx` — result classes

All save operations return a `StatusDocx` object indicating success or failure.

```csharp
StatusDocx status = processor.SaveXml(...);

if (!status.isCreated)
{
    if (status is StatusDocxError error)
        Console.WriteLine($"Error: {error.Exception}");
}
```

| Class               | Property    | Description                        |
|---------------------|-------------|------------------------------------|
| `StatusDocx`        | `isCreated` | `true` if the file was created     |
| `StatusDocxWarning` | `Warning`   | Non-critical warning message       |
| `StatusDocxError`   | `Exception` | Error message if creation failed   |

---

### Interfaces

| Interface               | Description                                                  |
|-------------------------|--------------------------------------------------------------|
| `IXMLParser`            | `GetXmlTree()` and `WriteXmlTree<T>()` for Content Controls |
| `IFile`                 | `Save(string nameFile)` — save the document                 |
| `IDocxService`          | `ExtractXml()` and `SaveXml()` — low-level XML operations   |
| `IDocxFileNameService`  | Reserved for file naming strategies                         |

---

## Real-world example

Generating a lease agreement with multiple document types:

```csharp
public async Task GenerateSubleaseContract(ContractData data)
{
    var processor = new XPathProcessor(
        pathToTemplate: _config["Templates:sublease-tov:contract"],
        filepath: _config["Paths:OutputFolder"]
    );

    processor.WriteXmlTree("DogovirSuborendu", data.Sublease.ContractNumber);
    processor.WriteXmlTree("DateTime", data.Sublease.ContractSigningDate.ToString("dd/MM/yyyy"));
    processor.WriteXmlTree("address_p", data.Group.Address);
    processor.WriteXmlTree("PIB", data.Counterparty.Fullname);
    processor.WriteXmlTree("rnokpp", data.Counterparty.Rnokpp);

    processor.Save($"{data.Sublease.ContractNumber}-{data.Group.Name}-договір");
}
```

---

## Project structure

```
XDoc/
├── XPathProcessor.cs          # Core implementation
├── IXmlParser.cs              # Interface: read/write Content Controls
├── IFile.cs                   # Interface: save document
├── IDocxService.cs            # Interface: low-level XML operations
├── IDocxFileNameService.cs    # Interface: file naming
├── Status.cs                  # StatusDocx, StatusDocxWarning, StatusDocxError
├── FileName.cs                # Internal helper
└── XDoc.csproj
```

---

## License

MIT © lxthxrgx
