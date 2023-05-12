Document .Net is .NET assembly (SDK) which gives API to Create, Read, Write, Edit, View, Convert, Merge, Sign Digitally, Find and Replace, Do Reporting with PDF, DOC, DOCX, HTML, RTF and Text documents, Rasterize to Image.

## Quick links
===============================
+ [Developer Guide](https://sautinsoft.com/products/document/help/net/developer-guide/create-document.php)
+ [API Reference](https://sautinsoft.com/products/document/help/net/api-reference/html/R_Project_Document__Net_-_API_Reference.htm)
+ [Document Object Model](https://sautinsoft.com/products/document/help/net/getting-started/document-object-model.php)

## Top Features
===============================
+ [Create a document in PDF, DOCX, RTF, HTML formats.](https://sautinsoft.com/products/document/help/net/developer-guide/create-document.php)
+ [Load a document in PDF, DOCX, RTF, HTML formats.](https://sautinsoft.com/products/document/help/net/developer-guide/load-document.php)
+ [Save a document as PDF, DOCX, RTF, HTML Flowing and Fixed, Text.](https://sautinsoft.com/products/document/help/net/developer-guide/save-document.php)
+ [Convert and Merge documents.](https://sautinsoft.com/products/document/help/net/developer-guide/convert-document.php)
+ [Protect and encrypt documents.](https://sautinsoft.com/products/document/help/net/developer-guide/security-options-net-csharp-vb.php)
+ [Digitally sign documents.](https://sautinsoft.com/products/document/help/net/developer-guide/digital-signature-net-csharp-vb.php)
+ [Make PDF/A compliance.](https://sautinsoft.com/products/document/help/net/developer-guide/create-and-save-document-in-pdf-a-format-net-csharp-vb.php)
+ [Perform Mail Merge process.](https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-simple-report-winforms-net-csharp-vb.php)
+ [Obvious Document Object Model.](https://sautinsoft.com/products/document/help/net/getting-started/document-object-model.php)
+ [Rasterize a document or specific pages to Image.](https://sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-document-pages-as-picture-net-csharp-vb.php)
+ [Get and set a Header / Footer, Page Setup.](https://sautinsoft.com/products/document/help/net/developer-guide/headersfooters.php)
+ [Add Page Numbering, Text Columns.](https://sautinsoft.com/products/document/help/net/developer-guide/page-numbering.php)
+ [Perform Pagination and get document pages.](https://sautinsoft.com/products/document/help/net/developer-guide/pagination.php)
+ [Find and Replace content.](https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-specific-page-after-specific-text-net-csharp-vb.php)
+ [Manipulate with Tables, Paragraphs and Text.](https://sautinsoft.com/products/document/help/net/developer-guide/contentrange-manipulation.php)
+ [Add and Extract Pictures.](https://sautinsoft.com/products/document/help/net/developer-guide/add-pictures.php)
+ [Work with Shapes, Shape groups and Geometry.](https://sautinsoft.com/products/document/help/net/developer-guide/geometry.php)
+ [Get and set Formatting and Styles.](https://sautinsoft.com/products/document/help/net/developer-guide/formatting-and-styles.php)
+ [Work with ordered and unordered Lists.](https://sautinsoft.com/products/document/help/net/developer-guide/create-multilevel-list-in-docx-document-net-csharp-vb.php)
+ [Insert and Update TOC - table of contents.](https://sautinsoft.com/products/document/help/net/developer-guide/update-table-of-contents-in-word-document-net-csharp-vb.php)
+ [Work with Forms and Fields.](https://sautinsoft.com/products/document/help/net/developer-guide/forms-and-fields.php)
+ [Change Document Properties: Author, Title, Creator, etc.](https://sautinsoft.com/products/document/help/net/developer-guide/document-properties-net-csharp-vb.php)

## System Requirement
===============================

* .NET Framework 4.6.1 - 4.8.1
* .NET Core 2.0 - 3.1, .NET 5, 6, 7, 8
* .NET Standard 2.0
* Windows, Linux, macOS, Android, iOS.

## Getting Started with Document .Net
===============================
Are you ready to give Document .NET a try? Simply execute `Install-Package sautinsoft.document` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Document .NET and want to upgrade the version, please execute `Update-Package sautinsoft.document` to get the latest version.

## Convert DOCX to PDF

```csharp
string inpFile = @"..\..\example.docx";
string outFile = @"Result.pdf";
DocumentCore dc = DocumentCore.Load(inpFile);
dc.Save(outFile);
```
## Create DOCX document on the fly

```csharp
Set a path to our document.
    string docPath = @"Result-DocumentBuilder.docx";

    // Create a new document and DocumentBuilder.
    DocumentCore dc = new DocumentCore();
    DocumentBuilder db = new DocumentBuilder(dc);

    // Set page size A4.
    Section section = db.Document.Sections[0];
    section.PageSetup.PaperType = PaperType.A4;

    // Add 1st paragraph with formatted text.
    db.CharacterFormat.FontName = "Verdana";
    db.CharacterFormat.Size = 16;
    db.CharacterFormat.FontColor = Color.Orange;
    db.Write("This is a first line in 1st paragraph!");
   
    // Save the document to the file in DOCX format.
    dc.Save(docPath, new DocxSaveOptions()
```

## Resources
===============================
+ **Website:** [www.sautinsoft.com](http://www.sautinsoft.com)
+ **Product Home:** [Document .Net](https://sautinsoft.com/products/document/)
+ [Download Document .Net](https://sautinsoft.com/products/document/download.php)
+ [Developer Guide](https://sautinsoft.com/products/document/help/net/developer-guide/create-document.php)
+ [API Reference](https://sautinsoft.com/products/document/help/net/api-reference/html/R_Project_Document__Net_-_API_Reference.htm)
+ [Document Object Model](https://sautinsoft.com/products/document/help/net/getting-started/document-object-model.php)
+ [Support Team](https://sautinsoft.com/support.php)
+ [License 📝](https://sautinsoft.com/products/document/help/net/getting-started/agreement.php)