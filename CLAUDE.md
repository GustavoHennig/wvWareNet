# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a complete rewrite of the wvWare Word97 document parser in C#. The project aims to create a 100% managed codebase with no native dependencies for extracting text from Microsoft Word .doc files (Word 6.0/95/97 binary format).

## Build and Test Commands

### Building the Project
```bash
dotnet build
```

### Running Tests
```bash
dotnet test
```

### Restoring Dependencies
```bash
dotnet restore
```

### Running the Console Application
```bash
dotnet run --project WvWareNetConsole -- <path-to-doc-file>
```

## Architecture and Code Structure

### Core Components

1. **WvDocExtractor** (`WvWareNet/WvDocExtractor.cs`): Main entry point for text extraction
   - Handles file format detection (Word95 vs Word97+)
   - Manages encryption detection and Word95 decryption
   - Coordinates with CFBF parser and WordDocumentParser

2. **CompoundFileBinaryFormatParser** (`WvWareNet/Parsers/CompoundFileBinaryFormatParser.cs`): 
   - Parses OLE Compound File Binary Format (CFBF) structure
   - Extracts WordDocument, Table, and other streams from .doc files

3. **WordDocumentParser** (`WvWareNet/Parsers/WordDocumentParser.cs`):
   - Implements MS-DOC specification paragraph boundary detection
   - Handles File Information Block (FIB) parsing
   - Manages piece table reconstruction for text extraction

4. **Core Data Models** (`WvWareNet/Core/`):
   - `DocumentModel.cs`: Document structure with sections, paragraphs, runs
   - `FileInformationBlock.cs`: FIB structure containing document metadata
   - `PieceTable.cs`: Manages document text reconstruction from pieces
   - `Stylesheet.cs`: Document style information
   - Properties classes for character, paragraph, and section formatting

5. **Utilities** (`WvWareNet/Utilities/`):
   - `ILogger.cs`: Logging interface with implementations
   - `Word95Decryptor.cs`: Handles Word95 encryption/decryption
   - `FileStreamHandler.cs`: File I/O utilities

### Key Technical Details

- **Target Framework**: .NET 9.0
- **Encoding**: Uses `System.Text.Encoding.CodePages` for legacy text encodings
- **Testing**: xUnit framework with extensive test coverage using example documents
- **Encryption**: Supports Word95 encrypted documents with password decryption

### Word Document Format Support

- **Word 6.0** (v6): Basic support
- **Word 95** (v7): Full support including encrypted documents
- **Word 97** (v8): Full support with complex document structures

### Text Extraction Strategy

1. **Primary Method**: Uses MS-DOC paragraph boundary detection algorithm
2. **Fallback Method**: Simple text extraction when paragraph boundaries fail
3. **Piece Table Reconstruction**: Handles complex documents with multiple text pieces

### Development Notes

- The project includes extensive example documents in `examples/` for testing
- Multiple working implementations exist in different directories (WvWareNet, WvWareNet95-97Working, WvWareNetWord97Working)
- The main implementation is in `WvWareNet/`
- Documentation files (PLAN.md, PROGRESS.md, GEMINI.md) track development progress
- When making changes, ensure MD files are updated per README requirements

### Testing

- Tests are located in `WvWareNet.Tests/`
- Use `DocParsingTests.cs` for document parsing functionality
- Use `ExamplesDocTests.cs` for testing against example documents
- Example documents have corresponding `.expected.txt` files for comparison

### Important Constraints

- No P/Invoke or native code dependencies
- Text extraction only (no layout, images, or embedded objects)
- GPL licensed (see COPYING file)
- Focus on text extraction accuracy over formatting preservation