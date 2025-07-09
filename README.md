# vwWare-toDotNet (TODO: rename me)


This project is a complete, from-scratch write of a Word97 document parser based on `wvWare` in C#.

- **No native code dependencies**
- **No P/Invoke or interop**
- **All logic implemented natively in C#**

See `PLAN.md` for the detailed rewrite roadmap and `PROGRESS.md` for current development updates.

## Goals

- 100% managed codebase
- Modern, maintainable, and idiomatic C#
- Inspired by the core functionality of `wvWare`
- Focus on extracting text, complete document structure, formatting, and relationships between elements

## Scope

- **Input**: Microsoft Word `.doc` files (pre-2007 binary format)
  - Word 6.0 (v6)
  - Word 95 (v7)
  - Word 97 (v8)
- **Output**: Complete text, document structure, formatting, and relationships between elements
- No support for layout, images, annotations, or embedded objects


## Key Design Constraints

- No P/Invoke
- No external dependencies unless absolutely necessary
- Built using .NET 9
- Unit tests are not a priority during early stages
- IMPORTANT: When changing the code, make sure the md files are properly updated.

### Compatibility

The parser attempts to read documents produced by Word 6, Word 95 and Word 97.
Encrypted files are detected and will trigger a clear error message.


## Reference Projects and Implementations

This project is inspired by and informed by several existing open-source implementations of the Word Binary Format:

| Name              | Language     | Description                                                              | Link                                                                  |
|-------------------|--------------|--------------------------------------------------------------------------|-----------------------------------------------------------------------|
| **wvWare**        | C            | Original GPL Word97 `.doc` text extractor                                | [SourceForge](https://sourceforge.net/projects/wvware/)               |
| **OnlyOffice**    | C++          | Proprietary editor with open-source core, includes DOC parsing           | [GitHub](https://github.com/ONLYOFFICE/core/tree/master/MsBinaryFile) |
| **Antiword**      | C            | Lightweight Word `.doc` to text/postscript converter                     | [GitHub Mirror](https://github.com/grobian/antiword)                         |
| **Apache POI**    | Java         | Java API for Microsoft Documents, includes Word97 support via HWPF       | [Apache POI - HWPF](https://poi.apache.org/hwpf/index.html)           |
| **b2xtranslator** | C#           | Open XML SDK-based translator, also parses legacy binary formats         | [GitHub](https://github.com/EvolutionJobs/b2xtranslator)              |
| **LibreOffice**   | C++          | Full office suite with robust support for legacy DOC files               | [GitHub](https://github.com/LibreOffice/core)                         |
| **Catdoc**        | C            | Lightweight Word `.doc` to text converter                                | [GitHub Mirror](https://github.com/petewarden/catdoc)                        |
| **DocToText**     | C++          | Lightweight any document file to text converter                          | [GitHub](https://github.com/tokgolich/doctotext)                      |




## Binary Format Specification

This implementation relies heavily on:

- **Microsoft Office Binary File Format Specification**  
  Included as `MS-DOC-spec-compressed.pdf` in this repository.  
  All rights to this document belong to Microsoft.


## License

Licensed under the GNU GPL. See COPYING for details.

## Running Tests

To run the tests for this project:

1. Restore dependencies:
   ```
   dotnet restore
   ```

2. Run the tests:
   ```
   dotnet test
   ```

3. Optionally, generate code coverage reports:
   ```
   dotnet test --collect:"XPlat Code Coverage"
   ```
