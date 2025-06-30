# vwWare-toDotNet


This project is a complete, from-scratch write of a Word97 document parser based on `wvWare` in C#.

- **No native code dependencies**
- **No P/Invoke or interop**
- **All logic implemented natively in C#**

See `PLAN.md` for the detailed rewrite roadmap and `PROGRESS.md` for current development updates.

## Goals

- 100% managed codebase
- Modern, maintainable, and idiomatic C#
- Inspired by the core functionality of `wvWare`
- Focus on text extraction feature only

## Scope

- Input: Microsoft Word `.doc` (binary, pre-2007) files
- Must support Word 6.0, Word 95 (7.0), and Word 97 (8.0)
- Output: Plain text extracted from the document
- Only extract-to-text feature from original wvWare is targeted
- Not a line-by-line port, but a clean architectural redesign in C#

## Key Design Constraints

- No P/Invoke
- No external dependencies unless absolutely necessary
- Built using .NET 9
- Unit tests are not a priority during early stages
- IMPORTANT: When changing the code, make sure the md files are properly updated.

### Compatibility

The parser attempts to read documents produced by Word 6, Word 95 and Word 97.
Encrypted files are detected and will trigger a clear error message.

## Reference Project

- wvWare ([wvWave on SourceForge](https://sourceforge.net/projects/wvware/)) was used as a main reference to understand the DOC format and as inspiration for this project. No code was copied or directly ported; all implementation is original to this repository.

## References

- The main reference for this project is the official Microsoft Word Binary File Format specification, included as `MS-DOC-spec-compressed.pdf` in this repository. All rights to this document belong to Microsoft.
- AbiWord, OnlyOffice and LibreOffice codebases were consulted as secondary references to help understand the DOC file format.

## License

This project is based on the original wvWave, licensed under the GNU GPL. See COPYING for details.

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
