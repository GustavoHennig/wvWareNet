# vwWare-toDotNet

## Project Status: Full C# Rewrite — 🚧 In Progress

This project is a complete, from-scratch rewrite of the original `wvWare` (Word document parser) in modern C#.

- **No native code dependencies**
- **No P/Invoke or interop**
- **All logic implemented natively in C#**

See `REWRITE_PLAN.md` for the detailed rewrite roadmap and `PROGRESS.md` for current development updates.

## Goals

- 100% managed codebase
- Modern, maintainable, and idiomatic C#
- Faithful to the original `wvWare` functionality
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

### Compatibility

The parser attempts to read documents produced by Word 6, Word 95 and Word 97.
Encrypted files are detected and will trigger a clear error message.

## Original Project

- Based on: [wvWave on SourceForge](https://sourceforge.net/projects/wvware/)
- This project is a modern reimplementation in C#.

## License

This project is based on the original wvWave, licensed under the GNU GPL. See COPYING for details.
