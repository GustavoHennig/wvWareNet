# vwWare-toDotNet Progress

## 2025-06-30

-   **Detailed Analysis of `decode_complex.c`**: Performed an in-depth comparison between the legacy C code and the current C# implementation.
    -   **Finding**: The C# `WordDocumentParser` correctly implements the high-level structure for parsing complex documents by reading the Piece Table (CLX) and locating property tables (PLCF).
    -   **Gap Identified**: The current implementation lacks a robust, generic mechanism for applying property modifications ("sprms"), which is the core of Word's formatting. The C# code currently uses a hardcoded, limited approach.
    -   **Conclusion**: To accurately handle text formatting (especially lists, indentation, and detailed styles), a full SPRM application framework is required.

-   **Created New Porting Plan**: A new, detailed plan has been created as `PORTING_PLAN.md`.
    -   This plan focuses specifically on porting the remaining text-centric features from `decode_complex.c`.
    -   It prioritizes the implementation of a SPRM handling framework, followed by list/list-override parsing, and finally section/document property integration.
    -   The old `PLAN.md` will be preserved for historical purposes.

## 2025-06-24

- Project direction changed: Full C# rewrite initiated
- README.md updated to reflect new goals
- Detailed rewrite plan created in PLAN.md
- .NET 9 solution and project structure created
- Coding standards established

## Current Status

-   Phase 1 (Project Setup) completed.
-   Phase 2 (Core Data Structures) is largely complete.
-   **Analysis of complex document parsing is complete.** The next stage of implementation is defined in `PORTING_PLAN.md`.
-   The core parsing logic in `WordDocumentParser` can successfully read the main document streams (`WordDocument`, `Table`) and the piece table (`CLX`).

### Solution Layout

```
vwWare-toDotNet.sln
└── WvWareNet/
    ├── WvWaveDotNet.csproj
    ├── Core/
    │   ├── FileInformationBlock.cs
    │   ├── PieceDescriptor.cs
    │   ├── PieceTable.cs
    │   ├── CharacterProperties.cs
    │   └── ParagraphProperties.cs
    ├── Parsers/
    │   └── CompoundFileBinaryFormatParser.cs
    ├── Utilities/
    │   ├── ILogger.cs
    │   ├── ConsoleLogger.cs
    │   └── FileStreamHandler.cs
    ├── bin/
    └── obj/
```

## Phase 3: File Parsing Progress (Completed)

- [x] CFBF reader with FAT chain traversal and sector reading
- [x] Directory entry parsing implemented
- [x] Stream reading logic (using FAT chain and sector size)
- [x] Document parsing scaffold
- [x] Text extraction logic (basic)
- [x] Document model structure

## Phase 4: Document Model (Completed)

- [x] Enhanced in-memory document representation to support sections, runs with formatting, headers, footers, and footnotes
- [x] Implemented SectionProperties class to handle section-level properties

## Phase 5: Text Extraction & Formatting — 🚧 In Progress

- [x] Implement text decoding logic
- [x] Handle character encoding conversions
- [x] Process piece tables to extract text
- [x] Build plain text output mechanism
- [x] Added FIB version detection and encryption check
- [x] Parse and extract headers, footers, and footnotes (basic implementation)
- [ ] **Implement comprehensive SPRM (property modification) handling framework**
    - [ ] Create core data structures for SPRMs (`Sprm.cs`, `SprmDefinitions.cs`).
    - [ ] Develop a `SprmApplicator` service to parse and apply `grpprl` byte arrays.
- [ ] **Parse and apply detailed paragraph formatting (PAPX)**
    - [ ] Enhance `ParagraphProperties.cs` to hold all text-relevant properties (indentation, justification, list info).
    - [ ] Apply paragraph SPRMs using the new applicator service.
- [ ] **Parse and apply detailed character formatting (CHPX)**
    - [ ] Enhance `CharacterProperties.cs` to hold all text-relevant properties (underline styles, etc.).
    - [ ] Apply character SPRMs using the new applicator service.
- [ ] **Implement full list and list override parsing**
    - [ ] Parse `PLCF LST` and `PLCF LFO` from the Table stream.
    - [ ] Create a `ListManager` to track list state and generate formatted list prefixes (e.g., "1.", "a)", "- ").

## Phase 6: API and Integration — 🚧 In Progress

- [x] Created `WvDocExtractor` public interface
- [x] Implemented `ExtractText()` method
- [x] Added basic error handling
- [x] Built console test application (in `WvWareNetConsole` project)

## Next Steps

The immediate focus is on building the foundational SPRM handling framework as outlined in `PLAN_DECODE_COMPLEX.md`.

1.  **Phase 1 (from `PLAN_DECODE_COMPLEX.md`): Foundational - Comprehensive SPRM Handling**
    -   Implement core SPRM data structures.
    -   Develop the `SprmApplicator` service.
    -   Implement the application logic for paragraph (PAP) and character (CHP) properties.

## Immediate Tasks

-   Create `Sprm.cs` and `SprmDefinitions.cs` to model property modifications.
-   Begin implementation of the `SprmApplicator.cs` service, starting with the logic to parse a `grpprl`.

## Legacy Port Tracking

This table tracks wvWare C files that were used as references to understand the DOC format and guide the C# implementation. No code was copied or directly ported; all implementation is original to this repository.

| C Source File      | Status    | Notes                                                                                                                            |
| ------------------ | --------- | -------------------------------------------------------------------------------------------------------------------------------- |
| wvWare.c           | [ ]       | Main logic entry point                                                                                                           |
| wvTextEngine.c     | [ ]       | Text extraction core                                                                                                             |
| wvparse.c          | [ ]       | Parser logic                                                                                                                     |
| **decode_complex.c** | **[✓] Analyzed** | **Analysis complete. Porting plan created in `PLAN_DECODE_COMPLEX.md`. Lacks robust SPRM handling, which is the next major task.** |
| fib.c              | [x]       | File Information Block                                                                                                           |
| plcf.c             | [x]       | Piece table logic                                                                                                                |
| text.c             | [ ]       | Text-related parsing                                                                                                             |
| unicode.c          | [ ]       | Unicode helpers                                                                                                                  |
| utf.c              | [ ]       | Encoding helpers                                                                                                                 |
| wvConfig.c         | [ ]       | Possibly minimal, for setup                                                                                                      |

For complete legacy file list, see original progress archive.
