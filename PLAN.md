# PLAN.md

# vwWare C# Rewrite Plan

## Project Goals

- Port only the "extract-to-text" feature from wvWare
- Create 100% managed C# implementation targeting .NET 9
- No P/Invoke or native dependencies
- Focus exclusively on Microsoft Word `.doc` (binary, pre-2007) files
- Output: Plain text extracted from documents
- Follow idiomatic C# patterns (not a line-by-line port)
- Unit tests are not a priority at this stage

## Feature Scope

- **Input**: Microsoft Word `.doc` (binary, pre-2007) files
- **Output**: Plain text extracted from documents

### Required Functionality

- Parse Compound File Binary Format (CFBF)
- Navigate Word document structure (FIB, piece tables, text runs)
- Decode text (Unicode, encodings, conversions)
- Output complete document structure with formatting

## Architecture

### Core Components
- **WvDocExtractor** (main class)
  - Loads `.doc` files
  - Exposes `ExtractText()` method
- **DocParser**
  - Handles CFBF parsing
  - Processes Word document structures
- **TextDecoder**
  - Manages encoding and Unicode conversion
- **DocumentModel**
  - In-memory representation of Word document structure

### Dependencies

- .NET `System.IO` for file access
- Minimal custom CFBF parser (since CFBF is not natively supported in .NET)
- No external or third-party dependencies unless strictly necessary

## Phase Breakdown

### Phase 1: Project Setup — ✅ Complete

- [x] Defined rewrite goals and scope
- [x] Updated documentation (`README`, `PROGRESS`)
- [x] Created .NET 9 solution/project structure
- [x] Established coding standards and conventions
- [x] Updated `README.md` to reflect current project status

### Phase 2: Core Data Structures

- [ ] Implement essential structs/classes:
  - File Information Block (FIB)
  - Character Properties (CHP)
  - Paragraph Properties (PAP)
  - Piece Table structures
- [ ] Create base parsing utilities

### Phase 3: File Parsing — ✅ Complete

- [x] Implement `.doc` file header parsing
- [x] Develop stream reading utilities
- [x] Create custom CFBF reader
- [x] Parse Word document streams (WordDocument, Table, etc.)

### Phase 4: Document Model — ✅ Complete

- [x] Build in-memory document representation
- [x] Implement text run extraction
- [x] Handle document structure navigation

### Phase 5: Text Extraction — 🚧 In Progress

- [x] Implement text decoding logic
- [x] Handle character encoding conversions
- [x] Process piece tables to extract text
- [x] Build document structure and formatting parser
- [x] Added FIB version detection and encryption check
- [ ] Parse and apply character formatting (CHPX/PLCFCHPX) to runs
- [ ] Convert CHPX to CharacterProperties for each run
- [ ] Parse and extract headers, footers, and footnotes
- [ ] Output text with formatting and document structure

### Phase 6: API and Integration — ✅ Complete

- [x] Create `WvDocExtractor` public interface
- [x] Implement `ExtractText()` method
- [x] Handle edge cases and error conditions
- [x] Build basic console test application
- [x] The output of "WvWareNetConsole\news-example.doc" must run without errors and the sentence "Existe uma Caixa de texto aqui" must be present in the output.

### Phase 7: Optimization & Finalization

- [ ] Performance tuning
- [ ] Memory optimization
- [ ] Error handling improvements
- [ ] Code cleanup and documentation

## Reference C Files

The following wvWare C source files were used as references to understand the DOC format and guide the architecture of this project. No code was copied or directly ported; all implementation is original to this repository.

- `wvWare.c` (main entry point)
- `wvTextEngine.c` (text extraction logic)
- `wvparse.c`, `fib.c`, `plcf.c`, `text.c` (parsing structures)
- `unicode.c`, `utf.c` (encoding helpers)
- `wvConfig.c` (configuration)

## Next Steps
1. Begin implementing core data structures (Phase 2)
2. Create CFBF parser prototype
3. Set up basic document parsing scaffold
