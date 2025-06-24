# vwWare C# Rewrite Plan

## Project Goals
- Port only the "extract-to-text" feature from wvWare
- Create 100% managed C# implementation (.NET 9)
- No P/Invoke or native dependencies
- Focus on Microsoft Word `.doc` (binary, pre-2007) files
- Output: Plain text extracted from documents
- Follow C# idiomatic patterns (not line-by-line port)
- Unit tests are not priority at this stage

## Feature Scope
- Input: Microsoft Word `.doc` (binary, pre-2007) files
- Output: Plain text extracted from documents
- Required functionality:
  - Parse Compound File Binary Format (CFBF)
  - Navigate document structure (FIB, piece tables, text runs)
  - Decode text (handling encodings, Unicode, etc.)
  - Output plain text

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
- .NET System.IO for file access
- Minimal custom CFBF parser (since CFBF not natively supported in .NET)
- No external dependencies

## Phase 1: Project Setup - COMPLETE
- [x] Define rewrite goals and scope
- [x] Update documentation (README, PROGRESS)
- [x] Create .NET 9 solution/project structure
- [x] Establish coding standards and conventions

## Phase 2: Core Data Structures
- [ ] Implement essential structs/classes:
  - File Information Block (FIB)
  - Character Properties (CHP)
  - Paragraph Properties (PAP)
  - Piece Table structures
- [ ] Create base parsing utilities

## Phase 3: File Parsing
- [ ] Implement .doc file header parsing
- [ ] Develop stream reading utilities
- [ ] Create CFBF reader
- [ ] Parse Word document streams (WordDocument, Table, etc.)

## Phase 4: Document Model
- [ ] Build in-memory document representation
- [ ] Implement text run extraction
- [ ] Handle document structure navigation

## Phase 5: Text Extraction
- [ ] Implement text decoding logic
- [ ] Handle character encoding conversions
- [ ] Process piece tables to extract text
- [ ] Build text output mechanism

## Phase 6: API and Integration
- [ ] Create WvDocExtractor public interface
- [ ] Implement ExtractText() method
- [ ] Handle edge cases and error conditions
- [ ] Basic console test application

## Phase 7: Optimization and Refinement
- [ ] Performance tuning
- [ ] Memory optimization
- [ ] Error handling improvements
- [ ] Code cleanup and documentation

## Key C Files to Reference
- `wvWare.c` (main entry point)
- `wvTextEngine.c` (text extraction logic)
- `wvparse.c`, `fib.c`, `plcf.c`, `text.c` (parsing structures)
- `unicode.c`, `utf.c` (encoding helpers)
- `wvConfig.c` (configuration)

## Next Steps
1. Begin implementing core data structures (Phase 2)
2. Create CFBF parser prototype
3. Set up basic document parsing scaffold
