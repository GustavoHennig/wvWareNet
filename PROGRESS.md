# vwWare-toDotNet Progress

## 2025-06-24

- Project direction changed: Full C# rewrite initiated
- README.md updated to reflect new goals
- Detailed rewrite plan created in REWRITE\_PLAN.md
- .NET 9 solution and project structure created
- Coding standards established

## Current Status

- Phase 1 (Project Setup) completed
- Phase 2 (Core Data Structures) in progress

### Solution Layout

```
vwWare-toDotNet.sln
└── WvWareNet/
    ├── WvWaveDotNet.csproj
    ├── Class1.cs (placeholder)
    ├── bin/
    └── obj/
```

## Next Steps

1. Implement core data structures (Phase 2):
   - File Information Block (FIB)
   - Character Properties (CHP)
   - Paragraph Properties (PAP)
   - Piece Table structures
2. Create CFBF reader prototype
3. Set up basic document parsing scaffold
4. Initialize text extraction logic

## Immediate Tasks

- Create FIB structure implementation
- Research .NET 9 features for binary parsing
- Setup logging infrastructure
- Create basic file stream handling

## Legacy Port Tracking

This table tracks original `wvWare` C files, useful for reference in C# port. All logic is to be rewritten idiomatically.

| C Source File  | Status | Notes                       |
| -------------- | ------ | --------------------------- |
| wvWare.c       | [ ]    | Main logic entry point      |
| wvTextEngine.c | [ ]    | Text extraction core        |
| wvparse.c      | [ ]    | Parser logic                |
| fib.c          | [ ]    | File Information Block      |
| plcf.c         | [ ]    | Piece table logic           |
| text.c         | [ ]    | Text-related parsing        |
| unicode.c      | [ ]    | Unicode helpers             |
| utf.c          | [ ]    | Encoding helpers            |
| wvConfig.c     | [ ]    | Possibly minimal, for setup |

For complete legacy file list, see original progress archive.
