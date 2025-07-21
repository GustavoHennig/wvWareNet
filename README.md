# vwWare-toDotNet – *A Failed Attempt*

> **Goal (2025):** Build a **pure C#** parser that extracts **plain text** from legacy Word `.doc` files—no native DLLs, no interop, just managed code.

---

## What on earth is this repo?

This repository is my personal logbook of a **frustrated/failed attempt** to port the classic **wvWare** parser to modern .NET.

I wanted to answer a burning question:

> **“Can today’s AI agents manage the task of porting a legacy C codebase, including its tricky binary formats?”**

### Short answer → *Almost.*

* Word 6 text extraction? **Perfect**  
* Standard Word 97 docs? **Mostly fine**  
* Fast-saved, corrupted, or otherwise “creative” docs? **Nope**—the AI-generated code collapsed under edge cases that the original C code handled just fine.

After weeks of patch-and-pray I pivoted:

1. I discovered **[b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)**—an abandoned but *well-structured* C# codebase that converts `.doc` → `.docx` (no text extraction).  
2. I poured the hard-won knowledge from **this** repo into **that** one.  
3. With AI agents’ help on specific pain points (fast-saved streams, damaged piece tables, etc.), **b2xtranslator** was resurrected and now produces reliable plain text. 🎉

> **Need a working DOC-to-text solution right now?**  
> Check out **<https://github.com/gustavohennig/b2xtranslator>**.  
> This repo remains online purely as a learning chronicle.

---

## Why keep a failed attempt public?

| Reason              | Detail                                                                                 |
|---------------------|----------------------------------------------------------------------------------------|
| **Historical value**| Shows where AI shines (boilerplate, repetitive structs) and where it stumbles (subtle state machines, excessive context windows, complex code) |
| **Reference snippets** | Some routines are correct and may help other ports                                 |
| **Storytelling**    | Failure, pivot, success—useful for anyone weighing AI pair-programming on binary formats |

---

## Original Scope

- **Input**: Microsoft Word `.doc` files (pre-2007 binary format)  
  - Word 6.0 (v6)  
  - Word 95 (v7)  
  - Word 97 (v8)  
- **Output**: Complete text extraction, document structure metadata, formatting relationships  
- No support for layout, images, annotations, or embedded objects

---

## Current State of the Code

| Component                  | Status    |
|----------------------------|-----------|
| Word 6 parser              | **Stable**   |
| Word 95 / 97 basic         | **Usable**   |
| Header / Footer / TextBox  | **Usable**   |
| Fast-saved documents       | **Unreliable** |
| Word track changes         | **Poor**     |
| Corruption handling        | **Poor**     |
| Maintenance                | **Frozen** (see b2xtranslator for ongoing work) |

---

## Design Constraints

* 100 % managed C# (.NET 9)  
* No P/Invoke or native DLLs  
* Minimal external dependencies  
* Readable first, performant later  

---

## Reference Projects that Guided Me

| Name              | Language | Description                                                          | Link                                                                  |
|-------------------|----------|----------------------------------------------------------------------|-----------------------------------------------------------------------|
| **wvWare**        | C        | Original GPL Word97 `.doc` text extractor                            | [SourceForge](https://sourceforge.net/projects/wvware/)               |
| **OnlyOffice**    | C++      | Proprietary editor with open-source core, includes DOC parsing       | [GitHub](https://github.com/ONLYOFFICE/core/tree/master/MsBinaryFile) |
| **Antiword**      | C        | Lightweight Word `.doc` to text/postscript converter                 | [GitHub Mirror](https://github.com/grobian/antiword)                  |
| **Apache POI**    | Java     | Java API for Microsoft Documents, includes Word97 support via HWPF   | [Apache POI - HWPF](https://poi.apache.org/hwpf/index.html)           |
| **b2xtranslator** | C#       | Open XML SDK-based translator, also parses legacy binary formats     | [GitHub](https://github.com/EvolutionJobs/b2xtranslator)              |
| **LibreOffice**   | C++      | Full office suite with robust support for legacy DOC files           | [GitHub](https://github.com/LibreOffice/core)                         |
| **Catdoc**        | C        | Simple Word `.doc` to text converter                                 | [GitHub Mirror](https://github.com/petewarden/catdoc)                 |
| **DocToText**     | C++      | Generic document-to-text converter                                   | [GitHub](https://github.com/tokgolich/doctotext)                      |

---

## Building & Running

```bash
dotnet restore
dotnet run -- sample.doc
```

---

## Binary Format Spec

The official **Microsoft Office Binary File Format Specification (MS-DOC)** PDF is included at `/docs/spec/MS-DOC-spec-compressed.pdf`.  
All trademarks belong to Microsoft; redistributed here under the spec’s license terms.

---

## License

GPL-2.0-only, identical to the original **wvWare**.

---

## Author

Gustavo Augusto Hennig
