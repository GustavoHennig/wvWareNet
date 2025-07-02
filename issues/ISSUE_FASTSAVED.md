The parser is not retrieving the following fragment from the file: `examples\fastsavedmix.doc`:
```
Fastsaved italic text
Fastsaved bold text
```

The complete expected result for the file `fastsavedmix.doc` is:
```
Italic normal text.
Bold normal text.
Big normal text
Fastsaved italic text
Fastsaved bold text
```

## To run

```shell
dotnet run --project .\WvWareNetConsole\WvWareNetConsole.csproj .\examples\fastsavedmix.doc
```

## Additional Info

Read the markdown files in this project if you need more information.
Please check `/wvWare` folder as reference to understand how the parse works (C version).
You can also check the folder `/OnlyOffice-MsBinaryFile` there is a totally different parser written in CPP.

## 🧠 What is "Fast Save" in Word 97?
"Fast Save" was a legacy feature that, when enabled, caused Word to append changes to the end of the .doc file, rather than rewriting the whole file. This improved performance for saving large documents on slow disks. The resulting file contains:

Original text and structure

A chain of edits and additions stored after the main text

A "save history" that can grow over time

So a fast-saved .doc file has an append-only log-like structure at the end, which must be merged to reconstruct the final state.

This is also why tools like wvWare, AbiWord, and LibreOffice need to detect and discard obsolete portions of the document, using markers in the FIB (File Information Block) and Piece Table.

### Other ways it’s referenced
In other tools or codebases, you might see references to:

"fWhichTblStm" or "fEncrypted" or fDot in the FIB flags — these bits sometimes mark fast saves

"clx" structure (Complex File) and "Piece Table" logic — both are central to decoding fast-saved content

"Undo tables", "save history", or "extra streams" — vague but related terms

### Good references
wvWare and AbiWord: implement logic to read through fast-saved documents by parsing the CLX and handling piece tables correctly

LibreOffice's MSWord filter in writerfilter also handles these with adjusted logic in its binary parser

The OnlyOffice MsBinaryFile parser uses terminology like PieceTable, ComplexFile, and FIB, but not "FastSave"

### More tips

You can ignore obsolete sections by using:

FibBase.nFib: version number

FibRgLw97.fcClx and lcbClx: offset and size of the piece table

In CLX, decode the PieceTable, which gives you the current "live" text layout

Ignore all earlier edits by trusting only the piece table.


## Important
- Do not hardcode solutions.
- Do not use reflection.
- This parser must work with any .doc file.

## Execution plan for AI (suggestion)

1.  **Understand Existing C# Parser:**
    *   **Analyze FIB Parsing:** Locate the C# code responsible for parsing the File Information Block (FIB). Specifically, examine how `FibBase.nFib`, `FibRgLw97.fcClx`, and `lcbClx` are currently handled.
    *   **CLX and Piece Table Status:** Determine if the Complex File (CLX) structure and Piece Table are currently being parsed and utilized in the C# project. If not, this is a critical area for implementation.
    *   **Cross-Reference with `wvWare` C:** Compare the C# implementation with the original `wvWare` C project's `decode_complex.c` and `clx.c` (and any other relevant files) to understand the established logic for handling fast-saved content.

2.  **Implement/Refine CLX and Piece Table Parsing:**
    *   **Robust CLX Parsing:** If not already present, implement or refine the parsing of the CLX structure to accurately extract its components.
    *   **Accurate Piece Table Interpretation:** Within the CLX parsing, focus on correctly extracting and interpreting the Piece Table. The Piece Table is essential for reconstructing the document's final state by mapping character positions (CPs) to their actual physical locations in the file, accounting for appended changes.
    *   **Discard Obsolete Sections:** Ensure the Piece Table logic correctly identifies and uses the "live" text layout, effectively ignoring or discarding obsolete sections introduced by fast saves.

3.  **Integrate Fast Save Detection and Handling:**
    *   **FIB Flag Utilization:** Use the FIB flags (e.g., `fWhichTblStm`, `fEncrypted`, `fDot`) as potential indicators of a fast-saved document. While not definitive, they can guide the parsing process.
    *   **Prioritize Piece Table:** When a fast-saved document is detected (or when the Piece Table is present and needs to be used), ensure the parser prioritizes the Piece Table's definition of the document structure over a linear read of the file.

4.  **Develop Comprehensive Tests:**
    *   **Dedicated Unit Test:** Create a dedicated unit test for `examples\fastsavedmix.doc` that specifically asserts the correct extraction of "Fastsaved italic text" and "Fastsaved bold text".
    *   **Expand Test Coverage:** If other fast-saved `.doc` files are available in the `examples` directory, use them to broaden test coverage and ensure robustness.
    *   **Regression Prevention:** Verify that the parser correctly handles both fast-saved and non-fast-saved documents without introducing regressions.

5.  **Refactor and Ensure Adherence to Guidelines:**
    *   **C# Conventions:** Adhere strictly to standard C# coding conventions and best practices.
    *   **Code Quality:** Ensure the resulting code is clean, readable, and maintainable.
    *   **Strategic Commenting:** Add comments where necessary to explain complex logic, particularly around the Piece Table and fast save handling, focusing on *why* something is done.
    *   **Avoid Hardcoding/Reflection:** Crucially, ensure the solution avoids hardcoding specific values or using reflection. The parser must be generic and work with any `.doc` file.