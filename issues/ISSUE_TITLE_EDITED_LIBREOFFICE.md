The parser output for `examples/title-edited-libreoffice.doc` is empty, even though all parsers are able to parse it.

**Expected output:**
```
I am a headinU1 title
I am some big heading2  test non-bold
I am normal bolded
I am normal italic
I am big normal
I am normal ariel
```

**Actual output:**
```
(empty)
```

## Problem summary

- The output is completely empty, while the expected text is present in the document.
- All available parsers (wvWare, OnlyOffice, etc.) are able to extract the expected text, but the current parser returns nothing.
- This is not a style or formatting issue; the parser fails to extract any content at all.

## To Reproduce

```
dotnet run --project WvWareNetConsole examples/title-edited-libreoffice.doc
```

## Additional Info

- Read the markdown files in this project if you need more information.
- Please check `/wvWare` folder as reference to understand how the parse works (C version).

## Important
- Do not hardcode solutions.
- Do not use reflection.
- This parser must work with any .doc file.
- Do not customize the parser for this specific file.

## What we found

- The document structure appears valid and is readable by other parsers.
- The file was likely created or edited with LibreOffice, which may introduce subtle differences in the binary structure or encoding.
- All text is present in the file and should be extractable.

## Execution plan for AI

1. **Analyze `examples/title-edited-libreoffice.doc`:**
    * Examine the internal structure of the file to identify any differences from standard Word-generated `.doc` files.
    * Check for encoding, piece table, or section differences that may cause the parser to skip or ignore content.

2. **Investigate `wvWare` C project:**
    * Review how `wvWare` handles files generated or edited by LibreOffice.
    * Pay attention to code paths that deal with piece tables, encoding detection, and text extraction.

4. **Formulate Hypothesis:**
    * Based on the analysis, hypothesize why the parser returns empty output. Consider if the parser is failing to recognize the piece table, misdetecting encoding, or skipping content due to unexpected flags or offsets.

5. **Propose Solution for `WvWareNet`:**
    * Suggest changes to ensure the parser can handle `.doc` files created or edited by LibreOffice.
    * Ensure the solution is general and does not rely on hardcoded values or reflection.

    **Hypothesis:** The most probable cause for the empty output is that the parser fails to recognize or correctly process the piece table or encoding used by LibreOffice. This may be due to differences in how LibreOffice structures the piece table, sets encoding flags, or stores text runs, causing the parser to skip all content.

    **Proposed Solution for `WvWareNet` (C# Port):**

    1. **Robust Piece Table Handling:** Ensure the parser can handle variations in the piece table structure, including those produced by LibreOffice. Validate that all pieces are processed and that their offsets and encodings are interpreted correctly.
    2. **Encoding Detection:** Improve encoding detection logic to account for files that may use different or mixed encodings, especially those set by LibreOffice.
    3. **Graceful Fallbacks:** If the parser encounters unknown or unexpected flags/structures, implement fallbacks to attempt extraction rather than skipping content entirely.
    4. **Cross-Parser Comparison:** Compare the output and parsing logic with wvWare and OnlyOffice to identify and address discrepancies.

6. **Implement and Verify:**
    * Implement the proposed solution in `WvWareNet`.
    * Verify the fix by running the parser on `examples/title-edited-libreoffice.doc` and comparing the output to the expected result.
    * Ensure the solution does not introduce regressions for other `.doc` files.
