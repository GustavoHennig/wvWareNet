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

## Execution plan for AI

TODO