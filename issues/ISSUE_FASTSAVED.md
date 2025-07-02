# 🐞 Bug – Fast-Saved text not extracted from `examples/fastsavedmix.doc`

The current parser skips the **incrementally-saved (“Fast Save” / “Quick Save”)** portion at the end of Word 97 documents.
As a result, two lines are missing from the extracted text.


**Actual output**  
```
Fastsaved italic text
Fastsaved bold text
```

**Expected output**  
The complete expected result for the file `fastsavedmix.doc` is:
```
Italic normal text.
Bold normal text.
Big normal text
Fastsaved italic text
Fastsaved bold text
```
---

## 💻 How to reproduce

```bash
dotnet run --project ./WvWareNetConsole/WvWareNetConsole.csproj ./examples/fastsavedmix.doc
```

---

## 🔍 Why this happens (technical summary)

* Word 97’s **Fast Save** mode appends edits to the end of the file instead of rewriting the body.
  The final document state is defined **only** by the *Piece Table* stored inside the **CLX** structure.
* Indicators in the **FIB** let us know the file has been saved incrementally:

  * `FibBase.cQuickSaves` (≤ nFib 0x00D9) or `FibRgCswNewData2000.cQuickSavesNew` (> 0x00D9)
  * Offsets `FibRgLw97.fcClx` / `lcbClx` point to the CLX.
* The C# parser currently reads the text stream linearly and ignores the Piece Table mapping, so it never reaches the appended runs that hold the “Fastsaved …” lines.

---

## 📋 Acceptance criteria

* [ ] Parser detects the presence of a CLX and parses its Piece Table.
* [ ] Character positions (CPs) are mapped through the Piece Table to their correct file offsets, including appended fast-saved ranges.
* [ ] The command above produces the **exact** expected output.
* [ ] No hard-coded offsets or reflection; works for any `.doc` (fast-saved or not).
* [ ] Existing non-fast-saved test files still pass.

---

## 🗂️ Reference material

| Resource                                              | Why it matters                                           |
| ----------------------------------------------------- | -------------------------------------------------------- |
| **\[MS-DOC] §2.5.12 FibRgCswNewData2000**, §2.5.3 CLX | Format details for `cQuickSavesNew`, CLX and Piece Table |
| `wvWare` (C) → `decode_complex.c`, `clx.c`            | Proven logic for merging fast-saved runs                 |
| OnlyOffice `MsBinaryFile` (C++)                       | Alternate parser illustrating CLX handling               |
| LibreOffice Writer filter                             | Another working implementation                           |

---

## 🚧 Suggested implementation steps (for assignee)

1. **Parse CLX → Piece Table** if `lcbClx > 0`.
2. Build CP→FC map; respect Unicode/compressed flags in each piece.
3. Stream text using the Piece Table instead of sequential file reading.
4. Add a unit test for `fastsavedmix.doc`; extend to any other Fast Save samples you have.
5. Refactor / document as needed; follow C# conventions.

---

*Please consult the markdown docs in the repo for a deeper primer on Word 97 internals.*
---


## 🧠 What is "Fast Save" in Word 97?
"Fast Save" was a legacy feature that, when enabled, caused Word to append changes to the end of the .doc file, rather than rewriting the whole file. This improved performance for saving large documents on slow disks. The resulting file contains:

Original text and structure

A chain of edits and additions stored after the main text

A "save history" that can grow over time

So a fast-saved .doc file has an append-only log-like structure at the end, which must be merged to reconstruct the final state.

This is also why tools like wvWare, AbiWord, and LibreOffice need to detect and discard obsolete portions of the document, using markers in the FIB (File Information Block) and Piece Table.


