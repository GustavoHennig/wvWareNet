# 🐞 Bug – Incomplete text extracted from `examples/other.doc`

The current parser does not extract all lines from the document. As a result, the output is missing content.

**Actual output**  
```
Bold simple line
ital
```

**Expected output**  
The complete expected result for the file `other.doc` is:
```
Big line
Bold simple line
Italic line
```
---

## 💻 How to reproduce

```bash
dotnet run --project ./WvWareNetConsole/WvWareNetConsole.csproj ./examples/other.doc
```

---

## 📋 Acceptance criteria

* [ ] Parser extracts all lines present in the document.
* [ ] The command above produces the **exact** expected output.
* [ ] No hard-coded offsets or reflection; works for any `.doc`.
* [ ] Existing test files still pass.

---

## 🚧 Suggested implementation steps (for assignee)

1. Investigate why some lines are missing from the output.
2. Review text extraction logic for edge cases in `other.doc`.
3. Add a unit test for `other.doc` to verify correct extraction.
4. Refactor / document as needed; follow C# conventions.

---
