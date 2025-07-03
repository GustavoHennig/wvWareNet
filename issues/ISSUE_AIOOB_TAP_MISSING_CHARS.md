# 🐞 Bug – Missing Characters and Line Breaks in `examples-local/AIOOB-Tap.doc` Output

When parsing `examples-local/AIOOB-Tap.doc`, the parser loses some characters and line breaks compared to the expected output in `examples-local/AIOOB-Tap.expected`.

---

## 💻 How to reproduce

```bash
dotnet run --project ./WvWareNetConsole/WvWareNetConsole.csproj ./examples-local/AIOOB-Tap.doc
```

Compare the output with `examples-local/AIOOB-Tap.expected`.

---

## **Problem summary**

- Some characters and line breaks are missing in the parsed output.
- The output does not match the expected result.
- This may indicate an issue with text extraction, encoding, or line break handling in the parser.

---

## 📋 Acceptance criteria

* [ ] Parser extracts all characters and line breaks present in the document.
* [ ] The command above produces the **exact** expected output.
* [ ] No hard-coded solutions; works for any `.doc` file.
* [ ] Existing test files still pass.

---

## 🚧 Suggested investigation steps

1. Compare actual vs expected output to identify missing characters and line breaks.
2. Review text extraction and line break handling logic in the parser.
3. Check for encoding or piece table issues that could cause loss of content.
4. Add or update unit tests to verify correct extraction for this and similar files.

---
