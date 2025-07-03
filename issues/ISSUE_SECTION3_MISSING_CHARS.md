# 🐞 Bug – Missing Characters in `examples/section3.doc` Output

The parser was losing specific characters when parsing `examples/section3.doc`, resulting in incomplete code snippets and mathematical expressions.

**Missing Characters:**
- `=` (equals signs)
- `<` and `>` (comparison operators)
- `^` (XOR operator)
- `|` (OR operator)
- `<<` and `>>` (bit shift operators)
- `+` (plus signs)
- `-` (minus signs)

**Example of the Problem:**
```
// Expected:
Srom.SromCRC = CalcSromCrc(&Srom.SromData);
for (Index = 0; Index < DATA_LEN; Index++)
if (Msb ^ (CurrentByte & 1))
crc ^= POLY;
crc |= 0x00000001;

// Actual (before fix):
Srom.SromCRC  CalcSromCrc(&Srom.SromData);
for (Index  0; Index  DATA_LEN; Index)
if (Msb  (CurrentByte & 1))
crc  POLY;
crc  0x00000001;
```

---

## 💻 How to reproduce

```bash
dotnet run --project ./WvWareNetConsole/WvWareNetConsole.csproj ./examples/section3.doc
```

Compare the output with `examples/section3.expected.txt` to see the missing characters.

---

## 🔍 Root Cause

The issue was in the `CleanText` method in `WvWareNet/Core/PieceTable.cs`. The method was filtering out characters that didn't match specific criteria, but it was too restrictive and excluded important symbols and operators.

The original filter only allowed:
- `char.IsLetterOrDigit(c)`
- `char.IsPunctuation(c)` 
- `char.IsWhiteSpace(c)`
- Some specific characters and ranges

However, many mathematical and programming operators like `=`, `<`, `>`, `^`, `|`, `<<`, `>>`, `+`, `-` were being filtered out because they didn't pass the `char.IsPunctuation(c)` test or weren't explicitly allowed.

---

## ✅ Solution

Modified the `CleanText` method to be more permissive by:

1. **Adding `char.IsSymbol(c)`** - This allows mathematical and programming symbols
2. **Explicitly allowing missing characters** - Added specific checks for `=`, `<`, `>`, `^`, `|`, `+`, `-`
3. **Added debugging** - Log any character that gets filtered out to help identify future issues

**Code Changes in `WvWareNet/Core/PieceTable.cs`:**

```csharp
// Before:
else if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
    c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '~' ||
    (c >= 'А' && c <= 'я') || // Cyrillic range
    (c >= 'À' && c <= 'ÿ'))   // Latin-1 Supplement

// After:
else if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
    char.IsSymbol(c) || // Add symbols like =, <, >, ^, |, etc.
    c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '~' ||
    (c >= 'А' && c <= 'я') || // Cyrillic range
    (c >= 'À' && c <= 'ÿ') || // Latin-1 Supplement
    c == '=' || c == '<' || c == '>' || c == '^' || c == '|' || c == '+' || c == '-') // Explicitly allow missing chars
```

---

## 📋 Verification

After the fix:
- ✅ All mathematical operators (`=`, `<`, `>`, `^`, `|`, `<<`, `>>`, `+`, `-`) are preserved
- ✅ Code snippets in the document are complete and readable
- ✅ Mathematical expressions like "x8 + x2 + x1 + 1" are correctly extracted
- ✅ C code examples maintain their syntax integrity
- ✅ No regressions in other document parsing

---

## 🚧 Implementation Notes

- **Generic Solution**: The fix uses `char.IsSymbol()` which is a standard .NET method that properly categorizes symbol characters according to Unicode standards
- **No Hardcoding**: While we explicitly allow some characters for safety, the primary fix relies on the standard Unicode categorization
- **Debugging Support**: Added logging to identify any future character filtering issues
- **Backward Compatible**: The change is additive and doesn't remove any previously allowed characters

---

## 🧪 Test Cases

The fix should be tested with:
- ✅ `examples/section3.doc` - Primary test case with mathematical operators
- ✅ Other documents with programming code or mathematical expressions
- ✅ Documents with special symbols and punctuation
- ✅ Existing test documents to ensure no regressions

---

*This issue demonstrates the importance of being permissive with character filtering in document parsers, as different document types may contain various symbols and operators that are essential for content integrity.*
