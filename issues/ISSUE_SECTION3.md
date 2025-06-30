The parser produces an unexpectedly huge output when parsing `examples/section3.doc`.  
**Expected output:**  
Should match the content of `examples/section3.expected.txt`:

**Actual output:**  
Matches `examples/section3.actual.txt`, which is extremely large (likely due to a wrong or runaway loop in the parser).

## Problem summary

- The output file is much larger than expected.
- The parser should produce output identical to `section3.expected.txt`, but instead produces the content of `section3.actual.txt`.
- This suggests a bug in the main parsing loop, possibly an infinite or incorrect iteration over document structures.

## Additional Info

- Read the markdown files in this project if you need more information.
- Please check `/wvWare` folder as reference to understand how the parse works (C version).
- You can also check the folder `/OnlyOffice-MsBinaryFile` for a different parser written in C++.

## Important
- Do not hardcode solutions.
- Do not use reflection.
- This parser must work with any .doc file.

## Execution plan for AI

1. **Analyze expected vs actual output**
   - Compare `examples/section3.expected.txt` and `examples/section3.actual.txt` to identify:
     - Where the output diverges
     - Patterns in the extra output (repetitions, garbage data, etc.)
     - Whether the parser is stuck in a loop or processing invalid data

2. **Identify the parsing entry point**
   - Locate the main document parsing method in `WvWareNet/Parsers/WordDocumentParser.cs`
   - Set breakpoints at the start of document processing and section handling

3. **Debug the section parsing flow**
   - Instrument the code to log:
     - Section start/end positions
     - Current file pointer position
     - Loop iteration counts
   - Focus on the section processing logic, especially:
     - Section boundary detection
     - Nested section handling
     - Loop exit conditions

4. **Compare with reference implementations**
   - Review the wvWare C implementation (`wvWare/text.c` - `wvGetText()` function)
     - Compare section handling logic
     - Note any loop guards or boundary checks
   - Examine OnlyOffice C++ implementation (`OnlyOffice-MsBinaryFile/MsBinaryFileReader.cpp`)
     - Analyze how it handles complex sections
     - Check for similar loop structures with better safeguards

5. **Add loop termination safeguards**
   - Implement maximum iteration limits for:
     - Section processing loops
     - Text extraction loops
     - Object parsing routines
   - Add position boundary checks to prevent reading beyond:
     - Document boundaries
     - Section boundaries
     - Known valid data ranges

6. **Handle embedded code blocks**
   - Identify special handling for embedded code blocks like:
     - `CalcSromCrc()` function in the hexdump
     - Other binary data sections
   - Implement logic to skip or properly delimit these blocks

7. **Test with problematic document**
   - Run the instrumented parser on `examples/section3.doc`
   - Analyze logs to identify:
     - Where the parser gets stuck
     - Which loop is running excessively
     - Boundary calculation errors
   - Verify output against expected results

8. **Implement fix and validate**
   - Based on findings, modify the loop exit conditions
   - Add additional boundary checks
   - Create test cases for complex section documents
   - Run full test suite to ensure no regressions

**Note:** Focus on making the solution generic - it should handle any .doc file structure, not just section3.doc specifically.
