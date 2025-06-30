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

TODO
