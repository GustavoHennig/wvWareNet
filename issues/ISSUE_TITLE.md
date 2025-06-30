The parser output for `examples/title.doc` is not as expected.  
**Expected output:**  
```
I am a heading1 title
I am some big heading2  test non-bold
I am normal bolded
I am normal italic
I am big normal
I am normal ariel
```

**Actual output:**  
```
I am a title
I am some big test non-bold
I am normal bolded
I am normal italic
I am normal
I am normal ariel
```

## Problem summary

- The output does not match the expected result.
- The parser should produce output identical to `title.expected.txt`, but instead produces the content of `title.actual.txt`.
- Very important: This is NOT related to headings or style. It is only a coincidence that the missing words are "heading1", "heading2", and "big". The problem is related to the source of these words—they are probably not contiguous in the document. This could come from a dictionary or be inserted with a different algorithm. It is imperative that you investigate how other parsers handle this, such as OnlyOffice and wvWare.

## Additional Info

- Read the markdown files in this project if you need more information.
- Please check `/wvWare` folder as reference to understand how the parse works (C version).
- You can also check the folder `/OnlyOffice-MsBinaryFile` for a different parser written in C++.

## Important
- Do not hardcode solutions.
- Do not use reflection.
- This parser must work with any .doc file.

## Execution plan for AI

1.  **Analyze `examples/title.doc`:**
    *   Attempt to understand the internal structure of `title.doc` to identify how the missing words ("heading1", "heading2", "big") are stored. This might involve looking for differences in how contiguous text is stored versus these specific words.
    *   Consider if these words are part of a specific document feature (e.g., fields, revisions, annotations) that is not being correctly parsed.

2.  **Investigate `wvWare` C project:**
    *   Examine the `wvWare` C source code (specifically `decode_complex.c`, `decode_simple.c`, `text.c`, `chp.c`, `pap.c`, `sep.c`, and any files related to fast saves or revisions) to understand how it extracts and reconstructs text from `.doc` files.
    *   Pay close attention to how `wvWare` handles text runs, character properties, paragraph properties, and section properties, as well as any mechanisms for handling non-contiguous text or document revisions.

3.  **Investigate `OnlyOffice-MsBinaryFile` C++ project:**
    *   Review the `OnlyOffice-MsBinaryFile` source code to gain insights into their approach to parsing `.doc` files and text extraction. This can provide an alternative perspective on handling complex document structures.

4.  **Formulate Hypothesis:**
    *   Based on the analysis of `title.doc` and the `wvWare`/`OnlyOffice` codebases, develop a hypothesis explaining why the words are missing in the current `WvWareNet` parser output. This hypothesis should address the underlying cause of the non-contiguous text or special insertion method.

5.  **Propose Solution for `WvWareNet`:**
    *   Outline a concrete solution for the `WvWareNet` project to correctly parse and include the missing words. This solution should align with the C# porting guidelines (no hardcoding, no reflection, general applicability).

6.  **Implement and Verify:**
    *   Implement the proposed solution in `WvWareNet`.
    *   Verify the fix by running the parser on `examples/title.doc` and comparing the output to the expected result.
    *   Ensure the solution does not introduce regressions for other `.doc` files.
