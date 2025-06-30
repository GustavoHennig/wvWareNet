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

TODO
