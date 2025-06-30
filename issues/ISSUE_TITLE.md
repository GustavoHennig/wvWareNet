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

## To Reproduce

```
dotnet run --project WvWareNetConsole examples/title.doc
```

## Additional Info

- Read the markdown files in this project if you need more information.
- Please check `/wvWare` folder as reference to understand how the parse works (C version).
- You can also check the folder `/OnlyOffice-MsBinaryFile` for a different parser written in C++.

## Important
- Do not hardcode solutions.
- Do not use reflection.
- This parser must work with any .doc file.

## What we found 

If you see, the missing words are located at the position 0x5222 to 0x524C in the file `examples/title.doc`. The parser is not extracting these words correctly, which is why they are missing from the output.
You need to find out how to place these words in the right place.


```
/* P:\Dev\Experiments\vwWare-toDotNet\examples\title.doc (30/06/2025 17:24:49)
   StartOffset(h): 00005222, EndOffset(h): 0000524C, Length(h): 0000002B */

byte[] rawData = {
	0x68, 0x00, 0x65, 0x00, 0x61, 0x00, 0x64, 0x00, 0x69, 0x00, 0x6E, 0x00,
	0x67, 0x00, 0x31, 0x00, 0x20, 0x00, 0x68, 0x00, 0x65, 0x00, 0x61, 0x00,
	0x64, 0x00, 0x69, 0x00, 0x6E, 0x00, 0x67, 0x00, 0x32, 0x00, 0x20, 0x00,
	0x20, 0x00, 0x62, 0x00, 0x69, 0x00, 0x67
};

```


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

    **Hypothesis:** The most probable cause for the missing words ("heading1", "heading2", "big") in `examples/title.doc` is that these words are stored as 16-bit Unicode characters, but `wvGuess16bit` incorrectly determines the document's overall encoding to be 8-bit. This leads to `wvNormFC` miscalculating file offsets for 16-bit characters, causing them to be skipped or misinterpreted.

    **Proposed Solution for `WvWareNet` (C# Port):**

    The C# port should implement a more robust character encoding detection mechanism that goes beyond a simple global guess. Here are a few approaches:

    1.  **Prioritize `PCD.fc` flags:** The `wvNormFC` function in the C code explicitly checks `fc & 0x40000000UL` to determine if a character is 8-bit or 16-bit. This flag within the `fc` itself is a direct indicator. The C# port should ensure this flag is correctly interpreted and used to determine the character width for each `fc` value, overriding any global guesses if present. This is the most reliable method as it's embedded in the file format.

    2.  **Per-Piece Encoding Determination:**
        *   Instead of a single global guess, attempt to determine the encoding (8-bit or 16-bit) for each individual piece (`PCD` entry) in the `CLX`.
        *   This might involve analyzing the content of each piece for common Unicode byte order marks (BOMs) or patterns indicative of 16-bit characters, or by relying on specific flags within the `PCD` or related structures if the Word file format provides such information.
        *   The `wvNormFC` function in C# would then need to take the piece's determined encoding into account when calculating file offsets.

    3.  **Heuristic Refinement:**
        *   If a per-piece determination is not feasible or reliable, refine the guessing heuristic. Instead of immediately defaulting to 8-bit on the first overlap, consider:
            *   **Majority Rule:** Count the number of pieces that would cause an overlap if assumed 16-bit vs. those that wouldn't. Choose the encoding that results in fewer "overlaps" or inconsistencies.
            *   **Fallback Mechanism:** If the initial 16-bit guess leads to significant parsing errors (e.g., unreadable characters, unexpected document structure), automatically re-attempt parsing with an 8-bit assumption.

    4.  **Leverage `fExtChar` in `FIB`:**
        *   The `FIB` (File Information Block) has a `fExtChar` field, which indicates if the document contains extended characters (potentially 16-bit). While `wvWare` uses this, ensure its interpretation in `WvWareNet` is accurate and consistently applied. This flag could be a strong hint for the initial encoding assumption.

6.  **Implement and Verify:**
    *   Implement the proposed solution in `WvWareNet`.
    *   Verify the fix by running the parser on `examples/title.doc` and comparing the output to the expected result.
    *   Ensure the solution does not introduce regressions for other `.doc` files.
