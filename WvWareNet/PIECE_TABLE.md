In Microsoft Word documents, a PieceTable is a crucial internal data structure that manages the document's text content. It doesn't store
  the text directly in a single contiguous block. Instead, it stores "pieces" of text, which are references to locations within the original
  WordDocument stream.

  Here's a breakdown:


   * What is a PieceTable?
       * Imagine a document where you frequently insert, delete, or reformat text. If the entire document text were stored contiguously, every
         change would require shifting large amounts of data, which is inefficient.
       * The PieceTable solves this by acting as a "map" or "index" to the document's text. It's a list of PieceDescriptor objects.
       * Each PieceDescriptor points to a specific segment (a "piece") of text within the raw WordDocument stream. It also stores information
         about whether that piece is Unicode (16-bit characters) or ANSI (8-bit characters).
       * When you edit a document, Word doesn't rewrite the entire text. Instead, it creates new pieces for the changed content and updates the
         PieceTable to reflect these new pieces and their positions. Old, unused pieces might remain in the stream but are no longer referenced
         by the PieceTable.


   * Relation to Paragraphs:
       * Paragraphs are logical blocks of text within a document. They have associated formatting (like alignment, indentation, line spacing).
       * The PieceTable provides the raw text content, but it doesn't inherently define paragraph boundaries or their formatting.
       * In the WordDocumentParser.cs file, you'll see that after the PieceTable is parsed, there's a section for "Paragraph boundary
         detection using PLCF for paragraphs (PAPX)".
       * PLCF (Piecewise Linear Control File) for PAPX (Paragraph Property Exceptions) is another structure that defines the character
         positions (CPs) where paragraph formatting changes or new paragraphs begin.
       * The WordDocumentParser uses the PieceTable to retrieve the actual text content for a given range of CPs, and then it uses the PAPX
         information to segment that text into Paragraph objects and apply their respective formatting.


  In essence:
   * The PieceTable tells you where the text is located in the raw document stream and how it's encoded (Unicode/ANSI).
   * The PAPX (and other formatting structures) tell you how that text should be structured into paragraphs and how it should be formatted.

     The WordDocumentParser combines information from the PieceTable (for text content) and the PAPX (for paragraph structure and formatting)
  to build the DocumentModel with its Sections and Paragraphs.



  A PieceTable is unique to each Word document. It's essentially the document's internal index or map to its text content. Every .doc
  file will have its own PieceTable (or equivalent structure in older formats) that describes how its specific text is laid out.


  Regarding reuse of entries:
   * Entries (PieceDescriptors) themselves are not "reused" in the sense that a single `PieceDescriptor` object would point to different parts
     of the document at different times. Once a PieceDescriptor is created, it points to a fixed range of text within the WordDocument stream.
   * However, the concept of a piece is that it's a segment of text. When you make edits, Word creates new PieceDescriptor entries for the new
     or modified text. It doesn't alter existing PieceDescriptor entries. Instead, it updates the PieceTable (the list of PieceDescriptors) to
     reflect the new arrangement of text. Old, unused pieces might remain in the underlying stream but are no longer referenced by the active
     PieceTable. This is what makes Word's editing model efficient for large documents.


  A "Piece" structure, as implemented in WvWareNet, refers to the PieceDescriptor class. Based on the PieceDescriptor.cs file, its key
  properties are:


   * `FilePosition`: This is the starting byte offset within the WordDocument stream where this piece of text begins.
   * `IsUnicode`: A boolean flag indicating whether the text in this piece is encoded as Unicode (16-bit characters) or ANSI (8-bit
     characters). This is crucial for correct decoding.
   * `HasFormatting`: A boolean flag indicating if this piece has associated character formatting (CHPX).
   * `CpStart`: The starting character position (CP) of this piece within the overall logical document text.
   * `CpEnd`: The ending character position (CP) of this piece within the overall logical document text.
   * `FcStart`: The starting file character position (FC) of this piece within the WordDocument stream. This is essentially the same as
     FilePosition but often used in the context of character offsets.
   * `FcEnd`: The ending file character position (FC) of this piece within the WordDocument stream.


  So, a PieceDescriptor is a record that tells the parser: "From character position CpStart to CpEnd in the document, the text is located in
  the WordDocument stream from byte offset FcStart to FcEnd, and it's encoded as IsUnicode."



  1. `1Table` or `0Table` (the `tableEntry`): These are indeed mandatory streams within the Compound File Binary Format (CFBF) file structure
      for Word documents. They contain crucial information about the document's structure, including the piece table. Word 97 and later
      typically use 1Table, while older versions might use 0Table. If neither is found, it's usually an indication of a corrupted or
      unsupported document.


   2. `tableStream`: Yes, the content of the identified 1Table or 0Table stream is read entirely into a byte[] array, which you correctly refer
      to as tableStream.


   3. `clxData`: This is where the nuance comes in. clxData is not the entire `tableStream`. Instead, clxData is a specific portion of the
      tableStream that contains the Complex File Layout (CLX) structure.
       * For Word 97 and later documents, the File Information Block (FIB) contains fields (fib.FcClx and fib.LcbClx) that explicitly tell the
         parser the byte offset (FcClx) and length (LcbClx) of the CLX data within the tableStream.
       * For older Word 6/95 documents, the CLX data is typically found at the beginning of the tableStream, and its length might be determined
         heuristically.
       * The CLX structure itself can contain various blocks, but the one most relevant to text layout is the `PlcPcd` (Piece Table) block.


   4. `pieceTable` (the object) and `PieceDescriptor`s: The clxData (specifically, the PlcPcd within it) is then parsed by the PieceTable class
      (your WvWareNet.Core.PieceTable object). This parsing populates the _pieces list within the PieceTable object with PieceDescriptor
      instances. Each PieceDescriptor represents a "piece" of the document's text, providing its starting and ending character positions (CPs)
      and its corresponding file offsets (FCs) within the main WordDocument stream, along with its encoding (Unicode or ANSI).

  In summary:


  The 1Table or 0Table stream (tableStream) is a container. Within this container, there's a specific sub-structure called CLX (clxData),
  and within CLX, there's the PlcPcd structure, which is the actual raw data that defines the individual text "pieces" (PieceDescriptors)
  that make up the document's content. The PieceTable object then interprets this PlcPcd data to build its internal map of the document's
  text.