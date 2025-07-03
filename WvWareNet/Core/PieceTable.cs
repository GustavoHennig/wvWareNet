using System;
using System.Collections.Generic;
using System.IO;
using WvWareNet.Utilities;

namespace WvWareNet.Core;

public class PieceTable
{
    private readonly ILogger _logger;
    private readonly List<PieceDescriptor> _pieces = new();

    public IReadOnlyList<PieceDescriptor> Pieces => _pieces;

    public PieceTable(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Assigns CHPX (character formatting) data to each piece descriptor.
    /// </summary>
    public void AssignChpxToPieces(List<byte[]> chpxList)
    {
        // Assumes chpxList.Count == _pieces.Count or chpxList.Count == _pieces.Count - 1
        int count = Math.Min(_pieces.Count, chpxList.Count);
        for (int i = 0; i < count; i++)
        {
            _pieces[i].Chpx = chpxList[i];
        }
    }

    public void Parse(byte[] clxData, uint fcMin, uint fcMac, ushort nFib)
    {
        _pieces.Clear();
        if (clxData == null || clxData.Length == 0)
        {
            _logger.LogWarning("Attempted to parse empty piece table data");
            return;
        }

        // DEBUG: Print first 32 bytes of clxData and its length
        Console.WriteLine($"[DEBUG] PieceTable.Parse: data.Length={clxData.Length}");
        Console.WriteLine($"[DEBUG] PieceTable.Parse: first 32 bytes: {BitConverter.ToString(clxData, 0, Math.Min(32, clxData.Length))}");
        Console.WriteLine("[DEBUG] PieceTable.Parse: full CLX data:");
        for (int dbg_i = 0; dbg_i < clxData.Length; dbg_i++)
        {
            Console.Write($"{clxData[dbg_i]:X2} ");
            if ((dbg_i + 1) % 16 == 0) Console.WriteLine();
        }
        Console.WriteLine();

        // --- CLX parsing logic ---
        // CLX can be a sequence of [0x01 Prc] and [0x02 PlcPcd] blocks.
        // We want to find the 0x02 block and parse it as the piece table.
        // However, some documents may have different formats or padding.

        int plcPcdOffset = -1;
        int plcPcdLength = -1;

        int i = 0;
        
        // Skip leading zeros (padding)
        while (i < clxData.Length && clxData[i] == 0)
            i++;

        while (i < clxData.Length)
        {
            byte clxType = clxData[i];
            if (clxType == 0x02)
            {
                // This is the Pcdt block. The next four bytes give the
                // length of the PlcPcd structure that follows.
                if (i + 5 > clxData.Length)
                    break;
                plcPcdLength = BitConverter.ToInt32(clxData, i + 1);
                plcPcdOffset = i + 5;
                _logger.LogInfo($"[DEBUG] Found PlcPcd at offset {i}, length {plcPcdLength}");
                break;
            }
            else if (clxType == 0x01)
            {
                // This is a Prc (property modifier) block, skip it
                if (i + 3 > clxData.Length)
                    break;
                ushort prcSize = BitConverter.ToUInt16(clxData, i + 1);
                _logger.LogInfo($"[DEBUG] Skipping Prc block at offset {i}, size {prcSize}");
                i += 1 + 2 + prcSize;
            }
            else
            {
                // For some documents, the piece table may start directly without clxt prefix
                // Try to parse the remainder as a piece table
                _logger.LogInfo($"[DEBUG] No clxt prefix found, trying direct piece table parse from offset {i}");
                plcPcdOffset = i;
                plcPcdLength = clxData.Length - i;
                break;
            }
        }

        if (plcPcdOffset == -1)
        {
            _logger.LogWarning("No PlcPcd (piece table) found in CLX data. Falling back to single piece.");
            SetSinglePiece(fcMin, fcMac, nFib);
            return;
        }

        // Now parse the PlcPcd as the piece table
        try
        {
            using var stream = new MemoryStream(clxData, plcPcdOffset, plcPcdLength);
            using var reader = new BinaryReader(stream);

            _logger.LogInfo($"[DEBUG] Attempting to parse piece table at offset {plcPcdOffset}, length {plcPcdLength}");
            
            // For debugging, log the first 16 bytes of what we're trying to parse
            if (plcPcdLength >= 16)
            {
                byte[] debugBytes = new byte[16];
                Array.Copy(clxData, plcPcdOffset, debugBytes, 0, 16);
                _logger.LogInfo($"[DEBUG] First 16 bytes of piece table data: {BitConverter.ToString(debugBytes)}");
            }

            // The piece table (PlcPcd) consists of an array of character
            // positions followed by an array of piece descriptors.  
            // Each piece: 4 bytes CP, 8 bytes descriptor (total 12 bytes per piece)
            // However, the structure might be different for some documents
            
            // Try to detect if this looks like valid piece table data
            // Valid CP values should be reasonable (0 to document length)
            // Let's try different interpretations
            
            bool validPieceTable = false;
            int pieceCount = 0;
            
            // Try standard interpretation first
            if (plcPcdLength >= 16) // Minimum for 1 piece (4+4+8)
            {
                int testPieceCount = (plcPcdLength - 4) / 12;
                if (testPieceCount > 0 && testPieceCount < 20) // Reasonable range
                {
                    // Test if the first few CP values look reasonable
                    stream.Position = 0;
                    int cp1 = reader.ReadInt32();
                    int cp2 = reader.ReadInt32();
                    
                    _logger.LogInfo($"[DEBUG] Test CP values: cp1={cp1}, cp2={cp2}");
                    
                    if (cp1 >= 0 && cp1 < 100000 && cp2 >= cp1 && cp2 < 100000)
                    {
                        validPieceTable = true;
                        pieceCount = testPieceCount;
                        _logger.LogInfo($"[DEBUG] Standard piece table format detected, {pieceCount} pieces");
                    }
                }
            }
            
            if (!validPieceTable)
            {
                _logger.LogWarning($"[DEBUG] Piece table data doesn't match expected format, falling back to single piece");
                SetSinglePiece(fcMin, fcMac, nFib);
                return;
            }

            // Parse the piece table using standard format
            stream.Position = 0;
            var cpArray = new int[pieceCount + 1];
            for (int j = 0; j < pieceCount + 1; j++)
            {
                cpArray[j] = reader.ReadInt32();
                _logger.LogInfo($"[DEBUG] CP[{j}] = {cpArray[j]}");
            }

            for (int j = 0; j < pieceCount; j++)
            {
                uint fcValue = reader.ReadUInt32();
                uint prm = reader.ReadUInt32(); // PRM (property modifier)
                
                _logger.LogInfo($"[DEBUG] Piece {j}: fcValue=0x{fcValue:X8}, prm=0x{prm:X8}");

                // According to MS-DOC spec, the FC value interpretation:
                // - If bit 30 (0x40000000) is SET: 8-bit characters, fc = (fcValue & ~0x40000000) / 2
                // - If bit 30 is CLEAR: 16-bit characters, fc = fcValue & ~0x40000000
                
                bool isCompressed = (fcValue & 0x40000000) != 0;
                bool isUnicode = !isCompressed;
                uint fc;
                
                if (isCompressed)
                {
                    // 8-bit characters: clear the compression bit and divide by 2
                    fc = (fcValue & 0xBFFFFFFF) / 2;
                }
                else
                {
                    // 16-bit characters: just clear any high bits
                    fc = fcValue & 0x3FFFFFFF;
                }

                _logger.LogInfo($"fcValue=0x{fcValue:X8}, isCompressed={isCompressed}, isUnicode={isUnicode}, fc={fc}");

                int cpStart = cpArray[j];
                int cpEnd = cpArray[j + 1];
                int fcStart = (int)fc;
                int fcEnd = fcStart + (isUnicode ? (cpEnd - cpStart) * 2 : (cpEnd - cpStart));

                _logger.LogInfo($"[DEBUG] Piece {j}: CP {cpStart}-{cpEnd}, FC {fcStart}-{fcEnd}, Unicode={isUnicode}");

                var descriptor = new PieceDescriptor
                {
                    FilePosition = fc,
                    IsUnicode = isUnicode,
                    HasFormatting = false,
                    CpStart = cpStart,
                    CpEnd = cpEnd,
                    FcStart = fcStart,
                    FcEnd = fcEnd
                };

                _pieces.Add(descriptor);
            }

            _logger.LogInfo($"Parsed {_pieces.Count} piece descriptors from piece table");
        }
        catch (Exception ex)
        {
            _logger.LogError("Error parsing piece table data", ex);
            _logger.LogWarning("Falling back to single piece due to parsing error");
            SetSinglePiece(fcMin, fcMac, nFib);
        }
    }

    public string GetTextForPiece(int index, Stream documentStream, int? codePage = null)
    {
        if (index < 0 || index >= _pieces.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        var piece = _pieces[index];
        return GetTextForRange(piece.FcStart, piece.FcEnd, documentStream, piece.IsUnicode, codePage);
    }

    /// <summary>
    /// Retrieve text for an arbitrary file position range. The range may span
    /// multiple pieces and does not need to align to piece boundaries.
    /// </summary>
    public string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, int? codePage = null)
    {
        if (fcEnd <= fcStart)
            return string.Empty;

        var sb = new System.Text.StringBuilder();

        foreach (var piece in _pieces)
        {
            int start = Math.Max(fcStart, piece.FcStart);
            int end = Math.Min(fcEnd, piece.FcEnd);
            if (start >= end)
                continue;

            sb.Append(GetTextForRange(start, end, documentStream, piece.IsUnicode, codePage));
        }

        return sb.ToString();
    }

    private string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, bool isUnicode, int? codePage = null)
    {
        _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart={fcStart}, fcEnd={fcEnd}, isUnicode={isUnicode}, streamLength={documentStream.Length}, codePage={codePage}");
        int length = fcEnd - fcStart;
        if (length <= 0)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Invalid length {length}, returning empty string");
            return string.Empty;
        }

        if (fcStart >= documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart {fcStart} >= streamLength {documentStream.Length}, returning empty string");
            return string.Empty;
        }

        // Clamp to stream bounds
        if (fcEnd > documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Clamping fcEnd from {fcEnd} to {documentStream.Length}");
            fcEnd = (int)documentStream.Length;
            length = fcEnd - fcStart;
        }

        using var reader = new BinaryReader(documentStream, System.Text.Encoding.UTF8, leaveOpen: true);
        documentStream.Seek(fcStart, SeekOrigin.Begin);
        byte[] bytes = reader.ReadBytes(length);
        
        _logger.LogInfo($"[DEBUG] GetTextForRange: Read {bytes.Length} bytes from offset {fcStart}");
        if (bytes.Length > 0)
        {
            string hexPreview = BitConverter.ToString(bytes, 0, Math.Min(32, bytes.Length));
            _logger.LogInfo($"[DEBUG] GetTextForRange: First bytes: {hexPreview}");
        }

        string text;
        if (isUnicode)
        {
            text = System.Text.Encoding.Unicode.GetString(bytes);
            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as UTF-16LE");
        }
        else
        {
            // Use code page if provided, otherwise default to Windows-1252
            var encoding = codePage.HasValue ? System.Text.Encoding.GetEncoding(codePage.Value) : System.Text.Encoding.GetEncoding(1252);
            text = encoding.GetString(bytes);
            _logger.LogInfo($"[DEBUG] GetTextForRange: Decoded as {(codePage.HasValue ? $"code page {codePage.Value}" : "Windows-1252")}");
        }

        string processedText = ProcessFieldCodes(text);
        _logger.LogInfo($"[DEBUG] GetTextForRange raw text length: {text.Length}, processed length: {processedText.Length}");
        _logger.LogInfo($"[DEBUG] GetTextForRange returning: '{processedText.Substring(0, Math.Min(50, processedText.Length))}'");
        return CleanText(processedText);
    }

    /// <summary>
    /// Process Word field codes and extract only the field results.
    /// Field structure: [0x13][field code][0x14][field result][0x15]
    /// We want to keep only the field result part.
    /// For HYPERLINK fields without field result, extract the URL from the field code.
    /// </summary>
    private static string ProcessFieldCodes(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        int i = 0;
        
        while (i < input.Length)
        {
            char c = input[i];
            
            if (c == (char)0x13) // Field begin (0x13)
            {
                // Find the field separator (0x14) and field end (0x15)
                int separatorPos = -1;
                int endPos = -1;
                int depth = 1; // Track nested fields
                
                for (int j = i + 1; j < input.Length; j++)
                {
                    if (input[j] == (char)0x13) // Nested field begin
                    {
                        depth++;
                    }
                    else if (input[j] == (char)0x15) // Field end
                    {
                        depth--;
                        if (depth == 0)
                        {
                            endPos = j;
                            break;
                        }
                    }
                    else if (input[j] == (char)0x14 && depth == 1 && separatorPos == -1) // Field separator at our level
                    {
                        separatorPos = j;
                    }
                }
                
                if (endPos != -1)
                {
                    if (separatorPos != -1)
                    {
                        // Extract field result (between separator and end)
                        string fieldResult = input.Substring(separatorPos + 1, endPos - separatorPos - 1);
                        sb.Append(fieldResult);
                    }
                    else
                    {
                        // No field separator found, extract field code and check if it's a HYPERLINK
                        string fieldCode = input.Substring(i + 1, endPos - i - 1);
                        
                        // Handle HYPERLINK fields
                        if (fieldCode.StartsWith(" HYPERLINK ", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract URL from HYPERLINK field
                            string url = fieldCode.Substring(11).Trim(); // Remove " HYPERLINK "
                            
                            // Remove any additional parameters or quotes
                            int spaceIndex = url.IndexOf(' ');
                            if (spaceIndex > 0)
                            {
                                url = url.Substring(0, spaceIndex);
                            }
                            
                            // Remove quotes if present
                            url = url.Trim('"');
                            
                            sb.Append(url);
                        }
                        else
                        {
                            // For other field types without separator, skip the field
                        }
                    }
                    // Skip to after the field end
                    i = endPos + 1;
                }
                else
                {
                    // Malformed field, just append the character and continue
                    sb.Append(c);
                    i++;
                }
            }
            else if (c == (char)0x14 || c == (char)0x15) // Standalone field separators/ends (shouldn't happen in well-formed text)
            {
                // Skip these control characters
                i++;
            }
            else
            {
                // Regular character, append it
                sb.Append(c);
                i++;
            }
        }
        
        return sb.ToString();
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            // Handle special Word control characters
            if (c == (char)0x07) // Word tab character (0x07) -> convert to standard tab
            {
                sb.Append('\t');
            }
            else if (c == (char)0x0B) // Word line break (0x0B) -> convert to newline
            {
                sb.Append('\n');
            }
            // Allow a much wider range of characters - be more permissive
            else if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
                char.IsSymbol(c) ||  
                // TODO: This seems redundany, review it
                c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '~' ||
                (c >= 'А' && c <= 'я') || // Cyrillic range
                (c >= 'À' && c <= 'ÿ') || // Latin-1 Supplement
                c == '=' || c == '<' || c == '>' || c == '^' || c == '|' || c == '+' || c == '-') // Explicitly allow missing chars
            {
                sb.Append(c);
            }
            else if (c == '\0' || c < ' ') // Replace null chars and control chars
            {
                // Skip these characters
            }
            else
            {
                // For debugging: log any character that's being filtered out
                Console.WriteLine($"[DEBUG] CleanText filtering out character: '{c}' (U+{(int)c:X4})");
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Attempt to determine if the pieces in a Word95 document contain
    /// 16-bit text. This mirrors the heuristic used by wvGuess16bit in the
    /// original wvWare project.
    /// </summary>
    /// <param name="fcValues">File positions for each piece.</param>
    /// <param name="cpArray">Character position array from the piece table.</param>
    private static bool GuessPiecesAre16Bit(uint[] fcValues, int[] cpArray)
    {
        var tuples = new List<(uint Fc, uint Offset)>();
        for (int i = 0; i < fcValues.Length; i++)
        {
            uint offset = (uint)(cpArray[i + 1] - cpArray[i]) * 2u;
            tuples.Add((fcValues[i], offset));
        }

        tuples.Sort((a, b) => a.Fc.CompareTo(b.Fc));

        for (int i = 0; i < tuples.Count - 1; i++)
        {
            if (tuples[i].Fc + tuples[i].Offset > tuples[i + 1].Fc)
                return false; // Overlap means 8-bit text
        }

        return true; // No overlap detected -> assume 16-bit
    }

    /// <summary>
    /// Replace the current table with a single Unicode piece using the
    /// supplied file positions. Used as a fallback when the piece table is
    /// corrupt or not present.
    /// </summary>
    public void SetSinglePiece(uint fcMin, uint fcMac, ushort nFib)
    {
        _pieces.Clear();
        // Always use 8-bit encoding for fallback, as in the working version.
        bool isUnicode = false;
        _pieces.Add(new PieceDescriptor
        {
            FilePosition = fcMin,
            IsUnicode = isUnicode,
            HasFormatting = false,
            CpStart = 0,
            CpEnd = (int)(fcMac - fcMin),
            FcStart = (int)fcMin,
            FcEnd = (int)fcMac
        });
    }
}
