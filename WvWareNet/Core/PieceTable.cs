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

    public void Parse(byte[] data, uint fcMin, uint fcMac, ushort nFib)
    {
        _pieces.Clear();
        if (data == null || data.Length == 0)
        {
            _logger.LogWarning("Attempted to parse empty piece table data");
            return;
        }

        // DEBUG: Print first 32 bytes of clxData and its length
        Console.WriteLine($"[DEBUG] PieceTable.Parse: data.Length={data.Length}");
        Console.WriteLine($"[DEBUG] PieceTable.Parse: first 32 bytes: {BitConverter.ToString(data, 0, Math.Min(32, data.Length))}");
        Console.WriteLine("[DEBUG] PieceTable.Parse: full CLX data:");
        for (int dbg_i = 0; dbg_i < data.Length; dbg_i++)
        {
            Console.Write($"{data[dbg_i]:X2} ");
            if ((dbg_i + 1) % 16 == 0) Console.WriteLine();
        }
        Console.WriteLine();

        // --- CLX parsing logic ---
        // CLX can be a sequence of [0x01 Prc] and [0x02 PlcPcd] blocks.
        // We want to find the 0x02 block and parse it as the piece table.

        int plcPcdOffset = -1;
        int plcPcdLength = -1;

        int i = 0;
        while (i < data.Length)
        {
            byte clxType = data[i];
            if (clxType == 0x02)
            {
                // This is the Pcdt block. The next four bytes give the
                // length of the PlcPcd structure that follows.
                if (i + 5 > data.Length)
                    break;
                plcPcdLength = BitConverter.ToInt32(data, i + 1);
                plcPcdOffset = i + 5;
                break;
            }
            else if (clxType == 0x01)
            {
                // This is a Prc (property modifier) block, skip it
                if (i + 3 > data.Length)
                    break;
                ushort prcSize = BitConverter.ToUInt16(data, i + 1);
                i += 1 + 2 + prcSize;
            }
            else
            {
                // Unknown block, abort
                break;
            }
        }

        if (plcPcdOffset == -1)
        {
            _logger.LogWarning("No PlcPcd (piece table) found in CLX data. Falling back to single piece.");
            SetSinglePiece(fcMin, fcMac, nFib); // Use the new overload
            return;
        }

        // Now parse the PlcPcd as the piece table
        try
        {
            using var stream = new MemoryStream(data, plcPcdOffset, plcPcdLength);
            using var reader = new BinaryReader(stream);

            // The piece table (PlcPcd) consists of an array of character
            // positions followed by an array of piece descriptors.  
            // Each piece: 4 bytes CP, 8 bytes descriptor (total 12 bytes per piece)
            int pieceCount = (plcPcdLength - 4) / 12;
            if (pieceCount <= 0)
            {
                _logger.LogWarning($"Invalid piece table length {plcPcdLength}. Using fallback single-piece table.");
                SetSinglePiece(fcMin, fcMac, nFib);
                return;
            }

            var cpArray = new int[pieceCount + 1];
            for (int j = 0; j < pieceCount + 1; j++)
                cpArray[j] = reader.ReadInt32();

            for (int j = 0; j < pieceCount; j++)
            {
                uint fcValue = reader.ReadUInt32();
                reader.ReadUInt32(); // skip PRM for now

                bool isUnicode = (fcValue & 0x40000000) != 0;
                uint fc = fcValue & 0x3FFFFFFF;

                int cpStart = cpArray[j];
                int cpEnd = cpArray[j + 1];
                int fcStart = (int)fc;
                int fcEnd = fcStart + (isUnicode ? (cpEnd - cpStart) * 2 : (cpEnd - cpStart));

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
            throw new InvalidDataException("Failed to parse piece table data", ex);
        }
    }

    public string GetTextForPiece(int index, Stream documentStream)
    {
        if (index < 0 || index >= _pieces.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        var piece = _pieces[index];
        return GetTextForRange(piece.FcStart, piece.FcEnd, documentStream, piece.IsUnicode);
    }

    /// <summary>
    /// Retrieve text for an arbitrary file position range. The range may span
    /// multiple pieces and does not need to align to piece boundaries.
    /// </summary>
    public string GetTextForRange(int fcStart, int fcEnd, Stream documentStream)
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

            sb.Append(GetTextForRange(start, end, documentStream, piece.IsUnicode));
        }

        return sb.ToString();
    }

    private string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, bool isUnicode)
    {
        _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart={fcStart}, fcEnd={fcEnd}, isUnicode={isUnicode}");
        int length = fcEnd - fcStart;
        if (length <= 0)
            return string.Empty;

        using var reader = new BinaryReader(documentStream, System.Text.Encoding.UTF8, leaveOpen: true);
        documentStream.Seek(fcStart, SeekOrigin.Begin);
        byte[] bytes = reader.ReadBytes(length);

        string text;
        if (isUnicode)
        {
            text = System.Text.Encoding.Unicode.GetString(bytes);
            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as UTF-16LE");
        }
        else
        {
            // Try 1252, then 1251, then ISO-8859-1 if output looks garbled
            text = System.Text.Encoding.GetEncoding(1252).GetString(bytes);
            if (text.Count(c => c == '\uFFFD' || c == '?') > text.Length / 4)
            {
                try
                {
                    var alt = System.Text.Encoding.GetEncoding(1251).GetString(bytes);
                    if (alt.Count(c => c == '\uFFFD' || c == '?') < text.Count(c => c == '\uFFFD' || c == '?'))
                    {
                        text = alt;
                        _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as Windows-1251");
                    }
                    else
                    {
                        alt = System.Text.Encoding.GetEncoding("iso-8859-1").GetString(bytes);
                        if (alt.Count(c => c == '\uFFFD' || c == '?') < text.Count(c => c == '\uFFFD' || c == '?'))
                        {
                            text = alt;
                            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as ISO-8859-1");
                        }
                        else
                        {
                            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as Windows-1252 (fallback, garbled)");
                        }
                    }
                }
                catch
                {
                    _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as Windows-1252 (fallback, exception)");
                }
            }
            else
            {
                _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as Windows-1252");
            }
        }

        _logger.LogInfo($"[DEBUG] GetTextForRange returning: '{text}'");
        return CleanText(text);
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            // Allow a wider range of characters including Cyrillic and special punctuation
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
                c == '\r' || c == '\n' || c == '\t' || c == '\v' ||
                (c >= '\u0400' && c <= '\u04FF') || // Cyrillic range
                (c >= '\u00A0' && c <= '\u00FF'))   // Latin-1 Supplement
            {
                sb.Append(c);
            }
            else if (c == '\uFFFD' || c < ' ') // Replace replacement chars and control chars
            {
                // Skip these characters
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Attempt to determine if the pieces in a Word95 document contain
    /// 16‑bit text. This mirrors the heuristic used by wvGuess16bit in the
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
