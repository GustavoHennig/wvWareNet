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
            // positions followed by an array of piece descriptors.  The
            // number of pieces can be derived from the total length.
            int numCps = (plcPcdLength - 4) / 4; // Each CP is 4 bytes
            int pieceCount = (numCps - 1) / 2;
            if (pieceCount <= 0)
            {
                _logger.LogWarning($"Invalid piece table length {plcPcdLength}. Using fallback single-piece table.");
                SetSinglePiece(fcMin, fcMac, nFib);
                return;
            }

            var cpArray = new int[pieceCount + 1];
            for (int j = 0; j < pieceCount + 1; j++)
                cpArray[j] = reader.ReadInt32();

            var fcValues = new uint[pieceCount];
            var unicodeFlags = new bool[pieceCount];
            for (int j = 0; j < pieceCount; j++)
            {
                uint fcValue = reader.ReadUInt32();
                reader.ReadUInt32(); // skip PRM for now

                unicodeFlags[j] = (fcValue & 0x40000000) != 0;
                fcValues[j] = fcValue & 0x3FFFFFFF;
            }

            bool anyUnicode = false;
            for (int j = 0; j < pieceCount; j++)
                if (unicodeFlags[j]) { anyUnicode = true; break; }

            if (nFib == 0x0065 && !anyUnicode)
            {
                bool guess = GuessPiecesAre16Bit(fcValues, cpArray);
                for (int j = 0; j < pieceCount; j++)
                    unicodeFlags[j] = guess;
            }

            for (int j = 0; j < pieceCount; j++)
            {
                int cpStart = cpArray[j];
                int cpEnd = cpArray[j + 1];
                int fcStart = (int)fcValues[j];
                int fcEnd = fcStart + (unicodeFlags[j] ? (cpEnd - cpStart) * 2 : (cpEnd - cpStart));

                var descriptor = new PieceDescriptor
                {
                    FilePosition = fcValues[j],
                    IsUnicode = unicodeFlags[j],
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

        string text = isUnicode
            ? System.Text.Encoding.Unicode.GetString(bytes)
            : System.Text.Encoding.GetEncoding(1252).GetString(bytes);

        _logger.LogInfo($"[DEBUG] GetTextForRange returning: '{text}'");
        return CleanText(text);
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            // Allow only printable ASCII, whitespace, and standard Unicode letters/digits/punctuation
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
                c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '\u2028' || c == '\u2029')
            {
                sb.Append(c);
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
        // Word 97 (nFib 0x0076) and later are typically Unicode.
        // Word 95 (nFib 0x0065) and earlier are typically not Unicode.
        bool isUnicode = nFib >= 0x0076;

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
