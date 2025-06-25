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

    public void Parse(byte[] data, uint fcMin, uint fcMac)
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
                if (i + 5 > data.Length)
                    break;
                int prcSize = BitConverter.ToInt32(data, i + 1);
                i += 1 + 4 + prcSize;
            }
            else
            {
                // Unknown block, abort
                break;
            }
        }

        if (plcPcdOffset == -1)
        {
            _logger.LogError("No PlcPcd (piece table) found in CLX data");
            throw new InvalidDataException("No PlcPcd (piece table) found in CLX data");
        }

        // Now parse the PlcPcd as the piece table
        try
        {
            using var stream = new MemoryStream(data, plcPcdOffset, plcPcdLength);
            using var reader = new BinaryReader(stream);

            // The piece table (PlcPcd) consists of an array of character
            // positions followed by an array of piece descriptors.  The
            // number of pieces can be derived from the total length.
            int pieceCount = (plcPcdLength - 4) / 12;
            if (pieceCount <= 0)
            {
                _logger.LogWarning($"Invalid piece table length {plcPcdLength}. Using fallback single-piece table.");
                var descriptor = new PieceDescriptor
                {
                    FilePosition = fcMin,
                    IsUnicode = true,
                    HasFormatting = false,
                    CpStart = 0,
                    CpEnd = (int)(fcMac - fcMin)
                };
                _pieces.Add(descriptor);
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
        using var reader = new BinaryReader(documentStream);
        documentStream.Seek(piece.FcStart, SeekOrigin.Begin);

        uint length = (uint)(piece.FcEnd - piece.FcStart);

        if (piece.IsUnicode)
        {
            // When the piece is marked as Unicode the text is stored as
            // little-endian UTF-16. BinaryReader by default uses UTF-8
            // which produces garbage characters.  Read the raw bytes and
            // decode them explicitly using Encoding.Unicode.
            byte[] bytes = reader.ReadBytes((int)length);
            var text = System.Text.Encoding.Unicode.GetString(bytes);
            return CleanText(text);
        }
        else
        {
            // Non Unicode pieces are stored using the Windows code page of
            // the document.  CP1252 is a sensible default for Western
            // documents and matches the behaviour of the original wvWare
            // library.
            byte[] bytes = reader.ReadBytes((int)length);
            var text = System.Text.Encoding.GetEncoding(1252).GetString(bytes);
            return CleanText(text);
        }
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            if (c == '\r' || c == '\n' || c == '\t' || c >= ' ')
                sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Replace the current table with a single Unicode piece using the
    /// supplied file positions. Used as a fallback when the piece table is
    /// corrupt or not present.
    /// </summary>
    public void SetSinglePiece(uint fcMin, uint fcMac)
    {
        _pieces.Clear();
        _pieces.Add(new PieceDescriptor
        {
            FilePosition = fcMin,
            IsUnicode = false,
            HasFormatting = false,
            CpStart = 0,
            CpEnd = (int)(fcMac - fcMin),
            FcStart = (int)fcMin,
            FcEnd = (int)fcMac
        });
    }
}
