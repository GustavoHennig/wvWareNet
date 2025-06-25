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

    public void Parse(byte[] data)
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
                // This is the PlcPcd (piece table) block
                plcPcdOffset = i + 1;
                plcPcdLength = data.Length - plcPcdOffset;
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

            // Read the character position array
            uint cpCount = reader.ReadUInt32();
            Console.WriteLine($"[DEBUG] PieceTable.Parse: cpCount={cpCount}");
            int cpArrayBytes = (int)cpCount * 4;
            int pieceDescriptorBytes = (int)(cpCount - 1) * 9;
            int expectedMinLength = 4 + cpArrayBytes + pieceDescriptorBytes;
            Console.WriteLine($"[DEBUG] PieceTable.Parse: expectedMinLength={expectedMinLength}, plcPcdLength={plcPcdLength}");

            if (expectedMinLength > plcPcdLength)
            {
                throw new InvalidDataException($"PieceTable.Parse: Not enough data for piece table: expected {expectedMinLength}, got {plcPcdLength}");
            }

            var cpArray = new int[cpCount];
            for (int j = 0; j < cpCount; j++)
            {
                cpArray[j] = reader.ReadInt32();
            }

            // Read piece descriptors and assign cpStart/cpEnd
            for (int j = 0; j < cpCount - 1; j++)
            {
                var descriptor = new PieceDescriptor
                {
                    FilePosition = reader.ReadUInt32(),
                };

                // Read flags from the last byte of the descriptor
                byte flags = reader.ReadByte();
                descriptor.IsUnicode = (flags & 0x40) != 0;
                descriptor.HasFormatting = (flags & 0x80) != 0;
                descriptor.ReservedFlags = (byte)(flags & 0x3F); // Lower 6 bits are reserved

                descriptor.CpStart = cpArray[j];
                descriptor.CpEnd = cpArray[j + 1];

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
        documentStream.Seek(piece.FilePosition, SeekOrigin.Begin);

        // For simplicity, we're just reading to the next piece
        // In a real implementation, we would use the character position array
        uint length = (index < _pieces.Count - 1) 
            ? _pieces[index + 1].FilePosition - piece.FilePosition 
            : (uint)(documentStream.Length - piece.FilePosition);

        if (piece.IsUnicode)
        {
            return new string(reader.ReadChars((int)length / 2));
        }
        else
        {
            byte[] bytes = reader.ReadBytes((int)length);
            return System.Text.Encoding.Default.GetString(bytes);
        }
    }
}
