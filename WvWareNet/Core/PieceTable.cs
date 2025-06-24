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

    public void Parse(byte[] data)
    {
        _pieces.Clear();
        if (data == null || data.Length == 0)
        {
            _logger.LogWarning("Attempted to parse empty piece table data");
            return;
        }

        try
        {
            using var stream = new MemoryStream(data);
            using var reader = new BinaryReader(stream);

            // Read the character position array
            uint cpCount = reader.ReadUInt32();
            for (int i = 0; i < cpCount; i++)
            {
                uint cp = reader.ReadUInt32();
                // We'll store this information later when we process the pieces
            }

            // Read piece descriptors
            while (stream.Position < stream.Length - 4)
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
