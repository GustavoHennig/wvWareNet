using System;
using System.IO;

namespace WvWareNet.Parsers;

public class CompoundFileBinaryFormatParser
{
    private readonly BinaryReader _reader;

    public CompoundFileBinaryFormatParser(Stream stream)
    {
        _reader = new BinaryReader(stream);
    }

    public void ParseHeader()
    {
        // Read and validate header signature
        byte[] signature = _reader.ReadBytes(8);
        if (!IsValidSignature(signature))
        {
            throw new InvalidDataException("Invalid CFBF file signature");
        }

        // Read header fields
        Guid clsid = new Guid(_reader.ReadBytes(16));
        ushort minorVersion = _reader.ReadUInt16();
        ushort majorVersion = _reader.ReadUInt16();
        ushort byteOrder = _reader.ReadUInt16();
        ushort sectorShift = _reader.ReadUInt16();
        ushort miniSectorShift = _reader.ReadUInt16();
        _reader.ReadBytes(6); // Reserved
        uint directorySectorCount = _reader.ReadUInt32();
        uint fatSectorCount = _reader.ReadUInt32();
        uint firstDirectorySectorLocation = _reader.ReadUInt32();
        uint transactionSignatureNumber = _reader.ReadUInt32();
        uint miniStreamCutoffSize = _reader.ReadUInt32();
        uint firstMiniFatSectorLocation = _reader.ReadUInt32();
        uint miniFatSectorCount = _reader.ReadUInt32();
        uint firstDifatSectorLocation = _reader.ReadUInt32();
        uint difatSectorCount = _reader.ReadUInt32();

        // TODO: Process header information and initialize parser state
    }

    private bool IsValidSignature(byte[] signature)
    {
        // CFBF signature should be: D0 CF 11 E0 A1 B1 1A E1
        byte[] expectedSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        return signature.SequenceEqual(expectedSignature);
    }

    // Directory Entry structure (partial, for scaffolding)
    public class DirectoryEntry
    {
        public string Name { get; set; }
        public byte EntryType { get; set; }
        public uint StartingSectorLocation { get; set; }
        public ulong StreamSize { get; set; }
        // Add more fields as needed
    }

    public List<DirectoryEntry> ParseDirectoryEntries()
    {
        // Assumes header has been parsed and directory sector location is known.
        // For demonstration, this reads a single directory entry at the current position.
        var entries = new List<DirectoryEntry>();
        try
        {
            while (_reader.BaseStream.Position < _reader.BaseStream.Length)
            {
                long entryStart = _reader.BaseStream.Position;
                byte[] nameBytes = _reader.ReadBytes(64);
                ushort nameLength = _reader.ReadUInt16();
                string name = System.Text.Encoding.Unicode.GetString(nameBytes, 0, nameLength - 2);
                byte entryType = _reader.ReadByte();
                _reader.BaseStream.Position = entryStart + 116; // Skip to StartingSectorLocation
                uint startingSector = _reader.ReadUInt32();
                ulong streamSize = _reader.ReadUInt64();
                _reader.BaseStream.Position = entryStart + 128; // Move to next entry

                entries.Add(new DirectoryEntry
                {
                    Name = name,
                    EntryType = entryType,
                    StartingSectorLocation = startingSector,
                    StreamSize = streamSize
                });

                if (_reader.BaseStream.Position + 128 > _reader.BaseStream.Length)
                    break;
            }
        }
        catch (EndOfStreamException)
        {
            // End of directory entries
        }
        return entries;
    }

    public byte[] ReadStream(DirectoryEntry entry)
    {
        // NOTE: This is a simplified implementation for standard streams only.
        // In a full implementation, FAT and sector size would be tracked from the header.
        const int sectorSize = 512; // Typical for CFBF, but should be read from header
        List<uint> sectorChain = GetFatChain(entry.StartingSectorLocation);
        using var ms = new MemoryStream();
        foreach (uint sector in sectorChain)
        {
            long sectorOffset = (sector + 1) * sectorSize; // +1 to skip header
            _reader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin);
            byte[] data = _reader.ReadBytes(sectorSize);
            ms.Write(data, 0, data.Length);
        }
        // Truncate to actual stream size
        return ms.ToArray().AsSpan(0, (int)entry.StreamSize).ToArray();
    }

    public List<uint> GetFatChain(uint startSector)
    {
        // NOTE: This is a simplified placeholder. In a full implementation,
        // the FAT should be read from the file and stored in a field.
        // Here, we simulate a single-sector stream for demonstration.
        return new List<uint> { startSector };
    }
}
