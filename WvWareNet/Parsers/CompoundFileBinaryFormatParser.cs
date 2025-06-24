using System;
using System.IO;

namespace WvWareNet.Parsers;

public class CompoundFileBinaryFormatParser
{
    private readonly BinaryReader _reader;
    private uint[] _fat;
    private uint _sectorSize;
    private uint _fatSectorCount;
    private uint _firstFatSectorLocation;

    // Constants
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;
    private const uint DifatSector = 0xFFFFFFFC;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint SectorSize = 512; // Default sector size

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
        _fatSectorCount = _reader.ReadUInt32();
        uint firstDirectorySectorLocation = _reader.ReadUInt32();
        uint transactionSignatureNumber = _reader.ReadUInt32();
        uint miniStreamCutoffSize = _reader.ReadUInt32();
        uint firstMiniFatSectorLocation = _reader.ReadUInt32();
        uint miniFatSectorCount = _reader.ReadUInt32();
        _firstFatSectorLocation = _reader.ReadUInt32();
        uint difatSectorCount = _reader.ReadUInt32();

        // Calculate sector size based on sector shift
        _sectorSize = (uint)(1 << sectorShift);
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
        List<uint> sectorChain = GetFatChain(entry.StartingSectorLocation);
        using var ms = new MemoryStream();
        foreach (uint sector in sectorChain)
        {
            // Calculate sector position: header (512 bytes) + sector * sector size
            long sectorOffset = 512 + (sector * _sectorSize);
            _reader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin);
            byte[] data = _reader.ReadBytes((int)_sectorSize);
            ms.Write(data, 0, data.Length);
        }
        // Truncate to actual stream size
        return ms.ToArray().AsSpan(0, (int)entry.StreamSize).ToArray();
    }

    public List<uint> GetFatChain(uint startSector)
    {
        var chain = new List<uint>();
        uint current = startSector;
        const uint EndOfChain = 0xFFFFFFFE;

        // Read FAT from header or cache
        if (_fat == null)
        {
            ReadFatTable();
        }

        while (current != EndOfChain)
        {
            if (current >= _fat.Length)
            {
                throw new InvalidDataException($"Invalid FAT index: {current}");
            }
            
            chain.Add(current);
            current = _fat[current];
        }
        return chain;
    }

    private void ReadFatTable()
    {
        if (_fat != null) return;

        // Calculate number of entries per sector
        uint entriesPerSector = _sectorSize / sizeof(uint);
        uint totalEntries = entriesPerSector * _fatSectorCount;
        _fat = new uint[totalEntries];
        
        uint currentSector = _firstFatSectorLocation;
        int fatIndex = 0;
        
        for (int i = 0; i < _fatSectorCount; i++)
        {
            // Calculate sector position (sectors start after header)
            long position = (currentSector + 1) * _sectorSize;
            _reader.BaseStream.Seek(position, SeekOrigin.Begin);
            
            // Read all entries in this FAT sector
            for (int j = 0; j < entriesPerSector; j++)
            {
                _fat[fatIndex++] = _reader.ReadUInt32();
            }
            
            // Get next FAT sector from the chain
            if (currentSector >= _fat.Length)
            {
                throw new InvalidDataException($"Invalid FAT index: {currentSector}");
            }
            currentSector = _fat[currentSector];
            
            // Check for end of chain
            if (currentSector == EndOfChain || currentSector == FreeSector) 
                break;
        }
    }
}
