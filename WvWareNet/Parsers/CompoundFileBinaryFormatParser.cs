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
    private uint _firstDirectorySectorLocation;
    private uint[] _difat = new uint[109];

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
        Console.WriteLine("[DEBUG] Signature bytes: " + BitConverter.ToString(signature));
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
        _firstDirectorySectorLocation = _reader.ReadUInt32();
        uint transactionSignatureNumber = _reader.ReadUInt32();
        uint miniStreamCutoffSize = _reader.ReadUInt32();
        uint firstMiniFatSectorLocation = _reader.ReadUInt32();
        uint miniFatSectorCount = _reader.ReadUInt32();
        _firstFatSectorLocation = _reader.ReadUInt32();
        uint difatSectorCount = _reader.ReadUInt32();

        // Read DIFAT array (109 entries)
        for (int i = 0; i < 109; i++)
        {
            _difat[i] = _reader.ReadUInt32();
        }

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
        // Read the directory stream using the FAT chain
        List<uint> dirChain = GetFatChain(_firstDirectorySectorLocation);
        using var dirStream = new MemoryStream();
        foreach (uint sector in dirChain)
        {
            long sectorOffset = 512 + (sector * _sectorSize);
            _reader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin);
            byte[] data = _reader.ReadBytes((int)_sectorSize);
            dirStream.Write(data, 0, data.Length);
        }

        // Parse 128-byte directory entries from the directory stream
        var entries = new List<DirectoryEntry>();
        dirStream.Position = 0;
        long dirLen = dirStream.Length;
        using var dirReader = new BinaryReader(dirStream);
        while (dirStream.Position + 128 <= dirLen)
        {
            byte[] nameBytes = dirReader.ReadBytes(64);
            ushort nameLength = dirReader.ReadUInt16();
            string name = "";
            if (nameLength >= 2 && nameLength <= 64 && (nameLength - 2) <= nameBytes.Length)
            {
                name = System.Text.Encoding.Unicode.GetString(nameBytes, 0, nameLength - 2);
                name = name.TrimEnd('\0', ' ', '\t', '\r', '\n');
            }
            byte entryType = dirReader.ReadByte();
            dirStream.Position = dirStream.Position - 65 + 116; // Move to StartingSectorLocation
            uint startingSector = dirReader.ReadUInt32();
            ulong streamSize = dirReader.ReadUInt64();
            dirStream.Position = dirStream.Position - 12 + 128; // Move to next entry

            Console.WriteLine($"[DEBUG] DirectoryEntry: '{name}'");
            entries.Add(new DirectoryEntry
            {
                Name = name,
                EntryType = entryType,
                StartingSectorLocation = startingSector,
                StreamSize = streamSize
            });
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
        Console.WriteLine($"[DEBUG] FAT: _fatSectorCount={_fatSectorCount}, _sectorSize={_sectorSize}, entriesPerSector={entriesPerSector}, totalEntries={totalEntries}");
        _fat = new uint[totalEntries];

        // Collect FAT sector locations from DIFAT array
        var fatSectors = new List<uint>();
        for (int i = 0; i < Math.Min(_fatSectorCount, 109); i++)
        {
            if (_difat[i] != FreeSector)
                fatSectors.Add(_difat[i]);
        }

        // TODO: For files with more than 109 FAT sectors, follow DIFAT sector chain (not implemented here)

        int fatIndex = 0;
        foreach (var sector in fatSectors)
        {
            long position = (sector + 1) * _sectorSize;
            _reader.BaseStream.Seek(position, SeekOrigin.Begin);

            for (int j = 0; j < entriesPerSector; j++)
            {
                if (fatIndex < _fat.Length)
                    _fat[fatIndex++] = _reader.ReadUInt32();
            }
        }
    }
}
