using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WvWareNet.Utilities;

namespace WvWareNet.Core
{
    /// <summary>
    /// Failed refactoring of PieceTable class.
    /// </summary>
        public class PieceTableRefactFail
    {
        private readonly ILogger _logger;
        private readonly List<PieceDescriptor> _pieces = new();
        private int[] _cpArray = Array.Empty<int>();

        public IReadOnlyList<PieceDescriptor> Pieces => _pieces;

        public PieceTableRefactFail(ILogger logger)
        {
            _logger = logger;
        }

        public static PieceTable CreateFromStreams(ILogger logger, FileInformationBlock fib, byte[]? tableStream, byte[] wordDocStream)
        {
            var pieceTable = new PieceTable(logger);
            byte[]? clxData = null;

            // Try to extract CLX data from table stream
            if (tableStream != null)
            {
                if (fib.LcbClx > 0 && fib.FcClx >= 0 && (fib.FcClx + fib.LcbClx) <= tableStream.Length)
                {
                    clxData = new byte[fib.LcbClx];
                    Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);
                    logger.LogInfo($"[DEBUG] Extracted CLX data: offset={fib.FcClx}, size={fib.LcbClx}");
                }
                else if (fib.NFib <= 0x00D0) // Word 6/95
                {
                    // For Word 6/95, CLX might be at the beginning of table stream
                    int assumedClxLength = Math.Min(tableStream.Length, 512);
                    clxData = new byte[assumedClxLength];
                    Array.Copy(tableStream, 0, clxData, 0, assumedClxLength);
                    logger.LogInfo($"[DEBUG] Word 6/95: Assuming CLX at start of table stream, size={assumedClxLength}");
                }
            }

            if (clxData != null && clxData.Length > 0)
            {
                pieceTable.Parse(clxData, fib);
            }

            // If no pieces were parsed, create a single piece
            if (pieceTable.Pieces.Count == 0)
            {
                logger.LogWarning("No piece table found. Creating single piece for entire document.");
                pieceTable.SetSinglePiece(fib);
            }

            return pieceTable;
        }

        public void Parse(byte[] clxData, FileInformationBlock fib)
        {
            _pieces.Clear();
            _cpArray = Array.Empty<int>();

            if (clxData.Length == 0) return;

            // Log CLX data for debugging
            _logger.LogInfo($"[DEBUG] CLX data (first 32 bytes): {BitConverter.ToString(clxData.Take(Math.Min(32, clxData.Length)).ToArray())}");

            int plcPcdOffset = -1;
            int plcPcdLength = -1;

            // Find the piece table (PlcPcd) in the CLX data
            int i = 0;
            while (i < clxData.Length)
            {
                byte clxType = clxData[i];
                if (clxType == 0x02) // Pcdt block
                {
                    if (i + 5 > clxData.Length) break;
                    plcPcdLength = BitConverter.ToInt32(clxData, i + 1);
                    plcPcdOffset = i + 5;
                    _logger.LogInfo($"[DEBUG] Found PlcPcd at offset {plcPcdOffset}, length {plcPcdLength}");
                    break;
                }
                else if (clxType == 0x01) // Prc block
                {
                    if (i + 3 > clxData.Length) break;
                    ushort prcSize = BitConverter.ToUInt16(clxData, i + 1);
                    i += 1 + 2 + prcSize;
                }
                else if (clxType == 0x00) // Skip null bytes
                {
                    i++;
                }
                else
                {
                    // No CLX structure found, try to parse as raw piece table
                    plcPcdOffset = i;
                    plcPcdLength = clxData.Length - i;
                    _logger.LogInfo($"[DEBUG] No CLX structure, parsing from offset {i}");
                    break;
                }
            }

            if (plcPcdOffset == -1 || plcPcdLength < 16) // Minimum size for one piece
            {
                _logger.LogWarning("Invalid or missing piece table data");
                return;
            }

            try
            {
                // Ensure we don't read beyond CLX data bounds
                if (plcPcdOffset + plcPcdLength > clxData.Length)
                {
                    plcPcdLength = clxData.Length - plcPcdOffset;
                }

                // Log the piece table data
                _logger.LogInfo($"[DEBUG] Piece table data: {BitConverter.ToString(clxData.Skip(plcPcdOffset).Take(plcPcdLength).ToArray())}");

                using var stream = new MemoryStream(clxData, plcPcdOffset, plcPcdLength);
                using var reader = new BinaryReader(stream);

                // Calculate piece count
                // PlcPcd structure: array of CPs (4 bytes each) followed by array of PCDs (8 bytes each)
                // So for n pieces: (n+1)*4 bytes for CPs + n*8 bytes for PCDs = total
                // Therefore: 4n + 4 + 8n = total => 12n + 4 = total => n = (total - 4) / 12
                int pieceCount = (plcPcdLength - 4) / 12;
                if (pieceCount <= 0) return;

                _logger.LogInfo($"[DEBUG] Calculated piece count: {pieceCount}");

                // Read CP array
                _cpArray = new int[pieceCount + 1];
                for (int j = 0; j <= pieceCount; j++)
                {
                    _cpArray[j] = reader.ReadInt32();
                    _logger.LogInfo($"[DEBUG] CP[{j}] = {_cpArray[j]}");
                }

                // Read piece descriptors
                for (int j = 0; j < pieceCount; j++)
                {
                    // According to [MS-DOC], a PCD is 8 bytes:
                    // First 2 bytes are ABCDEFGH IJKLMNOP
                    // where A = fNoParaLast, B = fPaphNil, C = fCopied
                    // D-P = reserved (must be 0)
                    byte byte1 = reader.ReadByte();
                    byte byte2 = reader.ReadByte();

                    // Next 4 bytes are the FC structure
                    uint fc = reader.ReadUInt32();

                    // Last 2 bytes are the PRM
                    ushort prm = reader.ReadUInt16();

                    var piece = new PieceDescriptor
                    {
                        RawFc = fc,
                        Prm = prm,
                        CpStart = _cpArray[j],
                        CpEnd = _cpArray[j + 1]
                    };

                    _pieces.Add(piece);
                    _logger.LogInfo($"[DEBUG] Piece {j}: Flags={byte1:X2}{byte2:X2}, RawFC=0x{fc:X8}, PRM=0x{prm:X4}, CP {_cpArray[j]}-{_cpArray[j + 1]}");
                }

                // Process FC values according to Word version
                ProcessPieceDescriptors(fib);

                _logger.LogInfo($"Parsed {_pieces.Count} piece descriptors from piece table.");
            }
            catch (Exception ex)
            {
                _logger.LogError("Error parsing piece table data", ex);
                _pieces.Clear();
                _cpArray = Array.Empty<int>();
            }
        }

        private void ProcessPieceDescriptors(FileInformationBlock fib)
        {
            // Process each piece to determine encoding and FC values
            for (int i = 0; i < _pieces.Count; i++)
            {
                var piece = _pieces[i];
                uint rawFc = piece.RawFc;

                // For Word 6/95 documents with fExtChar=0, apply special handling
                if (fib.NFib <= 0x00D0 && !fib.FExtChar)
                {
                    // The C code multiplies FC by 2 and sets bit 30
                    rawFc = (rawFc * 2) | 0x40000000U;
                    piece.RawFc = rawFc;
                    _logger.LogInfo($"[DEBUG] Word 6/95: Modified piece {i} FC to 0x{rawFc:X8}");
                }

                // Determine encoding based on bit 30
                // Bit 30 set (0x40000000) = 8-bit text (compressed)
                // Bit 30 clear = 16-bit Unicode text
                piece.IsUnicode = (rawFc & 0x40000000) == 0;

                // Extract actual file position
                if (!piece.IsUnicode)
                {
                    // 8-bit text: clear bit 30 and divide by 2
                    piece.FcStart = (int)((rawFc & 0xBFFFFFFF) / 2);
                }
                else
                {
                    // 16-bit Unicode text
                    // The FC value seems to have extra bits set in some documents
                    // If the raw FC is suspiciously large (> 0x1000000), extract just the lower bits
                    if (rawFc > 0x1000000)
                    {
                        // Extract the actual FC from the lower portion
                        // Based on the log showing 0x10000078, the FC is 0x78
                        piece.FcStart = (int)(rawFc & 0xFFFF);

                        // If that's still too large, try just the lowest byte
                        if (piece.FcStart > 4096 && (rawFc & 0xFF) < 4096)
                        {
                            piece.FcStart = (int)(rawFc & 0xFF);
                        }

                        _logger.LogInfo($"[DEBUG] Adjusted large FC from 0x{rawFc:X8} to 0x{piece.FcStart:X8}");
                    }
                    else
                    {
                        piece.FcStart = (int)(rawFc & 0x3FFFFFFF);
                    }
                }

                _logger.LogInfo($"[DEBUG] Piece {i}: RawFC=0x{rawFc:X8}, FC=0x{piece.FcStart:X8}, Unicode={piece.IsUnicode}, CP {piece.CpStart}-{piece.CpEnd}");
            }
        }

        public int GetPieceIndexFromCp(int cp)
        {
            for (int i = 0; i < _pieces.Count; i++)
            {
                if (cp >= _pieces[i].CpStart && cp < _pieces[i].CpEnd)
                {
                    return i;
                }
            }

            // Handle edge case for last character position
            if (_pieces.Count > 0 && cp == _pieces[^1].CpEnd)
            {
                return _pieces.Count - 1;
            }

            return -1;
        }

        public int ConvertCpToFc(int cp)
        {
            int pieceIndex = GetPieceIndexFromCp(cp);
            if (pieceIndex == -1)
            {
                _logger.LogWarning($"Could not find piece for CP={cp}");
                return -1;
            }

            var piece = _pieces[pieceIndex];
            int offsetInPiece = cp - piece.CpStart;

            // Calculate FC based on whether text is Unicode or ANSI
            int fc = piece.FcStart + (piece.IsUnicode ? offsetInPiece * 2 : offsetInPiece);

            _logger.LogInfo($"[DEBUG] ConvertCpToFc: CP {cp} -> Piece {pieceIndex} (CP {piece.CpStart}-{piece.CpEnd}), offset {offsetInPiece}, FcStart {piece.FcStart}, FC {fc}");
            return fc;
        }

        public void SetSinglePiece(FileInformationBlock fib)
        {
            _pieces.Clear();
            _cpArray = new int[] { 0, (int)fib.CcpText };

            // For fallback single piece, determine encoding based on Word version
            bool isUnicode = false;

            // Word 97 and later default to Unicode unless specified otherwise
            if (fib.NFib >= 0x00C1)
            {
                isUnicode = fib.FExtChar;
            }

            int fcStart = (int)fib.FcMin;
            int fcEnd = (int)fib.FcMac;

            _pieces.Add(new PieceDescriptor
            {
                IsUnicode = isUnicode,
                CpStart = 0,
                CpEnd = (int)fib.CcpText,
                FcStart = fcStart,
                RawFc = (uint)(isUnicode ? fcStart : (fcStart * 2) | 0x40000000U)
            });

            _logger.LogInfo($"[DEBUG] Single piece: FC {fcStart}-{fcEnd}, CP 0-{fib.CcpText}, Unicode={isUnicode}");
        }

        public string GetTextForRange(int startCp, int endCp, Stream documentStream)
        {
            if (endCp <= startCp) return string.Empty;

            var sb = new System.Text.StringBuilder();
            int currentCp = startCp;

            while (currentCp < endCp)
            {
                int pieceIndex = GetPieceIndexFromCp(currentCp);
                if (pieceIndex == -1)
                {
                    _logger.LogWarning($"Could not find piece for CP={currentCp}");
                    break;
                }

                var piece = _pieces[pieceIndex];

                // Calculate how many characters to read from this piece
                int cpInPieceStart = currentCp - piece.CpStart;
                int cpInPieceEnd = Math.Min(endCp - piece.CpStart, piece.CpEnd - piece.CpStart);
                int cpCount = cpInPieceEnd - cpInPieceStart;

                if (cpCount <= 0)
                {
                    currentCp = piece.CpEnd;
                    continue;
                }

                // Convert CP to FC for this position
                int fc = piece.FcStart + (piece.IsUnicode ? cpInPieceStart * 2 : cpInPieceStart);
                int byteCount = piece.IsUnicode ? cpCount * 2 : cpCount;

                // Validate FC and byte count
                if (fc < 0 || fc >= documentStream.Length)
                {
                    _logger.LogWarning($"Invalid FC {fc} for piece {pieceIndex}. Stream length: {documentStream.Length}");
                    break;
                }

                if (fc + byteCount > documentStream.Length)
                {
                    _logger.LogWarning($"Truncating read at FC {fc}: requested {byteCount} bytes, available {documentStream.Length - fc}");
                    byteCount = (int)(documentStream.Length - fc);
                    if (byteCount <= 0)
                    {
                        break;
                    }
                }

                // Read the text
                try
                {
                    documentStream.Seek(fc, SeekOrigin.Begin);
                    byte[] bytes = new byte[byteCount];
                    int bytesRead = documentStream.Read(bytes, 0, byteCount);

                    if (bytesRead < byteCount)
                    {
                        _logger.LogWarning($"Could only read {bytesRead} of {byteCount} bytes at FC {fc}");
                        Array.Resize(ref bytes, bytesRead);
                    }

                    string text = piece.IsUnicode
                        ? System.Text.Encoding.Unicode.GetString(bytes)
                        : System.Text.Encoding.GetEncoding(1252).GetString(bytes);

                    sb.Append(text);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error reading text at FC {fc}: {ex.Message}");
                    break;
                }

                // Move to next piece or position
                currentCp += cpCount;
            }

            return sb.ToString();
        }
    }
}