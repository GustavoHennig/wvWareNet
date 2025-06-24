using System;
using System.IO;

namespace WvWareNet.Utilities;

public class FileStreamHandler : IDisposable
{
    private readonly ILogger _logger;
    private FileStream _fileStream;
    private BinaryReader _binaryReader;

    public FileStreamHandler(ILogger logger)
    {
        _logger = logger;
    }

    public void OpenFile(string filePath)
    {
        try
        {
            _fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            _binaryReader = new BinaryReader(_fileStream);
            _logger.LogInfo($"Successfully opened file: {filePath}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error opening file: {filePath}", ex);
            throw;
        }
    }

    public byte[] ReadBytes(int count)
    {
        if (_binaryReader == null)
            throw new InvalidOperationException("File not opened. Call OpenFile first.");

        try
        {
            return _binaryReader.ReadBytes(count);
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error reading {count} bytes from file", ex);
            throw;
        }
    }

    public void Seek(long offset, SeekOrigin origin = SeekOrigin.Begin)
    {
        if (_fileStream == null)
            throw new InvalidOperationException("File not opened. Call OpenFile first.");

        try
        {
            _fileStream.Seek(offset, origin);
            _logger.LogDebug($"Seeked to position: {_fileStream.Position}");
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error seeking to position {offset}", ex);
            throw;
        }
    }

    public long Position => _fileStream?.Position ?? -1;
    public long Length => _fileStream?.Length ?? -1;

    public void Dispose()
    {
        _binaryReader?.Dispose();
        _fileStream?.Dispose();
        _logger.LogInfo("File stream closed");
    }
}
