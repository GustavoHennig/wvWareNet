import sys

def hexdump(data, start=0, length=None):
    if length is None:
        length = len(data) - start
    
    end = min(start + length, len(data))
    
    for i in range(start, end, 16):
        hex_str = ' '.join(f'{data[j]:02X}' for j in range(i, min(i + 16, end)))
        ascii_str = ''.join(chr(data[j]) if 32 <= data[j] <= 126 else '.' for j in range(i, min(i + 16, end)))
        print(f'{i:08X}: {hex_str:<48} {ascii_str}')

def search_text(data, text):
    text_bytes = text.encode('ascii', errors='ignore')
    text_utf16 = text.encode('utf-16le', errors='ignore')
    
    print(f"Searching for '{text}' as ASCII and UTF-16LE...")
    
    # Search for ASCII
    pos = data.find(text_bytes)
    if pos != -1:
        print(f"Found ASCII '{text}' at offset 0x{pos:X}")
        hexdump(data, max(0, pos-32), 64)
        return pos
        
    # Search for UTF-16LE
    pos = data.find(text_utf16)
    if pos != -1:
        print(f"Found UTF-16LE '{text}' at offset 0x{pos:X}")
        hexdump(data, max(0, pos-32), 64)
        return pos
        
    print(f"'{text}' not found")
    return -1

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python debug_doc.py <filename>")
        sys.exit(1)
        
    filename = sys.argv[1]
    
    with open(filename, 'rb') as f:
        data = f.read()
    
    print(f"File size: {len(data)} bytes")
    print()
    
    # Show first 512 bytes
    print("First 512 bytes:")
    hexdump(data, 0, 512)
    print()
    
    # Search for expected text
    search_text(data, "Software Configuration Management Solution")
    print()
    search_text(data, "Software")
    print()
    search_text(data, "Management")
    print()
    search_text(data, "Solution")
