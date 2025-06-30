import sys

def hexdump(file_path):
    try:
        with open(file_path, 'rb') as f:
            offset = 0
            while True:
                chunk = f.read(16)
                if not chunk:
                    break
                
                # Format the offset
                print(f"{offset:08x}  ", end='')
                
                # Format hex bytes
                hex_str = ' '.join(f"{b:02x}" for b in chunk)
                hex_str += '   ' * (16 - len(chunk))  # Padding for incomplete lines
                print(hex_str[:23], end=' ')
                print(hex_str[23:], end=' ')
                
                # Format ASCII representation
                ascii_str = ''.join(chr(b) if 32 <= b <= 126 else '.' for b in chunk)
                print(f"|{ascii_str}|")
                
                offset += len(chunk)
    except Exception as e:
        print(f"Error reading file: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python hexdump.py <file>")
        sys.exit(1)
    hexdump(sys.argv[1])
