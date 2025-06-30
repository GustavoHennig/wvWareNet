import sys
import argparse

def main():
    parser = argparse.ArgumentParser(description='Hex dump of a file section')
    parser.add_argument('file', help='File to dump')
    parser.add_argument('--offset', type=int, default=0, help='Start offset')
    parser.add_argument('--length', type=int, default=256, help='Number of bytes to dump')
    args = parser.parse_args()

    with open(args.file, "rb") as f:
        f.seek(args.offset)
        data = f.read(args.length)
        print(data.hex())

if __name__ == '__main__':
    main()
