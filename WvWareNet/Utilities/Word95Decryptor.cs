using System;
using System.IO;
using System.Text;

namespace WvWareNet.Utilities
{
    public static class Word95Decryptor
    {
        public static byte[] Decrypt(byte[] encryptedData, string password, uint lKey)
        {
            if (string.IsNullOrEmpty(password))
                throw new ArgumentException("Password is required for Word95 decryption");

            byte[] pw = new byte[16];
            byte[] key = new byte[16];
            byte[] pwkey = new byte[2];
            
            // Prepare password and key
            pwkey[0] = (byte)((lKey >> 16) & 0xFF);
            pwkey[1] = (byte)((lKey >> 24) & 0xFF);

            // Pad password with special bytes
            int len = Math.Min(password.Length, 16);
            Array.Copy(Encoding.ASCII.GetBytes(password), pw, len);
            
            for (int i = len, j = 0; i < 16; i++, j++)
            {
                switch (j % 15)
                {
                    case 0: pw[i] = 0xbb; break;
                    case 1: case 2: pw[i] = 0xff; break;
                    case 3: pw[i] = 0xba; break;
                    case 4: case 5: pw[i] = 0xff; break;
                    case 6: pw[i] = 0xb9; break;
                    case 7: pw[i] = 0x80; break;
                    case 8: pw[i] = 0x00; break;
                    case 9: pw[i] = 0xbe; break;
                    case 10: pw[i] = 0x0f; break;
                    case 11: pw[i] = 0x00; break;
                    case 12: pw[i] = 0xbf; break;
                    case 13: pw[i] = 0x0f; break;
                    case 14: pw[i] = 0x00; break;
                }
            }

            // Generate decryption key
            ushort h = 0xce4b;
            for (int i = 0; i < 16; i++)
            {
                byte g = (byte)(pw[i] ^ pwkey[i & 1]);
                g = RotateLeft(g, 7, 8);
                h ^= (ushort)(RotateLeft(pw[i], i + 1, 15) ^ (i + 1) ^ i);
                key[i] = g;
            }

            // Verify password hash
            ushort hash = (ushort)(lKey & 0xFFFF);
            if (h != hash)
                throw new InvalidDataException("Invalid password for Word95 document");

            // Decrypt the data
            using (var input = new MemoryStream(encryptedData))
            using (var output = new MemoryStream())
            {
                // Skip first 0x30 bytes (header)
                input.Position = 0x30;
                
                byte[] block = new byte[16];
                int bytesRead;
                while ((bytesRead = input.Read(block, 0, 16)) > 0)
                {
                    for (int i = 0; i < bytesRead; i++)
                    {
                        if (block[i] != 0)
                            block[i] ^= key[i];
                        else
                            block[i] = 0;
                    }
                    output.Write(block, 0, bytesRead);
                }

                return output.ToArray();
            }
        }

        private static byte RotateLeft(byte value, int shift, int bitLength)
        {
            return (byte)(((value << shift) | (value >> (bitLength - shift))) & 0xFF);
        }
    }
}
