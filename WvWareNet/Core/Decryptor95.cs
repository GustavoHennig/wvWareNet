namespace WvWareNet.Core;

public static class Decryptor95
{
    public static byte[] Decrypt(byte[] encryptedData, string password, uint lKey)
    {
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("Password is required for Word95 decryption");

        byte[] pw = new byte[16];
        byte[] key = new byte[16];
        byte[] pwkey = new byte[2];

        pwkey[0] = (byte)((lKey >> 16) & 0xFF);
        pwkey[1] = (byte)((lKey >> 24) & 0xFF);

        int len = Math.Min(password.Length, 16);
        var pwBytes = System.Text.Encoding.ASCII.GetBytes(password);
        Array.Copy(pwBytes, 0, pw, 0, len);

        for (int i = len, j = 0; i < 16; i++, j++)
        {
            switch (j % 15)
            {
                case 0: pw[i] = 0xbb; break;
                case 1: pw[i] = 0xff; break;
                case 2: pw[i] = 0xff; break;
                case 3: pw[i] = 0xba; break;
                case 4: pw[i] = 0xff; break;
                case 5: pw[i] = 0xff; break;
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

        ushort h = 0xce4b;
        for (int i = 0; i < 16; i++)
        {
            byte g = (byte)(pw[i] ^ pwkey[i & 1]);
            g = RotateLeft(g, 7, 8);
            h ^= (ushort)(RotateLeft(pw[i], i + 1, 15) ^ (i + 1) ^ i);
            key[i] = g;
        }

        ushort hash = (ushort)(lKey & 0xFFFF);
        if (h != hash)
            throw new InvalidDataException("Invalid password for Word95 document");

        byte[] output = new byte[encryptedData.Length];
        int length = encryptedData.Length;
        int pos = 0;
        while (pos < 0x30 && pos < length)
        {
            output[pos] = encryptedData[pos];
            pos++;
        }

        for (int block = pos; block < length; block += 16)
        {
            for (int i = 0; i < 16 && block + i < length; i++)
            {
                byte b = encryptedData[block + i];
                output[block + i] = b != 0 ? (byte)(b ^ key[i]) : (byte)0;
            }
        }

        return output;
    }

    private static byte RotateLeft(byte value, int shift, int bitLength)
    {
        return (byte)(((value << shift) | (value >> (bitLength - shift))) & 0xFF);
    }
}
