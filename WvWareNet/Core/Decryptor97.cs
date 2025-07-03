using System;
using System.IO;
using System.Security.Cryptography;
using System.Collections;

namespace WvWareNet.Core
{
    public class DecryptionResult
    {
        public byte[] TableStream { get; set; }
        public byte[] MainStream { get; set; }
    }

    public static class Decryptor97
    {
        public static DecryptionResult Decrypt(
            byte[] tableEncrypted,
            byte[] mainEncrypted,
            string password,
            byte[] docId,
            byte[] salt,
            byte[] hashedSalt)
        {
            if (string.IsNullOrEmpty(password))
                throw new ArgumentException("Password is required for Word97 decryption", nameof(password));

            // Expand password into 64-byte array
            byte[] pwArray = new byte[64];
            ExpandPassword(password, pwArray);

            // Verify password and compute 128-bit hashed password into valDigest
            byte[] valDigest = VerifyPassword(pwArray, docId, salt, hashedSalt);

            // Decrypt both streams
            var tableDecrypted = DecryptStream(tableEncrypted, pwArray, valDigest);
            var mainDecrypted = DecryptStream(mainEncrypted, pwArray, valDigest);

            return new DecryptionResult
            {
                TableStream = tableDecrypted,
                MainStream = mainDecrypted
            };
        }

        // Expand the Unicode (16-bit) password into the 64-byte array with padding
        private static void ExpandPassword(string password, byte[] output)
        {
            // Clear
            Array.Clear(output, 0, 64);

            // Fill with UTF-16LE characters until null or 16 characters
            int i = 0;
            foreach (char c in password)
            {
                if (i >= 16) break;
                ushort code = c;
                output[2 * i] = (byte)(code & 0xFF);
                output[2 * i + 1] = (byte)(code >> 8);
                i++;
            }

            // Padding marker
            output[2 * i] = 0x80;
            output[56] = (byte)(i << 4);
        }

        // Verifies password; returns the 16-byte MD5 digest (valDigest) on success or throws
        private static byte[] VerifyPassword(
            byte[] pwArray,
            byte[] docId,
            byte[] salt,
            byte[] hashedSalt)
        {
            // Step 1: MD5(pwArray)
            byte[] md1;
            using (var md5 = MD5.Create())
            {
                md1 = md5.ComputeHash(pwArray);
            }

            // Step 2: Build valDigest by iterating over 64-byte blocks mixing md1 and docId
            byte[] valDigest;
            using (var md5 = MD5.Create())
            {
                int offset = 0, keyOffset = 0;
                uint toCopy = 5;
                // Initialize
                md5.Initialize();
                while (offset < 16)
                {
                    // copy md1 bytes into pwArray until full or a full block
                    if (64 - offset < 5)
                        toCopy = (uint)(64 - offset);

                    Array.Copy(md1, keyOffset, pwArray, offset, toCopy);
                    offset += (int)toCopy;

                    if (offset == 64)
                    {
                        md5.TransformBlock(pwArray, 0, 64, null, 0);
                        keyOffset = (int)toCopy;
                        toCopy = 5 - toCopy;
                        offset = 0;
                        continue;
                    }

                    keyOffset = 0;
                    toCopy = 5;
                    Array.Copy(docId, 0, pwArray, offset, 16);
                    offset += 16;
                }

                // Pad and finalize
                pwArray[16] = 0x80;
                Array.Clear(pwArray, 17, 47);
                pwArray[56] = 0x80;
                pwArray[57] = 0x0A;

                md5.TransformBlock(pwArray, 0, 64, null, 0);
                md5.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                valDigest = md5.Hash;
            }

            // Step 3: Generate 40-bit RC4 key from valDigest, then decrypt salt and hashedSalt
            var rc4Key = MakeKey(0, valDigest);

            // Decrypt salt and hashedSalt in place
            rc4Key.Process(salt);
            rc4Key.Process(hashedSalt);

            // Step 4: MD5(salt)
            byte[] checkDigest;
            using (var md5 = MD5.Create())
            {
                // pad salt to 64 bytes
                var buf = new byte[64];
                Array.Copy(salt, buf, 16);
                buf[16] = 0x80;
                // rest is already zero
                md5.Initialize();
                md5.TransformFinalBlock(buf, 0, buf.Length);
                checkDigest = md5.Hash;
            }

            // Validate
            if (!StructuralComparisons.StructuralEqualityComparer.Equals(checkDigest, hashedSalt))
                throw new InvalidDataException("Invalid password for Word97 document");

            return valDigest;
        }

        // Decrypts a full stream using RC4 with rekey every 0x200 bytes
        private static byte[] DecryptStream(byte[] encrypted, byte[] pwArray, byte[] valDigest)
        {
            int length = encrypted.Length;
            var output = new byte[length];
            uint block = 0;
            var rc4 = MakeKey(block, valDigest);

            for (int pos = 0; pos < length; pos += 16)
            {
                int chunk = Math.Min(16, length - pos);
                rc4.Process(encrypted, pos, chunk, output, pos);

                // rekey every 0x200 bytes
                if (((pos + chunk) & 0x1FF) == 0)
                {
                    block++;
                    rc4 = MakeKey(block, valDigest);
                }
            }

            return output;
        }

        // Constructs an RC4 cipher keyed by MD5(pwArray + block || padding)
        private static RC4Cipher MakeKey(uint block, byte[] valDigest)
        {
            // Build 64-byte array
            var buf = new byte[64];
            Array.Copy(valDigest, buf, 5);
            buf[5] = (byte)(block & 0xFF);
            buf[6] = (byte)((block >> 8) & 0xFF);
            buf[7] = (byte)((block >> 16) & 0xFF);
            buf[8] = (byte)((block >> 24) & 0xFF);
            buf[9] = 0x80;
            buf[56] = 0x48;

            // MD5 of buffer
            byte[] digest;
            using (var md5 = MD5.Create())
            {
                digest = md5.ComputeHash(buf);
            }

            // Create RC4 with first 16 bytes of digest
            var key = new byte[16];
            Array.Copy(digest, 0, key, 0, 16);
            return new RC4Cipher(key);
        }

        // Simple RC4 implementation with persistent state
        private class RC4Cipher
        {
            private readonly byte[] S = new byte[256];
            private int i, j;

            public RC4Cipher(byte[] key)
            {
                for (int k = 0; k < 256; k++)
                {
                    S[k] = (byte)k;
                }
                j = 0;
                for (int k = 0; k < 256; k++)
                {
                    j = (j + S[k] + key[k % key.Length]) & 0xFF;
                    Swap(k, j);
                }
                i = j = 0;
            }

            public void Process(byte[] data)
            {
                Process(data, 0, data.Length, data, 0);
            }

            public void Process(byte[] input, int inOffset, int length, byte[] output, int outOffset)
            {
                for (int n = 0; n < length; n++)
                {
                    i = (i + 1) & 0xFF;
                    j = (j + S[i]) & 0xFF;
                    Swap(i, j);
                    byte k = S[(S[i] + S[j]) & 0xFF];
                    output[outOffset + n] = (byte)(input[inOffset + n] ^ k);
                }
            }

            private void Swap(int a, int b)
            {
                byte tmp = S[a];
                S[a] = S[b];
                S[b] = tmp;
            }
        }
    }
}
