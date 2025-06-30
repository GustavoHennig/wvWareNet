using System;
using System.Text;

namespace WvWareNet.Core;

public static class FastsavedTextFormatter
{
    public static string FormatFastsavedText(string input)
    {
        var sb = new StringBuilder(input.Length);
        int i = 0;
        
        while (i < input.Length)
        {
            char c = input[i];
            
            if (c == (char)0x01) // Start of fastsaved text
            {
                int start = i + 1;
                int end = input.IndexOf((char)0x01, start);
                if (end == -1) end = input.Length;
                
                string fastSavedText = input.Substring(start, end - start);
                
                if (end + 1 < input.Length)
                {
                    char formatFlag = input[end + 1];
                    
                    switch (formatFlag)
                    {
                        case (char)0x02: // Bold
                            fastSavedText = "**" + fastSavedText + "**";
                            i = end + 2;
                            break;
                        case (char)0x03: // Italic
                            fastSavedText = "_" + fastSavedText + "_";
                            i = end + 2;
                            break;
                        case (char)0x13: // Bold+Italic
                            fastSavedText = "_**" + fastSavedText + "**_";
                            i = end + 2;
                            break;
                        default:
                            i = end + 1;
                            break;
                    }
                }
                else
                {
                    i = end + 1;
                }
                
                sb.Append(fastSavedText);
            }
            else
            {
                sb.Append(c);
                i++;
            }
        }
        
        return sb.ToString();
    }
}
