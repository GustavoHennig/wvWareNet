using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WvWareNet.Core
{
    public class Stylesheet
    {
        public List<Style> Styles { get; set; } = new List<Style>();

        public static Stylesheet Parse(byte[] stshData)
        {
            var stylesheet = new Stylesheet();
            if (stshData == null || stshData.Length < 2)
                return stylesheet;

            using var ms = new MemoryStream(stshData);
            using var reader = new BinaryReader(ms);

            try
            {
                // STSH structure starts with count of styles
                ushort styleCount = reader.ReadUInt16();
                
                // Skip reserved fields (varies by Word version)
                ushort moreStyleCount = reader.ReadUInt16();
                ushort maxStyleCount = reader.ReadUInt16();
                ushort reserved1 = reader.ReadUInt16();
                ushort reserved2 = reader.ReadUInt16();
                ushort reserved3 = reader.ReadUInt16();

                // Calculate actual number of styles to read
                int actualStyleCount = Math.Min(styleCount, moreStyleCount);
                
                // Read style descriptors
                for (int i = 0; i < actualStyleCount && ms.Position < ms.Length - 10; i++)
                {
                    try
                    {
                        var style = ParseStyleDescriptor(reader, i);
                        if (style != null)
                            stylesheet.Styles.Add(style);
                    }
                    catch
                    {
                        // Skip malformed style entries
                        break;
                    }
                }
            }
            catch
            {
                // Return partial stylesheet if parsing fails
            }

            return stylesheet;
        }

        private static Style? ParseStyleDescriptor(BinaryReader reader, int index)
        {
            long startPos = reader.BaseStream.Position;
            
            // Each style descriptor has a variable length
            // Basic structure: length, name, formatting info
            
            if (reader.BaseStream.Position >= reader.BaseStream.Length - 4)
                return null;

            try
            {
                // Skip style formatting info (variable length)
                ushort styleInfo = reader.ReadUInt16();
                
                // Style name length
                byte nameLength = reader.ReadByte();
                if (nameLength == 0 || nameLength > 64)
                {
                    // Use default names for built-in styles
                    return CreateBuiltInStyle(index);
                }

                // Read style name
                byte[] nameBytes = reader.ReadBytes(nameLength);
                string styleName = Encoding.UTF8.GetString(nameBytes).TrimEnd('\0');
                
                if (string.IsNullOrEmpty(styleName))
                    styleName = CreateBuiltInStyle(index)?.Name ?? $"Style{index}";

                return new Style
                {
                    Index = index,
                    Name = styleName,
                    Type = StyleType.Paragraph
                };
            }
            catch
            {
                // Return built-in style as fallback
                return CreateBuiltInStyle(index);
            }
        }

        private static Style? CreateBuiltInStyle(int index)
        {
            // Map common built-in style indices to names
            string name = index switch
            {
                0 => "Normal",
                1 => "heading 1",
                2 => "heading 2", 
                3 => "heading 3",
                4 => "heading 4",
                5 => "heading 5",
                6 => "heading 6",
                7 => "heading 7",  
                8 => "heading 8",
                9 => "heading 9",
                10 => "List",
                11 => "List 2",
                12 => "List 3",
                13 => "List 4",
                14 => "List 5",
                15 => "Header",
                16 => "Footer",
                _ => $"Style{index}"
            };

            return new Style
            {
                Index = index,
                Name = name,
                Type = StyleType.Paragraph
            };
        }

        public Style? GetStyle(int index)
        {
            return Styles.Find(s => s.Index == index);
        }

        public string GetStyleName(int index)
        {
            var style = GetStyle(index);
            return style?.Name ?? $"Style{index}";
        }
    }

    public class Style
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
        public StyleType Type { get; set; }
        public ParagraphProperties? ParagraphProperties { get; set; }
        public CharacterProperties? CharacterProperties { get; set; }
    }

    public enum StyleType
    {
        Paragraph,
        Character,
        Table,
        List
    }
}
