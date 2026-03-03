namespace MarpToPptx.Pptx.Rendering;

internal static class ImageMetadataReader
{
    public static bool TryReadSize(string path, out int width, out int height)
    {
        width = 0;
        height = 0;

        using var stream = File.OpenRead(path);
        using var reader = new BinaryReader(stream);

        Span<byte> header = stackalloc byte[8];
        if (stream.Read(header) < header.Length)
        {
            return false;
        }

        stream.Position = 0;

        if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
        {
            return TryReadPng(reader, out width, out height);
        }

        if (header[0] == 0xFF && header[1] == 0xD8)
        {
            return TryReadJpeg(reader, out width, out height);
        }

        if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
        {
            return TryReadGif(reader, out width, out height);
        }

        return false;
    }

    private static bool TryReadPng(BinaryReader reader, out int width, out int height)
    {
        width = 0;
        height = 0;
        reader.BaseStream.Position = 16;
        width = ReadInt32BigEndian(reader);
        height = ReadInt32BigEndian(reader);
        return width > 0 && height > 0;
    }

    private static bool TryReadGif(BinaryReader reader, out int width, out int height)
    {
        width = 0;
        height = 0;
        reader.BaseStream.Position = 6;
        width = reader.ReadUInt16();
        height = reader.ReadUInt16();
        return width > 0 && height > 0;
    }

    private static bool TryReadJpeg(BinaryReader reader, out int width, out int height)
    {
        width = 0;
        height = 0;
        reader.BaseStream.Position = 2;

        while (reader.BaseStream.Position < reader.BaseStream.Length)
        {
            if (reader.ReadByte() != 0xFF)
            {
                continue;
            }

            byte marker;
            do
            {
                marker = reader.ReadByte();
            }
            while (marker == 0xFF);

            if (marker is 0xD8 or 0xD9)
            {
                continue;
            }

            var segmentLength = ReadUInt16BigEndian(reader);
            if (segmentLength < 2)
            {
                return false;
            }

            if (marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF)
            {
                _ = reader.ReadByte();
                height = ReadUInt16BigEndian(reader);
                width = ReadUInt16BigEndian(reader);
                return width > 0 && height > 0;
            }

            reader.BaseStream.Seek(segmentLength - 2, SeekOrigin.Current);
        }

        return false;
    }

    private static int ReadInt32BigEndian(BinaryReader reader)
    {
        Span<byte> buffer = stackalloc byte[4];
        reader.Read(buffer);
        if (BitConverter.IsLittleEndian)
        {
            buffer.Reverse();
        }

        return BitConverter.ToInt32(buffer);
    }

    private static ushort ReadUInt16BigEndian(BinaryReader reader)
    {
        Span<byte> buffer = stackalloc byte[2];
        reader.Read(buffer);
        if (BitConverter.IsLittleEndian)
        {
            buffer.Reverse();
        }

        return BitConverter.ToUInt16(buffer);
    }
}
