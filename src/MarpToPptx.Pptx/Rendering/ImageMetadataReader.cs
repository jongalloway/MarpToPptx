namespace MarpToPptx.Pptx.Rendering;

internal static class ImageMetadataReader
{
    public static bool TryReadSize(string path, out int width, out int height)
    {
        width = 0;
        height = 0;

        using var stream = File.OpenRead(path);
        using var reader = new BinaryReader(stream);

        Span<byte> header = stackalloc byte[12];
        if (stream.Read(header) < 8)
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

        if (header[0] == 0x42 && header[1] == 0x4D)
        {
            return TryReadBmp(reader, out width, out height);
        }

        // RIFF????WEBP
        if (header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46
            && header[8] == 0x57 && header[9] == 0x45 && header[10] == 0x42 && header[11] == 0x50)
        {
            return TryReadWebP(reader, out width, out height);
        }

        // SVG: XML-based text format. Binary magic bytes don't apply; read a larger
        // prefix (up to 512 bytes) so that files with many leading whitespace bytes,
        // newlines, or a UTF-8 BOM before <?xml or <svg are still detected correctly.
        {
            const int SvgProbeSize = 512;
            var probeSize = stream.CanSeek ? Math.Min(SvgProbeSize, (int)stream.Length) : SvgProbeSize;
            var probe = new byte[probeSize];
            stream.Position = 0;
            var probeRead = stream.Read(probe, 0, probe.Length);
            var svgPrefix = System.Text.Encoding.UTF8.GetString(probe, 0, probeRead).TrimStart();
            if (svgPrefix.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase)
                || svgPrefix.StartsWith("<svg", StringComparison.OrdinalIgnoreCase))
            {
                stream.Position = 0;
                return TryReadSvg(stream, out width, out height);
            }
        }

        return false;
    }

    /// <summary>
    /// Reads the intrinsic size of an SVG document supplied as a string (e.g., an in-memory rendered SVG).
    /// Returns <c>true</c> and sets <paramref name="width"/>/<paramref name="height"/> when the
    /// <c>viewBox</c> or explicit <c>width</c>/<c>height</c> attributes can be parsed.
    /// </summary>
    public static bool TryReadSvgStringSize(string svgText, out int width, out int height)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(svgText);
        return TryReadSvgBytesSize(bytes, out width, out height);
    }

    /// <summary>
    /// Reads the intrinsic size of an SVG document supplied as a UTF-8 byte array.
    /// Returns <c>true</c> and sets <paramref name="width"/>/<paramref name="height"/> when the
    /// <c>viewBox</c> or explicit <c>width</c>/<c>height</c> attributes can be parsed.
    /// </summary>
    public static bool TryReadSvgBytesSize(byte[] svgBytes, out int width, out int height)
    {
        using var stream = new MemoryStream(svgBytes, writable: false);
        return TryReadSvg(stream, out width, out height);
    }

    /// <summary>
    /// Detects the MIME content type of the image from its magic bytes.
    /// Returns null if the format is not recognized.
    /// </summary>
    public static string? TryDetectContentType(Stream stream)
    {
        var position = stream.Position;
        try
        {
            Span<byte> header = stackalloc byte[12];
            var read = stream.Read(header);
            if (read < 2)
            {
                return null;
            }

            if (read >= 4 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
            {
                return "image/png";
            }

            if (header[0] == 0xFF && header[1] == 0xD8)
            {
                return "image/jpeg";
            }

            if (read >= 3 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
            {
                return "image/gif";
            }

            if (header[0] == 0x42 && header[1] == 0x4D)
            {
                return "image/bmp";
            }

            if (read >= 12 && header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46
                && header[8] == 0x57 && header[9] == 0x45 && header[10] == 0x42 && header[11] == 0x50)
            {
                return "image/webp";
            }

            // SVG: look for XML/SVG marker in first bytes
            if (read >= 5)
            {
                Span<byte> svgCheck = stackalloc byte[10];
                stream.Position = position;
                var svgRead = stream.Read(svgCheck);
                var text = System.Text.Encoding.UTF8.GetString(svgCheck[..svgRead]).TrimStart();
                if (text.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) || text.StartsWith("<svg", StringComparison.OrdinalIgnoreCase))
                {
                    return "image/svg+xml";
                }
            }

            return null;
        }
        finally
        {
            stream.Position = position;
        }
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

    private static bool TryReadBmp(BinaryReader reader, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (reader.BaseStream.Length < 26)
        {
            return false;
        }

        reader.BaseStream.Position = 18;
        width = reader.ReadInt32();
        height = Math.Abs(reader.ReadInt32()); // height can be negative (top-down bitmap)
        return width > 0 && height > 0;
    }

    private static bool TryReadWebP(BinaryReader reader, out int width, out int height)
    {
        width = 0;
        height = 0;

        // After "RIFF????WEBP" (12 bytes), read the sub-chunk descriptor
        reader.BaseStream.Position = 12;

        if (reader.BaseStream.Length < 30)
        {
            return false;
        }

        Span<byte> chunkFourCC = stackalloc byte[4];
        if (reader.Read(chunkFourCC) < 4)
        {
            return false;
        }

        // VP8 (lossy): "VP8 "
        if (chunkFourCC[0] == 0x56 && chunkFourCC[1] == 0x50 && chunkFourCC[2] == 0x38 && chunkFourCC[3] == 0x20)
        {
            // Skip chunk size (4 bytes) + 3 bytes frame tag, then 3-byte start code, then dimensions
            reader.BaseStream.Position = 26;
            var w = reader.ReadUInt16() & 0x3FFF;
            var h = reader.ReadUInt16() & 0x3FFF;
            width = (int)w;
            height = (int)h;
            return width > 0 && height > 0;
        }

        // VP8L (lossless): "VP8L"
        if (chunkFourCC[0] == 0x56 && chunkFourCC[1] == 0x50 && chunkFourCC[2] == 0x38 && chunkFourCC[3] == 0x4C)
        {
            // Skip chunk size (4 bytes) + signature byte (0x2F)
            reader.BaseStream.Position = 21;
            // Next 28 bits: 14 bits (width-1) + 14 bits (height-1)
            Span<byte> bits = stackalloc byte[4];
            if (reader.Read(bits) < 4)
            {
                return false;
            }

            var packed = (uint)(bits[0] | (bits[1] << 8) | (bits[2] << 16) | (bits[3] << 24));
            width = (int)(packed & 0x3FFF) + 1;
            height = (int)((packed >> 14) & 0x3FFF) + 1;
            return width > 0 && height > 0;
        }

        // VP8X (extended): "VP8X"
        if (chunkFourCC[0] == 0x56 && chunkFourCC[1] == 0x50 && chunkFourCC[2] == 0x38 && chunkFourCC[3] == 0x58)
        {
            // Skip chunk size (4 bytes) + flags (4 bytes) = 8 bytes past chunk tag → position 24
            reader.BaseStream.Position = 24;
            // Canvas width minus one: 24-bit LE
            Span<byte> cw = stackalloc byte[3];
            Span<byte> ch = stackalloc byte[3];
            if (reader.Read(cw) < 3 || reader.Read(ch) < 3)
            {
                return false;
            }

            width = (cw[0] | (cw[1] << 8) | (cw[2] << 16)) + 1;
            height = (ch[0] | (ch[1] << 8) | (ch[2] << 16)) + 1;
            return width > 0 && height > 0;
        }

        return false;
    }

    private static bool TryReadSvg(Stream stream, out int width, out int height)
    {
        width = 0;
        height = 0;

        var bufferSize = (int)Math.Min(4096, stream.Length);
        var buffer = new byte[bufferSize];
        var read = stream.Read(buffer, 0, bufferSize);
        var text = System.Text.Encoding.UTF8.GetString(buffer, 0, read);

        if (!text.Contains("<svg", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        // Try viewBox="minX minY width height"
        var viewBoxMatch = System.Text.RegularExpressions.Regex.Match(
            text, @"viewBox\s*=\s*[""']([^""']+)[""']", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (viewBoxMatch.Success)
        {
            var parts = viewBoxMatch.Groups[1].Value.Split(
                [' ', ','], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 4 &&
                double.TryParse(parts[2], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var vw) &&
                double.TryParse(parts[3], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var vh) &&
                vw > 0 && vh > 0)
            {
                width = (int)Math.Ceiling(vw);
                height = (int)Math.Ceiling(vh);
                return true;
            }
        }

        // Try explicit width/height attributes on <svg>
        var wMatch = System.Text.RegularExpressions.Regex.Match(
            text, @"<svg[^>]+\bwidth\s*=\s*[""']([^""']+)[""']", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var hMatch = System.Text.RegularExpressions.Regex.Match(
            text, @"<svg[^>]+\bheight\s*=\s*[""']([^""']+)[""']", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wMatch.Success && hMatch.Success &&
            double.TryParse(
                new string(wMatch.Groups[1].Value.TakeWhile(static c => char.IsDigit(c) || c == '.').ToArray()),
                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var w) &&
            double.TryParse(
                new string(hMatch.Groups[1].Value.TakeWhile(static c => char.IsDigit(c) || c == '.').ToArray()),
                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var h) &&
            w > 0 && h > 0)
        {
            width = (int)Math.Ceiling(w);
            height = (int)Math.Ceiling(h);
            return true;
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
