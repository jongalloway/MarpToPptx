using System.IO.Compression;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml.Linq;

namespace MarpToPptx.Tests;

internal static class PptxGoldenPackage
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        WriteIndented = true,
    };

    private static readonly string[] FixedXmlPaths =
    [
        "[Content_Types].xml",
        "_rels/.rels",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    ];

    public static void AssertMatchesFixture(string pptxPath, string fixtureFileName)
    {
        var actual = CreateSnapshot(pptxPath);
        var actualJson = Serialize(actual);
        var fixturePath = GetFixturePath(fixtureFileName);

        if (ShouldUpdateFixtures())
        {
            Directory.CreateDirectory(Path.GetDirectoryName(fixturePath)!);
            File.WriteAllText(fixturePath, actualJson);
        }

        Assert.True(File.Exists(fixturePath), $"Golden fixture '{fixtureFileName}' was not found at '{fixturePath}'. Set UPDATE_GOLDEN_PACKAGES=1 and rerun the tests to create it intentionally.");

        var expectedJson = NormalizeLineEndings(File.ReadAllText(fixturePath));
        Assert.Equal(expectedJson, actualJson);
    }

    private static PptxPackageSnapshot CreateSnapshot(string pptxPath)
    {
        using var archive = ZipFile.OpenRead(pptxPath);
        var entryPaths = archive.Entries
            .Where(entry => !string.IsNullOrEmpty(entry.Name))
            .Select(entry => entry.FullName.Replace('\\', '/'))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        var xmlPaths = FixedXmlPaths
            .Concat(entryPaths.Where(path => path.StartsWith("ppt/slideLayouts/_rels/", StringComparison.Ordinal) && path.EndsWith(".rels", StringComparison.Ordinal)))
            .Concat(entryPaths.Where(path => path.StartsWith("ppt/slides/_rels/", StringComparison.Ordinal) && path.EndsWith(".rels", StringComparison.Ordinal)))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        var normalizedXmlParts = xmlPaths.ToDictionary(
            path => path,
            path => NormalizeXml(ReadArchiveEntry(archive, path)),
            StringComparer.Ordinal);

        return new PptxPackageSnapshot(entryPaths, normalizedXmlParts);
    }

    private static string ReadArchiveEntry(ZipArchive archive, string path)
    {
        using var stream = archive.GetEntry(path)?.Open()
            ?? throw new InvalidOperationException($"Expected archive entry '{path}' was not found.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string NormalizeXml(string xml)
    {
        var document = XDocument.Parse(xml, LoadOptions.None);
        if (document.Root is null)
        {
            return string.Empty;
        }

        NormalizeRelationshipIds(document.Root);
        NormalizeElement(document.Root);
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static void NormalizeRelationshipIds(XElement root)
    {
        if (!string.Equals(root.Name.LocalName, "Relationships", StringComparison.Ordinal))
        {
            return;
        }

        var nextId = 1;
        foreach (var relationship in root.Elements().Where(element => string.Equals(element.Name.LocalName, "Relationship", StringComparison.Ordinal)))
        {
            var idAttribute = relationship.Attribute("Id");
            if (idAttribute is not null)
            {
                idAttribute.Value = $"rId{nextId}";
                nextId++;
            }
        }
    }

    private static void NormalizeElement(XElement element)
    {
        var attributes = element.Attributes()
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .ToArray();

        element.ReplaceAttributes(attributes);

        foreach (var child in element.Elements())
        {
            NormalizeElement(child);
        }
    }

    private static string Serialize(PptxPackageSnapshot snapshot)
        => NormalizeLineEndings(JsonSerializer.Serialize(snapshot, JsonOptions));

    private static string NormalizeLineEndings(string text)
        => text.ReplaceLineEndings("\n");

    private static bool ShouldUpdateFixtures()
        => string.Equals(Environment.GetEnvironmentVariable("UPDATE_GOLDEN_PACKAGES"), "1", StringComparison.Ordinal);

    private static string GetFixturePath(string fixtureFileName)
    {
        var root = FindRepositoryRoot();
        return Path.Combine(root, "tests", "MarpToPptx.Tests", "Fixtures", fixtureFileName);
    }

    private static string FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            if (File.Exists(Path.Combine(current.FullName, "MarpToPptx.slnx")))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the repository root from the test output directory.");
    }

    private sealed record PptxPackageSnapshot(
        string[] Entries,
        IReadOnlyDictionary<string, string> NormalizedXmlParts);
}