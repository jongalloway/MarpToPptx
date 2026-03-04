namespace MarpToPptx.Core.Themes;

public sealed record ThemeDefinition
{
    public static ThemeDefinition Default => new();

    public string Name { get; init; } = "default";

    public string FontFamily { get; init; } = "Aptos";

    public string MonospaceFontFamily { get; init; } = "Cascadia Mono";

    public string TextColor { get; init; } = "#1F2937";

    public string BackgroundColor { get; init; } = "#FFFFFF";

    public string AccentColor { get; init; } = "#0F766E";

    public string? BackgroundImage { get; init; } = null;

    public string? BackgroundSize { get; init; } = null;

    public string? BackgroundPosition { get; init; } = null;

    public Spacing SlidePadding { get; init; } = new(48, 60, 48, 60);

    public TextStyle Body { get; init; } = new(24, "#1F2937", "Aptos", false);

    public TextStyle Code { get; init; } = new(18, "#F8FAFC", "Cascadia Mono", false, "#0F172A");

    public IReadOnlyDictionary<int, TextStyle> Headings { get; init; } = new Dictionary<int, TextStyle>
    {
        [1] = new(30, "#0F172A", "Aptos Display", true),
        [2] = new(26, "#0F172A", "Aptos Display", true),
        [3] = new(22, "#0F172A", "Aptos", true),
        [4] = new(20, "#0F172A", "Aptos", true),
        [5] = new(18, "#0F172A", "Aptos", true),
        [6] = new(16, "#0F172A", "Aptos", true),
    };
}

public sealed record TextStyle(double FontSize, string Color, string FontFamily, bool Bold, string? BackgroundColor = null, double? LineHeight = null, double? LetterSpacing = null, string? TextTransform = null);

public sealed record Spacing(double Top, double Right, double Bottom, double Left)
{
    public static Spacing Uniform(double value) => new(value, value, value, value);
}
