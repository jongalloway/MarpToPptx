using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;

return await ProgramEntry.RunAsync(args);

internal static class ProgramEntry
{
	public static Task<int> RunAsync(string[] args)
	{
		if (args.Length == 0 || args.Contains("-h", StringComparer.OrdinalIgnoreCase) || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
		{
			PrintUsage();
			return Task.FromResult(0);
		}

		string? inputPath = null;
		string? outputPath = null;
		string? templatePath = null;
		string? themeCssPath = null;
		var allowRemoteAssets = false;

		for (var index = 0; index < args.Length; index++)
		{
			var arg = args[index];
			switch (arg)
			{
				case "-o":
				case "--output":
					outputPath = RequireValue(args, ref index, arg);
					break;
				case "--template":
					templatePath = RequireValue(args, ref index, arg);
					break;
				case "--theme-css":
					themeCssPath = RequireValue(args, ref index, arg);
					break;
				case "--allow-remote-assets":
					allowRemoteAssets = true;
					break;
				default:
					if (arg.StartsWith('-'))
					{
						throw new ArgumentException($"Unknown option '{arg}'.");
					}

					inputPath ??= arg;
					break;
			}
		}

		if (string.IsNullOrWhiteSpace(inputPath))
		{
			throw new ArgumentException("An input Markdown file is required.");
		}

		inputPath = Path.GetFullPath(inputPath);
		if (!File.Exists(inputPath))
		{
			throw new FileNotFoundException("Input Markdown file was not found.", inputPath);
		}

		outputPath ??= Path.ChangeExtension(inputPath, ".pptx");
		outputPath = Path.GetFullPath(outputPath);

		var markdown = File.ReadAllText(inputPath);
		var themeCss = string.IsNullOrWhiteSpace(themeCssPath) ? null : File.ReadAllText(Path.GetFullPath(themeCssPath));

		var compiler = new MarpCompiler();
		var deck = compiler.Compile(markdown, inputPath, themeCss);

		var renderer = new OpenXmlPptxRenderer();
		renderer.Render(deck, outputPath, new PptxRenderOptions
		{
			TemplatePath = string.IsNullOrWhiteSpace(templatePath) ? null : Path.GetFullPath(templatePath),
			SourceDirectory = Path.GetDirectoryName(inputPath),
			AllowRemoteAssets = allowRemoteAssets,
		});

		Console.WriteLine($"Generated '{outputPath}'.");
		return Task.FromResult(0);
	}

	private static string RequireValue(string[] args, ref int index, string option)
	{
		if (index + 1 >= args.Length)
		{
			throw new ArgumentException($"Option '{option}' requires a value.");
		}

		index++;
		return args[index];
	}

	private static void PrintUsage()
	{
		Console.WriteLine("marp2pptx <input.md> [-o output.pptx] [--template theme.pptx] [--theme-css theme.css] [--allow-remote-assets]");
		Console.WriteLine();
		Console.WriteLine("Options:");
		Console.WriteLine("  -o, --output      Output .pptx path. Defaults to the input file name with a .pptx extension.");
		Console.WriteLine("  --template        Existing .pptx template to copy masters/themes from before rendering slides.");
		Console.WriteLine("  --theme-css       CSS file to parse for Marp-style theme values.");
		Console.WriteLine("  --allow-remote-assets  Enable HTTP/HTTPS image downloads during rendering.");
	}
}
