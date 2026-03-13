using MarpToPptx.Core;
using MarpToPptx.Pptx.Rendering;

return await ProgramEntry.RunAsync(args);

internal static class ProgramEntry
{
	public static Task<int> RunAsync(string[] args)
	{
		try
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
						{
							if (arg.StartsWith('-'))
							{
								throw new ArgumentException($"Unknown option '{arg}'.");
							}

							inputPath ??= arg;
							break;
						}
				}
			}

			if (string.IsNullOrWhiteSpace(inputPath))
			{
				throw new ArgumentException("An input Markdown file is required.");
			}

			inputPath = RequireExistingFile(inputPath, "Input Markdown");
			if (!string.IsNullOrWhiteSpace(themeCssPath))
			{
				themeCssPath = RequireExistingFile(themeCssPath, "Theme CSS");
			}

			if (!string.IsNullOrWhiteSpace(templatePath))
			{
				templatePath = RequireExistingFile(templatePath, "Template PPTX");
			}

			outputPath ??= Path.ChangeExtension(inputPath, ".pptx");
			outputPath = Path.GetFullPath(outputPath);

			var markdown = File.ReadAllText(inputPath);
			var themeCss = string.IsNullOrWhiteSpace(themeCssPath) ? null : File.ReadAllText(themeCssPath);

			var compiler = new MarpCompiler();
			var deck = compiler.Compile(markdown, inputPath, themeCss);

			var renderer = new OpenXmlPptxRenderer();
			renderer.Render(deck, outputPath, new PptxRenderOptions
			{
				TemplatePath = templatePath,
				SourceDirectory = Path.GetDirectoryName(inputPath),
				AllowRemoteAssets = allowRemoteAssets,
			});

			Console.WriteLine($"Generated '{outputPath}'.");
			return Task.FromResult(0);
		}
		catch (ArgumentException ex)
		{
			Console.Error.WriteLine($"Error: {ex.Message}");
			Console.Error.WriteLine();
			PrintUsage();
			return Task.FromResult(1);
		}
		catch (FileNotFoundException ex)
		{
			Console.Error.WriteLine($"Error: {ex.Message}");
			if (!string.IsNullOrWhiteSpace(ex.FileName))
			{
				Console.Error.WriteLine($"Path: {ex.FileName}");
			}

			return Task.FromResult(1);
		}
	}

	private static string RequireExistingFile(string path, string description)
	{
		var fullPath = Path.GetFullPath(path);
		if (!File.Exists(fullPath))
		{
			throw new FileNotFoundException($"{description} file was not found.", fullPath);
		}

		return fullPath;
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
