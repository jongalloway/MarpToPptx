using MarpToPptx.Pptx.Extraction;

return ProgramEntry.Run(args);

internal static class ProgramEntry
{
	public static int Run(string[] args)
	{
		try
		{
			if (args.Length == 0 || args.Contains("-h", StringComparer.OrdinalIgnoreCase) || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
			{
				PrintUsage();
				return 0;
			}

			string? inputPath = null;
			string? outputPath = null;
			string? assetsDirectory = null;
			var includeNotes = true;

			for (var index = 0; index < args.Length; index++)
			{
				var arg = args[index];
				switch (arg)
				{
					case "-o":
					case "--output":
						outputPath = RequireValue(args, ref index, arg);
						break;
					case "--assets-dir":
						assetsDirectory = RequireValue(args, ref index, arg);
						break;
					case "--no-notes":
						includeNotes = false;
						break;
					default:
						if (arg.StartsWith('-'))
						{
							throw new CliArgumentException($"Unknown option '{arg}'.");
						}

						inputPath ??= arg;
						break;
				}
			}

			if (string.IsNullOrWhiteSpace(inputPath))
			{
				throw new CliArgumentException("An input PowerPoint file is required.");
			}

			inputPath = RequireExistingFile(inputPath, "Input PowerPoint");
			outputPath ??= Path.ChangeExtension(inputPath, ".md");
			outputPath = Path.GetFullPath(outputPath);

			assetsDirectory = string.IsNullOrWhiteSpace(assetsDirectory)
				? Path.Combine(Path.GetDirectoryName(outputPath)!, Path.GetFileNameWithoutExtension(outputPath) + ".assets")
				: Path.GetFullPath(assetsDirectory);

			var assetPathPrefix = Path.GetRelativePath(Path.GetDirectoryName(outputPath)!, assetsDirectory)
				.Replace('\\', '/');
			if (string.IsNullOrWhiteSpace(assetPathPrefix))
			{
				assetPathPrefix = Path.GetFileName(assetsDirectory);
			}

			var exporter = new PptxMarkdownExporter();
			exporter.Export(inputPath, outputPath, new PptxMarkdownExportOptions
			{
				AssetsDirectory = assetsDirectory,
				AssetPathPrefix = assetPathPrefix,
				IncludeNotes = includeNotes,
			});

			Console.WriteLine($"Extracted '{outputPath}'.");
			return 0;
		}
		catch (CliArgumentException ex)
		{
			Console.Error.WriteLine($"Error: {ex.Message}");
			Console.Error.WriteLine();
			PrintUsage();
			return 1;
		}
		catch (FileNotFoundException ex)
		{
			Console.Error.WriteLine($"Error: {ex.Message}");
			if (!string.IsNullOrWhiteSpace(ex.FileName))
			{
				Console.Error.WriteLine($"Path: {ex.FileName}");
			}

			return 1;
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
			throw new CliArgumentException($"Option '{option}' requires a value.");
		}

		index++;
		return args[index];
	}

	private static void PrintUsage()
	{
		Console.WriteLine("PptxToMarp <input.pptx> [-o output.md] [--assets-dir path]");
		Console.WriteLine();
		Console.WriteLine("Options:");
		Console.WriteLine("  -o, --output      Output .md path. Defaults to the input file name with a .md extension.");
		Console.WriteLine("  --assets-dir      Directory to write extracted images. Defaults to <output>.assets.");
		Console.WriteLine("  --no-notes        Do not include speaker notes in the extracted Markdown.");
	}

	private sealed class CliArgumentException(string message) : ArgumentException(message);
}
