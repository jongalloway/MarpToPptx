using MarpToPptx.Core;
using MarpToPptx.Pptx.Contrast;
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
			string? contrastReportPath = null;
			var allowRemoteAssets = false;
			var warnLowContrast = false;

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
					case "--warn-low-contrast":
						warnLowContrast = true;
						break;
					case "--contrast-report":
						contrastReportPath = RequireValue(args, ref index, arg);
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
			if (!string.IsNullOrWhiteSpace(contrastReportPath))
			{
				contrastReportPath = Path.GetFullPath(contrastReportPath);
			}

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

			if (warnLowContrast || !string.IsNullOrWhiteSpace(contrastReportPath))
			{
				RunContrastAuditDiagnostics(outputPath, contrastReportPath);
			}

			return Task.FromResult(0);
		}
		catch (CliArgumentException ex)
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
			throw new CliArgumentException($"Option '{option}' requires a value.");
		}

		index++;
		return args[index];
	}

	private sealed class CliArgumentException : ArgumentException
	{
		public CliArgumentException(string message)
			: base(message)
		{
		}
	}

	private static void PrintUsage()
	{
		Console.WriteLine("marp2pptx <input.md> [-o output.pptx] [--template theme.pptx] [--theme-css theme.css] [--allow-remote-assets] [--warn-low-contrast] [--contrast-report report.txt]");
		Console.WriteLine();
		Console.WriteLine("Options:");
		Console.WriteLine("  -o, --output      Output .pptx path. Defaults to the input file name with a .pptx extension.");
		Console.WriteLine("  --template        Existing .pptx template to copy masters/themes from before rendering slides.");
		Console.WriteLine("  --theme-css       CSS file to parse for Marp-style theme values.");
		Console.WriteLine("  --allow-remote-assets  Enable HTTP/HTTPS image downloads during rendering.");
		Console.WriteLine("  --warn-low-contrast    Audit the generated deck and print warnings for low-contrast text/background pairs.");
		Console.WriteLine("  --contrast-report      Write the contrast audit summary to a text file. Implies a contrast audit run.");
	}

	private static void RunContrastAuditDiagnostics(string outputPath, string? contrastReportPath)
	{
		try
		{
			var auditor = new ContrastAuditor();
			var results = auditor.Audit(outputPath);
			var failures = results.Where(result => result.IsFailing)
				.OrderBy(result => result.SlideNumber)
				.ThenBy(result => result.ShapeContext, StringComparer.Ordinal)
				.ToArray();

			var reportLines = BuildContrastAuditReport(outputPath, results.Count, failures);
			if (!string.IsNullOrWhiteSpace(contrastReportPath))
			{
				var reportDirectory = Path.GetDirectoryName(contrastReportPath);
				if (!string.IsNullOrWhiteSpace(reportDirectory))
				{
					Directory.CreateDirectory(reportDirectory);
				}

				File.WriteAllLines(contrastReportPath, reportLines);
				Console.WriteLine($"Contrast audit report written to '{contrastReportPath}'.");
			}

			foreach (var line in reportLines)
			{
				if (failures.Length > 0)
				{
					Console.Error.WriteLine(line);
				}
				else
				{
					Console.WriteLine(line);
				}
			}
		}
		catch (Exception ex)
		{
			Console.Error.WriteLine($"Warning: Contrast audit could not be completed for '{outputPath}': {ex.Message}");
		}
	}

	private static string[] BuildContrastAuditReport(string outputPath, int resultCount, IReadOnlyList<ContrastAuditResult> failures)
	{
		if (failures.Count == 0)
		{
			if (resultCount == 0)
			{
				return
				[
					$"Contrast audit found no auditable color pairs in '{outputPath}'. Theme or inherited colors may not be resolvable from solid fills alone."
				];
			}

			return
			[
				$"Contrast audit passed for '{outputPath}'. {resultCount} color pair(s) checked."
			];
		}

		var lines = new List<string>
		{
			$"Warning: Contrast audit found {failures.Count} low-contrast text/background pair(s) in '{outputPath}'."
		};

		foreach (var failure in failures)
		{
			var textLabel = failure.IsLargeText ? "large text" : "normal text";
			lines.Add(
				$"  Slide {failure.SlideNumber} - {failure.ShapeContext}: " +
				$"#{failure.ForegroundColor} on #{failure.BackgroundColor} = {failure.ContrastRatio:F2}:1 " +
				$"(requires {failure.MinimumRequiredRatio:F1}:1 for {textLabel})");
		}

		return lines.ToArray();
	}
}
