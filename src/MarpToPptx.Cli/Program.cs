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
			string? existingDeckPath = null;
			var updateExisting = false;
			var allowRemoteAssets = false;
			var contrastWarningMode = ContrastWarningMode.Off;

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
					case "--update-existing":
						// Update the existing output deck in place rather than rebuilding it
						// from scratch. When set without an argument the output path is used.
						// Accepts an optional explicit path: --update-existing [path]
						updateExisting = true;
						if (index + 1 < args.Length && !args[index + 1].StartsWith('-'))
						{
							index++;
							existingDeckPath = args[index];
						}
						break;
					case "--contrast-warnings":
						contrastWarningMode = ParseContrastWarningMode(RequireValue(args, ref index, arg));
						break;
					case "--warn-low-contrast":
						contrastWarningMode = ContrastWarningMode.Detailed;
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

			// Resolve the existing deck path for update mode.
			// When --update-existing was specified without an explicit path, use the output path.
			if (updateExisting)
			{
				existingDeckPath = string.IsNullOrWhiteSpace(existingDeckPath)
					? outputPath
					: Path.GetFullPath(existingDeckPath);
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
				ExistingDeckPath = existingDeckPath,
			});

			var updateMode = !string.IsNullOrEmpty(existingDeckPath) && File.Exists(existingDeckPath);
			Console.WriteLine(updateMode
				? $"Updated '{outputPath}'."
				: $"Generated '{outputPath}'.");

			if (contrastWarningMode != ContrastWarningMode.Off || !string.IsNullOrWhiteSpace(contrastReportPath))
			{
				RunContrastAuditDiagnostics(outputPath, contrastWarningMode, contrastReportPath);
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
		Console.WriteLine("marp2pptx <input.md> [-o output.pptx] [--template theme.pptx] [--theme-css theme.css] [--allow-remote-assets] [--update-existing [path]] [--contrast-warnings off|summary|detailed] [--contrast-report report.txt]");
		Console.WriteLine();
		Console.WriteLine("Options:");
		Console.WriteLine("  -o, --output      Output .pptx path. Defaults to the input file name with a .pptx extension.");
		Console.WriteLine("  --template        Existing .pptx template to copy masters/themes from before rendering slides.");
		Console.WriteLine("  --theme-css       CSS file to parse for Marp-style theme values.");
		Console.WriteLine("  --allow-remote-assets  Enable HTTP/HTTPS image downloads during rendering.");
		Console.WriteLine("  --update-existing [path]  Update an existing MarpToPptx-generated deck instead of rebuilding from");
		Console.WriteLine("                    scratch. Preserves manually added slides. Optionally specify a source path;");
		Console.WriteLine("                    defaults to the output file when omitted.");
		Console.WriteLine("  --contrast-warnings  Contrast warning mode: off, summary, or detailed.");
		Console.WriteLine("  --warn-low-contrast  Backward-compatible alias for '--contrast-warnings detailed'.");
		Console.WriteLine("  --contrast-report    Write a detailed contrast audit report to a text file. Implies a contrast audit run.");
	}

	private static ContrastWarningMode ParseContrastWarningMode(string value)
	{
		return value.ToLowerInvariant() switch
		{
			"off" => ContrastWarningMode.Off,
			"summary" => ContrastWarningMode.Summary,
			"detailed" => ContrastWarningMode.Detailed,
			_ => throw new CliArgumentException("Option '--contrast-warnings' requires one of: off, summary, detailed.")
		};
	}

	private static void RunContrastAuditDiagnostics(string outputPath, ContrastWarningMode contrastWarningMode, string? contrastReportPath)
	{
		try
		{
			var auditor = new ContrastAuditor();
			var results = auditor.Audit(outputPath);
			var failures = results.Where(result => result.IsFailing)
				.OrderBy(result => result.SlideNumber)
				.ThenBy(result => result.ShapeContext, StringComparer.Ordinal)
				.ToArray();

			var reportLines = BuildDetailedContrastAuditReport(outputPath, results.Count, failures);
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

			var consoleLines = BuildConsoleContrastAuditReport(outputPath, results.Count, failures, contrastWarningMode);
			foreach (var line in consoleLines)
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

	private static string[] BuildConsoleContrastAuditReport(string outputPath, int resultCount, IReadOnlyList<ContrastAuditResult> failures, ContrastWarningMode contrastWarningMode)
	{
		if (contrastWarningMode == ContrastWarningMode.Off)
		{
			return [];
		}

		if (contrastWarningMode == ContrastWarningMode.Summary)
		{
			if (failures.Count == 0)
			{
				return resultCount == 0
					? [$"Contrast audit found no auditable color pairs in '{outputPath}'."]
					: [$"Contrast audit passed for '{outputPath}'. {resultCount} color pair(s) checked."];
			}

			var slideList = string.Join(", ", failures.Select(failure => failure.SlideNumber).Distinct().OrderBy(slideNumber => slideNumber));
			return [$"Warning: Slides {slideList} may have low-contrast accessibility issues."];
		}

		return BuildDetailedContrastAuditReport(outputPath, resultCount, failures);
	}

	private static string[] BuildDetailedContrastAuditReport(string outputPath, int resultCount, IReadOnlyList<ContrastAuditResult> failures)
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

	private enum ContrastWarningMode
	{
		Off,
		Summary,
		Detailed,
	}
}
