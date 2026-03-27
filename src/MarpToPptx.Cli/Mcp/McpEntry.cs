using MarpToPptx.Cli.Mcp;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;

internal static class McpEntry
{
    public static async Task<int> RunAsync(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);

        builder.Logging.AddConsole(options =>
        {
            options.LogToStandardErrorThreshold = LogLevel.Trace;
        });

        builder.Services.AddMcpServer(options =>
        {
            options.ServerInfo = new Implementation
            {
                Name = "marp2pptx",
                Version = typeof(McpEntry).Assembly.GetName().Version?.ToString() ?? "0.0.0",
                Title = "MarpToPptx MCP Server",
                Description = "MCP server providing AI assistants with direct access to the MarpToPptx Markdown-to-PowerPoint pipeline.",
                WebsiteUrl = "https://github.com/jongalloway/MarpToPptx"
            };
        })
        .WithStdioServerTransport()
        .WithTools<MarpToPptxTools>();

        await builder.Build().RunAsync();
        return 0;
    }
}
