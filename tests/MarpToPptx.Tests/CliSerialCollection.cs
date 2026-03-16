namespace MarpToPptx.Tests;

/// <summary>
/// xUnit collection that disables parallel execution for CLI tests that redirect
/// <see cref="Console.Out"/> / <see cref="Console.Error"/> to capture output.
/// Placing all such tests in this collection prevents race conditions when multiple
/// tests run concurrently and modify the shared global console state.
/// </summary>
[CollectionDefinition(Name)]
public sealed class CliSerialCollection : ICollectionFixture<CliSerialCollection>
{
    public const string Name = "CLI serial tests";
}
