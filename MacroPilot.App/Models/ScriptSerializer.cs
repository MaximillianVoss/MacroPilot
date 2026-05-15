using System.Text.Json;
using System.Text.Json.Serialization;

namespace MacroPilot.App.Models;

public static class ScriptSerializer
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    public static async Task SaveAsync(MacroScript script, string path, CancellationToken cancellationToken = default)
    {
        await using FileStream stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, script, Options, cancellationToken);
    }

    public static async Task<MacroScript> LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        await using FileStream stream = File.OpenRead(path);
        MacroScript? script = await JsonSerializer.DeserializeAsync<MacroScript>(stream, Options, cancellationToken);
        return script ?? throw new InvalidDataException("Script file is empty or has an invalid format.");
    }
}
