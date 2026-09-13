using System.IO;
using System.Reflection;
using System.Security;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace TanssAddin.ManifestGenerator;

/// <summary>Ein Manifest, das so nicht ausgeliefert werden darf.</summary>
public sealed class ManifestError(string message) : Exception(message);

/// <summary>
/// Setzt die Vorlage zum Manifest zusammen und prueft das Ergebnis.
/// </summary>
/// <remarks>
/// Geprueft wird IMMER, nicht auf Wunsch. Es soll keinen Weg geben, ein Manifest mit
/// bekannter Schwaeche zu erzeugen: Es wird zentral verteilt und erreicht jeden Benutzer
/// auf einmal.
/// </remarks>
public static partial class Renderer
{
    private const string NoteBegin = "<!-- @vorlage:begin@ -->";
    private const string NoteEnd = "<!-- @vorlage:end@ -->";
    private const string NsApp = "http://schemas.microsoft.com/office/appforoffice/1.1";

    [GeneratedRegex(@"\$\{[a-z_]+\}")]
    private static partial Regex PlaceholderPattern { get; }

    /// <summary>Die Vorlage aus dem Programm selbst.</summary>
    public static string Template()
    {
        using var stream = Assembly.GetExecutingAssembly()
            .GetManifestResourceStream("manifest.xml.tmpl")
            ?? throw new ManifestError(
                "Die Manifestvorlage fehlt im Programm. Das ist ein Baufehler, kein "
                + "Bedienfehler - bitte das Werkzeug neu bauen.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>Erzeugt das Manifest.</summary>
    public static string Render(Options options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var text = Template();
        text = RemoveBlock(text, NoteBegin, NoteEnd);

        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["base_url"] = options.AssetBase,
            ["page_url"] = options.PageUrl,
            ["addin_id"] = AddinId.For(options.TanssApi).ToString(),
            ["version"] = options.Version,
            ["display_name"] = options.DisplayName,
            ["provider_name"] = options.ProviderName,
            ["support_url"] = (options.SupportUrl?.ToString() ?? options.AssetBase),
            ["tanss_frontend_origin"] = options.FrontendOrigin,
        };

        foreach (var (key, value) in values)
        {
            // Jeder eingesetzte Wert wird escaped. Ein Firmenname mit einem Und-Zeichen
            // ergaebe sonst eine Datei, die kein XML mehr ist - und der Fehler traete
            // erst beim Hochladen im Admin Center auf.
            text = text.Replace($"${{{key}}}", SecurityElement.Escape(value), StringComparison.Ordinal);
        }

        var leftover = PlaceholderPattern.Matches(text);
        if (leftover.Count > 0)
        {
            var names = string.Join(", ", leftover.Select(m => m.Value).Distinct());
            throw new ManifestError(
                $"Die Vorlage enthaelt Platzhalter, fuer die es keinen Wert gibt: {names}. "
                + "Vorlage und Werkzeug passen nicht zusammen.");
        }

        var problems = Check(text);
        if (problems.Count > 0)
        {
            throw new ManifestError(
                "Das erzeugte Manifest ist nicht auslieferbar:"
                + Environment.NewLine
                + string.Join(Environment.NewLine, problems.Select(p => $"  - {p}")));
        }

        return text.TrimStart();
    }

    /// <summary>Strukturpruefung eines fertigen Manifests.</summary>
    public static IReadOnlyList<string> Check(string xml)
    {
        var problems = new List<string>();

        XDocument document;
        try
        {
            document = XDocument.Parse(xml, LoadOptions.None);
        }
        catch (XmlException error)
        {
            return [$"Die Datei ist kein wohlgeformtes XML: {error.Message}"];
        }

        var root = document.Root;
        if (root is null || root.Name.LocalName != "OfficeApp")
        {
            return ["Das Wurzelelement heisst nicht OfficeApp."];
        }
        if (root.Name.NamespaceName != NsApp)
        {
            problems.Add($"Falscher Namensraum am Wurzelelement: {root.Name.NamespaceName}");
        }

        var id = Value(root, "Id");
        if (!Guid.TryParse(id, out _))
        {
            problems.Add($"Die Id \"{id}\" ist keine GUID.");
        }

        if (Value(root, "Version").Length == 0) problems.Add("Die Version fehlt.");
        if (Value(root, "ProviderName").Length == 0) problems.Add("Der Herausgeber fehlt.");

        // ReadWriteItem ist genau die Stufe, die dieses Add-in braucht: Es schreibt
        // Eigenschaften am geoeffneten Element. Eine hoehere Stufe kostet Zustimmung,
        // die niemand braucht - und der Benutzer liest sie im Zustimmungsdialog.
        var permissions = Value(root, "Permissions");
        if (permissions != "ReadWriteItem")
        {
            problems.Add($"Die Berechtigungsstufe ist \"{permissions}\", erwartet ist ReadWriteItem.");
        }

        // Jede Adresse ueber TLS. Eine einzige http-Adresse laesst Outlook den
        // betroffenen Teil stillschweigend nicht laden.
        foreach (var attribute in document.Descendants().Attributes("DefaultValue"))
        {
            var value = attribute.Value;
            if (value.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            {
                problems.Add($"Adresse ohne TLS: {value}");
            }
        }

        foreach (var domain in document.Descendants().Where(e => e.Name.LocalName == "AppDomain"))
        {
            if (domain.Value.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            {
                problems.Add($"AppDomain ohne TLS: {domain.Value}");
            }
        }

        // Die drei Schaltflaechen muessen auf drei VERSCHIEDENE Adressen zeigen. Waeren
        // sie gleich, oeffneten alle drei dieselbe Seite - ein Fehler, den man nur
        // bemerkt, wenn man alle drei ausprobiert.
        var urls = document.Descendants()
            .Where(e => e.Name.LocalName == "Url")
            .Select(e => e.Attribute("DefaultValue")?.Value ?? "")
            .ToList();
        if (urls.Count > 0 && urls.Distinct(StringComparer.Ordinal).Count() != urls.Count)
        {
            problems.Add("Zwei Schaltflaechen zeigen auf dieselbe Adresse.");
        }

        return problems;
    }

    private static string Value(XElement root, string localName) =>
        root.Elements().FirstOrDefault(e => e.Name.LocalName == localName)?.Value.Trim() ?? "";

    /// <summary>Entfernt einen markierten Bereich samt seiner Marken.</summary>
    private static string RemoveBlock(string text, string begin, string end)
    {
        var start = text.IndexOf(begin, StringComparison.Ordinal);
        var stop = text.IndexOf(end, StringComparison.Ordinal);
        if (start < 0 || stop < 0 || stop < start) return text;
        return text.Remove(start, stop - start + end.Length);
    }
}
