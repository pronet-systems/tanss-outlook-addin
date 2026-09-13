namespace TanssAddin.ManifestGenerator;

/// <summary>Eine Angabe, die der Benutzer korrigieren kann - kein Programmfehler.</summary>
/// <remarks>
/// <paramref name="field"/> benennt das Feld, damit die Oberflaeche die Meldung dorthin
/// stellen kann, wo der Fehler steht, statt sie in einen Sammelkasten zu legen.
/// </remarks>
public sealed class OptionsError(string field, string message) : Exception(message)
{
    public string Field { get; } = field;
}

/// <summary>
/// Die Angaben, aus denen ein Manifest entsteht.
/// </summary>
/// <remarks>
/// Geprueft wird hier und nicht beim Rendern: Ein Manifest mit bekannter Schwaeche soll
/// gar nicht erst entstehen koennen. Es wird zentral verteilt und erreicht damit jeden
/// Benutzer auf einmal - eine Korrektur danach kostet einen erneuten Upload durch den
/// Administrator jedes betroffenen Mandanten.
/// </remarks>
public sealed record Options
{
    /// <summary>Wo die Dateien des Panes liegen.</summary>
    public required Uri AddinBase { get; init; }

    /// <summary>
    /// Die TANSS-API des Kunden, etwa <c>https://tanss.kunde.de/backend</c>.
    /// </summary>
    /// <remarks>
    /// Sie muss NICHT oeffentlich erreichbar sein. Der Hoster spricht nie mit TANSS;
    /// jeder Aufruf geht vom Browser des Technikers aus, und der steht im Firmennetz.
    /// </remarks>
    public required Uri TanssApi { get; init; }

    /// <summary>Die TANSS-Weboberflaeche - Ziel der Verweise "in TANSS oeffnen".</summary>
    public required Uri TanssFrontend { get; init; }

    public required string Version { get; init; }

    public string DisplayName { get; init; } = DefaultDisplayName;

    public string ProviderName { get; init; } = DefaultProviderName;

    public Uri? SupportUrl { get; init; }

    /// <summary>
    /// Anwendungs-Id der Entra-Registrierung, oder leer.
    /// </summary>
    /// <remarks>
    /// Sie wandert als Parameter in die Seitenadresse, nicht in einen Manifestblock: Das
    /// Pane holt sein Token ueber die verschachtelte Anmeldung, und die verlangt dafuer
    /// nichts im Manifest. Ohne Id setzt das Pane die Nachricht aus Bordmitteln zusammen -
    /// eine Rekonstruktion statt des Originals.
    /// </remarks>
    public string EntraClientId { get; init; } = "";

    /// <summary>
    /// Der Mandant, gegen den die Registrierung geprueft wird - Domaene oder Kennung.
    /// </summary>
    /// <remarks>
    /// Er steht NICHT im Manifest und aendert am Add-in nichts. Er wird ausschliesslich
    /// fuer den Prueflauf gebraucht: Ohne Mandanten weist der Anmeldedienst eine Anfrage
    /// ohne Browsersitzung ab, bevor er die Anwendungs-Id ueberhaupt ansieht - die
    /// Pruefung koennte dann nichts feststellen und muesste das auch so sagen.
    /// </remarks>
    public string TenantHint { get; init; } = "";

    /// <summary>
    /// Ob die TANSS-Adresse in die Seitenaufrufe geschrieben wird.
    /// </summary>
    /// <remarks>
    /// Das ist der Mechanismus, mit dem EINE Ablage beliebig viele Kunden bedient: Jeder
    /// bringt sein eigenes Manifest mit, und darin steht seine Instanz. Abschaltbar fuer
    /// den Fall, dass die Instanz ohnehin in der config.json neben den Dateien steht.
    /// </remarks>
    public bool EmbedTanssInUrl { get; init; } = true;

    public const string DefaultDisplayName = "TANSS";
    public const string DefaultProviderName = "ProNet Systems GmbH";
    public const string DefaultVersion = "1.0.0.0";

    /// <summary>Die Adresse der Seite, wie sie ins Manifest kommt.</summary>
    public string PageUrl
    {
        get
        {
            var basis = $"{AddinBase.GetLeftPart(UriPartial.Path).TrimEnd('/')}/";
            var teile = new List<string>();
            if (EmbedTanssInUrl)
            {
                teile.Add($"tanss={Uri.EscapeDataString(TanssApi.ToString().TrimEnd('/'))}");
            }
            if (EntraClientId.Length > 0)
            {
                teile.Add($"entra={Uri.EscapeDataString(EntraClientId)}");
            }
            return teile.Count == 0 ? basis : $"{basis}?{string.Join("&", teile)}";
        }
    }

    /// <summary>Die Basis fuer Symbole - sie darf niemals einen Parameter tragen.</summary>
    public string AssetBase => AddinBase.GetLeftPart(UriPartial.Path).TrimEnd('/');

    /// <summary>Herkunft der TANSS-Oberflaeche, wie AppDomains sie erwartet.</summary>
    public string FrontendOrigin => TanssFrontend.GetLeftPart(UriPartial.Authority);

    /// <summary>
    /// Baut die Angaben aus Formularfeldern. Wirft <see cref="OptionsError"/> mit dem
    /// Namen des Feldes, das nicht stimmt.
    /// </summary>
    public static Options Create(
        string addinBase,
        string tanssApi,
        string tanssFrontend,
        string version,
        string displayName,
        string providerName,
        string supportUrl,
        string entraClientId,
        string tenantHint,
        bool embedTanssInUrl)
    {
        var basis = RequireHttps(nameof(AddinBase), addinBase);
        var api = RequireHttps(nameof(TanssApi), tanssApi);

        var frontend = string.IsNullOrWhiteSpace(tanssFrontend)
            ? new Uri(api.GetLeftPart(UriPartial.Authority))
            : RequireHttps(nameof(TanssFrontend), tanssFrontend);

        var stand = string.IsNullOrWhiteSpace(version) ? DefaultVersion : version.Trim();
        RequireVersion(stand);

        var id = (entraClientId ?? "").Trim();
        if (id.Length > 0 && !Guid.TryParse(id, out _))
        {
            throw new OptionsError(nameof(EntraClientId),
                "Die Anwendungs-Id ist keine GUID. Sie steht in der Entra-Registrierung "
                + "unter \"Anwendungs-ID (Client)\". Leer lassen, wenn es keine gibt.");
        }

        return new Options
        {
            AddinBase = basis,
            TanssApi = api,
            TanssFrontend = frontend,
            Version = stand,
            DisplayName = Fallback(displayName, DefaultDisplayName),
            ProviderName = Fallback(providerName, DefaultProviderName),
            SupportUrl = string.IsNullOrWhiteSpace(supportUrl)
                ? null
                : RequireHttps(nameof(SupportUrl), supportUrl),
            EntraClientId = id,
            TenantHint = (tenantHint ?? "").Trim(),
            EmbedTanssInUrl = embedTanssInUrl,
        };
    }

    private static string Fallback(string value, string fallback) =>
        string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();

    /// <summary>
    /// Nur HTTPS. Outlook laedt ein Taskpane ausschliesslich ueber TLS, und der
    /// Techniker tippt sein TANSS-Kennwort hinein - ueber http liefe es im Klartext.
    /// Eine solche Adresse wird deshalb verweigert statt stillschweigend eingesetzt.
    /// </summary>
    private static Uri RequireHttps(string field, string raw)
    {
        var value = (raw ?? "").Trim();
        if (value.Length == 0)
        {
            throw new OptionsError(field, "Diese Angabe fehlt.");
        }
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            throw new OptionsError(field,
                $"\"{value}\" ist keine vollstaendige Adresse. Erwartet wird etwas wie "
                + "https://beispiel.de/pfad.");
        }
        if (uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new OptionsError(field,
                $"Nur https ist zulaessig, angegeben war \"{uri.Scheme}\". Outlook laedt "
                + "ein Taskpane ausschliesslich verschluesselt.");
        }
        return uri;
    }

    /// <summary>
    /// Bis zu vier Zahlen, jede im 16-Bit-Bereich, und mindestens 1.0.0.0.
    /// </summary>
    /// <remarks>
    /// Die Untergrenze ist keine Vorliebe: Microsofts Pruefdienst weist ein Manifest mit
    /// kleinerer Version ab. Und die Version muss von Erzeugung zu Erzeugung WACHSEN -
    /// bei gleicher Version uebergeht Microsoft die Aktualisierung stillschweigend, und
    /// die Benutzer behalten die alte Version. Das kann dieses Werkzeug nicht pruefen;
    /// es kennt die zuletzt ausgelieferte Version nicht.
    /// </remarks>
    private static void RequireVersion(string version)
    {
        var parts = version.Split('.');
        if (parts.Length is < 2 or > 4)
        {
            throw new OptionsError(nameof(Version),
                $"\"{version}\" braucht zwei bis vier durch Punkt getrennte Zahlen.");
        }

        foreach (var part in parts)
        {
            if (!int.TryParse(part, out var number) || number is < 0 or > 65535)
            {
                throw new OptionsError(nameof(Version),
                    $"\"{part}\" ist keine Zahl zwischen 0 und 65535.");
            }
        }

        if (!int.TryParse(parts[0], out var major) || major < 1)
        {
            throw new OptionsError(nameof(Version),
                $"\"{version}\" ist zu klein. Microsoft weist alles unter 1.0.0.0 ab.");
        }
    }
}
