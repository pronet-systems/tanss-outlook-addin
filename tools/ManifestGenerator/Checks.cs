using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;

namespace TanssAddin.ManifestGenerator;

/// <summary>Wie eine Pruefung ausgegangen ist.</summary>
public enum Verdict
{
    /// <summary>Gemessen und in Ordnung.</summary>
    Ok,

    /// <summary>
    /// Gemessen, aber nicht eindeutig - oder eindeutig und nicht schlimm.
    /// </summary>
    /// <remarks>
    /// Dieser Ausgang ist wichtig genug fuer einen eigenen Wert: Eine Pruefung, die
    /// keine Antwort bekommen hat, darf nicht als bestanden durchgehen. "Unklar" ist
    /// eine ehrliche Aussage, "in Ordnung" waere eine erfundene.
    /// </remarks>
    Unclear,

    /// <summary>Gemessen, und so geht das Add-in nicht in Betrieb.</summary>
    Fail,
}

/// <summary>Das Ergebnis einer einzelnen Pruefung.</summary>
/// <param name="Area">Womit sie zu tun hat - fuer die Gruppierung in der Anzeige.</param>
/// <param name="Name">Was geprueft wurde, in einem halben Satz.</param>
/// <param name="Verdict">Wie es ausging.</param>
/// <param name="Detail">Was gemessen wurde und was daraus folgt.</param>
public sealed record CheckResult(string Area, string Name, Verdict Verdict, string Detail);

/// <summary>
/// Prueft die Angaben gegen die Wirklichkeit, bevor ein Manifest an alle Benutzer geht.
/// </summary>
/// <remarks>
/// <para>
/// Jede dieser Pruefungen deckt einen Fehler ab, der ohne sie erst beim Techniker
/// auffaellt - und dort als "das Add-in geht nicht" ankommt, ohne Hinweis darauf, woran
/// es liegt. Ein leeres Taskpane sagt nicht, dass eine Kopfzeile fehlt.
/// </para>
/// <para>
/// Alle Pruefungen sind LESEND und ohne Anmeldung. Es werden keine Zugangsdaten
/// verlangt, nichts veraendert und nichts angelegt.
/// </para>
/// </remarks>
public static class Checks
{
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(15);

    /// <summary>Wie nah ein Zertifikatsablauf sein darf, bevor er gemeldet wird.</summary>
    private static readonly TimeSpan CertWarning = TimeSpan.FromDays(21);

    /// <summary>
    /// Laeuft alle Pruefungen und liefert sie einzeln, sobald sie fertig sind.
    /// </summary>
    public static async IAsyncEnumerable<CheckResult> RunAsync(
        Options options,
        [System.Runtime.CompilerServices.EnumeratorCancellation]
        CancellationToken cancel = default)
    {
        ArgumentNullException.ThrowIfNull(options);

        foreach (var result in Local(options))
        {
            yield return result;
        }

        yield return await Reachable(options.AddinBase, "Ablage", "Die Dateien des Panes", cancel);
        yield return await FramingAllowed(options, cancel);
        yield return await IconReachable(options, cancel);
        yield return await Certificate(options.AddinBase, "Ablage", cancel);

        yield return await TanssReachable(options, cancel);
        yield return await CorsAllowed(options, cancel);
        yield return await Certificate(options.TanssApi, "TANSS", cancel);

        if (options.EntraClientId.Length > 0)
        {
            yield return await EntraAppExists(options, cancel);
            yield return await EntraRedirectRegistered(options, cancel);
        }
        else
        {
            yield return new CheckResult("Anmeldung", "Microsoft-Anmeldung",
                Verdict.Unclear,
                "Keine Anwendungs-Id angegeben. Das Pane setzt die E-Mail dann aus "
                + "Office-Bordmitteln zusammen - eine Rekonstruktion statt des Originals: "
                + "Empfangskette und Signaturkoepfe fehlen.");
        }
    }

    /// <summary>Pruefungen, die ohne Netz auskommen.</summary>
    private static IEnumerable<CheckResult> Local(Options options)
    {
        string? xml = null;
        string? error = null;
        try
        {
            xml = Renderer.Render(options);
        }
        catch (ManifestError failure)
        {
            error = failure.Message;
        }

        yield return xml is not null
            ? new CheckResult("Manifest", "Struktur des Manifests", Verdict.Ok,
                $"Wohlgeformt, Berechtigungsstufe ReadWriteItem, alle Adressen ueber TLS. "
                + $"Kennung {AddinId.For(options.TanssApi)}.")
            : new CheckResult("Manifest", "Struktur des Manifests", Verdict.Fail, error ?? "");

        // Ein Manifest, dessen Ablage und TANSS-Instanz dieselbe Herkunft haben, ist
        // zulaessig - aber dann teilen sich Pane und TANSS-Oberflaeche den Speicher, in
        // dem das Anmeldetoken liegt.
        if (string.Equals(options.AddinBase.Host, options.TanssApi.Host,
            StringComparison.OrdinalIgnoreCase))
        {
            yield return new CheckResult("Manifest", "Getrennte Herkunft", Verdict.Unclear,
                "Pane und TANSS liegen unter demselben Namen. Das funktioniert, aber "
                + "Skripte der TANSS-Oberflaeche koennen dann das Anmeldetoken des "
                + "Technikers lesen - beides liegt im selben Speicher.");
        }
        else
        {
            yield return new CheckResult("Manifest", "Getrennte Herkunft", Verdict.Ok,
                $"Pane auf {options.AddinBase.Host}, TANSS auf {options.TanssApi.Host}. "
                + "Speicher und Inhaltsrichtlinie sind damit getrennt.");
        }
    }

    /* ------------------------------------------------------------------ Ablage */

    private static async Task<CheckResult> Reachable(Uri target, string area, string what,
        CancellationToken cancel)
    {
        using var client = Client();
        try
        {
            using var response = await client.GetAsync(target, cancel);
            return response.IsSuccessStatusCode
                ? new CheckResult(area, $"{what} sind erreichbar", Verdict.Ok,
                    $"HTTP {(int)response.StatusCode} von {target}")
                : new CheckResult(area, $"{what} sind erreichbar", Verdict.Fail,
                    $"HTTP {(int)response.StatusCode} von {target}. Erwartet wird eine "
                    + "Erfolgsantwort - liegen die Dateien schon dort?");
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return new CheckResult(area, $"{what} sind erreichbar", Verdict.Fail,
                $"Keine Antwort von {target}: {Short(error)}");
        }
    }

    /// <summary>
    /// Die Kopfzeile, an der das Add-in am haeufigsten scheitert.
    /// </summary>
    /// <remarks>
    /// Outlook zeigt das Taskpane in einem Rahmen FREMDER Herkunft an. Sendet der
    /// Webserver <c>X-Frame-Options</c> - gleich welchen Wertes -, bleibt das Pane weiss,
    /// und zwar ohne Fehlermeldung, die irgendwo auftauchte. Der Techniker sieht einen
    /// leeren Bereich und meldet "das Add-in geht nicht".
    /// </remarks>
    private static async Task<CheckResult> FramingAllowed(Options options, CancellationToken cancel)
    {
        const string name = "Einbettung in Outlook erlaubt";
        using var client = Client();
        try
        {
            using var response = await client.GetAsync(options.AddinBase, cancel);

            if (response.Headers.TryGetValues("X-Frame-Options", out var values))
            {
                return new CheckResult("Ablage", name, Verdict.Fail,
                    $"Der Webserver sendet X-Frame-Options: {string.Join(", ", values)}. "
                    + "Damit bleibt das Taskpane in Outlook weiss - ohne Fehlermeldung. "
                    + "Die Kopfzeile muss fuer diese Ablage entfernt werden.");
            }

            var csp = response.Headers.TryGetValues("Content-Security-Policy", out var policy)
                ? string.Join(" ", policy)
                : "";
            if (csp.Contains("frame-ancestors", StringComparison.OrdinalIgnoreCase))
            {
                var office = csp.Contains("office.com", StringComparison.OrdinalIgnoreCase)
                    || csp.Contains("outlook.com", StringComparison.OrdinalIgnoreCase);
                return office
                    ? new CheckResult("Ablage", name, Verdict.Ok,
                        "Kein X-Frame-Options, und frame-ancestors nennt die Office-Herkuenfte.")
                    : new CheckResult("Ablage", name, Verdict.Fail,
                        "frame-ancestors ist gesetzt, nennt aber keine Office-Herkunft. "
                        + "Damit bleibt das Taskpane in Outlook weiss.");
            }

            return new CheckResult("Ablage", name, Verdict.Ok,
                "Kein X-Frame-Options. Das Pane laesst sich in Outlook anzeigen. "
                + "Eine frame-ancestors-Richtlinie ist nicht gesetzt - die Einbettung ist "
                + "damit auch fremden Seiten erlaubt.");
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return new CheckResult("Ablage", name, Verdict.Unclear,
                $"Nicht pruefbar, die Ablage antwortet nicht: {Short(error)}");
        }
    }

    private static async Task<CheckResult> IconReachable(Options options, CancellationToken cancel)
    {
        const string name = "Symbole sind erreichbar";
        var icon = new Uri($"{options.AssetBase}/assets/icon-64.png");
        using var client = Client();
        try
        {
            using var response = await client.GetAsync(icon, cancel);
            if (!response.IsSuccessStatusCode)
            {
                return new CheckResult("Ablage", name, Verdict.Fail,
                    $"HTTP {(int)response.StatusCode} fuer {icon}. Ohne Symbole zeigt "
                    + "Outlook die Schaltflaechen mit einem Platzhalterbild.");
            }
            var type = response.Content.Headers.ContentType?.MediaType ?? "";
            return type.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
                ? new CheckResult("Ablage", name, Verdict.Ok, $"{icon} liefert {type}.")
                : new CheckResult("Ablage", name, Verdict.Fail,
                    $"{icon} liefert \"{type}\" statt eines Bildes.");
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return new CheckResult("Ablage", name, Verdict.Unclear,
                $"Nicht pruefbar: {Short(error)}");
        }
    }

    /* ------------------------------------------------------------------- TANSS */

    /// <summary>
    /// Antwortet unter der angegebenen Adresse wirklich eine TANSS-API?
    /// </summary>
    /// <remarks>
    /// Ohne Token ist die erwartete Antwort eine Abweisung - und genau die ist der
    /// Nachweis: Sie zeigt, dass dort etwas steht, das Anmeldung verlangt. Eine
    /// Erfolgsantwort waere verdaechtig, ein HTTP 404 hiesse, der Pfad stimmt nicht.
    /// </remarks>
    private static async Task<CheckResult> TanssReachable(Options options, CancellationToken cancel)
    {
        const string name = "TANSS-API antwortet";
        var probe = new Uri($"{options.TanssApi.ToString().TrimEnd('/')}/api/v1/employees/technicians");
        using var client = Client();
        try
        {
            using var response = await client.GetAsync(probe, cancel);
            var status = (int)response.StatusCode;
            return status switch
            {
                401 or 403 => new CheckResult("TANSS", name, Verdict.Ok,
                    $"HTTP {status} ohne Token - so antwortet die API. Adresse und Pfad stimmen."),
                404 => new CheckResult("TANSS", name, Verdict.Fail,
                    $"HTTP 404 fuer {probe}. Unter dieser Adresse liegt keine TANSS-API - "
                    + "fehlt der Pfad, meist /backend?"),
                >= 200 and < 300 => new CheckResult("TANSS", name, Verdict.Unclear,
                    $"HTTP {status} OHNE Token. Das ist unerwartet: Die API sollte eine "
                    + "Anmeldung verlangen."),
                _ => new CheckResult("TANSS", name, Verdict.Unclear,
                    $"HTTP {status} von {probe}."),
            };
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return new CheckResult("TANSS", name, Verdict.Fail,
                $"Keine Antwort von {probe}: {Short(error)}. Das ist nur dann in Ordnung, "
                + "wenn dieser Rechner im Netz der TANSS-Instanz steht - die Arbeitsplaetze "
                + "der Techniker muessen sie in jedem Fall erreichen.");
        }
    }

    /// <summary>
    /// Laesst TANSS den fremden Ursprung des Panes zu?
    /// </summary>
    /// <remarks>
    /// Liegt das Pane unter einem anderen Namen als die API - der Regelfall -, fragt der
    /// Browser vorher nach. Fehlt in dieser Antwort die Erlaubnis fuer den Kopf
    /// <c>apiToken</c>, scheitert jeder Aufruf, und der Browser meldet es nur in seiner
    /// eigenen Konsole: Im Pane steht dann "nicht erreichbar".
    /// </remarks>
    private static async Task<CheckResult> CorsAllowed(Options options, CancellationToken cancel)
    {
        const string name = "Fremder Ursprung zugelassen";
        var probe = new Uri($"{options.TanssApi.ToString().TrimEnd('/')}/api/v1/employees/technicians");
        var origin = options.AddinBase.GetLeftPart(UriPartial.Authority);

        using var client = Client();
        using var request = new HttpRequestMessage(HttpMethod.Options, probe);
        request.Headers.TryAddWithoutValidation("Origin", origin);
        request.Headers.TryAddWithoutValidation("Access-Control-Request-Method", "GET");
        request.Headers.TryAddWithoutValidation("Access-Control-Request-Headers", "apitoken");

        try
        {
            using var response = await client.SendAsync(request, cancel);

            var allowOrigin = Header(response, "Access-Control-Allow-Origin");
            var allowHeaders = Header(response, "Access-Control-Allow-Headers");

            if (allowOrigin.Length == 0)
            {
                return new CheckResult("TANSS", name, Verdict.Fail,
                    $"Die Vorabfrage beantwortet TANSS mit HTTP {(int)response.StatusCode}, "
                    + "aber ohne Access-Control-Allow-Origin. Ein Pane unter einem anderen "
                    + "Namen kann die API damit nicht aufrufen.");
            }

            var originOk = allowOrigin == "*"
                || allowOrigin.Contains(origin, StringComparison.OrdinalIgnoreCase);
            var tokenOk = allowHeaders.Contains("apitoken", StringComparison.OrdinalIgnoreCase);

            if (originOk && tokenOk)
            {
                return new CheckResult("TANSS", name, Verdict.Ok,
                    $"Allow-Origin: {allowOrigin}; der Kopf apiToken ist zugelassen.");
            }
            if (!tokenOk)
            {
                return new CheckResult("TANSS", name, Verdict.Fail,
                    $"Der Kopf apiToken fehlt in Access-Control-Allow-Headers "
                    + $"({allowHeaders}). Ohne ihn wird jeder Aufruf vom Browser "
                    + "abgewiesen, bevor er TANSS erreicht.");
            }
            return new CheckResult("TANSS", name, Verdict.Fail,
                $"Access-Control-Allow-Origin lautet \"{allowOrigin}\" und deckt "
                + $"\"{origin}\" nicht ab.");
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return new CheckResult("TANSS", name, Verdict.Unclear,
                $"Die Vorabfrage blieb unbeantwortet: {Short(error)}");
        }
    }

    /* -------------------------------------------------------------- Zertifikat */

    /// <summary>
    /// Wann laeuft das Zertifikat ab?
    /// </summary>
    /// <remarks>
    /// Ein abgelaufenes Zertifikat nimmt das Add-in fuer ALLE Benutzer gleichzeitig
    /// ausser Betrieb, und niemand kann es an seinem Arbeitsplatz umgehen. Deshalb ist
    /// der Ablauf hier eine eigene Pruefung und keine Fussnote.
    /// </remarks>
    private static async Task<CheckResult> Certificate(Uri target, string area,
        CancellationToken cancel)
    {
        var name = $"Zertifikat {target.Host}";
        X509Certificate2? seen = null;

        using var handler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (_, certificate, _, errors) =>
            {
                if (certificate is not null) seen = new X509Certificate2(certificate);
                // Nicht durchwinken: Ein Zertifikat, dem dieser Rechner nicht traut,
                // traut auch der Arbeitsplatz des Technikers nicht.
                return errors == System.Net.Security.SslPolicyErrors.None;
            },
        };
        using var client = new HttpClient(handler) { Timeout = Timeout };

        try
        {
            using var response = await client.GetAsync(
                new Uri(target.GetLeftPart(UriPartial.Authority)), cancel);
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            if (seen is null)
            {
                return new CheckResult(area, name, Verdict.Unclear,
                    $"Nicht pruefbar: {Short(error)}");
            }
            return new CheckResult(area, name, Verdict.Fail,
                $"Das Zertifikat wird von diesem Rechner nicht anerkannt "
                + $"(Aussteller: {seen.Issuer}). Dem Arbeitsplatz des Technikers ginge es "
                + "genauso, und das Pane bliebe leer.");
        }

        if (seen is null)
        {
            return new CheckResult(area, name, Verdict.Unclear, "Kein Zertifikat gesehen.");
        }

        var rest = seen.NotAfter - DateTime.Now;
        var detail = $"Aussteller {Issuer(seen)}, gueltig bis {seen.NotAfter:dd.MM.yyyy} "
            + $"({(int)rest.TotalDays} Tage).";
        return rest < CertWarning
            ? new CheckResult(area, name, Verdict.Unclear,
                detail + " Laeuft es ab, faellt das Add-in fuer alle Benutzer gleichzeitig aus.")
            : new CheckResult(area, name, Verdict.Ok, detail);
    }

    private static string Issuer(X509Certificate2 certificate)
    {
        var parts = certificate.Issuer.Split(',', StringSplitOptions.TrimEntries);
        var organisation = parts.FirstOrDefault(p => p.StartsWith("O=", StringComparison.Ordinal));
        return organisation?[2..] ?? certificate.Issuer;
    }

    /* ---------------------------------------------------------------- Anmeldung */

    /// <summary>
    /// Gibt es die Entra-Anwendung ueberhaupt?
    /// </summary>
    /// <remarks>
    /// Gefragt wird der Anmeldedienst mit einer absichtlich NICHT eingetragenen
    /// Rueckadresse. Die Antwort unterscheidet die beiden Faelle sauber: Meldet er, dass
    /// er die Anwendung nicht kennt, ist die Id falsch. Beanstandet er die Rueckadresse,
    /// hat er die Anwendung gefunden - und genau das wollten wir wissen.
    /// </remarks>
    private static async Task<CheckResult> EntraAppExists(Options options, CancellationToken cancel)
    {
        const string name = "Entra-Anwendung vorhanden";
        var answer = await AskEntra(options.EntraClientId,
            "https://localhost/nicht-eingetragen", cancel);

        return answer.Code switch
        {
            "AADSTS700016" => new CheckResult("Anmeldung", name, Verdict.Fail,
                $"Der Anmeldedienst kennt die Anwendungs-Id {options.EntraClientId} nicht. "
                + "Bitte in der Entra-Registrierung unter \"Anwendungs-ID (Client)\" nachsehen."),
            "AADSTS50011" => new CheckResult("Anmeldung", name, Verdict.Ok,
                "Die Anwendung ist eingetragen."),
            "" => new CheckResult("Anmeldung", name, Verdict.Unclear,
                $"Keine eindeutige Antwort vom Anmeldedienst. {answer.Detail}"),
            _ => new CheckResult("Anmeldung", name, Verdict.Unclear,
                $"Der Anmeldedienst antwortete mit {answer.Code}. Die Anwendung scheint zu "
                + "existieren, die Antwort ist aber nicht eindeutig."),
        };
    }

    /// <summary>
    /// Ist die Rueckadresse eingetragen, die ein Add-in braucht?
    /// </summary>
    /// <remarks>
    /// Die Anmeldung aus einem Taskpane heraus laeuft nicht ueber eine gewoehnliche
    /// Webadresse, sondern ueber eine eigene, die den Office-Wirt benennt. Fehlt sie in
    /// der Registrierung, scheitert die Anmeldung erst beim Techniker - und die Meldung
    /// dort nennt nicht die Ursache.
    /// </remarks>
    private static async Task<CheckResult> EntraRedirectRegistered(Options options,
        CancellationToken cancel)
    {
        const string name = "Rueckadresse des Panes eingetragen";
        var redirect = $"brk-multihub://{options.AddinBase.Host}";
        var answer = await AskEntra(options.EntraClientId, redirect, cancel);

        return answer.Code switch
        {
            "AADSTS50011" => new CheckResult("Anmeldung", name, Verdict.Fail,
                $"\"{redirect}\" ist in der Entra-Registrierung nicht eingetragen. Sie "
                + "gehoert dort unter die Plattform \"Einseitige Anwendung (SPA)\"."),
            "AADSTS700016" => new CheckResult("Anmeldung", name, Verdict.Fail,
                "Die Anwendungs-Id ist unbekannt - siehe vorige Pruefung."),
            "" => new CheckResult("Anmeldung", name, Verdict.Ok,
                $"\"{redirect}\" wird nicht beanstandet."),
            _ => new CheckResult("Anmeldung", name, Verdict.Unclear,
                $"Der Anmeldedienst antwortete mit {answer.Code}; die Rueckadresse wurde "
                + "dabei nicht beanstandet."),
        };
    }

    /// <summary>Stellt eine Anmeldeanfrage und liest den Fehlercode aus der Antwort.</summary>
    private static async Task<(string Code, string Detail)> AskEntra(string clientId,
        string redirect, CancellationToken cancel)
    {
        var url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
            + $"?client_id={Uri.EscapeDataString(clientId)}"
            + "&response_type=code"
            + $"&redirect_uri={Uri.EscapeDataString(redirect)}"
            + "&scope=openid"
            + "&response_mode=query"
            + "&prompt=none";

        using var handler = new HttpClientHandler { AllowAutoRedirect = false };
        using var client = new HttpClient(handler) { Timeout = Timeout };
        try
        {
            using var response = await client.GetAsync(new Uri(url), cancel);

            // Der Fehler steht entweder in der Umleitung oder im Text der Fehlerseite.
            var haystack = response.Headers.Location?.ToString() ?? "";
            if (haystack.Length == 0)
            {
                haystack = await response.Content.ReadAsStringAsync(cancel);
            }

            var marker = haystack.IndexOf("AADSTS", StringComparison.Ordinal);
            if (marker < 0) return ("", $"HTTP {(int)response.StatusCode}, kein Fehlercode.");

            var code = new string(haystack[marker..]
                .TakeWhile(c => char.IsLetterOrDigit(c))
                .ToArray());
            return (code, "");
        }
        catch (Exception error) when (error is HttpRequestException or TaskCanceledException)
        {
            return ("", $"Der Anmeldedienst ist von hier nicht erreichbar: {Short(error)}");
        }
    }

    /* ------------------------------------------------------------- Hilfsmittel */

    private static HttpClient Client()
    {
        var client = new HttpClient { Timeout = Timeout };
        client.DefaultRequestHeaders.UserAgent.Add(
            new ProductInfoHeaderValue("TanssManifest", "1.0"));
        client.DefaultRequestHeaders.CacheControl =
            new CacheControlHeaderValue { NoCache = true, NoStore = true };
        return client;
    }

    private static string Header(HttpResponseMessage response, string name) =>
        response.Headers.TryGetValues(name, out var values) ? string.Join(", ", values) : "";

    /// <summary>Die innerste Ursache in einem Satz - nicht der ganze Stapel.</summary>
    private static string Short(Exception error)
    {
        var innermost = error;
        while (innermost.InnerException is not null) innermost = innermost.InnerException;
        return innermost.Message;
    }
}
