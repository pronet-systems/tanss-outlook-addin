using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace TanssAddin.ManifestGenerator;

/// <summary>
/// Die Bedienung. Sie enthaelt keine Fachlogik - die steht in
/// <see cref="Options"/>, <see cref="Renderer"/> und <see cref="AddinId"/> und laesst
/// sich dort ohne Fenster pruefen.
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Zeigt die abgeleitete Kennung, waehrend getippt wird.
    /// </summary>
    /// <remarks>
    /// Das ist mehr als Zierde: Wer ein Manifest fuer eine bestehende Installation neu
    /// erzeugt, sieht hier sofort, ob dieselbe Kennung herauskommt wie beim letzten Mal.
    /// Eine abweichende Kennung waere fuer Outlook ein anderes Add-in, und die alte
    /// Installation bliebe bei jedem Benutzer daneben stehen.
    /// </remarks>
    private void AufTanssGeaendert(object sender, TextChangedEventArgs e)
    {
        if (KennungFeld is null) return;

        var eingabe = (TanssFeld.Text ?? "").Trim();
        if (Uri.TryCreate(eingabe, UriKind.Absolute, out var uri)
            && uri.Scheme == Uri.UriSchemeHttps)
        {
            KennungFeld.Text = AddinId.For(uri).ToString();
        }
        else
        {
            KennungFeld.Text = "— wird aus der TANSS-Adresse abgeleitet —";
        }
    }

    private void AufErzeugen(object sender, RoutedEventArgs e)
    {
        Options options;
        try
        {
            options = AusDenFeldern();
        }
        catch (OptionsError error)
        {
            Beanstanden(error);
            return;
        }

        string xml;
        try
        {
            xml = Renderer.Render(options);
        }
        catch (ManifestError error)
        {
            Sagen(error.Message, gut: false);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = "Manifest speichern",
            FileName = "manifest.xml",
            DefaultExt = ".xml",
            Filter = "Manifestdatei (*.xml)|*.xml",
            OverwritePrompt = true,
        };
        if (dialog.ShowDialog(this) != true) return;

        try
        {
            File.WriteAllText(dialog.FileName, xml);
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            Sagen($"Die Datei liess sich nicht schreiben: {error.Message}", gut: false);
            return;
        }

        Sagen($"Geschrieben: {Path.GetFileName(dialog.FileName)}", gut: true);
        NaechsteSchritte(dialog.FileName, options);
    }

    /// <summary>
    /// Misst die Angaben gegen die Wirklichkeit.
    /// </summary>
    /// <remarks>
    /// Jede dieser Pruefungen deckt einen Fehler ab, der sonst erst beim Techniker
    /// auffaellt - und dort als "das Add-in geht nicht" ankommt. Ein leeres Taskpane
    /// sagt nicht, dass eine Kopfzeile fehlt oder dass ein Ursprung nicht zugelassen ist.
    /// Die Ergebnisse erscheinen einzeln, sobald sie vorliegen; ein haengender Aufruf
    /// haelt die uebrigen damit nicht auf.
    /// </remarks>
    private async void AufPruefen(object sender, RoutedEventArgs e)
    {
        Options options;
        try
        {
            options = AusDenFeldern();
        }
        catch (OptionsError error)
        {
            Beanstanden(error);
            return;
        }

        _laufendePruefung?.Cancel();
        var laufen = new CancellationTokenSource();
        _laufendePruefung = laufen;

        var zeilen = new ObservableCollection<Pruefzeile>();
        Pruefliste.ItemsSource = zeilen;
        PruefenKnopf.IsEnabled = false;
        ErzeugenKnopf.IsEnabled = false;
        Sagen("Prüfung läuft…", gut: true);

        var schlecht = 0;
        var unklar = 0;
        try
        {
            await foreach (var ergebnis in Checks.RunAsync(options, laufen.Token))
            {
                zeilen.Add(Pruefzeile.Aus(ergebnis));
                if (ergebnis.Verdict == Verdict.Fail) schlecht++;
                if (ergebnis.Verdict == Verdict.Unclear) unklar++;
            }
        }
        catch (OperationCanceledException)
        {
            return;
        }
        finally
        {
            PruefenKnopf.IsEnabled = true;
            ErzeugenKnopf.IsEnabled = true;
            if (ReferenceEquals(_laufendePruefung, laufen)) _laufendePruefung = null;
            laufen.Dispose();
        }

        // "Nichts beanstandet" wird nur gesagt, wenn wirklich nichts offen ist. Ein
        // unklarer Ausgang ist keine bestandene Pruefung.
        Sagen(
            schlecht > 0
                ? $"{schlecht} Beanstandung(en) — so geht das Add-in nicht in Betrieb."
                : unklar > 0
                    ? $"Keine Beanstandung, aber {unklar} Punkt(e) blieben unklar."
                    : "Alles geprüft, nichts beanstandet.",
            gut: schlecht == 0);
    }

    private CancellationTokenSource? _laufendePruefung;

    private Options AusDenFeldern() => Options.Create(
        addinBase: AblageFeld.Text,
        tanssApi: TanssFeld.Text,
        tanssFrontend: OberflaecheFeld.Text,
        version: VersionFeld.Text,
        displayName: NameFeld.Text,
        providerName: HerausgeberFeld.Text,
        supportUrl: SupportFeld.Text,
        entraClientId: EntraFeld.Text,
        tenantHint: MandantFeld.Text,
        ticketTypes: TypenFeld.Text,
        embedTanssInUrl: AdresseEinbetten.IsChecked == true);

    /// <summary>Ein Prüfergebnis, wie die Liste es zeigt.</summary>
    public sealed record Pruefzeile(string Bereich, string Name, string Detail,
        string Zeichen, Brush Farbe)
    {
        public static Pruefzeile Aus(CheckResult ergebnis) => ergebnis.Verdict switch
        {
            Verdict.Ok => new Pruefzeile(ergebnis.Area, ergebnis.Name, ergebnis.Detail,
                "✓", new SolidColorBrush(Color.FromRgb(0x1B, 0x6B, 0x2A))),
            Verdict.Unclear => new Pruefzeile(ergebnis.Area, ergebnis.Name, ergebnis.Detail,
                "!", new SolidColorBrush(Color.FromRgb(0xB0, 0x6A, 0x00))),
            _ => new Pruefzeile(ergebnis.Area, ergebnis.Name, ergebnis.Detail,
                "✕", new SolidColorBrush(Color.FromRgb(0xB0, 0x00, 0x20))),
        };
    }

    /// <summary>Stellt die Meldung an das Feld, in dem der Fehler steht.</summary>
    private void Beanstanden(OptionsError error)
    {
        var feld = error.Field switch
        {
            nameof(Options.TicketTypes) => TypenFeld,
            nameof(Options.AddinBase) => AblageFeld,
            nameof(Options.TanssApi) => TanssFeld,
            nameof(Options.TanssFrontend) => OberflaecheFeld,
            nameof(Options.Version) => VersionFeld,
            nameof(Options.SupportUrl) => SupportFeld,
            nameof(Options.EntraClientId) => EntraFeld,
            _ => null,
        };
        feld?.Focus();
        Sagen(error.Message, gut: false);
    }

    private void Sagen(string text, bool gut)
    {
        Meldung.Text = text;
        Meldung.Foreground = gut
            ? (Brush)FindResource("Gut")
            : (Brush)FindResource("Fehler");
    }

    /// <summary>
    /// Was nach dem Speichern zu tun ist.
    /// </summary>
    /// <remarks>
    /// Der Hinweis auf die Version steht hier und nicht in der Anleitung, weil er den
    /// stillsten aller Fehler betrifft: Bleibt die Version gleich, uebergeht Microsoft
    /// die Aktualisierung wortlos, und niemand bemerkt, dass die Benutzer weiter mit der
    /// alten Fassung arbeiten.
    /// </remarks>
    private void NaechsteSchritte(string pfad, Options options)
    {
        var text =
            $"Die Datei liegt unter:{Environment.NewLine}{pfad}{Environment.NewLine}{Environment.NewLine}"
            + $"Kennung:  {AddinId.For(options.TanssApi)}{Environment.NewLine}"
            + $"Version:  {options.Version}{Environment.NewLine}{Environment.NewLine}"
            + "Weiter im Microsoft 365 Admin Center:" + Environment.NewLine
            + "  Einstellungen > Integrierte Apps > Add-Ins" + Environment.NewLine
            + "  > Add-In bereitstellen > Benutzerdefiniertes Add-In" + Environment.NewLine
            + "  > Manifestdatei hochladen" + Environment.NewLine + Environment.NewLine
            + "Die Datei wird hochgeladen, nicht verlinkt. Der Hoster muss fuer Microsoft "
            + "also nie erreichbar sein." + Environment.NewLine + Environment.NewLine
            + "Bei einer spaeteren Aenderung die Version erhoehen. Bleibt sie gleich, "
            + "uebergeht Microsoft die Aktualisierung, und die Benutzer behalten "
            + "stillschweigend die alte Fassung.";

        MessageBox.Show(this, text, "Manifest erzeugt", MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}
