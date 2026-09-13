using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace TanssAddin.ManifestGenerator;

public partial class App : Application
{
    /// <summary>
    /// Ob schon ein Fehler gemeldet wurde.
    /// </summary>
    /// <remarks>
    /// Der wichtigste Zustand dieser Klasse. Ein Meldungsfenster oeffnet eine eigene
    /// Nachrichtenschleife; tritt der Fehler beim Zeichnen auf, entsteht er darin
    /// SOFORT WIEDER, und es erscheint ein zweites Fenster, ein drittes, ein viertes.
    /// Das Programm laesst sich dann nicht mehr schliessen - man kommt gegen die
    /// Fenster nicht an. Deshalb gibt es genau eine Meldung, und danach endet das
    /// Programm, statt in diese Schleife zu laufen.
    /// </remarks>
    private bool _gemeldet;

    /// <summary>Wohin der vollstaendige Fehler geschrieben wird.</summary>
    private static string LogPath => Path.Combine(
        Path.GetTempPath(), "tanss-manifest-fehler.txt");

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnUnhandled;
    }

    private void OnUnhandled(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        // Zuerst schreiben, dann zeigen. Der Stapel ist das, womit sich der Fehler
        // finden laesst; ein Meldungsfenster kann man wegklicken, eine Datei nicht.
        var protokoll = Schreiben(e.Exception);

        if (_gemeldet)
        {
            // Zweiter Fehler waehrend der Meldung des ersten: nicht noch ein Fenster,
            // sondern Schluss. Alles andere endete in einer Schleife.
            e.Handled = true;
            Shutdown(1);
            return;
        }
        _gemeldet = true;
        e.Handled = true;

        MessageBox.Show(
            $"Unerwarteter Fehler:{Environment.NewLine}{Environment.NewLine}"
            + $"{e.Exception.Message}{Environment.NewLine}{Environment.NewLine}"
            + $"Einzelheiten stehen in:{Environment.NewLine}{protokoll}{Environment.NewLine}{Environment.NewLine}"
            + "Das Programm wird beendet.",
            "TANSS-Manifest",
            MessageBoxButton.OK,
            MessageBoxImage.Error);

        Shutdown(1);
    }

    /// <summary>Schreibt den Fehler samt Stapel. Wirft selbst nie.</summary>
    private static string Schreiben(Exception error)
    {
        try
        {
            File.AppendAllText(LogPath,
                $"--- {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---{Environment.NewLine}"
                + $"{error}{Environment.NewLine}{Environment.NewLine}");
            return LogPath;
        }
        catch
        {
            return "(keine Datei schreibbar)";
        }
    }
}
