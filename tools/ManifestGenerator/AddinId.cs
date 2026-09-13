using System.Security.Cryptography;
using System.Text;

namespace TanssAddin.ManifestGenerator;

/// <summary>
/// Die Add-in-Kennung - abgeleitet, nicht gewuerfelt.
/// </summary>
/// <remarks>
/// <para>
/// Fuer Outlook ist ein Add-in mit anderer <c>&lt;Id&gt;</c> ein ANDERES Add-in. Wird beim
/// zweiten Erzeugen eine neue Id vergeben, bleibt die alte Installation bei jedem
/// Benutzer stehen, beide funktionieren nebeneinander, und niemand bemerkt es. Das ist
/// der stillste Schaden, den dieses Werkzeug anrichten koennte.
/// </para>
/// <para>
/// Dagegen hilft keine Datei, die man aufbewahren muss - die geht verloren. Die Id wird
/// deshalb aus der TANSS-Adresse ABGELEITET: Dieselbe Instanz ergibt immer dieselbe Id,
/// auf jedem Rechner, ohne dass irgendwo etwas gemerkt wird. Wer das Manifest in zwei
/// Jahren neu erzeugt, bekommt genau dieselbe Kennung wieder.
/// </para>
/// <para>
/// Verfahren ist der namensbasierte UUID-Typ 5: SHA-1 ueber Namensraum und Name, dann
/// werden Versions- und Variantenbits gesetzt. Die Wahl von SHA-1 ist hier keine
/// Sicherheitsaussage - es geht um Reproduzierbarkeit, nicht um Faelschungssicherheit.
/// Die Id ist ohnehin oeffentlich; sie steht in jedem ausgelieferten Manifest.
/// </para>
/// </remarks>
public static class AddinId
{
    /// <summary>
    /// Namensraum dieses Projekts. Frei gewaehlt und unveraenderlich: Eine Aenderung
    /// vergaebe fuer jede bestehende Instanz eine neue Id und damit fuer Outlook ein
    /// neues Add-in.
    /// </summary>
    private static readonly Guid Namespace = new("6f2a1c84-3b7e-4d59-9c21-8e4f0a7b6d33");

    /// <summary>
    /// Leitet die Kennung aus der Adresse der TANSS-Instanz ab.
    /// </summary>
    /// <param name="tanssApiBase">
    /// Die Adresse, gegen die das Add-in arbeitet. Verwendet wird nur ihr URSPRUNG -
    /// Schema, Name und Port. Ein spaeter geaenderter Pfad (etwa <c>/backend</c> nach
    /// <c>/api</c>) darf die Kennung nicht verschieben, denn es ist dieselbe Instanz.
    /// Gross- und Kleinschreibung des Namens spielt keine Rolle.
    /// </param>
    public static Guid For(Uri tanssApiBase)
    {
        ArgumentNullException.ThrowIfNull(tanssApiBase);
        var origin = $"{tanssApiBase.Scheme}://{tanssApiBase.Host.ToLowerInvariant()}:{tanssApiBase.Port}";
        return Version5(Namespace, origin);
    }

    /// <summary>Namensbasierte UUID nach Typ 5.</summary>
    public static Guid Version5(Guid ns, string name)
    {
        var namespaceBytes = ToBigEndian(ns);
        var nameBytes = Encoding.UTF8.GetBytes(name);

        var input = new byte[namespaceBytes.Length + nameBytes.Length];
        namespaceBytes.CopyTo(input, 0);
        nameBytes.CopyTo(input, namespaceBytes.Length);

#pragma warning disable CA5350 // Das Verfahren schreibt SHA-1 vor; es geht um Reproduzierbarkeit.
        var hash = SHA1.HashData(input);
#pragma warning restore CA5350

        var result = new byte[16];
        Array.Copy(hash, result, 16);

        // Versionsbits auf 5 und Variantenbits auf RFC 4122. Ohne diese beiden Zeilen
        // entstuende eine Kennung, die zufaellig aussieht, aber keine gueltige UUID ist.
        result[6] = (byte)((result[6] & 0x0F) | 0x50);
        result[8] = (byte)((result[8] & 0x3F) | 0x80);

        return FromBigEndian(result);
    }

    /// <summary>
    /// <see cref="Guid"/> in die Netzreihenfolge bringen.
    /// </summary>
    /// <remarks>
    /// .NET legt die ersten drei Felder einer GUID in der Bytereihenfolge des Rechners
    /// ab, das UUID-Verfahren erwartet sie in Netzreihenfolge. Ohne diese Umsortierung
    /// kaeme auf einem anderen Rechner eine andere Kennung heraus - und genau die
    /// Reproduzierbarkeit ist der Zweck der ganzen Klasse.
    /// </remarks>
    private static byte[] ToBigEndian(Guid value)
    {
        var bytes = value.ToByteArray();
        Array.Reverse(bytes, 0, 4);
        Array.Reverse(bytes, 4, 2);
        Array.Reverse(bytes, 6, 2);
        return bytes;
    }

    private static Guid FromBigEndian(byte[] bytes)
    {
        var copy = (byte[])bytes.Clone();
        Array.Reverse(copy, 0, 4);
        Array.Reverse(copy, 4, 2);
        Array.Reverse(copy, 6, 2);
        return new Guid(copy);
    }
}
