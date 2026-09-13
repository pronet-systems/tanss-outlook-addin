/**
 * Die Kennung, an der ein Outlook-Termin und ein TANSS-Einsatz zusammengehalten werden.
 *
 * Outlook liefert die `iCalUId` als lange Hex-Kette - eine Global Object ID, in der die
 * urspruengliche iCalendar-UID EINGEBETTET steckt. Wer die Hex-Kette selbst vergleicht,
 * findet eine bestehende Kopplung nicht wieder: Dieselbe Besprechung traegt in zwei
 * Postfaechern verschiedene Global Object IDs, aber dieselbe eingebettete UID.
 *
 * Und der Vergleich selbst ist nicht `===`. TANSS vergleicht ab einer gewissen Laenge
 * nur Anfang und Ende der Zeichenkette. Wer hier strenger vergleicht als der Server,
 * findet Kopplungen nicht, die es gibt - die Aufloesung fiele still auf den
 * Zeitfenster-Abgleich zurueck und traefe damit eine Heuristik statt einer Tatsache.
 */

/** Marke, hinter der Outlook die urspruengliche UID einbettet. */
const VCAL_MARKER = "vCal-Uid";

/** Nach der Marke folgen vier Byte Laengen- und Versionsfeld, dann die UID. */
const MARKER_SKIP = VCAL_MARKER.length + 4;

/** Vergleichsregel des Servers: unterhalb dieser Laenge gilt strikte Gleichheit. */
const STRICT_BELOW = 100;
const PREFIX_LEN = 32;
const SUFFIX_LEN = 70;

const SEQ_SEPARATOR = "|";

/** Ein Einzeltermin, keine Wiederholung. */
export const NO_SEQUENCE = -1;

/**
 * Die kanonische UID zu einer `iCalUId`.
 *
 * Absichtlich nachsichtig: Alles, was sich nicht als Hex lesen laesst, wird unveraendert
 * zurueckgegeben. Ein unbrauchbarer Wert waere hier ein schlechterer Ausgang als ein
 * unveraenderter - die Aufloesung hat noch andere Stufen.
 */
export function canonicalUid(icalUid) {
  if (!icalUid) return "";
  const raw = String(icalUid).trim();
  return embeddedUid(raw) || raw;
}

/** Sucht die eingebettete UID. Leer, wenn keine vorhanden ist. */
function embeddedUid(value) {
  // Nur reine Hex-Ketten gerader Laenge koennen eine Global Object ID sein.
  if (value.length < 2 * MARKER_SKIP || value.length % 2 !== 0) return "";
  if (!/^[0-9a-fA-F]+$/.test(value)) return "";

  let decoded = "";
  for (let i = 0; i < value.length; i += 2) {
    decoded += String.fromCharCode(Number.parseInt(value.slice(i, i + 2), 16));
  }

  const marker = decoded.indexOf(VCAL_MARKER);
  if (marker < 0) return "";

  let payload = decoded.slice(marker + MARKER_SKIP);
  // Die eingebettete UID ist nullterminiert; danach folgt Fuellmaterial.
  const end = payload.indexOf("\u0000");
  if (end >= 0) payload = payload.slice(0, end);

  const trimmed = payload.trim();
  // Nur druckbares ASCII - alles andere ist kein UID-Text, sondern Fuellmaterial.
  // eslint-disable-next-line no-control-regex
  return /^[ -~]+$/.test(trimmed) ? trimmed : "";
}

/**
 * Vergleicht zwei UIDs nach der Regel, die der Server anwendet.
 *
 * Ist EINE der beiden Ketten kuerzer als die Grenze, gilt strikte Gleichheit. Sind
 * BEIDE mindestens so lang, werden nur Anfang und Ende verglichen - der Mittelteil
 * wird ignoriert.
 */
export function uidMatches(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length < STRICT_BELOW || b.length < STRICT_BELOW) return false;
  return a.slice(0, PREFIX_LEN) === b.slice(0, PREFIX_LEN)
    && a.slice(-SUFFIX_LEN) === b.slice(-SUFFIX_LEN);
}

/**
 * Ist diese UID so lang, dass der Server zwei verschiedene fuer gleich halten kann?
 *
 * Die Antwort gehoert in die Oberflaeche: Eine Kopplung, die auf einem unscharfen
 * Vergleich beruht, soll als solche kenntlich sein.
 */
export function isAmbiguouslyLong(value) {
  return Boolean(value) && value.length >= STRICT_BELOW;
}

/**
 * Zerlegt einen Kopplungswert in UID und Sequenz.
 *
 * Ein Suffix `|0` ist gueltig und bedeutet Wiederholung Nummer 0 - der Wert darf nicht
 * als "nicht gesetzt" verschluckt werden. Ein Trennzeichen, das nicht zu einer Sequenz
 * gehoert, kommt in UIDs aus Fremdsystemen vor; dann ist der ganze Wert die UID.
 */
export function parseSyncGroup(value) {
  if (!value) return { uid: "", sequence: NO_SEQUENCE };
  const cut = value.lastIndexOf(SEQ_SEPARATOR);
  if (cut < 0) return { uid: value, sequence: NO_SEQUENCE };

  const tail = value.slice(cut + 1);
  if (!/^-?\d+$/.test(tail)) return { uid: value, sequence: NO_SEQUENCE };
  return { uid: value.slice(0, cut), sequence: Number.parseInt(tail, 10) };
}

/** Baut einen Kopplungswert. Sequenz kleiner null ergibt nur die UID. */
export function formatSyncGroup(uid, sequence = NO_SEQUENCE) {
  return sequence < 0 ? uid : `${uid}${SEQ_SEPARATOR}${sequence}`;
}

/**
 * Ob ein gelesener Kopplungswert zu unserer kanonischen UID gehoert.
 *
 * Das Sequenz-Suffix schliesst nur dann aus, wenn BEIDE Seiten eine Sequenz tragen und
 * sie sich unterscheiden. Beim Lesen traegt der Wert in aller Regel keine - expandierte
 * Wiederholungen erben die Angaben des Ursprungstermins unveraendert.
 */
export function syncGroupMatches(syncGroup, uid, sequence = NO_SEQUENCE) {
  const read = parseSyncGroup(syncGroup);
  if (!uidMatches(read.uid, uid)) return false;
  if (read.sequence >= 0 && sequence >= 0) return read.sequence === sequence;
  return true;
}
