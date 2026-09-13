/**
 * Textaufbereitung fuer die Ticketmaske.
 *
 * Aus einem Mailbetreff wird ein Tickettitel, und aus einem Mailtext der Inhalt. Beides
 * geschieht hier und nicht in TANSS - der Server schneidet einen zu langen Titel
 * WORTLOS mitten im Wort ab, und dann stuende im Ticket etwas anderes als das, was der
 * Techniker abgeschickt hat.
 */

/** Laengengrenze des Tickettitels. Darueber schneidet der Server selbst. */
export const TITLE_MAX_LENGTH = 100;

const ELLIPSIS = "…";

/**
 * Praefixe, die Outlook, Gmail und die deutschen Mailprogramme voranstellen.
 * "Re[2]:" und "AW :" kommen real vor - daher die optionale Zahl und das optionale
 * Leerzeichen. Die Wiederholung am Ende faengt verschachtelte Ketten ab.
 */
const PREFIX_RE = /^(?:\s*(?:RE|AW|WG|FW|FWD|ANTW)\s*(?:\[\d+\])?\s*:)+\s*/i;

/** Macht aus beliebigem Zeilenumbruch- und Leerraumsalat eine Zeile. */
export function collapseWhitespace(text) {
  if (!text) return "";
  // Das geschuetzte Leerzeichen ausdruecklich und als Fluchtfolge: Als Zeichen
  // geschrieben waere es im Quelltext unsichtbar, und wer die Zeile spaeter
  // umbaut, entfernte es versehentlich.
  return String(text).replace(/\u00a0/g, " ").split(/\s+/).filter(Boolean).join(" ");
}

/** Entfernt die Antwort- und Weiterleitungspraefixe, auch verschachtelte. */
export function stripReplyPrefixes(subject) {
  if (!subject) return "";
  return collapseWhitespace(subject).replace(PREFIX_RE, "").trim();
}

/**
 * Kuerzt an der WORTGRENZE und setzt ein Auslassungszeichen.
 *
 * Das Ergebnis ist nie laenger als `limit` Zeichen, das Auslassungszeichen
 * eingerechnet. Gekuerzt wird nur an einem Leerzeichen, das noch genug Titel uebrig
 * laesst - sonst bliebe von "Stoerung Telefonanlage Hauptstandort" bei knappem Limit
 * nur "S…" stehen.
 */
export function shortenTitle(text, limit = TITLE_MAX_LENGTH) {
  const line = collapseWhitespace(text);
  if (line.length <= limit) return line;
  if (limit <= 1) return ELLIPSIS.slice(0, limit);

  let head = line.slice(0, limit - 1);
  const space = head.lastIndexOf(" ");
  if (space >= Math.floor(limit / 2)) head = head.slice(0, space);
  return head.replace(/[ ,;:\-–—]+$/, "") + ELLIPSIS;
}

/** Betreff zu Tickettitel: Praefixe weg, eine Zeile, hoechstens `limit` Zeichen. */
export function titleFromSubject(subject, limit = TITLE_MAX_LENGTH) {
  return shortenTitle(stripReplyPrefixes(subject), limit);
}
