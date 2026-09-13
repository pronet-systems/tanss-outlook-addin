/**
 * Der Fehlertyp dieser Schicht und der Umschlag, in dem jede Auskunft die Oberflaeche
 * erreicht.
 *
 * Ein Umschlag, nicht zwei. Zwei nebeneinander gefuehrte Fehlerformen enden im selben
 * Fehler: Der Leser greift auf ein Feld zu, das diesmal fehlt, wirft im eigenen `catch`
 * und die Seite bleibt fuer immer im Ladezustand. Jede Aussenfunktion gibt deshalb
 * `{ok:true, data}` oder `{ok:false, error:{code, message}}` zurueck - auch bei
 * Netzabbruch, Zeitueberschreitung und Proxy-Fehlerseite. Es fliegt nichts nach aussen.
 *
 * Diese Datei kennt die Uebersetzungstabelle bewusst NICHT. `ApiError` traegt nur den
 * Code; der Satz dazu entsteht erst in `envelope.js`. So laesst sich die ganze
 * TANSS-Schicht ohne Oberflaeche pruefen.
 */

/**
 * Ein Fehler mit genau einem Code aus dem geschlossenen Bestand der Oberflaeche.
 *
 * `detail` traegt den unuebersetzten Namen aus TANSS, wo es einen gibt. Er ist fuer die
 * Diagnoseseite gedacht und wird dem Techniker nie als Erklaerung vorgesetzt.
 */
export class ApiError extends Error {
  constructor(code, message = "", { status = 0, detail = "", cause = null } = {}) {
    super(message || code);
    this.name = "ApiError";
    this.code = code;
    this.status = status;
    this.detail = detail;
    if (cause) this.cause = cause;
  }
}

/**
 * Der Grund eines Ausfalls in wenigen Worten - fuer die Anzeige, nicht fuer die Logik.
 *
 * Route, Code und HTTP-Status beschreiben die SCHNITTSTELLE. Ein Antwortkoerper
 * beschriebe den Vorgang, und der geht niemanden ausser dem Techniker etwas an; er kommt
 * hier deshalb nicht vor.
 *
 * Der Satz landet in einer Meldung, die weitergereicht wird. Er muss einem Administrator
 * die Ursache nennen, ohne dass jemand ein Protokoll aufmachen muss - "HTTP 404" und
 * "HTTP 403" fuehren zu ganz verschiedenen Suchen.
 */
export function reasonOf(error) {
  if (!error) return "ohne Angabe";
  const teile = [];
  if (error.status) teile.push(`HTTP ${error.status}`);
  if (error.code && error.code !== "INTERNAL") teile.push(String(error.code));
  if (error.detail) teile.push(String(error.detail).slice(0, 80));
  if (teile.length === 0) {
    const name = error.name && error.name !== "Error" ? error.name : "";
    const text = String(error.message || "").slice(0, 80);
    return [name, text].filter(Boolean).join(": ") || "ohne Angabe";
  }
  return teile.join(" · ");
}

/** Ist dieser Fehler ein abgelaufenes oder verworfenes Token? */
export function isUnauthenticated(error) {
  return error instanceof ApiError && error.code === "UNAUTHENTICATED";
}
