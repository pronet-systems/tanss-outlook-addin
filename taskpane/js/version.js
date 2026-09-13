/**
 * Die Fassung des Panes.
 *
 * Sie steht in einer eigenen Datei, weil sie an genau zwei Stellen gebraucht wird - der
 * Diagnoseseite und der Identitaetsauskunft - und sonst in beiden haengen bliebe. Bei
 * einer Auslieferung wird sie hier erhoeht.
 *
 * Sie hat NICHTS mit der Manifestversion zu tun: Die zaehlt Microsoft mit und
 * entscheidet, ob eine Aktualisierung des Menuebands uebernommen wird. Diese hier sagt
 * nur, welcher Stand der Dateien gerade ausgeliefert ist.
 */
export const VERSION = "0.1.0";
