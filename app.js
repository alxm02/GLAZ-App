/* ==========================================================================
   GLAZ-Korrekturbuchungsliste — Anwendungslogik der Handy-Fassung

   Der Aufbau folgt einer Grundentscheidung: Die Fachlogik wird NICHT in
   JavaScript nachgebaut. Sie liegt weiterhin in glaz/validation.py,
   glaz/excel_engine.py, glaz/filename.py und glaz/mailtext.py und laeuft
   hier unter Pyodide — demselben CPython, nur als WebAssembly. Eine zweite
   Umsetzung derselben Regeln waere eine zweite Quelle fuer dieselbe
   Wahrheit, und eine davon liefe der anderen frueher oder spaeter hinterher.

   JavaScript hat deshalb genau drei Aufgaben:
     1. die Oberflaeche zeichnen und bedienen,
     2. den Zustand im Geraet halten (localStorage),
     3. Daten als JSON zu Python schicken und Ergebnisse zurueckholen.

   Alles, was eine fachliche Entscheidung ist — ob eine Zeile gueltig ist,
   wie die Datei heisst, was in der Mail steht — faellt in Python.
   ========================================================================== */

"use strict";

/* --------------------------------------------------------------------------
   Konstanten
   -------------------------------------------------------------------------- */

const XLSX_TYP =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/* Die Kernmodule, die ins virtuelle Dateisystem von Pyodide wandern.
   Bewusst eine ausgeschriebene Liste und kein Verzeichnis-Abruf: der
   Service-Worker-Cache muss dieselben Pfade kennen, und eine Liste, die man
   lesen kann, ist einer zur Laufzeit erratenen vorzuziehen. Wer hier ein
   Modul ergaenzt, ergaenzt es auch in tools/baue_pwa.py. */
const KERNMODULE = [
  "__init__.py",
  "model.py",
  "timeconv.py",
  "filename.py",
  "excel_engine.py",
  "validation.py",
  "resources_util.py",
  "vorgaben.py",
  "mailtext.py",
  "portabel.py",
];

/* Die Bruecke nach Python. Alles geht als JSON-Text hin und zurueck: das
   erspart uns PyProxy-Objekte, die man von Hand freigeben muesste, und
   haelt die Schnittstelle so schmal, dass sie in einen Blick passt. */
const BRUECKE_PY = `
import json, sys, traceback
sys.path.insert(0, "/py")
from glaz import portabel
from glaz.resources_util import ressource

def bruecke_vorlage():
    return str(ressource("template.xlsx"))

def bruecke_pruefe(roh):
    return json.dumps(portabel.pruefe(json.loads(roh)), ensure_ascii=False)

def bruecke_erzeuge(roh, ziel):
    erg = portabel.erzeuge_xlsx(json.loads(roh), bruecke_vorlage(), ziel)
    return json.dumps(erg, ensure_ascii=False)

def bruecke_ergaenze(roh, pfad, start_row=None, ueberschreibe=0):
    erg = portabel.ergaenze_xlsx(
        json.loads(roh), pfad,
        start_row=start_row or None,
        ueberschreibe_zeilen=int(ueberschreibe or 0),
    )
    return json.dumps(erg, ensure_ascii=False)

def bruecke_belegung(pfad):
    return json.dumps(portabel.belegung(pfad), ensure_ascii=False)

def bruecke_pruefe_ziel(roh, pfad, modus):
    return json.dumps(portabel.pruefe_ziel(json.loads(roh), pfad, modus), ensure_ascii=False)

def bruecke_mailtexte(roh, empfaenger, betreff, body):
    return json.dumps(
        portabel.mailtexte(json.loads(roh), empfaenger, betreff, body),
        ensure_ascii=False,
    )
`;

const SPEICHER_SCHLUESSEL = "glaz.zustand.v1";
const ANSTUPSER_SCHLUESSEL = "glaz.anstupser.gesehen";

const LEERES_PROFIL = {
  name: "",
  personalnummer: "",
  vorname: "",
  nachname: "",
  abteilung: "",
  gruppenleiter_name: "",
  gruppenleiter_email: "",
  zielordner: "",
};

const ZEITFELDER = [
  "passiv_von", "passiv_bis",
  "aktiv_von", "aktiv_bis",
  "grenz_anreise", "grenz_rueckreise",
  "arbeit_von", "arbeit_bis",
];

const MAX_ZEILEN = 14;

/* Die Seiten des Assistenten, in der Reihenfolge, in der man sie durchlaeuft.
   Aus der einen langen Seite wurden vier kurze: Auf 390 px Breite scrollte
   man sonst durch Profil, Vorgang und vierzehn Reisetage, bevor der
   Abschluss ueberhaupt in Sicht kam. ``id`` ist die section in index.html,
   ``name`` steht in der Fortschrittsleiste. Die Nummer ist der Index + 1 --
   sie wird gespeichert und angezeigt, deshalb nicht 0-basiert. */
const SCHRITTE = [
  { id: "profil-block",  name: "Profil" },
  { id: "vorgang-block", name: "Vorgang" },
  { id: "reisetage",     name: "Reisetage" },
  { id: "abschluss",     name: "Abschluss" },
];

/* Bereiche, die aus dem Menue heraus geoeffnet werden und nicht zum Ablauf
   gehoeren. Sie ersetzen den aktuellen Schritt, bis "Fertig" gedrueckt wird. */
const SONDERSEITEN = ["einstellungen", "selbsttest", "zwischenstaende"];

//: Ruhetext unter den Profilfeldern. Er sagt, was ohnehin passiert -- und
//: macht damit den Speichern-Knopf zu einer Bestaetigung statt zu einer
//: Bedingung.
const PROFIL_HINWEIS_STANDARD = "Änderungen werden sofort im Gerät gesichert.";

/* --------------------------------------------------------------------------
   Kleine Helfer
   -------------------------------------------------------------------------- */

const el = (id) => document.getElementById(id);
const alle = (auswahl, wurzel = document) => [...wurzel.querySelectorAll(auswahl)];

function heuteAlsText() {
  // Bewusst nicht toISOString(): das rechnet nach UTC um und liefert vor
  // 01:00 Ortszeit den Vortag. Bei einer Reisekostenliste ist ein um einen
  // Tag verschobenes Vorbelegungsdatum kein Schoenheitsfehler.
  const d = new Date();
  const zwei = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${zwei(d.getMonth() + 1)}-${zwei(d.getDate())}`;
}

function minutenAusZeit(text) {
  if (!text) return null;
  const teile = String(text).split(":");
  if (teile.length < 2) return null;
  const h = Number(teile[0]);
  const m = Number(teile[1]);
  if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
  return h * 60 + m;
}

function entprellt(fn, ms) {
  let kennung = null;
  return (...args) => {
    clearTimeout(kennung);
    kennung = setTimeout(() => fn(...args), ms);
  };
}

/* --------------------------------------------------------------------------
   Zustand — alles, was der Nutzer eingegeben hat, lebt hier und im
   localStorage des Geraets. Es verlaesst das Telefon nie.
   -------------------------------------------------------------------------- */

const zustand = {
  profile: [{ ...LEERES_PROFIL, name: "Mein Profil" }],
  aktivesProfil: 0,
  /* Der Vorrat gepflegter Gruppenleiter -- dasselbe wie die Stammdatenliste
     der Desktop-App. Verwaltet wird er in web/gruppenleiter.js. */
  gruppenleiter: [],
  einsatzart: "",
  einsatzarten: [],
  reisetyp: "inland",
  wiederhole_stammdaten: false,
  zeilen: [],
  modus: "neu",
  mailMitsenden: true,
  /* Wie die Mail das Telefon verlaesst: "outlook" oeffnet die Outlook-App mit
     fertig eingetragenem An, CC, Betreff und Text; "teilen" gibt die Datei ans
     Teilen-Menue. Outlook ist die Vorgabe, weil es genau das Tippen erspart,
     das am Telefon am meisten stoert -- die Adressen. Den Anhang kann es
     dafuer nicht mitnehmen (siehe versendeUeberOutlook). */
  versandweg: "outlook",
  /* Beiseitegelegte Vorgaenge (siehe "Zwischenstaende" unten). Der laufende
     Stand selbst steht wie bisher in einsatzart/zeilen/...; die Liste haelt
     die, an denen gerade nicht gearbeitet wird. */
  zwischenstaende: [],
  /* Kennung des Eintrags, zu dem der laufende Stand gehoert -- oder null.
     Sichern aktualisiert dann diesen Eintrag, statt einen zweiten anzulegen. */
  zwischenstandId: null,
  /* Aktuelle Seite des Assistenten (1 bis SCHRITTE.length). Wird mitgesichert,
     damit ein Neustart -- etwa weil iOS die App im Hintergrund beendet hat --
     dort weitergeht, wo man aufgehoert hat, und nicht wieder beim Profil. */
  schritt: 1,
  einstellungen: {
    empfaenger: "",
    betreff_vorlage: "",
    body_vorlage: "",
    /* "dunkel" oder "hell". Vorgabe ist dunkel, weil die Gestaltung darauf
       hin entworfen ist -- die Systemeinstellung des Telefons sagt nichts
       darueber, wie diese Anwendung aussehen soll. */
    thema: "dunkel",
  },
};

/* Die gewaehlte Zieldatei lebt absichtlich NICHT im gespeicherten Zustand:
   ein File-Objekt ueberlebt keinen Neustart, und ein Pfad, auf den wir beim
   naechsten Start nicht mehr zugreifen koennen, waere ein leeres Versprechen. */
let zieldatei = null;
/* Getrennt von `zieldatei`, weil das File-Objekt schon gewaehlt sein kann,
   waehrend das Einlesen nach /ziel.xlsx noch laeuft. Die laufende Pruefung
   darf erst danach darauf zugreifen. */
let zieldateiGeladen = false;
/* Was der letzte erfolgreiche Lauf im Modus "Ergaenzen" geschrieben hat.
   Ohne diese Merker haengt ein zweiter Druck denselben Block ein zweites Mal
   an -- genau der Schaden, gegen den die Desktop-App ihren
   Aktualisieren-Zustand hat. */
let letzterAnhang = null;

/**
 * Datum, mit dem eine neu angelegte Reisezeile vorbelegt wird.
 *
 * Erste Wahl ist der zuletzt eingetragene Tag -- NICHT der Folgetag. Genau so
 * haelt es ``_vorschlagsdatum`` in der Desktop-App, und das aus gutem Grund:
 * Ein Reisetag verteilt sich regelmaessig auf mehrere Zeilen (arbeiten,
 * reisen, arbeiten, reisen passt nicht in eine). Wer dort jedes Mal den
 * Folgetag vorgesetzt bekommt, verschiebt unbemerkt den Zeitraum -- und der
 * steht im Dateinamen und im Betreff der Mail.
 *
 * Verlaesslich bleibt die Vorbelegung, weil eine Zeile ohne Uhrzeiten nicht
 * als Datenzeile zaehlt (siehe TravelRow.ist_leer): ein unbeachtetes
 * Vorschlagsdatum kann weder in die Excel noch in den Zeitraum geraten.
 */
function vorschlagsdatum() {
  for (let i = zustand.zeilen.length - 1; i >= 0; i--) {
    if (zustand.zeilen[i].datum) return zustand.zeilen[i].datum;
  }
  return heuteAlsText();
}

function leereZeile(folgezeile = false) {
  const z = { datum: folgezeile ? null : vorschlagsdatum(), ist_folgezeile: folgezeile };
  for (const f of ZEITFELDER) z[f] = null;
  return z;
}

function aktuellesProfil() {
  return zustand.profile[zustand.aktivesProfil] || zustand.profile[0];
}

/**
 * Beschriftung eines Profils in der Auswahlliste.
 *
 * Einzige Quelle fuer diesen Text -- dasselbe Vorgehen wie in
 * ``Profile.anzeige_label`` der Desktop-App: ohne eigenen Namen tritt der
 * Personenname ein. Wichtig ist, dass wirklich alle Stellen ihn hierher
 * holen; sonst heisst ein Profil in der Liste "Erika Mustermann", seine Kopie
 * aber "Profil (Kopie)", und niemand versteht, woher das kommt.
 */
function profilLabel(p) {
  return (
    (p.name || "").trim() ||
    `${p.vorname} ${p.nachname}`.trim() ||
    "(ohne Namen)"
  );
}

/**
 * Sorgt dafuer, dass kein zweites Profil denselben Namen traegt.
 *
 * Zwei gleich benannte Eintraege in der Auswahlliste sind schlimmer als ein
 * unschoener Name: Man kann sie nicht auseinanderhalten und waehlt frueher
 * oder spaeter das falsche Profil. Angehaengt wird deshalb " (2)", " (3)" und
 * so fort -- dieselbe Regel, die die Desktop-App in
 * ``glaz/settings.py`` anwendet. (Bewusst nachgebaut statt von dort geholt:
 * Die Profile der Handy-Fassung leben im localStorage und begegnen denen der
 * Desktop-App nie, es gibt also nichts, was auseinanderlaufen koennte.)
 *
 * Ein leerer Name bleibt leer -- dafuer tritt in der Liste der Personenname
 * ein, genau wie in ``Profile.anzeige_label``.
 *
 * @param {string} wunsch       gewuenschter Name
 * @param {number} eigenerIndex Platz des Profils, das umbenannt wird (-1 fuer
 *                              ein noch nicht eingefuegtes)
 */
function eindeutigerProfilname(wunsch, eigenerIndex) {
  const name = (wunsch || "").trim();
  if (!name) return "";

  const vergeben = new Set(
    zustand.profile
      .filter((_, i) => i !== eigenerIndex)
      .map((p) => (p.name || "").trim())
      .filter(Boolean)
  );
  if (!vergeben.has(name)) return name;

  for (let n = 2; n < 1000; n++) {
    const versuch = `${name} (${n})`;
    if (!vergeben.has(versuch)) return versuch;
  }
  return name;
}

/** Der Vorgang in genau der Form, die glaz/portabel.py erwartet. */
function vorgangDict() {
  const p = aktuellesProfil() || LEERES_PROFIL;
  return {
    profil: { ...LEERES_PROFIL, ...p },
    einsatzart: zustand.einsatzart,
    reisetyp: zustand.reisetyp,
    wiederhole_stammdaten: zustand.wiederhole_stammdaten,
    // Im Inland sind die Grenzfelder ausgeblendet. Ihre Werte bleiben im
    // Zustand, damit ein Hin- und Zurueckschalten nichts verliert -- an die
    // Pruefung gehen sie aber nicht: Sonst meldete sie einen Fehler an einem
    // Feld, das man weder sieht noch leeren kann.
    zeilen: zustand.zeilen.map((z) =>
      zustand.reisetyp === "ausland"
        ? { ...z }
        : { ...z, grenz_anreise: null, grenz_rueckreise: null }
    ),
  };
}

function speichern() {
  try {
    localStorage.setItem(SPEICHER_SCHLUESSEL, JSON.stringify(zustand));
    return true;
  } catch (e) {
    // Voller oder gesperrter Speicher darf die Eingabe nicht abwuergen.
    // Der Nutzer erfaehrt es an der Statuszeile, nicht per Absturz.
    melde("Der Zwischenstand konnte nicht gesichert werden.", "warnung");
    return false;
  }
}

function laden() {
  let roh = null;
  try {
    roh = localStorage.getItem(SPEICHER_SCHLUESSEL);
  } catch (e) {
    roh = null;
  }
  if (!roh) return;
  try {
    const gelesen = JSON.parse(roh);
    // Feld fuer Feld uebernehmen statt Object.assign: eine aeltere oder von
    // Hand verbogene Fassung soll nicht den ganzen Zustand vergiften.
    if (Array.isArray(gelesen.profile) && gelesen.profile.length) {
      // Jedes Feld als Text: Ein null oder eine Zahl aus einer aelteren
      // Fassung liesse sonst profilLabel() beim Start an .trim() scheitern --
      // und die App bliebe leer.
      zustand.profile = gelesen.profile
        .filter((p) => p && typeof p === "object")
        .map((p) => {
          const profil = { ...LEERES_PROFIL };
          for (const feld of Object.keys(LEERES_PROFIL)) {
            if (p[feld] !== undefined && p[feld] !== null) profil[feld] = String(p[feld]);
          }
          return profil;
        });
      if (!zustand.profile.length) zustand.profile = [{ ...LEERES_PROFIL, name: "Mein Profil" }];
    }
    if (Number.isInteger(gelesen.aktivesProfil)) {
      zustand.aktivesProfil = Math.min(
        Math.max(0, gelesen.aktivesProfil),
        zustand.profile.length - 1
      );
    }
    if (typeof gelesen.einsatzart === "string") zustand.einsatzart = gelesen.einsatzart;
    if (Array.isArray(gelesen.einsatzarten)) {
      zustand.einsatzarten = gelesen.einsatzarten.filter((a) => typeof a === "string");
    }
    if (gelesen.reisetyp === "ausland" || gelesen.reisetyp === "inland") {
      zustand.reisetyp = gelesen.reisetyp;
    }
    if (Array.isArray(gelesen.zeilen)) {
      zustand.zeilen = gelesen.zeilen
        .filter((z) => z && typeof z === "object")
        .slice(0, MAX_ZEILEN)
        .map((z) => ({ ...leereZeile(), ...z }));
    }
    if (typeof gelesen.mailMitsenden === "boolean") {
      zustand.mailMitsenden = gelesen.mailMitsenden;
    }
    if (gelesen.versandweg === "outlook" || gelesen.versandweg === "teilen") {
      zustand.versandweg = gelesen.versandweg;
    }
    if (Array.isArray(gelesen.zwischenstaende)) {
      zustand.zwischenstaende = gelesen.zwischenstaende
        .map(bereinigeZwischenstand)
        .filter(Boolean)
        .slice(0, MAX_ZWISCHENSTAENDE);
    }
    if (typeof gelesen.zwischenstandId === "string"
        && zustand.zwischenstaende.some((e) => e.id === gelesen.zwischenstandId)) {
      zustand.zwischenstandId = gelesen.zwischenstandId;
    }
    if (Number.isInteger(gelesen.schritt)) {
      zustand.schritt = Math.min(Math.max(1, gelesen.schritt), SCHRITTE.length);
    }
    if (Array.isArray(gelesen.gruppenleiter)) {
      zustand.gruppenleiter = gelesen.gruppenleiter
        .filter((g) => g && typeof g === "object")
        .map((g) => ({
          name: String(g.name || ""),
          abteilung: String(g.abteilung || ""),
          email: String(g.email || ""),
        }));
    }
    if (gelesen.einstellungen && typeof gelesen.einstellungen === "object") {
      for (const feld of ["empfaenger", "betreff_vorlage", "body_vorlage", "thema"]) {
        const wert = gelesen.einstellungen[feld];
        if (typeof wert === "string") zustand.einstellungen[feld] = wert;
      }
      // Ein unbekannter Wert wuerde die Seite unlesbar machen, deshalb hier
      // und nicht erst beim Anwenden abfangen.
      if (zustand.einstellungen.thema !== "hell") zustand.einstellungen.thema = "dunkel";
    }
  } catch (e) {
    // Kaputter Speicherinhalt: lieber frisch anfangen als halb geladen
    // weiterarbeiten. Die Vorgaben stehen bereits im Zustand.
  }
}

/* --------------------------------------------------------------------------
   Python-Kern
   -------------------------------------------------------------------------- */

let py = null;              // die Pyodide-Instanz
let kernBereit = false;
let kernFehler = "";

async function starteKern() {
  setzeBereitschaft("startet", "startet");
  try {
    py = await loadPyodide({ indexURL: "vendor/pyodide/" });

    py.FS.mkdirTree("/py/glaz/resources");
    for (const name of KERNMODULE) {
      const antwort = await fetch(`py/glaz/${name}`);
      if (!antwort.ok) throw new Error(`${name} fehlt (${antwort.status})`);
      py.FS.writeFile(`/py/glaz/${name}`, await antwort.text());
    }
    const vorlage = await fetch("py/glaz/resources/template.xlsx");
    if (!vorlage.ok) throw new Error(`Excel-Vorlage fehlt (${vorlage.status})`);
    py.FS.writeFile(
      "/py/glaz/resources/template.xlsx",
      new Uint8Array(await vorlage.arrayBuffer())
    );

    py.runPython(BRUECKE_PY);

    kernBereit = true;
    setzeBereitschaft("bereit", "bereit");
    pruefeJetzt();
    pruefeEinstellungen();
    // Erst jetzt ist bekannt, ob im Profil etwas fehlt.
    if (zustand.schritt === 1) klappeProfilAufWennUnvollstaendig();
  } catch (e) {
    kernFehler = e && e.message ? e.message : String(e);
    setzeBereitschaft("fehler", "Kern fehlt");
    melde(`Die Prüfung ist nicht verfügbar: ${kernFehler}`, "fehler");
  }
}

/**
 * Macht aus einem Pyodide-Fehler die Meldung, die die Nutzerin lesen soll.
 *
 * e.message ist der komplette Python-Traceback. Die Fachmodule formulieren
 * ihre Fehler bewusst deutsch und handlungsleitend (siehe MailtextFehler) --
 * die stehen in der letzten Zeile, hinter "modul.Klasse: ".
 */
function pythonMeldung(e) {
  const text = String(e && e.message ? e.message : e).trim();
  const zeilen = text.split("\n").map((z) => z.trim()).filter(Boolean);
  const letzte = zeilen[zeilen.length - 1] || text;
  return letzte.replace(/^[\w.]+(Error|Fehler|Exception):\s*/, "");
}

function rufe(name, ...args) {
  const fn = py.globals.get(name);
  try {
    return JSON.parse(fn(...args));
  } catch (e) {
    throw new Error(pythonMeldung(e));
  } finally {
    // PyProxy-Objekte haelt der Browser sonst bis zum Neuladen fest.
    if (fn && typeof fn.destroy === "function") fn.destroy();
  }
}

/* --------------------------------------------------------------------------
   Oberflaeche: Reisetage
   -------------------------------------------------------------------------- */

function zeichneZeilen() {
  const liste = el("zeilen-liste");
  const vorlage = el("vorlage-reisetag");
  liste.textContent = "";

  zustand.zeilen.forEach((zeile, index) => {
    const knoten = vorlage.content.firstElementChild.cloneNode(true);
    knoten.dataset.index = String(index);
    knoten.dataset.folgezeile = String(!!zeile.ist_folgezeile);

    knoten.querySelector(".tag-nummer").textContent = String(index + 1);
    // "Arbeitszeit von" gibt es vierzehnmal. Fuer VoiceOver bekommt jede
    // Beschriftung der Karte ihren Reisetag vorangestellt.
    alle("[aria-label]", knoten).forEach((f) => {
      f.setAttribute("aria-label", `Reisetag ${index + 1}: ${f.getAttribute("aria-label")}`);
    });

    const datum = knoten.querySelector(".tag-datum");
    datum.value = zeile.datum || "";
    datum.disabled = !!zeile.ist_folgezeile;
    datum.addEventListener("change", () => {
      zustand.zeilen[index].datum = datum.value || null;
      nachEingabe();
    });

    alle(".zeit", knoten).forEach((feld) => {
      const name = feld.dataset.feld;
      feld.value = zeile[name] || "";
      feld.dataset.gefuellt = String(!!zeile[name]);
      feld.addEventListener("input", () => {
        zustand.zeilen[index][name] = feld.value || null;
        feld.dataset.gefuellt = String(!!feld.value);
        zeichneBand(knoten, zustand.zeilen[index]);
        nachEingabe();
      });
    });

    // Grenzuebertritte gehoeren zur Auslandsreise — bei Inland waeren sie
    // ein Feld, das nie ausgefuellt werden darf. Solche Felder zeigt man nicht.
    knoten.querySelector(".paar-grenze").hidden = zustand.reisetyp !== "ausland";

    const menue = knoten.querySelector(".tag-menue");
    const aktionen = knoten.querySelector(".tag-aktionen");
    menue.addEventListener("click", () => {
      const auf = aktionen.hidden;
      aktionen.hidden = !auf;
      menue.setAttribute("aria-expanded", String(auf));
    });

    // Die Beschriftung sagt, was der Druck bewirkt -- nicht, in welchem
    // Zustand die Zeile gerade ist. Ein Schalter, der seinen Zustand
    // beschriftet, wird regelmaessig falsch herum gelesen.
    knoten.querySelector('[data-tat="rueckreise"]').textContent =
      zeile.ist_folgezeile ? "Markierung aufheben" : "Als Rückreise markieren";

    alle(".tag-aktionen button", knoten).forEach((knopf) => {
      knopf.addEventListener("click", () => {
        fuehreZeilenaktion(knopf.dataset.tat, index);
      });
    });

    zeichneBand(knoten, zeile);
    liste.append(knoten);
  });

  el("zeilen-zaehler").textContent = `${zustand.zeilen.length} von ${MAX_ZEILEN}`;
  el("knopf-zeile-neu").disabled = zustand.zeilen.length >= MAX_ZEILEN;
  el("knopf-folgezeile").disabled = zustand.zeilen.length >= MAX_ZEILEN
    || zustand.zeilen.length === 0;
}

function fuehreZeilenaktion(tat, index) {
  const zeilen = zustand.zeilen;
  if (tat === "rueckreise") {
    const zeile = zeilen[index];
    zeile.ist_folgezeile = !zeile.ist_folgezeile;
    // Eine Folgezeile traegt kein eigenes Datum -- so steht es in der
    // Excel-Vorlage, und so leert es auch die Desktop-App beim Umschalten.
    // Beim Zuruecknehmen bekommt die Zeile wieder eine Vorbelegung, sonst
    // stuende dort ein leeres Pflichtfeld.
    zeile.datum = zeile.ist_folgezeile ? null : vorschlagsdatum();
  } else if (tat === "entfernen") {
    zeilen.splice(index, 1);
  } else if (tat === "duplizieren") {
    if (zeilen.length >= MAX_ZEILEN) return;
    zeilen.splice(index + 1, 0, { ...zeilen[index] });
  } else if (tat === "hoch" && index > 0) {
    [zeilen[index - 1], zeilen[index]] = [zeilen[index], zeilen[index - 1]];
  } else if (tat === "runter" && index < zeilen.length - 1) {
    [zeilen[index + 1], zeilen[index]] = [zeilen[index], zeilen[index + 1]];
  }
  zeichneZeilen();
  nachEingabe();
}

/** Zeichnet das Tagesband: erfasste Zeiten massstaeblich auf der Tagesachse. */
function zeichneBand(knoten, zeile) {
  const spur = knoten.querySelector(".band-spur");
  spur.textContent = "";

  const teile = [
    ["arbeit", zeile.arbeit_von, zeile.arbeit_bis],
    ["passiv", zeile.passiv_von, zeile.passiv_bis],
    ["aktiv", zeile.aktiv_von, zeile.aktiv_bis],
  ];

  for (const [art, von, bis] of teile) {
    const a = minutenAusZeit(von);
    const b = minutenAusZeit(bis);
    // Ein halb gefuelltes Paar ist eine Eingabe mitten im Tippen, kein
    // Fehler — wir zeigen es einfach noch nicht an.
    if (a === null || b === null || b <= a) continue;
    const stueck = document.createElement("div");
    stueck.className = "band-teil";
    stueck.dataset.art = art;
    stueck.style.left = `${(a / 1440) * 100}%`;
    stueck.style.width = `${((b - a) / 1440) * 100}%`;
    spur.append(stueck);
  }

  // Grenzuebertritte sind Zeitpunkte, keine Spannen — sie bekommen einen
  // Strich statt eines Balkens.
  for (const feld of ["grenz_anreise", "grenz_rueckreise"]) {
    const m = minutenAusZeit(zeile[feld]);
    if (m === null) continue;
    const strich = document.createElement("div");
    strich.className = "band-teil";
    strich.dataset.art = "grenze";
    strich.style.left = `${(m / 1440) * 100}%`;
    strich.style.width = "2px";
    spur.append(strich);
  }
}

/* --------------------------------------------------------------------------
   Assistent: Seiten, Fortschritt, Zurueck und Weiter
   -------------------------------------------------------------------------- */

/** Baut die vier antippbaren Segmente unter dem Fortschrittsbalken. Einmal beim Start. */
function baueSegmente() {
  const liste = el("fortschritt-segmente");
  liste.textContent = "";
  SCHRITTE.forEach((schritt, i) => {
    const punkt = document.createElement("li");
    const knopf = document.createElement("button");
    knopf.type = "button";
    knopf.dataset.schritt = String(i + 1);
    knopf.textContent = schritt.name;
    knopf.setAttribute("aria-label", `Schritt ${i + 1} von ${SCHRITTE.length}: ${schritt.name}`);
    knopf.addEventListener("click", () => zeigeSchritt(i + 1, { fokus: true }));
    punkt.append(knopf);
    liste.append(punkt);
  });
}

/**
 * Ordnet eine Pruefmeldung der Seite zu, auf der man sie beheben kann.
 *
 * Die Feldnamen kommen aus glaz/validation.py: ``profil.*`` fuer das Profil,
 * ``vorgang.einsatzart`` fuer den Vorgang, ``zeile[n].*`` und
 * ``vorgang.zeilen`` fuer die Reisetage, ``abschluss.zieldatei`` fuer die
 * gewaehlte Datei. Unbekanntes landet beim Abschluss -- dort steht auch die
 * Statuszeile, die es im Klartext nennt.
 */
function schrittFuerIssue(issue) {
  const feld = String(issue.feld || "");
  if (feld.startsWith("profil.")) return 1;
  if (feld === "vorgang.einsatzart") return 2;
  if (feld.startsWith("zeile[") || feld === "vorgang.zeilen") return 3;
  if (issue.zeile !== null && issue.zeile !== undefined) return 3;
  return 4;
}

/** Fehler je Schritt aus dem letzten Pruefergebnis, als Feld [n1, n2, n3, n4]. */
function fehlerJeSchritt(erg) {
  const zaehler = SCHRITTE.map(() => 0);
  if (!erg) return zaehler;
  for (const issue of erg.issues) {
    if (issue.schwere !== "fehler") continue;
    zaehler[schrittFuerIssue(issue) - 1] += 1;
  }
  return zaehler;
}

/**
 * Zeichnet die Fortschrittsleiste fuer den aktuellen Schritt.
 *
 * Der Balken fuellt sich mit ``schritt / anzahl``: Auf der letzten Seite
 * steht er voll, weil dort die Datei entstehen kann -- das ist die Frage, die
 * er beantwortet. Die Segmente tragen drei Zustaende (fertig, aktuell, offen)
 * und zusaetzlich die Fehlermarkierung aus der letzten Pruefung.
 */
function zeichneFortschritt() {
  const nr = zustand.schritt;
  const anzahl = SCHRITTE.length;
  const name = SCHRITTE[nr - 1].name;

  el("fortschritt-schritt").textContent = `Schritt ${nr} von ${anzahl}`;
  el("fortschritt-name").textContent = name;
  el("fortschritt-fuellung").style.width = `${(nr / anzahl) * 100}%`;

  const balken = el("fortschritt-balken");
  balken.setAttribute("aria-valuenow", String(nr));
  balken.setAttribute("aria-valuetext", `Schritt ${nr} von ${anzahl}: ${name}`);

  const fehler = fehlerJeSchritt(letztePruefung);
  alle("#fortschritt-segmente button").forEach((knopf, i) => {
    const eigener = i + 1;
    knopf.dataset.stand = eigener < nr ? "fertig" : eigener === nr ? "aktuell" : "offen";
    knopf.dataset.fehler = String(fehler[i] > 0);
    if (eigener === nr) knopf.setAttribute("aria-current", "step");
    else knopf.removeAttribute("aria-current");
  });
}

/**
 * Zeigt genau eine Seite des Assistenten und blendet alles andere aus.
 *
 * ``fokus`` setzt den Schreibcursor auf die Ueberschrift der neuen Seite:
 * Fuer Screenreader ist das die Ansage "du bist jetzt bei Reisetage", fuer
 * alle anderen ist es unsichtbar. Der Bildlauf springt nach oben, damit jede
 * Seite mit ihrer Ueberschrift beginnt und nicht mitten im Inhalt.
 */
function zeigeSchritt(nr, { fokus = false } = {}) {
  nr = Math.min(Math.max(1, nr), SCHRITTE.length);
  zustand.schritt = nr;

  SCHRITTE.forEach((schritt, i) => { el(schritt.id).hidden = i + 1 !== nr; });
  SONDERSEITEN.forEach((id) => { el(id).hidden = true; });

  const letzter = nr === SCHRITTE.length;
  el("knopf-zurueck").disabled = nr === 1;
  el("knopf-weiter").hidden = letzter;
  el("knopf-abschluss").hidden = !letzter;
  el("leiste-knoepfe").hidden = false;
  el("fortschritt").hidden = false;

  zeichneFortschritt();
  speichern();
  if (nr === 1) klappeProfilAufWennUnvollstaendig();

  window.scrollTo({ top: 0, behavior: "auto" });
  if (fokus) {
    const titel = el(SCHRITTE[nr - 1].id).querySelector("h2");
    if (titel) titel.focus({ preventScroll: true });
  }
}

/**
 * Oeffnet die Profilfelder, wenn im Profil noch etwas fehlt.
 *
 * Zugeklappt zeigt Seite 1 nur die Auswahlliste und eine Zeile Text -- fuer
 * jemanden mit fertigem Profil genau richtig, fuer den ersten Start eine fast
 * leere Seite mit einem roten Zaehler. Aufgerufen wird das nur beim Betreten
 * der Seite und einmal nach der ersten Pruefung, nicht bei jedem Tastendruck:
 * Wer die Felder trotz Fehler zuklappt, soll sie nicht sofort wieder offen
 * vorfinden.
 */
function klappeProfilAufWennUnvollstaendig() {
  if (!letztePruefung) return;
  if (fehlerJeSchritt(letztePruefung)[0] === 0) return;
  const felder = el("profil-felder");
  if (!felder.hidden) return;
  felder.hidden = false;
  el("knopf-profil-auf").setAttribute("aria-expanded", "true");
  el("knopf-profil-auf").textContent = "Fertig";
  zeigePruefung(letztePruefung);
}

/** Zeigt Einstellungen oder Selbsttest anstelle des aktuellen Schritts. */
function zeigeSonderseite(id) {
  SCHRITTE.forEach((schritt) => { el(schritt.id).hidden = true; });
  SONDERSEITEN.forEach((andere) => { el(andere).hidden = andere !== id; });
  // Weiter/Zurueck und die Fortschrittsleiste gehoeren zum Ablauf, nicht zu
  // den Einstellungen -- "Schritt 4 von 4" ueber den Mailvorlagen waere eine
  // falsche Auskunft. Die Statuszeile bleibt: Sie meldet auch hier, wenn der
  // Kern stolpert.
  el("leiste-knoepfe").hidden = true;
  el("fortschritt").hidden = true;
  window.scrollTo({ top: 0, behavior: "auto" });
  const titel = el(id).querySelector("h2");
  if (titel) { titel.tabIndex = -1; titel.focus({ preventScroll: true }); }
}

/* --------------------------------------------------------------------------
   Oberflaeche: Profil, Vorgang, Einstellungen
   -------------------------------------------------------------------------- */

function zeichneProfil() {
  const auswahl = el("profil-auswahl");
  auswahl.textContent = "";
  zustand.profile.forEach((p, i) => {
    const eintrag = document.createElement("option");
    eintrag.value = String(i);
    eintrag.textContent = profilLabel(p);
    auswahl.append(eintrag);
  });
  auswahl.value = String(zustand.aktivesProfil);

  const p = aktuellesProfil();
  // Nur zuweisen, wenn sich der Wert wirklich unterscheidet: Eine Zuweisung an
  // .value setzt in mehreren Browsern den Schreibcursor ans Ende. Beim Tippen
  // am Zeilenende faellt das nicht auf, beim Korrigieren mitten im Wort sehr
  // wohl -- und diese Funktion laeuft nach jedem Tastendruck.
  const setzeWenn = (id, wert) => {
    const feld = el(id);
    if (feld.value !== wert) feld.value = wert;
  };
  setzeWenn("f-profilname", p.name);
  setzeWenn("f-vorname", p.vorname);
  setzeWenn("f-nachname", p.nachname);
  setzeWenn("f-personalnummer", p.personalnummer);
  setzeWenn("f-abteilung", p.abteilung);
  setzeWenn("f-gl-name", p.gruppenleiter_name);
  setzeWenn("f-gl-email", p.gruppenleiter_email);

  const teile = [];
  const person = `${p.vorname} ${p.nachname}`.trim();
  if (person) teile.push(person);
  if (p.personalnummer) teile.push(`Pers.-Nr. ${p.personalnummer}`);
  if (p.abteilung) teile.push(p.abteilung);
  el("profil-zusammenfassung").textContent =
    teile.length ? teile.join(" · ") : "Noch nichts hinterlegt — tippe auf Bearbeiten.";

  if (typeof GLAZ_GRUPPENLEITER !== "undefined") GLAZ_GRUPPENLEITER.zeichne();
}

function zeichneEinsatzarten() {
  const liste = el("einsatzarten");
  liste.textContent = "";
  for (const art of zustand.einsatzarten) {
    const eintrag = document.createElement("option");
    eintrag.value = art;
    liste.append(eintrag);
  }
}

function zeichneUmschalter() {
  alle("[data-reisetyp]").forEach((k) =>
    k.setAttribute("aria-checked", String(k.dataset.reisetyp === zustand.reisetyp))
  );
  alle("[data-modus]").forEach((k) =>
    k.setAttribute("aria-checked", String(k.dataset.modus === zustand.modus))
  );
  el("seite-neu").hidden = zustand.modus !== "neu";
  el("seite-zieldatei").hidden = zustand.modus === "neu";
}

function zeichneEinstellungen() {
  el("f-empfaenger").value = zustand.einstellungen.empfaenger;
  el("f-betreff-vorlage").value = zustand.einstellungen.betreff_vorlage;
  el("f-body-vorlage").value = zustand.einstellungen.body_vorlage;
  pruefeEinstellungen();
}

/* Dasselbe Muster wie EMAIL_MUSTER in web/gruppenleiter.js und EMAIL_RE in
   glaz/model.py. Wer es dort aendert, aendert es auch hier. */
const EMAIL_MUSTER = /^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$/;

/** Die Adressen eines Empfaengerfelds, das mehrere mit , oder ; trennen darf. */
function adressenAus(text) {
  return String(text || "").split(/[,;]/).map((a) => a.trim()).filter(Boolean);
}

function ungueltigeAdressen(text) {
  return adressenAus(text).filter((a) => !EMAIL_MUSTER.test(a));
}

/**
 * Prueft die Mail-Einstellungen, waehrend man sie tippt -- statt erst beim
 * Senden. Ein unbekannter Platzhalter fiel bisher erst auf, nachdem die Datei
 * schon erzeugt war.
 *
 * Nebenbei zeigt das leere Empfaengerfeld die Vorgabe als Platzhalter: Leer
 * heisst "Vorgabe verwenden", und das soll man sehen.
 */
function pruefeEinstellungen() {
  const hinweis = el("einstellungen-hinweis");
  const zeige = (text, schwere) => {
    hinweis.textContent = text;
    if (schwere) hinweis.dataset.schwere = schwere;
    else delete hinweis.dataset.schwere;
  };

  const falsch = ungueltigeAdressen(zustand.einstellungen.empfaenger);
  el("f-empfaenger").toggleAttribute("aria-invalid", falsch.length > 0);
  if (falsch.length) {
    zeige(`„${falsch[0]}“ ist keine gültige E-Mail-Adresse.`, "fehler");
    return;
  }
  if (!kernBereit) {
    zeige("");
    return;
  }

  let texte;
  try {
    texte = mailtexte(JSON.stringify(vorgangDict()));
  } catch (e) {
    zeige(e && e.message ? e.message : String(e), "fehler");
    return;
  }
  const vorgabe = zustand.einstellungen.empfaenger.trim() ? "" : texte.an;
  el("f-empfaenger").placeholder = vorgabe ? `Vorgabe: ${vorgabe}` : "name@firma.de";
  if (!texte.an.trim()) {
    zeige("Noch kein Empfänger — ohne ihn lässt sich nicht über Outlook senden.", "warnung");
  } else {
    zeige("");
  }
}

const pruefeEinstellungenGleich = entprellt(pruefeEinstellungen, 250);

/* --------------------------------------------------------------------------
   Pruefung
   -------------------------------------------------------------------------- */

let letztePruefung = null;

let gehalteneMeldungBis = 0;

/** Wie melde(), aber die naechsten Sekunden ueberschreibt die Pruefung nicht. */
function meldeGehalten(text, schwere = "", ms = 5000) {
  gehalteneMeldungBis = Date.now() + ms;
  melde(text, schwere);
}

function melde(text, schwere = "") {
  const feld = el("leiste-meldung");
  // Die Statuszeile ist aria-live. Dieselbe Meldung nach jedem Tastendruck
  // neu zu setzen, liesse VoiceOver sie jedes Mal wieder vorlesen.
  if (feld.textContent === text && (feld.dataset.schwere || "") === schwere) return;
  feld.textContent = text;
  if (schwere) feld.dataset.schwere = schwere;
  else delete feld.dataset.schwere;
}

/**
 * Wendet das gewaehlte Erscheinungsbild an.
 *
 * Das Attribut am <html>-Element steuert die Palette in app.css. Gesetzt wird
 * es bereits im Kopf der Seite, damit beim Start nichts aufblitzt; diese
 * Funktion ist fuer das spaetere Umschalten zustaendig.
 *
 * Mitgefuehrt wird die Farbe der Statusleiste: Auf dem iPhone faerbt sie den
 * Bereich um die Uhrzeit. Bliebe sie stehen, saesse ueber einer hellen Seite
 * ein tiefblauer Balken.
 */
function wendeThemaAn(thema) {
  const gewaehlt = thema === "hell" ? "hell" : "dunkel";
  document.documentElement.dataset.thema = gewaehlt;

  const marke = document.querySelector('meta[name="theme-color"]');
  if (marke) marke.setAttribute("content", gewaehlt === "hell" ? "#EEF2F7" : "#060E1A");

  alle("[data-thema]", el("menue")).forEach((knopf) =>
    knopf.setAttribute("aria-checked", String(knopf.dataset.thema === gewaehlt))
  );
}

function setzeBereitschaft(stand, text) {
  el("bereitschaftspunkt").dataset.stand = stand;
  el("bereitschaftstext").textContent = text;
}

function pruefeJetzt() {
  if (!kernBereit) return;
  let ergebnis;
  const roh = JSON.stringify(vorgangDict());
  try {
    ergebnis = rufe("bruecke_pruefe", roh);

    // Die Zieldatei gehoert in dieselbe laufende Pruefung. Lief sie erst beim
    // Druck auf den Abschlussknopf, erfuhr man "die Datei ist voll" in dem
    // Moment, in dem man senden wollte -- und nicht, als man sie auswaehlte.
    if (zustand.modus !== "neu" && zieldateiGeladen) {
      const ziel = rufe("bruecke_pruefe_ziel", roh, "/ziel.xlsx", zustand.modus);
      ergebnis.issues = ergebnis.issues.concat(ziel.issues);
      ergebnis.absendbar = ergebnis.absendbar && ziel.absendbar;
    }
  } catch (e) {
    melde(`Die Prüfung ist gestolpert: ${e.message}`, "fehler");
    return;
  }
  letztePruefung = ergebnis;
  zeigePruefung(ergebnis);
}

const pruefeGleich = entprellt(pruefeJetzt, 180);

function zeigePruefung(erg) {
  // Meldungen an den Zeilen anschreiben
  alle(".tag").forEach((knoten) => {
    knoten.querySelector(".tag-meldungen").textContent = "";
    alle(".zeit, .tag-datum", knoten).forEach((f) => {
      f.removeAttribute("aria-invalid");
      delete f.dataset.warnung;
    });
  });

  const uebrige = [];
  for (const issue of erg.issues) {
    const knoten =
      issue.zeile !== null && issue.zeile !== undefined
        ? document.querySelector(`.tag[data-index="${issue.zeile}"]`)
        : null;
    if (!knoten) {
      uebrige.push(issue);
      continue;
    }
    const punkt = document.createElement("li");
    punkt.dataset.schwere = issue.schwere;
    punkt.textContent = issue.meldung;
    knoten.querySelector(".tag-meldungen").append(punkt);

    // Das betroffene Feld markieren, sofern die Meldung eines benennt.
    const treffer = /\.([a-z_]+)$/.exec(issue.feld || "");
    if (treffer) {
      const feld = knoten.querySelector(`[data-feld="${treffer[1]}"]`)
        || (treffer[1] === "datum" ? knoten.querySelector(".tag-datum") : null);
      if (feld) {
        if (issue.schwere === "fehler") feld.setAttribute("aria-invalid", "true");
        else feld.dataset.warnung = "true";
      }
    }
  }

  // Profilfelder markieren
  const profilFelder = {
    "profil.vorname": "f-vorname",
    "profil.nachname": "f-nachname",
    "profil.personalnummer": "f-personalnummer",
    "profil.abteilung": "f-abteilung",
    "profil.gruppenleiter_name": "f-gl-name",
    "profil.gruppenleiter_email": "f-gl-email",
    "vorgang.einsatzart": "f-einsatzart",
  };
  Object.values(profilFelder).forEach((id) => {
    const f = el(id);
    if (f) { f.removeAttribute("aria-invalid"); delete f.dataset.warnung; }
  });
  for (const issue of erg.issues) {
    const id = profilFelder[issue.feld];
    if (!id) continue;
    const f = el(id);
    if (!f) continue;
    if (issue.schwere === "fehler") f.setAttribute("aria-invalid", "true");
    else f.dataset.warnung = "true";
  }

  // Zaehler am Blockkopf, solange der Profilbereich zugeklappt ist.
  const profilFehler = erg.issues.filter(
    (i) => i.schwere === "fehler" && String(i.feld || "").startsWith("profil.")
  ).length;
  const zaehler = el("profil-fehlerzahl");
  const zugeklappt = el("profil-felder").hidden;
  zaehler.hidden = !(profilFehler && zugeklappt);
  zaehler.dataset.schwere = "fehler";
  zaehler.textContent =
    profilFehler === 1 ? "1 Angabe fehlt" : `${profilFehler} Angaben fehlen`;

  el("kopf-zeitraum").textContent = erg.zeitraum_text || "Noch keine Zeiten erfasst";
  el("dateiname-vorschau").textContent = erg.dateiname || "—";

  // Fehlermarkierungen an den Segmenten der Fortschrittsleiste nachfuehren.
  zeichneFortschritt();

  const hinweis = el("arbeitszeit-hinweis");
  hinweis.hidden = !erg.hat_arbeit_vor_aktivreise;
  if (erg.hat_arbeit_vor_aktivreise) {
    hinweis.textContent =
      "An mindestens einem Tag liegt Arbeitszeit vor der aktiven Reisezeit. "
      + "Die Mail bekommt dazu einen zusätzlichen Absatz.";
  }

  // Statuszeile: erst die blockierenden Fehler, dann die Warnungen.
  const fehler = erg.issues.filter((i) => i.schwere === "fehler");
  const warnungen = erg.issues.filter((i) => i.schwere === "warnung");
  const knopf = el("knopf-abschluss");
  // Waehrend eines Laufs bleibt der Knopf gesperrt, auch wenn eine
  // nachlaufende Pruefung "absendbar" meldet -- sonst startete ein zweiter
  // Tipp einen zweiten Anhang an dieselbe Datei.
  knopf.disabled = !erg.absendbar || !!knopf.dataset.laeuft;

  // Eine gerade gegebene Auskunft ("… geladen", "der vorige liegt unter
  // Zwischenstaende") ein paar Sekunden stehen lassen. Offene Punkte zeigen
  // die Felder selbst, und nach einem neuen Vorgang sind sie ohnehin erwartet.
  if (Date.now() < gehalteneMeldungBis) return;

  if (fehler.length) {
    const erster = uebrige.find((i) => i.schwere === "fehler") || fehler[0];
    melde(
      fehler.length === 1 ? erster.meldung : `${fehler.length} offene Punkte: ${erster.meldung}`,
      "fehler"
    );
  } else if (warnungen.length) {
    melde(
      warnungen.length === 1
        ? warnungen[0].meldung
        : `${warnungen.length} Hinweise — der Versand ist möglich.`,
      "warnung"
    );
  } else {
    melde("Alles vollständig.", "erfolg");
  }
}

let profilHinweisKennung = null;

/**
 * Zeigt eine Meldung unter den Profilfeldern und nimmt sie danach zurueck.
 *
 * Zurueckgenommen wird sie, weil eine stehengebliebene Erfolgsmeldung beim
 * naechsten Blick etwas behauptet, das laengst nicht mehr stimmt. Dieselbe
 * Ueberlegung wie in _setze_profil_hinweis der Desktop-App, dort ebenfalls
 * mit sechs Sekunden.
 */
function setzeProfilHinweis(text, schwere = "") {
  const feld = el("profil-hinweis");
  feld.textContent = text;
  if (schwere) feld.dataset.schwere = schwere;
  else delete feld.dataset.schwere;

  clearTimeout(profilHinweisKennung);
  if (text !== PROFIL_HINWEIS_STANDARD) {
    profilHinweisKennung = setTimeout(
      () => setzeProfilHinweis(PROFIL_HINWEIS_STANDARD),
      6000
    );
  }
}

/* --------------------------------------------------------------------------
   Zwischenstaende

   Der laufende Stand wird ohnehin bei jeder Eingabe gesichert (speichern()).
   Was fehlte, war ein Ort fuer Vorgaenge, an denen man gerade NICHT
   arbeitet: "Neuer Vorgang" warf den offenen Stand bisher weg. Jetzt legt es
   ihn hier ab, ebenso "Laden" eines anderen Stands. Ein Vorgang ist genau ein
   Eintrag -- wer einen Stand laedt und weiterarbeitet, aktualisiert beim
   naechsten Sichern denselben Eintrag, statt Duplikate anzuhaeufen.
   -------------------------------------------------------------------------- */

/* Genug fuer ein Vierteljahr Dienstreisen; aeltere fallen hinten heraus. */
const MAX_ZWISCHENSTAENDE = 30;

function neueZwischenstandId() {
  return `zs-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

/** Nur die Teile, die einen Vorgang ausmachen -- das Profil gilt dauerhaft. */
function aktuellerVorgangsinhalt() {
  return {
    einsatzart: zustand.einsatzart,
    reisetyp: zustand.reisetyp,
    modus: zustand.modus,
    zeilen: zustand.zeilen.map((z) => ({ ...z })),
  };
}

/** Wie hat_inhalt in glaz/session.py: das Datum zaehlt nicht, es ist vorbelegt. */
function hatVorgangsinhalt(inhalt) {
  return !!String(inhalt.einsatzart || "").trim()
    || inhalt.zeilen.some((z) => ZEITFELDER.some((f) => z[f]));
}

function gleicherInhalt(a, b) {
  const kern = (v) => JSON.stringify([v.einsatzart, v.reisetyp, v.modus, v.zeilen]);
  return kern(a) === kern(b);
}

/** Haertet einen Eintrag aus dem localStorage -- wie laden() den Rest. */
function bereinigeZwischenstand(e) {
  if (!e || typeof e !== "object" || typeof e.id !== "string") return null;
  const zeilen = Array.isArray(e.zeilen)
    ? e.zeilen.filter((z) => z && typeof z === "object").slice(0, MAX_ZEILEN)
      .map((z) => ({ ...leereZeile(), ...z }))
    : [];
  return {
    id: e.id,
    gesichert_am: typeof e.gesichert_am === "string" ? e.gesichert_am : "",
    versendet_am: typeof e.versendet_am === "string" ? e.versendet_am : "",
    einsatzart: typeof e.einsatzart === "string" ? e.einsatzart : "",
    reisetyp: e.reisetyp === "ausland" ? "ausland" : "inland",
    modus: ["neu", "ergaenzen", "nur_versenden"].includes(e.modus) ? e.modus : "neu",
    zeilen: zeilen.length ? zeilen : [leereZeile()],
  };
}

/**
 * Legt den laufenden Stand in der Liste ab (oder aktualisiert seinen
 * Eintrag). Liefert den Eintrag, oder null, wenn es nichts zu sichern gab.
 *
 * ``versendet`` markiert den Eintrag als verschickt. Aendert man danach noch
 * etwas und sichert erneut, faellt die Markierung weg -- sie galt dem alten
 * Inhalt.
 */
function sichereZwischenstand({ versendet = false } = {}) {
  const inhalt = aktuellerVorgangsinhalt();
  if (!hatVorgangsinhalt(inhalt)) return null;

  const id = zustand.zwischenstandId || neueZwischenstandId();
  const alt = zustand.zwischenstaende.find((e) => e.id === id);
  const jetzt = new Date().toISOString();
  const eintrag = {
    id,
    gesichert_am: jetzt,
    versendet_am: versendet
      ? jetzt
      : (alt && alt.versendet_am && gleicherInhalt(alt, inhalt) ? alt.versendet_am : ""),
    ...inhalt,
  };
  zustand.zwischenstaende = [eintrag]
    .concat(zustand.zwischenstaende.filter((e) => e.id !== id))
    .slice(0, MAX_ZWISCHENSTAENDE);
  zustand.zwischenstandId = id;
  speichern();
  return eintrag;
}

/** Setzt einen Vorgangsinhalt ins Formular (Neuer Vorgang, Laden). */
function setzeVorgang(inhalt, id) {
  zustand.einsatzart = inhalt.einsatzart;
  zustand.reisetyp = inhalt.reisetyp;
  zustand.modus = inhalt.modus;
  zustand.zeilen = inhalt.zeilen.map((z) => ({ ...z }));
  zustand.zwischenstandId = id;
  // Eine gewaehlte Zieldatei und der Merker des letzten Anhangs gehoeren
  // zum vorigen Vorgang.
  zieldatei = null;
  zieldateiGeladen = false;
  letzterAnhang = null;
  el("f-zieldatei").value = "";
  el("zieldatei-info").textContent = "Noch keine Datei gewählt.";
  el("f-einsatzart").value = zustand.einsatzart;
  zeichneUmschalter();
  zeichneZeilen();
  aktualisiereKnopftext();
}

function ladeZwischenstand(id) {
  const eintrag = zustand.zwischenstaende.find((e) => e.id === id);
  if (!eintrag) return;
  // Der offene Stand geht nicht verloren: er wird vorher abgelegt.
  const abgelegt = zustand.zwischenstandId !== id ? sichereZwischenstand() : null;
  setzeVorgang(eintrag, id);
  zeigeSchritt(eintrag.zeilen.length ? 3 : 2, { fokus: true });
  nachEingabe();
  meldeGehalten(
    abgelegt
      ? `„${zwischenstandTitel(eintrag)}“ geladen. Der vorige Stand liegt unter Zwischenstände.`
      : `„${zwischenstandTitel(eintrag)}“ geladen.`,
    "erfolg"
  );
}

function loescheZwischenstand(id) {
  const eintrag = zustand.zwischenstaende.find((e) => e.id === id);
  if (!eintrag) return;
  if (!confirm(`Zwischenstand „${zwischenstandTitel(eintrag)}“ löschen?`)) return;
  zustand.zwischenstaende = zustand.zwischenstaende.filter((e) => e.id !== id);
  // Der laufende Stand bleibt stehen; er gehoert nur keinem Eintrag mehr.
  if (zustand.zwischenstandId === id) zustand.zwischenstandId = null;
  speichern();
  zeichneZwischenstaende();
}

function zwischenstandTitel(e) {
  return String(e.einsatzart || "").trim() || "Ohne Einsatzart";
}

function kurzDatum(iso) {
  const [j, m, t] = String(iso).split("-");
  return t && m ? `${t}.${m}.${j}` : "";
}

/** "21.09.–23.09.2026 · 3 Reisetage" -- aus den Zeilen mit Zeiten. */
function zwischenstandZeitraum(e) {
  const aktive = e.zeilen.filter((z) => ZEITFELDER.some((f) => z[f]));
  const daten = aktive.map((z) => z.datum).filter(Boolean).sort();
  const teile = [];
  if (daten.length) {
    const von = kurzDatum(daten[0]);
    const bis = kurzDatum(daten[daten.length - 1]);
    teile.push(von === bis ? von : `${von.slice(0, 6)}–${bis}`);
  }
  teile.push(aktive.length === 1 ? "1 Reisetag" : `${aktive.length} Reisetage`);
  return teile.join(" · ");
}

function zeitpunktText(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const zwei = (n) => String(n).padStart(2, "0");
  return `${zwei(d.getDate())}.${zwei(d.getMonth() + 1)}. ${zwei(d.getHours())}:${zwei(d.getMinutes())}`;
}

function zeichneZwischenstaende() {
  const liste = el("zwischenstand-liste");
  liste.textContent = "";
  const hinweis = el("zwischenstand-hinweis");
  hinweis.textContent = zustand.zwischenstaende.length ? "" : "Noch keine Zwischenstände.";

  for (const e of zustand.zwischenstaende) {
    const punkt = document.createElement("li");
    punkt.dataset.stand = e.versendet_am ? "versendet" : "offen";
    const aktuell = e.id === zustand.zwischenstandId;
    punkt.dataset.aktuell = String(aktuell);

    const text = document.createElement("div");
    text.className = "zs-text";
    const titel = document.createElement("b");
    titel.textContent = zwischenstandTitel(e) + (aktuell ? " (in Bearbeitung)" : "");
    const zeile = document.createElement("span");
    const wann = e.versendet_am
      ? `versendet ${zeitpunktText(e.versendet_am)}`
      : `gesichert ${zeitpunktText(e.gesichert_am)}`;
    zeile.textContent = `${zwischenstandZeitraum(e)} · ${wann}`;
    text.append(titel, zeile);

    const knoepfe = document.createElement("div");
    knoepfe.className = "zs-knoepfe";
    if (!aktuell) {
      const laden = document.createElement("button");
      laden.type = "button";
      laden.className = "textknopf";
      laden.textContent = "Laden";
      laden.setAttribute("aria-label", `${zwischenstandTitel(e)} laden`);
      laden.addEventListener("click", () => ladeZwischenstand(e.id));
      knoepfe.append(laden);
    }
    const weg = document.createElement("button");
    weg.type = "button";
    weg.className = "textknopf textknopf-warnend";
    weg.textContent = "Löschen";
    weg.setAttribute("aria-label", `${zwischenstandTitel(e)} löschen`);
    weg.addEventListener("click", () => loescheZwischenstand(e.id));
    knoepfe.append(weg);

    punkt.append(text, knoepfe);
    liste.append(punkt);
  }
}

function nachEingabe() {
  // Was jetzt noch geaendert wurde, steht nicht in der Datei in den Downloads.
  verwerfeOutlookEntwurf();
  speichern();
  pruefeGleich();
}

/* --------------------------------------------------------------------------
   Abschluss: Datei erzeugen und ins Teilen-Menue geben
   -------------------------------------------------------------------------- */

async function abschluss() {
  if (el("knopf-abschluss").dataset.laeuft) return;
  if (!kernBereit) {
    melde("Der Rechenkern ist noch nicht bereit.", "warnung");
    return;
  }
  const knopf = el("knopf-abschluss");
  knopf.disabled = true;
  knopf.dataset.laeuft = "true";
  knopf.textContent = "Wird erzeugt …";

  try {
    // Eine noch ausstehende, entprellte Pruefung jetzt nachholen: Der
    // Dateiname kommt aus ihr, und er soll zur letzten Eingabe passen.
    pruefeJetzt();
    const roh = JSON.stringify(vorgangDict());
    let ergebnis;
    let pfad;

    // Beim Weg ueber Outlook wird der Empfaenger vor dem Erzeugen geprueft:
    // Ohne ihn waere das Einzige, was der Outlook-Weg dem Teilen-Menue voraus
    // hat, leer -- und die Datei laege umsonst in den Downloads.
    const ueberOutlook = zustand.mailMitsenden && zustand.versandweg === "outlook";
    const texte = ueberOutlook ? mailtexte(roh) : null;
    if (ueberOutlook && !texte.an.trim()) {
      melde(
        "Kein Empfänger hinterlegt. Trag ihn unter Menü → Einstellungen → Empfänger ein.",
        "fehler"
      );
      return;
    }
    const falsch = ueberOutlook ? ungueltigeAdressen(texte.an) : [];
    if (falsch.length) {
      melde(
        `Der Empfänger „${falsch[0]}“ ist keine gültige E-Mail-Adresse (Menü → Einstellungen).`,
        "fehler"
      );
      return;
    }

    if (zustand.modus === "neu") {
      // Der Zielpfad im virtuellen Dateisystem bestimmt, wie die Datei am Ende
      // heisst: erzeuge_xlsx liefert den Namen dieses Pfades zurueck, und der
      // wandert ins Teilen-Menue. Ein Arbeitsname wie "/ausgabe.xlsx" kaeme
      // also beim Empfaenger an -- obwohl die Vorschau darueber den richtigen
      // Namen anzeigt. Deshalb kommt der Name aus derselben Quelle wie die
      // Vorschau: glaz/filename.py, ueber das Ergebnis der Pruefung.
      const gewuenscht =
        (letztePruefung && letztePruefung.dateiname) || "GLAZ-Korrekturbuchungsliste.xlsx";
      pfad = "/" + gewuenscht;
      ergebnis = rufe("bruecke_erzeuge", roh, pfad);
    } else {
      if (!zieldatei) {
        melde("Wähle zuerst die vorhandene Liste aus.", "fehler");
        return;
      }
      // Eine Arbeitskopie statt /ziel.xlsx: bruecke_ergaenze schreibt in die
      // Datei, die es bekommt. Laege der neue Block danach in /ziel.xlsx,
      // zaehlte die laufende Pruefung ihn als belegt mit -- und sperrte den
      // Knopf genau fuer den Korrekturlauf, fuer den letzterAnhang da ist.
      pfad = "/arbeit.xlsx";
      py.FS.writeFile(pfad, new Uint8Array(await zieldatei.arrayBuffer()));

      const zielpruefung = rufe("bruecke_pruefe_ziel", roh, pfad, zustand.modus);
      const zielfehler = zielpruefung.issues.filter((i) => i.schwere === "fehler");
      if (zielfehler.length) {
        melde(zielfehler[0].meldung, "fehler");
        return;
      }
      if (zustand.modus === "ergaenzen") {
        // Liegt ein Block aus einem vorherigen Lauf in genau dieser Datei,
        // wird er ueberschrieben statt ein zweiter danebengesetzt. Dieselbe
        // Ueberlegung wie im Aktualisieren-Zustand der Desktop-App: Wer nach
        // dem Versand noch etwas korrigiert, will eine berichtigte Liste --
        // keine zweite Eintragung derselben Reise.
        const merker = letzterAnhang && letzterAnhang.name === zieldatei.name
          ? letzterAnhang
          : null;
        ergebnis = rufe(
          "bruecke_ergaenze", roh, pfad,
          merker ? merker.start_row : null,
          merker ? merker.zeilen : 0
        );
        letzterAnhang = {
          name: zieldatei.name,
          start_row: ergebnis.start_row,
          zeilen: ergebnis.geschriebene_zeilen,
        };
      } else {
        ergebnis = { pfad, dateiname: zieldatei.name, bytes: zieldatei.size };
      }
    }

    const bytes = py.FS.readFile(ergebnis.pfad || pfad);
    const blob = new Blob([bytes], { type: XLSX_TYP });
    // Beim Ergaenzen behaelt die Datei ihren Namen. Ein frisch erzeugter Name
    // waere hier falsch: der Empfaenger bekaeme dieselbe Liste unter wechselnden
    // Namen und koennte die Fassungen nicht mehr auseinanderhalten.
    const dateiname =
      zustand.modus === "neu"
        ? ergebnis.dateiname || "GLAZ-Korrekturbuchungsliste.xlsx"
        : (zieldatei && zieldatei.name) || ergebnis.dateiname
          || "GLAZ-Korrekturbuchungsliste.xlsx";

    if (ueberOutlook) {
      versendeUeberOutlook(blob, dateiname, texte);
    } else if (zustand.mailMitsenden) {
      await teileDatei(blob, dateiname, roh);
    } else {
      speichereDatei(blob, dateiname);
      melde(`${dateiname} wurde gesichert.`, "erfolg");
    }
  } catch (e) {
    melde(`Es hat nicht geklappt: ${e && e.message ? e.message : e}`, "fehler");
  } finally {
    // Nicht einfach die alte Beschriftung zurueck: Nach dem Outlook-Weg heisst
    // der Knopf jetzt "In Outlook öffnen" (siehe aktualisiereKnopftext).
    aktualisiereKnopftext();
    delete knopf.dataset.laeuft;
    // Die Pruefung stellt den Knopf wieder scharf -- und schriebe dabei
    // "Alles vollstaendig." ueber "Geteilt." Das Ergebnis des Laufs ist aber
    // die Auskunft, auf die man in diesem Moment wartet; sie bleibt stehen.
    const ergebnisText = el("leiste-meldung").textContent;
    const ergebnisSchwere = el("leiste-meldung").dataset.schwere || "";
    pruefeJetzt();
    if (ergebnisText) melde(ergebnisText, ergebnisSchwere);
  }
}

/** Adressen, Betreff und Text der Mail -- gefuellt von glaz/mailtext.py. */
function mailtexte(roh) {
  return rufe(
    "bruecke_mailtexte",
    roh,
    zustand.einstellungen.empfaenger || "",
    zustand.einstellungen.betreff_vorlage || "",
    zustand.einstellungen.body_vorlage || ""
  );
}

/* Der vorbereitete Outlook-Aufruf des letzten Laufs, oder null. Er lebt
   bewusst ausserhalb von `zustand`: Er gehoert zu genau der Datei, die gerade
   in den Downloads liegt, und wird mit der naechsten Eingabe wertlos. */
let outlookEntwurf = null;

/**
 * Baut den Aufruf, mit dem die Outlook-App eine neue Mail oeffnet.
 *
 * Das Schema ist Microsofts dokumentierter Einstieg fuer Outlook auf iOS und
 * Android. Es fuellt An, CC, Betreff und Text vor -- einen Anhang kennt es
 * nicht, und eine Webseite hat auch keinen anderen Weg, Outlook eine Datei
 * UND Adressen zugleich zu geben: Das Teilen-Menue reicht die Datei weiter,
 * aber keine Adressen. Deshalb bleibt ein Handgriff, die Bueroklammer.
 */
function outlookAdresse(texte) {
  const teile = [["to", texte.an]];
  if (texte.cc) teile.push(["cc", texte.cc]);
  teile.push(["subject", texte.betreff], ["body", texte.body]);
  return "ms-outlook://compose?" + teile
    .map(([schluessel, wert]) => `${schluessel}=${encodeURIComponent(wert)}`)
    .join("&");
}

/**
 * Legt die Datei in die Downloads und stellt den Outlook-Aufruf bereit.
 *
 * Outlook oeffnet erst der naechste Tipp auf den Abschlussknopf, nicht dieser
 * Lauf. Zwei Gruende: iOS laesst eine Seite eine andere App nur aus einer
 * frischen Fingergeste heraus oeffnen, und die ist nach dem Erzeugen der
 * Datei verbraucht. Und die Datei muss in den Downloads liegen, BEVOR man in
 * Outlook zur Bueroklammer greift.
 */
function versendeUeberOutlook(blob, dateiname, texte) {
  speichereDatei(blob, dateiname);
  outlookEntwurf = { adresse: outlookAdresse(texte), dateiname };
  melde(
    `${dateiname} liegt in den Downloads. Jetzt „In Outlook öffnen“ tippen — ` +
    "An, CC, Betreff und Text stehen dort schon drin.",
    "erfolg"
  );
}

function oeffneOutlook() {
  if (!outlookEntwurf) return;
  const { adresse, dateiname } = outlookEntwurf;

  // Ob Outlook aufgegangen ist, verraet nur, ob die Seite in den Hintergrund
  // geht. Bleibt sie sichtbar, fehlt vermutlich die App -- oder iOS wartet
  // noch auf "Öffnen", daher die vorsichtige Formulierung.
  let verlassen = false;
  const merke = () => { if (document.hidden) verlassen = true; };
  document.addEventListener("visibilitychange", merke);
  setTimeout(() => {
    document.removeEventListener("visibilitychange", merke);
    if (!verlassen && outlookEntwurf) {
      melde(
        "Outlook ist nicht aufgegangen? Dann fehlt die Outlook-App — " +
        "unter „Mail über“ lässt sich das Teilen-Menü wählen.",
        "warnung"
      );
    }
  }, 4000);

  melde(
    `In Outlook ${dateiname} über die Büroklammer aus „Downloads“ anhängen.`,
    "erfolg"
  );
  // Ab hier gilt der Vorgang als verschickt -- unter Zwischenstaende mit
  // Haken, falls man ihn spaeter noch einmal braucht.
  sichereZwischenstand({ versendet: true });
  window.location.href = adresse;
}

/** Vergisst den vorbereiteten Outlook-Aufruf -- die Datei ist veraltet. */
function verwerfeOutlookEntwurf() {
  if (!outlookEntwurf && !teilenEntwurf) return;
  outlookEntwurf = null;
  teilenEntwurf = null;
  aktualisiereKnopftext();
}

/* Die fertige Datei, falls iOS das Teilen-Menue nach dem Erzeugen verweigert
   hat (NotAllowedError: die Fingergeste war verbraucht). Der naechste Tipp
   auf den Abschlussknopf oeffnet es dann mit frischer Geste. */
let teilenEntwurf = null;

async function teileDatei(blob, dateiname, roh) {
  const texte = mailtexte(roh);
  const datei = new File([blob], dateiname, { type: XLSX_TYP });
  await teile(blob, datei, texte);
}

async function teileErneut() {
  if (!teilenEntwurf) return;
  const { blob, datei, texte } = teilenEntwurf;
  teilenEntwurf = null;
  aktualisiereKnopftext();
  await teile(blob, datei, texte);
}

async function teile(blob, datei, texte) {
  const dateiname = datei.name;

  // Zuerst in die Zwischenablage, dann teilen — und zwar in dieser
  // Reihenfolge. Safari erlaubt den Zugriff auf die Zwischenablage nur,
  // solange die Fingergeste des Nutzers noch "frisch" ist; nach dem await auf
  // navigator.share() ist sie das nicht mehr und der Schreibversuch scheitert
  // still. Der Text wird gebraucht, weil das Teilen-Blatt von iOS bei einem
  // Dateianhang den Begleittext nicht zuverlaessig an jede App weiterreicht:
  // einmal einfuegen ist zumutbar, ihn neu zu tippen nicht.
  const kopiert = await inZwischenablage(`${texte.betreff}\n\n${texte.body}`);
  const nachsatz = kopiert ? " Betreff und Text liegen in der Zwischenablage." : "";

  if (navigator.canShare && navigator.canShare({ files: [datei] })) {
    try {
      await navigator.share({
        files: [datei],
        title: texte.betreff,
        text: texte.body,
      });
      sichereZwischenstand({ versendet: true });
      melde(`Geteilt.${nachsatz}`, "erfolg");
      return;
    } catch (e) {
      // Ein Abbruch durch den Nutzer ist kein Fehler.
      if (e && e.name === "AbortError") {
        melde("Teilen abgebrochen. Die Datei ist erzeugt.", "warnung");
        return;
      }
      // Das Geraet kann teilen, nur nicht mehr aus diesem Tipp heraus: Das
      // Erzeugen hat zu lange gedauert. Kein Grund fuer "kann nicht teilen".
      if (e && e.name === "NotAllowedError") {
        teilenEntwurf = { blob, datei, texte };
        aktualisiereKnopftext();
        melde(`${dateiname} ist fertig. Tippe auf „Teilen-Menü öffnen“.`, "erfolg");
        return;
      }
    }
  }

  // Rueckfall fuer Browser ohne Dateifreigabe.
  speichereDatei(blob, dateiname);
  melde(
    `Dieses Gerät kann Dateien nicht teilen. Die Liste wurde gesichert.${nachsatz}`,
    "warnung"
  );
}

function speichereDatei(blob, dateiname) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = dateiname;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

async function inZwischenablage(text) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (e) { /* kein Zugriff — kein Beinbruch */ }
  return false;
}

/* --------------------------------------------------------------------------
   Selbsttest — beantwortet auf dem echten Geraet die Frage, ob es dort
   laeuft. Alles, was diese Liste behauptet, hat sie gerade ausprobiert.
   -------------------------------------------------------------------------- */

const PRUEFUNGEN = [
  {
    titel: "Auf dem Home-Bildschirm installiert",
    lauf: async () => {
      const eigenstaendig =
        window.matchMedia("(display-mode: standalone)").matches ||
        window.navigator.standalone === true;
      return eigenstaendig
        ? { stand: "gut", text: "Läuft als eigenständige App." }
        : {
            stand: "warnung",
            text: "Läuft im Browser. Ohne Installation räumt iOS die Daten nach sieben ungenutzten Tagen ab.",
          };
    },
  },
  {
    titel: "Offline-Verwalter aktiv",
    lauf: async () => {
      if (!("serviceWorker" in navigator)) {
        return { stand: "schlecht", text: "Dieser Browser kennt keine Service Worker." };
      }
      const reg = await navigator.serviceWorker.getRegistration();
      return reg && reg.active
        ? { stand: "gut", text: "Registriert und aktiv." }
        : { stand: "schlecht", text: "Nicht aktiv — die App braucht beim Start Netz." };
    },
  },
  {
    titel: "Offline-Vorrat gefüllt",
    lauf: async () => {
      if (!("caches" in window)) {
        return { stand: "schlecht", text: "Kein Cache-Speicher verfügbar." };
      }
      const namen = await caches.keys();
      let anzahl = 0;
      for (const name of namen) {
        anzahl += (await (await caches.open(name)).keys()).length;
      }
      return anzahl > 0
        ? { stand: "gut", text: `${anzahl} Dateien liegen im Gerät bereit.` }
        : { stand: "schlecht", text: "Der Vorrat ist leer." };
    },
  },
  {
    titel: "Speicherplatz",
    lauf: async () => {
      if (!navigator.storage || !navigator.storage.estimate) {
        return { stand: "warnung", text: "Das Gerät macht dazu keine Angabe." };
      }
      const { usage = 0, quota = 0 } = await navigator.storage.estimate();
      const mb = (n) => (n / 1048576).toFixed(1);
      return {
        stand: quota > 0 && usage < quota ? "gut" : "warnung",
        text: `${mb(usage)} MB belegt von ${mb(quota)} MB verfügbar.`,
      };
    },
  },
  {
    titel: "Rechenkern geladen",
    lauf: async () => {
      if (kernBereit) {
        const fassung = py.runPython("import sys; sys.version.split()[0]");
        return { stand: "gut", text: `Python ${fassung} läuft im Gerät.` };
      }
      return { stand: "schlecht", text: kernFehler || "Noch nicht bereit." };
    },
  },
  {
    titel: "Prüfung der Eingaben",
    lauf: async () => {
      if (!kernBereit) return { stand: "schlecht", text: "Ohne Rechenkern nicht möglich." };
      const erg = rufe("bruecke_pruefe", JSON.stringify(probeVorgang()));
      return {
        stand: "gut",
        text: `Antwortet: ${erg.issues.length} Meldungen, absendbar = ${erg.absendbar}.`,
      };
    },
  },
  {
    titel: "Excel-Datei erzeugen",
    lauf: async () => {
      if (!kernBereit) return { stand: "schlecht", text: "Ohne Rechenkern nicht möglich." };
      const erg = rufe("bruecke_erzeuge", JSON.stringify(probeVorgang()), "/selbsttest.xlsx");
      const bytes = py.FS.readFile("/selbsttest.xlsx");
      // Eine XLSX ist ein Zip-Paket. Die ersten beiden Bytes muessen "PK"
      // sein — sonst ist etwas entstanden, das Excel nicht oeffnen wird.
      const istZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
      return istZip
        ? { stand: "gut", text: `${erg.dateiname} — ${bytes.length} Bytes, gültiges Zip-Paket.` }
        : { stand: "schlecht", text: "Die erzeugte Datei ist kein gültiges Excel-Paket." };
    },
  },
  {
    titel: "An Outlook übergeben",
    lauf: async () => {
      const probe = new File([new Blob(["x"], { type: XLSX_TYP })], "probe.xlsx", {
        type: XLSX_TYP,
      });
      if (navigator.canShare && navigator.canShare({ files: [probe] })) {
        return { stand: "gut", text: "Das Teilen-Menü nimmt Dateien an." };
      }
      return {
        stand: "warnung",
        text: "Kein Dateiversand über das Teilen-Menü. Die Liste wird stattdessen gesichert.",
      };
    },
  },
  {
    titel: "Daten überleben den Neustart",
    lauf: async () => {
      try {
        localStorage.setItem("glaz.probe", "1");
        const gelesen = localStorage.getItem("glaz.probe");
        localStorage.removeItem("glaz.probe");
        return gelesen === "1"
          ? { stand: "gut", text: "Der Gerätespeicher nimmt Daten an." }
          : { stand: "schlecht", text: "Geschriebenes kam nicht zurück." };
      } catch (e) {
        return { stand: "schlecht", text: "Der Speicher ist gesperrt (privates Fenster?)." };
      }
    },
  },
];

function probeVorgang() {
  return {
    profil: {
      ...LEERES_PROFIL,
      personalnummer: "1234567",
      vorname: "Probe",
      nachname: "Lauf",
      abteilung: "XX YY ZZZ",
      gruppenleiter_name: "Probe Leitung",
      gruppenleiter_email: "probe@example.com",
    },
    einsatzart: "Selbsttest",
    reisetyp: "inland",
    wiederhole_stammdaten: false,
    zeilen: [
      {
        datum: heuteAlsText(),
        passiv_von: "07:00", passiv_bis: "09:30",
        aktiv_von: "09:30", aktiv_bis: "11:00",
        grenz_anreise: null, grenz_rueckreise: null,
        arbeit_von: "11:00", arbeit_bis: "17:00",
        ist_folgezeile: false,
      },
    ],
  };
}

async function starteSelbsttest() {
  const liste = el("pruefliste");
  liste.textContent = "";
  el("knopf-selbsttest-start").disabled = true;
  const ergebnisse = [];

  for (const pruefung of PRUEFUNGEN) {
    // Bewusst ueber DOM-Methoden statt innerHTML zusammengesetzt: hier landen
    // spaeter Fehlermeldungen aus Python und die Geraetekennung des Browsers.
    const punkt = document.createElement("li");
    punkt.dataset.stand = "laeuft";
    const huelle = document.createElement("div");
    const titel = document.createElement("b");
    titel.textContent = pruefung.titel;
    const text = document.createElement("span");
    text.textContent = "läuft …";
    huelle.append(titel, text);
    punkt.append(huelle);
    liste.append(punkt);

    let ergebnis;
    try {
      ergebnis = await pruefung.lauf();
    } catch (e) {
      ergebnis = { stand: "schlecht", text: e && e.message ? e.message : String(e) };
    }
    punkt.dataset.stand = ergebnis.stand;
    punkt.querySelector("span").textContent = ergebnis.text;
    ergebnisse.push(`${ergebnis.stand === "gut" ? "OK" : ergebnis.stand.toUpperCase()}  ${pruefung.titel}: ${ergebnis.text}`);
  }

  el("knopf-selbsttest-start").disabled = false;
  const teilen = el("knopf-selbsttest-teilen");
  teilen.hidden = false;
  teilen.onclick = async () => {
    const bericht =
      `GLAZ-Selbsttest\n${new Date().toLocaleString("de-DE")}\n`
      + `${navigator.userAgent}\n\n${ergebnisse.join("\n")}`;
    if (navigator.share) {
      try { await navigator.share({ title: "GLAZ-Selbsttest", text: bericht }); return; }
      catch (e) { /* weiter zur Zwischenablage */ }
    }
    if (await inZwischenablage(bericht)) {
      teilen.textContent = "In der Zwischenablage";
    }
  };

  el("geraet-angaben").textContent = navigator.userAgent;
}

/* --------------------------------------------------------------------------
   Service Worker und Installationshinweis
   -------------------------------------------------------------------------- */

function registriereServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    el("menue-fuss").textContent = "Offline-Betrieb: von diesem Browser nicht unterstützt.";
    return;
  }
  navigator.serviceWorker.register("sw.js").then(
    (registrierung) => {
      zeigeFassung();
      // Die Meldung "neue-version" geht verloren, wenn die neue Fassung
      // geladen wurde, waehrend die App nicht offen war. Ein wartender Worker
      // verraet es trotzdem.
      if (registrierung.waiting && navigator.serviceWorker.controller) {
        zeigeUpdateHinweis();
      }
    },
    (e) => { el("menue-fuss").textContent = `Offline-Verwalter nicht aktiv: ${e.message}`; }
  );

  navigator.serviceWorker.addEventListener("message", (ereignis) => {
    const nachricht = ereignis.data || {};
    if (nachricht.typ === "neue-version") zeigeUpdateHinweis();
    if (nachricht.typ === "version") {
      el("menue-fuss").textContent = `Fassung ${nachricht.version} · offline verfügbar`;
    }
  });

  // Nur nach "Jetzt aktualisieren" neu laden -- nie ungefragt mitten in einer
  // Eingabe (siehe sw.js, Kommentar vor meldeAllen).
  let neuLaden = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (neuLaden) location.reload();
  });
  el("knopf-update").addEventListener("click", async () => {
    const registrierung = await navigator.serviceWorker.getRegistration();
    if (registrierung && registrierung.waiting) {
      neuLaden = true;
      registrierung.waiting.postMessage({ typ: "uebernehmen" });
    } else {
      location.reload();
    }
  });
}

/**
 * Ein stehender Hinweis statt einer Zeile in der Statusleiste: Die schrieb
 * die naechste Pruefung sofort wieder ueber, und die neue Fassung kam erst
 * an, wenn iOS die App irgendwann ganz beendet hatte. Der Zwischenstand liegt
 * im localStorage, ein Neuladen verliert also nichts.
 */
function zeigeUpdateHinweis() {
  el("update-hinweis").hidden = false;
}

async function zeigeFassung() {
  // Die laufende Fassung weiss nur der aktive Worker. version.json holt er
  // aus dem Netz, sobald es eins gibt -- das waere die NEUE Nummer, waehrend
  // noch der alte Code laeuft. Die Antwort kommt als Nachricht "version".
  const aktiv = navigator.serviceWorker && navigator.serviceWorker.controller;
  if (aktiv) {
    aktiv.postMessage({ typ: "version-abfragen" });
    return;
  }
  try {
    const antwort = await fetch("version.json");
    const daten = await antwort.json();
    el("menue-fuss").textContent = `Fassung ${daten.version} · offline verfügbar`;
  } catch (e) {
    el("menue-fuss").textContent = "Fassung unbekannt";
  }
}

function vielleichtAnstupsen() {
  const istIOS =
    /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  const eigenstaendig =
    window.matchMedia("(display-mode: standalone)").matches ||
    window.navigator.standalone === true;
  let gesehen = false;
  try { gesehen = localStorage.getItem(ANSTUPSER_SCHLUESSEL) === "1"; } catch (e) {}

  if (istIOS && !eigenstaendig && !gesehen) {
    el("ios-anstupser").hidden = false;
  }
}

/* --------------------------------------------------------------------------
   Verdrahtung
   -------------------------------------------------------------------------- */

function verdrahte() {
  // --- Profil
  el("knopf-profil-auf").addEventListener("click", () => {
    const felder = el("profil-felder");
    felder.hidden = !felder.hidden;
    el("knopf-profil-auf").setAttribute("aria-expanded", String(!felder.hidden));
    el("knopf-profil-auf").textContent = felder.hidden ? "Bearbeiten" : "Fertig";
    if (letztePruefung) zeigePruefung(letztePruefung);
  });

  el("profil-auswahl").addEventListener("change", (e) => {
    zustand.aktivesProfil = Number(e.target.value);
    zeichneProfil();
    nachEingabe();
  });

  const profilFeldpaare = [
    ["f-profilname", "name"],
    ["f-vorname", "vorname"],
    ["f-nachname", "nachname"],
    ["f-personalnummer", "personalnummer"],
    ["f-abteilung", "abteilung"],
    ["f-gl-name", "gruppenleiter_name"],
    ["f-gl-email", "gruppenleiter_email"],
  ];
  for (const [id, schluessel] of profilFeldpaare) {
    el(id).addEventListener("input", (e) => {
      aktuellesProfil()[schluessel] = e.target.value;
      el("profil-zusammenfassung").textContent = "";
      zeichneProfil();
      nachEingabe();
    });
  }

  // Eindeutigkeit erst beim Verlassen des Feldes herstellen, nicht bei jedem
  // Tastendruck: Wer "Montage" tippen will, soll nicht nach dem ersten
  // Buchstaben ein "M (2)" vorgesetzt bekommen, weil ein "M (2)" schon
  // existiert.
  el("f-profilname").addEventListener("change", () => {
    const profil = aktuellesProfil();
    profil.name = eindeutigerProfilname(profil.name, zustand.aktivesProfil);
    zeichneProfil();
    speichern();
  });

  /** Oeffnet den Profilbereich und setzt den Schreibcursor in das Namensfeld. */
  function zeigeProfilfelder(mitFokus) {
    const felder = el("profil-felder");
    felder.hidden = false;
    el("knopf-profil-auf").setAttribute("aria-expanded", "true");
    el("knopf-profil-auf").textContent = "Fertig";
    if (mitFokus) {
      const feld = el("f-profilname");
      feld.focus();
      feld.select();
    }
  }

  el("knopf-profil-neu").addEventListener("click", () => {
    // Ein frisches Profil braucht als Erstes einen Namen — deshalb klappt der
    // Bereich auf und der Cursor steht bereits im richtigen Feld. Der
    // Vorschlag ist vorausgewaehlt, sodass Tippen ihn ersetzt.
    zustand.profile.push({
      ...LEERES_PROFIL,
      name: eindeutigerProfilname(`Profil ${zustand.profile.length + 1}`, -1),
    });
    zustand.aktivesProfil = zustand.profile.length - 1;
    zeichneProfil();
    zeigeProfilfelder(true);
    nachEingabe();
  });

  el("knopf-profil-duplizieren").addEventListener("click", () => {
    const quelle = aktuellesProfil();
    zustand.profile.push({
      ...quelle,
      // Von der sichtbaren Beschriftung ausgehen, nicht vom internen Feld:
      // Sonst wird aus der Kopie von "Erika Mustermann" ein "Profil (Kopie)".
      name: eindeutigerProfilname(`${profilLabel(quelle)} (Kopie)`, -1),
    });
    zustand.aktivesProfil = zustand.profile.length - 1;
    zeichneProfil();
    zeigeProfilfelder(true);
    nachEingabe();
  });

  el("knopf-profil-loeschen").addEventListener("click", () => {
    if (zustand.profile.length <= 1) {
      melde("Das letzte Profil lässt sich nicht löschen.", "warnung");
      return;
    }
    if (!confirm("Dieses Profil wirklich löschen?")) return;
    zustand.profile.splice(zustand.aktivesProfil, 1);
    zustand.aktivesProfil = Math.max(0, zustand.aktivesProfil - 1);
    zeichneProfil();
    nachEingabe();
  });

  // --- Vorgang
  el("f-einsatzart").addEventListener("input", (e) => {
    zustand.einsatzart = e.target.value;
    nachEingabe();
  });
  el("f-einsatzart").addEventListener("change", (e) => {
    const wert = e.target.value.trim();
    if (wert && !zustand.einsatzarten.includes(wert)) {
      zustand.einsatzarten.unshift(wert);
      // 20 wie MAX_EINSATZART_HISTORIE in glaz/settings.py.
      zustand.einsatzarten = zustand.einsatzarten.slice(0, 20);
      zeichneEinsatzarten();
      speichern();
    }
  });

  alle("[data-reisetyp]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      zustand.reisetyp = knopf.dataset.reisetyp;
      zeichneUmschalter();
      zeichneZeilen();
      nachEingabe();
    })
  );

  // --- Reisetage
  el("knopf-zeile-neu").addEventListener("click", () => {
    if (zustand.zeilen.length >= MAX_ZEILEN) return;
    zustand.zeilen.push(leereZeile());
    zeichneZeilen();
    nachEingabe();
  });

  el("knopf-folgezeile").addEventListener("click", () => {
    if (zustand.zeilen.length >= MAX_ZEILEN) return;
    zustand.zeilen.push(leereZeile(true));
    zeichneZeilen();
    nachEingabe();
  });

  // --- Abschluss
  alle("[data-modus]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      zustand.modus = knopf.dataset.modus;
      zeichneUmschalter();
      aktualisiereKnopftext();
      nachEingabe();
    })
  );

  el("f-zieldatei").addEventListener("change", async (e) => {
    zieldatei = e.target.files && e.target.files[0] ? e.target.files[0] : null;
    if (!zieldatei) {
      el("zieldatei-info").textContent = "Noch keine Datei gewählt.";
      return;
    }
    if (!/\.xlsx$/i.test(zieldatei.name)) {
      el("zieldatei-info").textContent =
        `${zieldatei.name} ist keine Excel-Liste (.xlsx). Bitte die GLAZ-Liste wählen.`;
      zieldatei = null;
      zieldateiGeladen = false;
      e.target.value = "";
      nachEingabe();
      return;
    }
    el("zieldatei-info").textContent = `${zieldatei.name} wird gelesen …`;
    zieldateiGeladen = false;
    // Eine andere Datei heisst: Der gemerkte Block gilt nicht mehr. Ihn
    // stehen zu lassen hiesse, in einer fremden Datei Zeilen zu ueberschreiben.
    if (!letzterAnhang || letzterAnhang.name !== zieldatei.name) letzterAnhang = null;
    try {
      py.FS.writeFile("/ziel.xlsx", new Uint8Array(await zieldatei.arrayBuffer()));
      zieldateiGeladen = true;
      const bel = rufe("bruecke_belegung", "/ziel.xlsx");
      const belegt = bel.belegte_zeilen ?? bel.belegt ?? bel.anzahl ?? null;
      el("zieldatei-info").textContent =
        belegt === null
          ? `${zieldatei.name} gelesen.`
          : `${zieldatei.name} — ${belegt} von ${MAX_ZEILEN} Zeilen belegt.`;
    } catch (err) {
      el("zieldatei-info").textContent = `${zieldatei.name} ließ sich nicht lesen: ${err.message}`;
    }
    nachEingabe();
  });

  el("f-mail-mitsenden").addEventListener("change", (e) => {
    zustand.mailMitsenden = e.target.checked;
    outlookEntwurf = null;
    teilenEntwurf = null;
    zeichneVersandweg();
    aktualisiereKnopftext();
    speichern();
  });

  alle("[data-versandweg]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      zustand.versandweg = knopf.dataset.versandweg;
      outlookEntwurf = null;
      teilenEntwurf = null;
      zeichneVersandweg();
      aktualisiereKnopftext();
      speichern();
    })
  );

  // Derselbe Knopf in zwei Rollen: erst erzeugen, dann Outlook oeffnen. Ein
  // zweiter farbiger Knopf daneben braeche die Regel "genau eine farbige
  // Schaltflaeche" der Aktionsleiste.
  el("knopf-abschluss").addEventListener("click", () => {
    if (outlookEntwurf) oeffneOutlook();
    else if (teilenEntwurf) teileErneut();
    else abschluss();
  });

  // --- Assistent: blaettern
  el("knopf-zurueck").addEventListener("click", () =>
    zeigeSchritt(zustand.schritt - 1, { fokus: true })
  );
  el("knopf-weiter").addEventListener("click", () =>
    zeigeSchritt(zustand.schritt + 1, { fokus: true })
  );
  // Die Sprungmarke fuer Tastatur und Screenreader fuehrt weiterhin zu den
  // Reisetagen -- die liegen jetzt auf Seite 3 statt weiter unten.
  document.querySelector(".sprungmarke").addEventListener("click", (e) => {
    e.preventDefault();
    zeigeSchritt(3, { fokus: true });
  });

  // --- Menue
  el("knopf-menue").addEventListener("click", () => {
    const menue = el("menue");
    menue.hidden = !menue.hidden;
    el("knopf-menue").setAttribute("aria-expanded", String(!menue.hidden));
  });
  const schliesseMenue = () => {
    if (el("menue").hidden) return;
    el("menue").hidden = true;
    el("knopf-menue").setAttribute("aria-expanded", "false");
  };
  document.addEventListener("click", (e) => {
    if (!el("menue").contains(e.target) && !el("knopf-menue").contains(e.target)) {
      schliesseMenue();
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") schliesseMenue();
  });

  alle(".menue [data-ziel]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      el("menue").hidden = true;
      el("knopf-menue").setAttribute("aria-expanded", "false");
      const ziel = knopf.dataset.ziel;
      if (ziel === "neuer-vorgang") {
        // Nichts geht verloren: Der offene Stand wird als Zwischenstand
        // abgelegt, deshalb auch keine Rueckfrage mehr.
        const abgelegt = sichereZwischenstand();
        // Genau das, was _neuer_vorgang in der Desktop-App zuruecksetzt:
        // Zeilen, Einsatzart und Reisetyp. Das Profil bleibt stehen -- es ist
        // der Teil, der ueber Vorgaenge hinweg gilt.
        setzeVorgang(
          { einsatzart: "", reisetyp: "inland", modus: "neu", zeilen: [leereZeile()] },
          null
        );
        // Auf Seite 2, nicht auf Seite 1: Das Profil bleibt ja stehen, der
        // neue Vorgang beginnt bei der Einsatzart.
        zeigeSchritt(2, { fokus: true });
        nachEingabe();
        if (abgelegt) {
          meldeGehalten("Neuer Vorgang. Der vorige liegt unter Menü → Zwischenstände.", "erfolg");
        }
        return;
      }
      if (ziel === "zwischenstaende") zeichneZwischenstaende();
      zeigeSonderseite(ziel);
    })
  );

  alle("[data-schliessen]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      // "Fertig" fuehrt auf die Seite zurueck, von der man kam.
      zeigeSchritt(zustand.schritt, { fokus: true });
    })
  );

  // --- Einstellungen
  // Jede Aenderung verwirft einen vorbereiteten Outlook-Aufruf: Er traegt
  // Adresse und Text von vorher.
  const einstellungGeaendert = () => {
    verwerfeOutlookEntwurf();
    speichern();
    pruefeEinstellungenGleich();
  };
  el("f-empfaenger").addEventListener("input", (e) => {
    zustand.einstellungen.empfaenger = e.target.value;
    einstellungGeaendert();
  });
  el("f-betreff-vorlage").addEventListener("input", (e) => {
    zustand.einstellungen.betreff_vorlage = e.target.value;
    einstellungGeaendert();
  });
  el("f-body-vorlage").addEventListener("input", (e) => {
    zustand.einstellungen.body_vorlage = e.target.value;
    einstellungGeaendert();
  });
  el("knopf-vorlagen-zuruecksetzen").addEventListener("click", () => {
    if (!confirm("Empfänger, Betreff und Mailtext auf die Vorgaben zurücksetzen?")) return;
    // Nur die drei Mailfelder -- das Erscheinungsbild gehoert nicht zu den
    // Vorlagen und blieb bisher beim Zuruecksetzen auf der Strecke.
    Object.assign(zustand.einstellungen, { empfaenger: "", betreff_vorlage: "", body_vorlage: "" });
    zeichneEinstellungen();
    einstellungGeaendert();
    melde("Die Vorlagen stehen wieder auf den Vorgabewerten.", "erfolg");
  });
  el("knopf-daten-loeschen").addEventListener("click", () => {
    if (!confirm("Profile, Zwischenstand und Einstellungen auf diesem Gerät löschen?")) return;
    try { localStorage.removeItem(SPEICHER_SCHLUESSEL); } catch (e) {}
    location.reload();
  });

  // --- Profil ausdruecklich sichern
  el("knopf-profil-speichern").addEventListener("click", () => {
    const profil = aktuellesProfil();
    profil.name = eindeutigerProfilname(profil.name, zustand.aktivesProfil);
    zeichneProfil();
    if (speichern()) {
      setzeProfilHinweis(
        `Profil „${profilLabel(profil)}“ gesichert – nur auf diesem Gerät.`,
        "erfolg"
      );
    } else {
      setzeProfilHinweis(
        "Das Profil ließ sich nicht sichern. Ist der Speicher des Browsers gesperrt " +
        "(privates Fenster)?",
        "fehler"
      );
    }
  });

  // --- Erscheinungsbild
  alle("[data-thema]", el("menue")).forEach((knopf) =>
    knopf.addEventListener("click", () => {
      zustand.einstellungen.thema = knopf.dataset.thema;
      wendeThemaAn(zustand.einstellungen.thema);
      speichern();
    })
  );

  // --- Selbsttest
  el("knopf-selbsttest-start").addEventListener("click", starteSelbsttest);

  // --- Zwischenstaende
  el("knopf-zwischenstand-sichern").addEventListener("click", () => {
    const eintrag = sichereZwischenstand();
    zeichneZwischenstaende();
    el("zwischenstand-hinweis").textContent = eintrag
      ? `„${zwischenstandTitel(eintrag)}“ gesichert.`
      : "Noch nichts eingegeben — es gibt nichts zu sichern.";
  });

  // --- Anstupser
  el("knopf-anstupser-weg").addEventListener("click", () => {
    el("ios-anstupser").hidden = true;
    try { localStorage.setItem(ANSTUPSER_SCHLUESSEL, "1"); } catch (e) {}
  });
}

function aktualisiereKnopftext() {
  const knopf = el("knopf-abschluss");
  if (outlookEntwurf) {
    knopf.textContent = "In Outlook öffnen";
    return;
  }
  if (teilenEntwurf) {
    knopf.textContent = "Teilen-Menü öffnen";
    return;
  }
  const mail = zustand.mailMitsenden;
  const wort = zustand.versandweg === "outlook" ? "senden" : "teilen";
  if (zustand.modus === "ergaenzen") {
    knopf.textContent = mail ? `Datei ergänzen & ${wort}` : "Datei ergänzen";
  } else if (zustand.modus === "nur_versenden") {
    knopf.textContent = mail ? `Datei ${wort}` : "Datei sichern";
  } else {
    knopf.textContent = mail ? `Excel erzeugen & ${wort}` : "Excel erzeugen";
  }
}

function zeichneVersandweg() {
  el("versandweg").hidden = !zustand.mailMitsenden;
  alle("[data-versandweg]").forEach((k) =>
    k.setAttribute("aria-checked", String(k.dataset.versandweg === zustand.versandweg))
  );
  el("versandweg-hinweis").textContent =
    zustand.versandweg === "outlook"
      ? "Öffnet Outlook mit Empfänger, CC (Gruppenleitung), Betreff und Text. " +
        "Die Datei liegt dann in „Downloads“ und kommt über die Büroklammer dazu."
      : "Öffnet das Teilen-Menü mit der Datei. Dort Mail-App, Empfänger und CC selbst wählen.";
}

/* --------------------------------------------------------------------------
   Start
   -------------------------------------------------------------------------- */

function start() {
  laden();
  if (!zustand.zeilen.length) zustand.zeilen = [leereZeile()];

  wendeThemaAn(zustand.einstellungen.thema);

  // Die Gruppenleiter-Verwaltung baut sich selbst in ihren Behaelter. Sie
  // bekommt ausschliesslich diese schmale Umgebung statt Zugriff auf den
  // ganzen Zustand -- so bleibt nachvollziehbar, was sie anfassen kann.
  if (typeof GLAZ_GRUPPENLEITER !== "undefined") {
    GLAZ_GRUPPENLEITER.initialisieren({
      liste: () => zustand.gruppenleiter,
      setzeListe: (neu) => { zustand.gruppenleiter = neu; speichern(); },
      aktuellesProfil,
      profilGeaendert: () => { zeichneProfil(); nachEingabe(); },
      melde: setzeProfilHinweis,
    });
  }

  baueSegmente();
  verdrahte();
  zeichneProfil();
  zeichneEinsatzarten();
  zeichneUmschalter();
  zeichneZeilen();
  zeichneEinstellungen();
  aktualisiereKnopftext();
  zeigeSchritt(zustand.schritt);

  el("f-einsatzart").value = zustand.einsatzart;
  el("f-mail-mitsenden").checked = zustand.mailMitsenden;
  zeichneVersandweg();

  melde("Der Rechenkern startet …");
  registriereServiceWorker();
  vielleichtAnstupsen();
  starteKern();
}

document.addEventListener("DOMContentLoaded", start);
