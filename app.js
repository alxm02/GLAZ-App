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

def bruecke_ergaenze(roh, pfad):
    return json.dumps(portabel.ergaenze_xlsx(json.loads(roh), pfad), ensure_ascii=False)

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
  einsatzart: "",
  einsatzarten: [],
  reisetyp: "inland",
  wiederhole_stammdaten: false,
  zeilen: [],
  modus: "neu",
  mailMitsenden: true,
  einstellungen: { empfaenger: "", betreff_vorlage: "", body_vorlage: "" },
};

/* Die gewaehlte Zieldatei lebt absichtlich NICHT im gespeicherten Zustand:
   ein File-Objekt ueberlebt keinen Neustart, und ein Pfad, auf den wir beim
   naechsten Start nicht mehr zugreifen koennen, waere ein leeres Versprechen. */
let zieldatei = null;

function leereZeile(folgezeile = false) {
  const z = { datum: folgezeile ? null : heuteAlsText(), ist_folgezeile: folgezeile };
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
    zeilen: zustand.zeilen.map((z) => ({ ...z })),
  };
}

function speichern() {
  try {
    localStorage.setItem(SPEICHER_SCHLUESSEL, JSON.stringify(zustand));
  } catch (e) {
    // Voller oder gesperrter Speicher darf die Eingabe nicht abwuergen.
    // Der Nutzer erfaehrt es an der Statuszeile, nicht per Absturz.
    melde("Der Zwischenstand konnte nicht gesichert werden.", "warnung");
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
      zustand.profile = gelesen.profile.map((p) => ({ ...LEERES_PROFIL, ...p }));
    }
    if (Number.isInteger(gelesen.aktivesProfil)) {
      zustand.aktivesProfil = Math.min(
        Math.max(0, gelesen.aktivesProfil),
        zustand.profile.length - 1
      );
    }
    if (typeof gelesen.einsatzart === "string") zustand.einsatzart = gelesen.einsatzart;
    if (Array.isArray(gelesen.einsatzarten)) zustand.einsatzarten = gelesen.einsatzarten;
    if (gelesen.reisetyp === "ausland" || gelesen.reisetyp === "inland") {
      zustand.reisetyp = gelesen.reisetyp;
    }
    if (Array.isArray(gelesen.zeilen)) {
      zustand.zeilen = gelesen.zeilen.slice(0, MAX_ZEILEN).map((z) => ({
        ...leereZeile(),
        ...z,
      }));
    }
    if (typeof gelesen.mailMitsenden === "boolean") {
      zustand.mailMitsenden = gelesen.mailMitsenden;
    }
    if (gelesen.einstellungen && typeof gelesen.einstellungen === "object") {
      Object.assign(zustand.einstellungen, gelesen.einstellungen);
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
  } catch (e) {
    kernFehler = e && e.message ? e.message : String(e);
    setzeBereitschaft("fehler", "Kern fehlt");
    melde(`Die Prüfung ist nicht verfügbar: ${kernFehler}`, "fehler");
  }
}

function rufe(name, ...args) {
  const fn = py.globals.get(name);
  try {
    return JSON.parse(fn(...args));
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
  if (tat === "entfernen") {
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
}

/* --------------------------------------------------------------------------
   Pruefung
   -------------------------------------------------------------------------- */

let letztePruefung = null;

function melde(text, schwere = "") {
  const feld = el("leiste-meldung");
  feld.textContent = text;
  if (schwere) feld.dataset.schwere = schwere;
  else delete feld.dataset.schwere;
}

function setzeBereitschaft(stand, text) {
  el("bereitschaftspunkt").dataset.stand = stand;
  el("bereitschaftstext").textContent = text;
}

function pruefeJetzt() {
  if (!kernBereit) return;
  let ergebnis;
  try {
    ergebnis = rufe("bruecke_pruefe", JSON.stringify(vorgangDict()));
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

  el("kopf-zeitraum").textContent = erg.zeitraum_text || "Noch keine Zeiten erfasst";
  el("dateiname-vorschau").textContent = erg.dateiname || "—";

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
  knopf.disabled = !erg.absendbar;

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

function nachEingabe() {
  speichern();
  pruefeGleich();
}

/* --------------------------------------------------------------------------
   Abschluss: Datei erzeugen und ins Teilen-Menue geben
   -------------------------------------------------------------------------- */

async function abschluss() {
  if (!kernBereit) {
    melde("Der Rechenkern ist noch nicht bereit.", "warnung");
    return;
  }
  const knopf = el("knopf-abschluss");
  knopf.disabled = true;
  knopf.dataset.laeuft = "true";
  const beschriftungVorher = knopf.textContent;
  knopf.textContent = "Wird erzeugt …";

  try {
    const roh = JSON.stringify(vorgangDict());
    let ergebnis;
    let pfad;

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
      pfad = "/ziel.xlsx";
      py.FS.writeFile(pfad, new Uint8Array(await zieldatei.arrayBuffer()));

      const zielpruefung = rufe("bruecke_pruefe_ziel", roh, pfad, zustand.modus);
      const zielfehler = zielpruefung.issues.filter((i) => i.schwere === "fehler");
      if (zielfehler.length) {
        melde(zielfehler[0].meldung, "fehler");
        return;
      }
      ergebnis =
        zustand.modus === "ergaenzen"
          ? rufe("bruecke_ergaenze", roh, pfad)
          : { pfad, dateiname: zieldatei.name, bytes: zieldatei.size };
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

    if (zustand.mailMitsenden) {
      await teileDatei(blob, dateiname, roh);
    } else {
      speichereDatei(blob, dateiname);
      melde(`${dateiname} wurde gesichert.`, "erfolg");
    }
  } catch (e) {
    melde(`Es hat nicht geklappt: ${e && e.message ? e.message : e}`, "fehler");
  } finally {
    knopf.textContent = beschriftungVorher;
    delete knopf.dataset.laeuft;
    pruefeJetzt();
  }
}

async function teileDatei(blob, dateiname, roh) {
  const texte = rufe(
    "bruecke_mailtexte",
    roh,
    zustand.einstellungen.empfaenger || "",
    zustand.einstellungen.betreff_vorlage || "",
    zustand.einstellungen.body_vorlage || ""
  );

  const datei = new File([blob], dateiname, { type: XLSX_TYP });

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
      melde(`Geteilt.${nachsatz}`, "erfolg");
      return;
    } catch (e) {
      // Ein Abbruch durch den Nutzer ist kein Fehler.
      if (e && e.name === "AbortError") {
        melde("Teilen abgebrochen. Die Datei ist erzeugt.", "warnung");
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
    () => zeigeFassung(),
    (e) => { el("menue-fuss").textContent = `Offline-Verwalter nicht aktiv: ${e.message}`; }
  );

  navigator.serviceWorker.addEventListener("message", (ereignis) => {
    const nachricht = ereignis.data || {};
    if (nachricht.typ === "neue-version") {
      melde("Eine neuere Fassung liegt bereit. Schließe die App und öffne sie erneut.", "warnung");
    }
  });
}

async function zeigeFassung() {
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
      zustand.einsatzarten = zustand.einsatzarten.slice(0, 12);
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
    // Datum der letzten Zeile plus einen Tag: eine Dienstreise geht
    // typischerweise ueber aufeinanderfolgende Tage.
    const letzte = zustand.zeilen.filter((z) => z.datum).slice(-1)[0];
    const neue = leereZeile();
    if (letzte && letzte.datum) {
      const d = new Date(`${letzte.datum}T12:00:00`);
      d.setDate(d.getDate() + 1);
      neue.datum = d.toISOString().slice(0, 10);
    }
    zustand.zeilen.push(neue);
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
    el("zieldatei-info").textContent = `${zieldatei.name} wird gelesen …`;
    try {
      py.FS.writeFile("/ziel.xlsx", new Uint8Array(await zieldatei.arrayBuffer()));
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
    aktualisiereKnopftext();
    speichern();
  });

  el("knopf-abschluss").addEventListener("click", abschluss);

  // --- Menue
  el("knopf-menue").addEventListener("click", () => {
    const menue = el("menue");
    menue.hidden = !menue.hidden;
    el("knopf-menue").setAttribute("aria-expanded", String(!menue.hidden));
  });

  alle(".menue [data-ziel]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      el("menue").hidden = true;
      el("knopf-menue").setAttribute("aria-expanded", "false");
      const ziel = knopf.dataset.ziel;
      if (ziel === "neuer-vorgang") {
        if (!confirm("Alle erfassten Reisetage verwerfen?")) return;
        zustand.zeilen = [leereZeile()];
        zieldatei = null;
        zustand.modus = "neu";
        zeichneUmschalter();
        zeichneZeilen();
        nachEingabe();
        return;
      }
      const abschnitt = el(ziel);
      abschnitt.hidden = false;
      abschnitt.scrollIntoView({ behavior: "smooth", block: "start" });
    })
  );

  alle("[data-schliessen]").forEach((knopf) =>
    knopf.addEventListener("click", () => {
      el(knopf.dataset.schliessen).hidden = true;
      window.scrollTo({ top: 0, behavior: "smooth" });
    })
  );

  // --- Einstellungen
  el("f-empfaenger").addEventListener("input", (e) => {
    zustand.einstellungen.empfaenger = e.target.value;
    speichern();
  });
  el("f-betreff-vorlage").addEventListener("input", (e) => {
    zustand.einstellungen.betreff_vorlage = e.target.value;
    speichern();
  });
  el("f-body-vorlage").addEventListener("input", (e) => {
    zustand.einstellungen.body_vorlage = e.target.value;
    speichern();
  });
  el("knopf-vorlagen-zuruecksetzen").addEventListener("click", () => {
    zustand.einstellungen = { empfaenger: "", betreff_vorlage: "", body_vorlage: "" };
    zeichneEinstellungen();
    speichern();
    melde("Die Vorlagen stehen wieder auf den Vorgabewerten.", "erfolg");
  });
  el("knopf-daten-loeschen").addEventListener("click", () => {
    if (!confirm("Profile, Zwischenstand und Einstellungen auf diesem Gerät löschen?")) return;
    try { localStorage.removeItem(SPEICHER_SCHLUESSEL); } catch (e) {}
    location.reload();
  });

  // --- Selbsttest
  el("knopf-selbsttest-start").addEventListener("click", starteSelbsttest);

  // --- Anstupser
  el("knopf-anstupser-weg").addEventListener("click", () => {
    el("ios-anstupser").hidden = true;
    try { localStorage.setItem(ANSTUPSER_SCHLUESSEL, "1"); } catch (e) {}
  });
}

function aktualisiereKnopftext() {
  const knopf = el("knopf-abschluss");
  const teilen = zustand.mailMitsenden;
  if (zustand.modus === "ergaenzen") {
    knopf.textContent = teilen ? "Datei ergänzen & teilen" : "Datei ergänzen";
  } else if (zustand.modus === "nur_versenden") {
    knopf.textContent = teilen ? "Datei teilen" : "Datei sichern";
  } else {
    knopf.textContent = teilen ? "Excel erzeugen & teilen" : "Excel erzeugen";
  }
}

/* --------------------------------------------------------------------------
   Start
   -------------------------------------------------------------------------- */

function start() {
  laden();
  if (!zustand.zeilen.length) zustand.zeilen = [leereZeile()];

  verdrahte();
  zeichneProfil();
  zeichneEinsatzarten();
  zeichneUmschalter();
  zeichneZeilen();
  zeichneEinstellungen();
  aktualisiereKnopftext();

  el("f-einsatzart").value = zustand.einsatzart;
  el("f-mail-mitsenden").checked = zustand.mailMitsenden;

  melde("Der Rechenkern startet …");
  registriereServiceWorker();
  vielleichtAnstupsen();
  starteKern();
}

document.addEventListener("DOMContentLoaded", start);
