/* ==========================================================================
   GLAZ-Korrekturbuchungsliste — Gruppenleiter-Stammdaten der Handy-Fassung

   Nachbau des Desktop-Bereichs "Gruppenleiter verwalten" (glaz/ui/main_window.py,
   _baue_gruppenleiter_panel und die _gruppenleiter_*-Methoden). Die Liste ist
   ein reiner Komfort-Vorrat: Sie fuellt die drei Profilfelder Abteilung,
   Gruppenleitung und deren E-Mail, damit man sie nicht bei jedem Vorgang neu
   eintippt. Die Wahrheit fuer Excel und Mail bleibt das Profil selbst — genau
   wie es der Klassenkommentar von glaz.model.Gruppenleiter festhaelt.

   Warum eine eigene Datei: web/app.js ist die Anwendungslogik, und ein
   Stammdaten-Verwalter, den man an drei Tagen im Jahr oeffnet, gehoert nicht
   mittenhinein. Die Datei wird als klassisches <script> VOR app.js geladen und
   stellt genau einen globalen Namen bereit — alles andere lebt in der
   umschliessenden Funktion und kann von app.js weder gelesen noch versehentlich
   ueberschrieben werden.

   Die Fachregeln sind hier bewusst nachgebaut statt aus Python geholt: Der
   Vorrat lebt im localStorage des Geraets und begegnet den Desktop-Daten nie,
   es gibt also nichts, was auseinanderlaufen koennte. Pyodide erst zu starten,
   um eine E-Mail-Adresse zu pruefen, waere zudem eine Sekunde Wartezeit fuer
   einen regulaeren Ausdruck. Dieselbe Abwaegung trifft app.js fuer die
   Profilnamen (siehe eindeutigerProfilname dort).
   ========================================================================== */

"use strict";

/**
 * Der einzige globale Name dieser Datei.
 *
 * app.js ruft `initialisieren(umgebung)` einmal beim Start und `zeichne()`,
 * wenn sich das Profil geaendert hat. Alles Uebrige — Zustand, Helfer, die
 * gebaute Oberflaeche — bleibt in dieser Funktion eingeschlossen.
 */
const GLAZ_GRUPPENLEITER = (() => {
  /* Zeichen fuer Zeichen dasselbe Muster wie EMAIL_RE in glaz/model.py. Wer
     es dort aendert, aendert es auch hier — ein Kommentar ist die einzige
     Klammer, die zwei Sprachen zusammenhalten kann. */
  const EMAIL_MUSTER = /^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$/;

  /* Der leere Behaelter, den web/index.html im Profilbereich bereitstellt. */
  const HALTER_ID = "gruppenleiter-halter";

  /* Eintrag 0 der Auswahlliste. Wie die Desktop-Combobox ist die Liste ein
     Uebernahme-Werkzeug und keine dauerhafte Zustandsanzeige: nach der
     Uebernahme springt sie sofort auf diesen Platzhalter zurueck. */
  const PLATZHALTER = "Gruppenleitung auswählen …";

  /** Von app.js gereichte Umgebung (liste, setzeListe, aktuellesProfil, …). */
  let umg = null;

  /** Die gebauten Knoten. null, solange nicht initialisiert wurde. */
  let teile = null;

  /** Ist der Verwaltungsbereich aufgeklappt? */
  let offen = false;

  /* Fingerabdruck des zuletzt gezeichneten Vorrats. zeichne() laeuft nach
     JEDEM Tastendruck im Profil (app.js ruft es aus seiner Zeichenroutine);
     die Auswahlliste jedes Mal neu aufzubauen wuerde eine gerade geoeffnete
     Auswahl des Systems zuklappen und kostet ohne Not Rechenzeit. */
  let letzterAbdruck = null;

  /* ------------------------------------------------------------------------
     Reine Helfer — Spiegel von glaz/model.py und glaz/settings.py
     ------------------------------------------------------------------------ */

  /** Zwingt einen aus dem Geraetespeicher gelesenen Wert auf Text.
   *  Gegenstueck zu ``_als_text`` in glaz/settings.py: eine von Hand
   *  bearbeitete Ablage darf die App nicht am ersten ``.trim()`` zerschellen
   *  lassen. */
  const alsText = (wert) =>
    typeof wert === "string" ? wert : wert == null ? "" : String(wert);

  const geputzt = (wert) => alsText(wert).trim();

  /** Vergleichsform der E-Mail-Adresse. Sie ist der Schluessel des Vorrats,
   *  und zwar unabhaengig von Gross- und Kleinschreibung — so wie es
   *  merke_gruppenleiter/finde_gruppenleiter/entferne_gruppenleiter in
   *  glaz/settings.py handhaben. */
  const schluessel = (email) => geputzt(email).toLowerCase();

  /** Gegenstueck zu ``Gruppenleiter.email_gueltig``. */
  const emailGueltig = (email) => EMAIL_MUSTER.test(geputzt(email));

  /** Gegenstueck zu ``Gruppenleiter.anzeige``: "Name – Abteilung", ohne
   *  Abteilung nur der Name. */
  function anzeige(gl) {
    const abteilung = geputzt(gl && gl.abteilung);
    const name = geputzt(gl && gl.name);
    return abteilung ? `${name} – ${abteilung}` : name;
  }

  /** Listenzeile wie _aktualisiere_gruppenleiter_liste sie baut: Anzeige plus
   *  E-Mail, damit zwei Namensgleiche unterscheidbar bleiben. */
  function listenzeile(gl) {
    const email = geputzt(gl && gl.email);
    return email ? `${anzeige(gl)} – ${email}` : anzeige(gl);
  }

  /** Gegenstueck zu ``finde_gruppenleiter``. */
  function finde(liste, email) {
    const gesucht = schluessel(email);
    return liste.find((g) => schluessel(g.email) === gesucht) || null;
  }

  /**
   * Gegenstueck zu ``merke_gruppenleiter``: aktualisieren statt verdoppeln.
   *
   * Liefert bewusst eine NEUE Liste, statt die alte zu veraendern — app.js
   * bekommt sie ueber setzeListe gereicht und speichert sie dabei; ein an
   * zwei Stellen geteiltes Feld waere eine Einladung, genau diesen Schritt
   * einmal zu vergessen.
   */
  function merke(liste, eintrag) {
    const gesucht = schluessel(eintrag.email);
    const platz = liste.findIndex((g) => schluessel(g.email) === gesucht);
    if (platz < 0) return [...liste, eintrag];
    const neu = [...liste];
    neu[platz] = eintrag;
    return neu;
  }

  /** Gegenstueck zu ``entferne_gruppenleiter``. Liefert die neue Liste oder
   *  null, wenn nichts zu entfernen war. */
  function entferne(liste, email) {
    const gesucht = schluessel(email);
    const platz = liste.findIndex((g) => schluessel(g.email) === gesucht);
    if (platz < 0) return null;
    const neu = [...liste];
    neu.splice(platz, 1);
    return neu;
  }

  /** Der Vorrat, immer als Feld — auch wenn app.js (noch) nichts liefert. */
  function vorrat() {
    const roh = umg && typeof umg.liste === "function" ? umg.liste() : null;
    return Array.isArray(roh) ? roh : [];
  }

  function melde(text, schwere = "") {
    if (umg && typeof umg.melde === "function") umg.melde(text, schwere);
  }

  /* ------------------------------------------------------------------------
     Oberflaeche bauen

     Alles per DOM-Methoden und textContent. Namen und Adressen kommen aus dem
     Geraetespeicher; als innerHTML eingesetzt wuerde ein Name wie
     "<b>Meier</b>" zu Markup statt zu Text — und was mit spitzen Klammern
     anfaengt, hoert dort erfahrungsgemaess nicht auf.

     Es entstehen ausschliesslich Klassen aus web/app.css. Die Tippflaechen
     ergeben sich daraus von selbst: .knopf-zweit und die Eingabefelder sind
     44 px hoch (--tippflaeche), .textknopf ist die flache Nebenaktion, die
     auch der uebrige Profilbereich verwendet.
     ------------------------------------------------------------------------ */

  function bauLabel(beschriftung, feldKnoten, extraKlasse) {
    const label = document.createElement("label");
    label.className = extraKlasse ? `feld ${extraKlasse}` : "feld";
    const span = document.createElement("span");
    span.textContent = beschriftung;
    label.append(span, feldKnoten);
    return label;
  }

  function bauKnopf(beschriftung, klasse, beiKlick) {
    const knopf = document.createElement("button");
    // Ohne type="button" ist ein Knopf innerhalb eines Formulars ein
    // Absende-Knopf und laedt die Seite neu. Der Profilbereich ist heute kein
    // <form> — aber darauf sollte sich eine Schaltflaeche nicht verlassen.
    knopf.type = "button";
    knopf.className = klasse;
    knopf.textContent = beschriftung;
    knopf.addEventListener("click", beiKlick);
    return knopf;
  }

  function bauEingabe(art, zusatz = {}) {
    const feld = document.createElement("input");
    feld.type = art;
    for (const [name, wert] of Object.entries(zusatz)) feld.setAttribute(name, wert);
    return feld;
  }

  function bauOberflaeche(halter) {
    halter.textContent = "";
    // Der Behaelter steht im Profilraster (.felder, zwei Spalten). Ohne eigene
    // Spaltenangabe quetschte er sich in eine halbe Spalte; .feld ist genau
    // die vorhandene Klasse, die "volle Breite" sagt — und ausserhalb eines
    // Rasters ein harmloser Block.
    halter.classList.add("feld");

    /* --- Kopfzeile mit Aufklapp-Knopf --- */
    const kopf = document.createElement("div");
    kopf.className = "block-kopf";
    const titel = document.createElement("span");
    titel.className = "feld-name";
    titel.textContent = "Gruppenleitungen";
    const umschalten = bauKnopf("Verwalten", "textknopf", () => setzeOffen(!offen));
    umschalten.setAttribute("aria-expanded", "false");
    umschalten.setAttribute("aria-controls", "gruppenleiter-bereich");
    kopf.append(titel, umschalten);

    /* --- Auswahlliste: das Uebernahme-Werkzeug ---
       Sie bleibt auch bei zugeklapptem Bereich sichtbar. Das ist der eine
       Handgriff, den man wirklich regelmaessig braucht — die Verwaltung
       darunter dagegen nur, wenn sich etwas aendert. */
    const auswahl = document.createElement("select");
    auswahl.setAttribute("aria-label", "Gespeicherte Gruppenleitung übernehmen");
    auswahl.addEventListener("change", ausgewaehlt);
    const auswahlLabel = bauLabel("Aus dem Vorrat übernehmen", auswahl);

    const zusammenfassung = document.createElement("p");
    zusammenfassung.className = "zusammenfassung";

    /* --- Verwaltungsbereich, zugeklappt --- */
    const bereich = document.createElement("div");
    bereich.className = "felder";
    bereich.id = "gruppenleiter-bereich";
    bereich.hidden = true;

    // Gespeicherte Eintraege als Knopfreihe: ein Tipp laedt den Eintrag in die
    // Felder darunter — das Gegenstueck zu "Bearbeiten" der Desktop-App, nur
    // ohne den Umweg ueber eine markierte Listenzeile. Auf 390 px ist eine
    // Liste, die man erst markieren und dann mit einem zweiten Knopf oeffnen
    // muss, zwei Handgriffe zu viel.
    const listenHalter = document.createElement("div");
    listenHalter.className = "feld";
    const listenTitel = document.createElement("span");
    listenTitel.className = "feld-name";
    listenTitel.textContent = "Gespeichert (zum Bearbeiten antippen)";
    const listenReihe = document.createElement("div");
    listenReihe.className = "feld-aktionen";
    listenHalter.append(listenTitel, listenReihe);

    const nameFeld = bauEingabe("text", { autocomplete: "name", enterkeyhint: "next" });
    const abteilungFeld = bauEingabe("text", { enterkeyhint: "next" });
    const emailFeld = bauEingabe("email", {
      inputmode: "email",
      autocomplete: "email",
      autocapitalize: "off",
      spellcheck: "false",
      enterkeyhint: "done",
    });

    const aktionen = document.createElement("div");
    aktionen.className = "feld-aktionen";
    aktionen.append(
      bauKnopf("Speichern", "knopf-zweit", speichern),
      bauKnopf("Aus Profil übernehmen", "textknopf", ausProfil),
      bauKnopf("Neu", "textknopf", () => {
        felderLeeren();
        melde("Felder geleert — bereit für einen neuen Eintrag.");
      }),
      bauKnopf("Löschen", "textknopf textknopf-warnend", loeschen)
    );

    bereich.append(
      listenHalter,
      bauLabel("Name", nameFeld, "feld-halb"),
      bauLabel("Abteilung", abteilungFeld, "feld-halb"),
      bauLabel("E-Mail", emailFeld),
      aktionen
    );

    halter.append(kopf, auswahlLabel, zusammenfassung, bereich);

    return {
      halter, umschalten, auswahl, auswahlLabel, zusammenfassung,
      bereich, listenHalter, listenReihe, nameFeld, abteilungFeld, emailFeld,
    };
  }

  /* ------------------------------------------------------------------------
     Zeichnen
     ------------------------------------------------------------------------ */

  function setzeOffen(neu) {
    offen = neu;
    teile.bereich.hidden = !offen;
    teile.umschalten.setAttribute("aria-expanded", String(offen));
    teile.umschalten.textContent = offen ? "Fertig" : "Verwalten";
    // Zugeklappt traegt die Zusammenfassung die Auskunft ("3 gespeichert"),
    // aufgeklappt steht die Liste selbst da — dann waere sie nur eine Zeile,
    // die dasselbe noch einmal sagt.
    teile.zusammenfassung.hidden = offen;
    if (offen) {
      // Wie _gruppenleiter_panel_umschalten: beim Oeffnen mit leeren Feldern
      // beginnen, damit der Bereich nicht mit einem halb getippten Eintrag von
      // vorgestern aufgeht.
      felderLeeren();
      zeichne();
    }
  }

  function felderLeeren() {
    if (!teile) return;
    teile.nameFeld.value = "";
    teile.abteilungFeld.value = "";
    teile.emailFeld.value = "";
    markiere(teile.nameFeld, true);
    markiere(teile.emailFeld, true);
  }

  /** Setzt die Fehlermarkierung eines Feldes. aria-invalid faerbt ueber
   *  app.css den Rand und sagt gleichzeitig der Vorlesesoftware Bescheid —
   *  eine rote Umrandung allein tut nur das Erste. */
  function markiere(feld, gut) {
    if (gut) feld.removeAttribute("aria-invalid");
    else feld.setAttribute("aria-invalid", "true");
  }

  function zeichne() {
    if (!teile) return;
    const liste = vorrat();

    const anzahl = liste.length;
    teile.zusammenfassung.textContent = anzahl
      ? anzahl === 1
        ? "1 Gruppenleitung gespeichert."
        : `${anzahl} Gruppenleitungen gespeichert.`
      : "Noch keine Gruppenleitung hinterlegt.";
    // Eine Auswahlliste, in der nur der Platzhalter steht, ist ein Feld, das
    // nichts kann. Sie verschwindet, bis es etwas zu waehlen gibt — genau wie
    // das Desktop-Panel seine leere Liste gegen einen Hinweis tauscht.
    teile.auswahlLabel.hidden = anzahl === 0;
    teile.listenHalter.hidden = anzahl === 0;

    const abdruck = JSON.stringify(
      liste.map((g) => [geputzt(g.name), geputzt(g.abteilung), geputzt(g.email)])
    );
    if (abdruck === letzterAbdruck) return;
    letzterAbdruck = abdruck;

    teile.auswahl.textContent = "";
    const platzhalter = document.createElement("option");
    platzhalter.value = "";
    platzhalter.textContent = PLATZHALTER;
    teile.auswahl.append(platzhalter);
    for (const gl of liste) {
      const eintrag = document.createElement("option");
      // Wert ist die E-Mail, nicht der Listenplatz: Plaetze verschieben sich,
      // sobald jemand einen Eintrag loescht, die Adresse bleibt der
      // Schluessel — genau wie in glaz/settings.py.
      eintrag.value = geputzt(gl.email);
      eintrag.textContent = anzeige(gl);
      teile.auswahl.append(eintrag);
    }
    teile.auswahl.value = "";

    teile.listenReihe.textContent = "";
    for (const gl of liste) {
      const email = geputzt(gl.email);
      teile.listenReihe.append(
        bauKnopf(listenzeile(gl), "knopf-zweit", () => bearbeiten(email))
      );
    }
  }

  /* ------------------------------------------------------------------------
     Aktionen
     ------------------------------------------------------------------------ */

  /**
   * Uebernimmt den gewaehlten Eintrag ins aktive Profil.
   *
   * Dieselben drei Felder wie _gruppenleiter_ausgewaehlt in der Desktop-App:
   * Abteilung, Gruppenleitung (Name), deren E-Mail. Die Abteilung wandert mit
   * — sie gehoert in der Vorlage zum Gruppenleiter-Block und nicht zur eigenen
   * Person. Danach springt die Auswahlliste auf den Platzhalter zurueck: sie
   * hat den Eintrag uebergeben, nicht gebunden.
   */
  function ausgewaehlt() {
    const email = teile.auswahl.value;
    teile.auswahl.value = "";
    if (!email) return;

    const gl = finde(vorrat(), email);
    if (!gl) return;

    const profil = umg.aktuellesProfil();
    if (!profil) return;
    profil.abteilung = geputzt(gl.abteilung);
    profil.gruppenleiter_name = geputzt(gl.name);
    profil.gruppenleiter_email = geputzt(gl.email);
    // app.js zeichnet daraufhin das Profil neu, prueft und speichert — das
    // Wissen darueber, was nach einer Profilaenderung zu tun ist, bleibt an
    // einer Stelle.
    umg.profilGeaendert();
    melde(`${geputzt(gl.name)} ins Profil übernommen.`, "erfolg");
  }

  /** Laedt einen gespeicherten Eintrag in die Eingabefelder
   *  (_gruppenleiter_bearbeiten). */
  function bearbeiten(email) {
    const gl = finde(vorrat(), email);
    if (!gl) return;
    teile.nameFeld.value = geputzt(gl.name);
    teile.abteilungFeld.value = geputzt(gl.abteilung);
    teile.emailFeld.value = geputzt(gl.email);
    markiere(teile.nameFeld, true);
    markiere(teile.emailFeld, true);
    melde(`${geputzt(gl.name)} zum Bearbeiten geladen.`);
  }

  /**
   * Legt einen Eintrag an oder aktualisiert einen vorhandenen
   * (_gruppenleiter_uebernehmen).
   *
   * Beides derselbe Knopf, weil es fuer den Vorrat derselbe Vorgang ist: Der
   * Schluessel ist die E-Mail-Adresse, und wer eine bereits bekannte Adresse
   * noch einmal speichert, meint deren Eintrag — nicht einen zweiten daneben.
   *
   * Liefert true, wenn gespeichert wurde. Den Rueckgabewert braucht
   * ``ausProfil``, das ueber denselben Weg geht.
   */
  function speichern() {
    const name = geputzt(teile.nameFeld.value);
    const abteilung = geputzt(teile.abteilungFeld.value);
    const email = geputzt(teile.emailFeld.value);

    // Beide Felder werden markiert, bevor abgebrochen wird: Wer Name UND
    // Adresse vergessen hat, soll das an beiden Raendern sehen und nicht
    // zweimal hintereinander abgewiesen werden.
    const nameGut = name !== "";
    const emailGut = emailGueltig(email);
    markiere(teile.nameFeld, nameGut);
    markiere(teile.emailFeld, emailGut);

    if (!nameGut) {
      melde("Name darf nicht leer sein.", "fehler");
      return false;
    }
    if (!emailGut) {
      melde("Bitte eine gültige E-Mail-Adresse angeben.", "fehler");
      return false;
    }

    const liste = vorrat();
    const kannte = finde(liste, email) !== null;
    umg.setzeListe(merke(liste, { name, abteilung, email }));
    zeichne();
    felderLeeren();
    melde(kannte ? `${name} aktualisiert.` : `${name} gespeichert.`, "erfolg");
    return true;
  }

  /**
   * Nimmt die im Profil stehende Gruppenleitung in den Vorrat auf.
   *
   * Der haeufigste Weg, wie ein Eintrag entsteht: Man hat die Adresse ohnehin
   * gerade oben eingetippt. Die Werte wandern erst in die Felder und dann
   * durch dieselbe Pruefung wie eine Handeingabe — eine zweite, stillere
   * Speicherroute waere eine zweite Stelle, an der die Pruefung fehlen kann.
   */
  function ausProfil() {
    const profil = umg.aktuellesProfil();
    if (!profil) return;
    const name = geputzt(profil.gruppenleiter_name);
    const email = geputzt(profil.gruppenleiter_email);
    if (!name && !email) {
      melde("Im Profil steht noch keine Gruppenleitung.", "warnung");
      return;
    }
    teile.nameFeld.value = name;
    teile.abteilungFeld.value = geputzt(profil.abteilung);
    teile.emailFeld.value = email;
    speichern();
  }

  /** Entfernt den Eintrag zur Adresse im E-Mail-Feld (_gruppenleiter_loeschen).
   *
   *  Der Bezug laeuft ueber das Feld und nicht ueber eine Listenmarkierung:
   *  Angetippt wird ein Eintrag ohnehin nur, um ihn in die Felder zu laden —
   *  so loescht "Löschen" immer genau das, was gerade sichtbar dasteht. */
  function loeschen() {
    const email = geputzt(teile.emailFeld.value);
    if (!email) {
      melde("Zum Löschen zuerst einen gespeicherten Eintrag antippen.", "warnung");
      return;
    }
    const neu = entferne(vorrat(), email);
    if (neu === null) {
      melde("Zu dieser E-Mail-Adresse ist nichts gespeichert.", "warnung");
      return;
    }
    umg.setzeListe(neu);
    zeichne();
    felderLeeren();
    melde("Gruppenleitung entfernt.", "erfolg");
  }

  /* ------------------------------------------------------------------------
     Schnittstelle nach app.js
     ------------------------------------------------------------------------ */

  return {
    /**
     * Baut die Oberflaeche in #gruppenleiter-halter. Einmal beim Start.
     *
     * Fehlt der Behaelter, geschieht nichts weiter und es kommt false zurueck
     * — eine Stammdatenliste ist kein Grund, den Start der App scheitern zu
     * lassen.
     */
    initialisieren(umgebung) {
      umg = umgebung;
      const halter = document.getElementById(HALTER_ID);
      if (!halter) return false;
      teile = bauOberflaeche(halter);
      setzeOffen(false);
      zeichne();
      return true;
    },

    /** Zeichnet Auswahlliste, Eintragsliste und Zusammenfassung neu. */
    zeichne() {
      zeichne();
    },
  };
})();
