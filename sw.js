/*
 * Service Worker der GLAZ-Korrekturbuchungsliste (PWA).
 *
 * Diese Datei ist kein Beiwerk, sondern das Herzstueck der App. Die Zusage
 * lautet: Seite einmal oeffnen, "Zum Home-Bildschirm hinzufuegen", und danach
 * laeuft alles OHNE Netz und OHNE eingeschalteten PC -- im Zug, im Keller, im
 * Ausland ohne Roaming. Erfuellt wird diese Zusage genau hier: Der Worker legt
 * beim ersten Besuch JEDE Datei der App in den Cache und bedient danach jede
 * Anfrage aus diesem Cache.
 *
 * ACHTUNG: Diese Datei ist eine Vorlage. Die Platzhalter 1.1.0+28780583b3 und
 * [
  "./",
  "./index.html",
  "./app.css",
  "./app.js",
  "./fonts/Manrope.ttf",
  "./fonts/OFL.txt",
  "./gruppenleiter.js",
  "./icons/app-icon.svg",
  "./icons/apple-touch-icon.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./manifest.webmanifest",
  "./py/glaz/__init__.py",
  "./py/glaz/excel_engine.py",
  "./py/glaz/filename.py",
  "./py/glaz/mailtext.py",
  "./py/glaz/model.py",
  "./py/glaz/portabel.py",
  "./py/glaz/resources/template.xlsx",
  "./py/glaz/resources_util.py",
  "./py/glaz/timeconv.py",
  "./py/glaz/validation.py",
  "./py/glaz/vorgaben.py",
  "./vendor/pyodide/pyodide-lock.json",
  "./vendor/pyodide/pyodide.asm.js",
  "./vendor/pyodide/pyodide.asm.wasm",
  "./vendor/pyodide/pyodide.js",
  "./vendor/pyodide/pyodide.mjs",
  "./vendor/pyodide/python_stdlib.zip",
  "./version.json"
] werden von tools/baue_pwa.py beim Bauen ersetzt. Direkt aus
 * web/ ausgeliefert funktioniert sie NICHT -- und das ist Absicht: Die Liste
 * der zu cachenden Dateien wird nie von Hand gepflegt, sondern aus dem
 * tatsaechlichen Inhalt von dist_web/ erzeugt. Eine handgepflegte Liste waere
 * die sicherste Art, die Offline-Zusage zu brechen: Man vergisst genau eine
 * Datei und merkt es erst, wenn kein Netz mehr da ist.
 */

/** Versionsstempel aus version.json: Anwendungsversion + Inhaltskuerzel. */
const VERSION = "1.1.0+28780583b3";

/*
 * Der Cache-Name traegt die Version. Dadurch legt jeder neue Build einen
 * frischen Cache an, statt in den alten hineinzuschreiben -- ein halb
 * aktualisierter Cache (neue index.html, alte app.js) waere schlimmer als gar
 * kein Update. Der alte Cache wird erst beim activate geloescht, also erst,
 * wenn der neue vollstaendig steht.
 */
const CACHE_NAME = `glaz-pwa-${VERSION}`;

/**
 * Alle Dateien der App, von tools/baue_pwa.py eingesetzt: Oberflaeche,
 * Symbole, Schrift, Pyodide (Interpreter + Standardbibliothek) und die
 * Excel-Vorlage. Die Pfade sind relativ zum Verzeichnis dieses Workers, damit
 * die App auch in einem Unterverzeichnis liegen kann (GitHub Pages).
 */
const PRECACHE = [
  "./",
  "./index.html",
  "./app.css",
  "./app.js",
  "./fonts/Manrope.ttf",
  "./fonts/OFL.txt",
  "./gruppenleiter.js",
  "./icons/app-icon.svg",
  "./icons/apple-touch-icon.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./manifest.webmanifest",
  "./py/glaz/__init__.py",
  "./py/glaz/excel_engine.py",
  "./py/glaz/filename.py",
  "./py/glaz/mailtext.py",
  "./py/glaz/model.py",
  "./py/glaz/portabel.py",
  "./py/glaz/resources/template.xlsx",
  "./py/glaz/resources_util.py",
  "./py/glaz/timeconv.py",
  "./py/glaz/validation.py",
  "./py/glaz/vorgaben.py",
  "./vendor/pyodide/pyodide-lock.json",
  "./vendor/pyodide/pyodide.asm.js",
  "./vendor/pyodide/pyodide.asm.wasm",
  "./vendor/pyodide/pyodide.js",
  "./vendor/pyodide/pyodide.mjs",
  "./vendor/pyodide/python_stdlib.zip",
  "./version.json"
];

/*
 * version.json wird bewusst NICHT aus dem Cache bedient, sondern zuerst aus
 * dem Netz: Genau daran erkennt die Seite, ob es eine neuere Fassung gibt.
 * Ohne Netz faellt auch diese Datei auf den Cache zurueck.
 */
const NETZ_ZUERST = "version.json";

/* ------------------------------------------------------------------ */
/* Installation: alles einsammeln                                      */
/* ------------------------------------------------------------------ */

self.addEventListener("install", (ereignis) => {
  ereignis.waitUntil(installiere());
});

async function installiere() {
  const cache = await caches.open(CACHE_NAME);

  /*
   * In Haeppchen statt mit einem einzigen addAll: Pyodide bringt eine
   * ~10-MB-WASM-Datei mit, und mobile Browser (vor allem iOS Safari) quittieren
   * sehr viele gleichzeitige Anfragen gern mit einem Abbruch. Sechs Dateien auf
   * einmal sind schnell genug und robust.
   */
  const haeppchen = 6;
  for (let i = 0; i < PRECACHE.length; i += haeppchen) {
    const teil = PRECACHE.slice(i, i + haeppchen);
    try {
      await cache.addAll(teil);
    } catch (fehler) {
      /*
       * Zweiter Versuch, Datei fuer Datei. So sehen wir in der Konsole, welche
       * Datei klemmt, statt nur "addAll failed". Bleibt es dabei, scheitert die
       * Installation absichtlich: Ein unvollstaendiger Cache wuerde die
       * Offline-Zusage brechen, und ein lautes Scheitern beim ersten Besuch
       * (mit Netz) ist allemal besser als eine tote App ohne Netz.
       */
      for (const pfad of teil) {
        try {
          await cache.add(pfad);
        } catch (einzelfehler) {
          console.error("[sw] Konnte nicht in den Cache legen:", pfad, einzelfehler);
          throw einzelfehler;
        }
      }
    }
  }

  /*
   * Kein skipWaiting() an dieser Stelle: Wenn gerade jemand ein Formular
   * ausfuellt, darf ihm nicht mitten im Satz eine neue Fassung untergeschoben
   * werden. Stattdessen meldet sich der Worker bei der Seite und wartet auf
   * ihr Einverstaendnis (Nachricht "uebernehmen", siehe unten).
   *
   * self.registration.active ist nur dann gesetzt, wenn schon eine aeltere
   * Fassung laeuft -- beim allerersten Besuch gibt es nichts zu melden.
   */
  if (self.registration.active) {
    await meldeAllen({ typ: "neue-version", version: VERSION });
  }
}

/* ------------------------------------------------------------------ */
/* Aktivierung: alte Caches raeumen                                     */
/* ------------------------------------------------------------------ */

self.addEventListener("activate", (ereignis) => {
  ereignis.waitUntil(aktiviere());
});

async function aktiviere() {
  // Alte Versionsstaende wegwerfen. Sonst sammelt sich auf dem Handy mit
  // jedem Update ein weiterer 20-MB-Block an -- und Speicher ist genau das,
  // woran mobile Browser Website-Daten als Erstes loeschen.
  const namen = await caches.keys();
  await Promise.all(
    namen
      .filter((name) => name.startsWith("glaz-pwa-") && name !== CACHE_NAME)
      .map((name) => caches.delete(name))
  );

  // Ab sofort auch fuer bereits offene Seiten zustaendig.
  await self.clients.claim();
  await meldeAllen({ typ: "aktiv", version: VERSION });
}

/* ------------------------------------------------------------------ */
/* Abrufe: offline zuerst                                              */
/* ------------------------------------------------------------------ */

self.addEventListener("fetch", (ereignis) => {
  const anfrage = ereignis.request;

  // Nur GET laesst sich sinnvoll cachen.
  if (anfrage.method !== "GET") return;

  // Fremde Herkunft (falls doch einmal etwas eingebunden wird) geht am Cache
  // vorbei -- die App selbst braucht das Netz nie.
  const adresse = new URL(anfrage.url);
  if (adresse.origin !== self.location.origin) return;

  if (adresse.pathname.endsWith(NETZ_ZUERST)) {
    ereignis.respondWith(netzZuerst(anfrage));
    return;
  }

  ereignis.respondWith(cacheZuerst(anfrage));
});

/**
 * Cache-first: Die App ist offline-first, nicht "offline zur Not". Was im
 * Cache liegt, wird von dort bedient -- ohne auf ein Netz zu warten, das im
 * Zweifel gar nicht da ist. Das Netz ist nur der Rueckfall.
 */
async function cacheZuerst(anfrage) {
  const cache = await caches.open(CACHE_NAME);

  const treffer = await cache.match(anfrage, { ignoreSearch: true });
  if (treffer) return treffer;

  try {
    const antwort = await fetch(anfrage);
    // Erfolgreich nachgeladenes wandert in den Cache, damit es beim naechsten
    // Mal offline verfuegbar ist (z. B. eine Datei, die erst spaeter dazukam).
    if (antwort && antwort.ok && antwort.type === "basic") {
      cache.put(anfrage, antwort.clone());
    }
    return antwort;
  } catch (fehler) {
    /*
     * Kein Netz und nichts im Cache. Fuer eine Navigation (der Nutzer tippt
     * das Symbol auf dem Home-Bildschirm an) ist die Startseite die richtige
     * Antwort: Adressen mit Anhaengsel oder ein tieferer Pfad landen so
     * trotzdem in der App.
     */
    if (anfrage.mode === "navigate") {
      const start =
        (await cache.match("./index.html")) || (await cache.match("./"));
      if (start) return start;
    }
    return new Response(
      "Offline und nicht im Zwischenspeicher: " + anfrage.url,
      { status: 504, headers: { "Content-Type": "text/plain; charset=utf-8" } }
    );
  }
}

/**
 * Netz zuerst, Cache als Rueckfall -- nur fuer version.json. Daran erkennt die
 * Seite eine neuere Fassung; aus dem Cache bedient waere die Abfrage sinnlos.
 */
async function netzZuerst(anfrage) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const antwort = await fetch(anfrage, { cache: "no-store" });
    if (antwort && antwort.ok) {
      cache.put(anfrage, antwort.clone());
      return antwort;
    }
  } catch (fehler) {
    // Offline ist der Normalfall, kein Fehler. Weiter unten aus dem Cache.
  }
  const treffer = await cache.match(anfrage, { ignoreSearch: true });
  return (
    treffer ||
    new Response(JSON.stringify({ version: VERSION }), {
      headers: { "Content-Type": "application/json; charset=utf-8" },
    })
  );
}

/* ------------------------------------------------------------------ */
/* Update-Pfad: die Seite entscheidet                                  */
/* ------------------------------------------------------------------ */

self.addEventListener("message", (ereignis) => {
  const nachricht = ereignis.data || {};

  if (nachricht.typ === "uebernehmen") {
    /*
     * Die Seite hat die neue Fassung bestaetigt. skipWaiting() macht diesen
     * Worker sofort zum aktiven; die Seite laedt sich danach ueblicherweise
     * selbst neu (controllerchange).
     */
    self.skipWaiting();
    return;
  }

  if (nachricht.typ === "version-abfragen") {
    // Damit die Oberflaeche anzeigen kann, welcher Stand gerade laeuft.
    const antwort = { typ: "version", version: VERSION };
    if (ereignis.source) {
      ereignis.source.postMessage(antwort);
    } else {
      meldeAllen(antwort);
    }
  }
});

/** Schickt eine Nachricht an alle offenen Seiten dieser App. */
async function meldeAllen(nachricht) {
  const seiten = await self.clients.matchAll({
    type: "window",
    includeUncontrolled: true,
  });
  for (const seite of seiten) {
    seite.postMessage(nachricht);
  }
}
