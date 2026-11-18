# FABOS – GitHub Pages Online-Testversion

**Quellbasis:** v0.20.16, PR #55, einschließlich der gestapelten FABOS-Design- und Entwicklungscommits ab main. Veröffentlichung aus einem **separaten Test-Branch**; kein produktiver Merge oder Supabase-Rollout.

## Ziel und UI

Die Anwendung wird aus derselben `schichtplan_mvp.html` und denselben CSS-/JS-Modulen gebaut. Auf dem veröffentlichten GitHub-Pages-Auftritt erscheinen weder der bisherige gelbe Entwicklungs-/Demo-Banner noch ergänzende technische Hinweiskarten. Beim Öffnen werden automatisch die fiktiven Produktionsdaten geladen, damit die eigentliche FABOS-Oberfläche direkt sichtbar ist.

Daten: acht Produktionsabteilungen, 50 Fantasiemitarbeiter sowie ein fiktiver Administrator, 48 Maschinen (sechs je Abteilung), 50 Betriebsaufträge mit mehr als 250 logisch verketteten OP, 20 Werkzeuge. Die gesonderte Seite `maschinenbelegung.html` enthält Maschinenfolge und OP-Anweisungen. Ein normaler Menülink im FABOS-Kopf führt dorthin.

**Die App ist für eine geschlossene fachliche Vorstellung geeignet, aber diese statische Ausgabe ist kein funktionsgleiches Produktionssystem:**

- GitHub Pages stellt nur HTML/JS/CSS bereit; es gibt dort weder serverseitige Anmeldung noch sichere Benutzerrechte.
- Jeder Besucher sieht die frei zugänglichen synthetischen Daten und erhält im Browser die fiktive Admin-Ansicht. **Keine echten Logins oder Passwörter auf dieser Site eingeben.**
- Lokale Testaktionen bearbeiten ausschließlich veränderbare Browser-Fixtures. Nach Neuladen sind Änderungen weg. RPC-Module ohne Backend liefern weiterhin einen `42501`-Fehler und sind nicht als funktionsfähige persistente Geschäftsprozesse abgenommen.
- Kein Netzwerkzugang zur produktiven Supabase-Datenbank. Der Build ersetzt die veröffentlichte Supabase-URL und den publishable Key durch funktionslose Platzhalter, lädt eigene Mock-Daten und lässt den echten Supabase-Client weg.
- Keine Original-Personaldaten in der Test-Datenbasis. Bekannte Teamnamen werden **nur in der generierten Testausgabe** neutralisiert; Repository-Quellen und die echten Funktionen bleiben unverändert.
- Der Datenstand ist eine deterministische Momentaufnahme, keine aktuelle Echtzeit-Disposition.

Diese Einschränkungen werden **in der Repository-Dokumentation und im Prüfbericht** offengelegt und bewusst nicht als störende UI-Banner dargestellt. Echte Authentifizierung, Rollenrechte, gemeinsame persistente Daten und serverseitige Fertigungsrückmeldungen sind eine andere Hosting-/Integrationsaufgabe.

## Reproduzierbarer Build und Prüfung

```sh
npm ci --include=dev --ignore-scripts
node scripts/pages/build.js
node --test tests/pages/pages-demo.test.js
```

Der Build erstellt `dist-pages/` mit sämtlichen benötigten Komponenten, einer lokalen Tailwind-CSS-Kopie, Client-Mock und OP-Plan. Die Testpipeline öffnet FABOS unter Desktop- und Smartphone-Chromium und prüft tatsächlich sichtbare Daten, Navigation und das Fehlen des Warnbanners. Andere bestehenden 853 Funktionstests, 33 Bildreferenzen, QA-Assertions und Produktivcode bleiben unverändert.

GitHub Actions Workflow: `.github/workflows/fabos-pages.yml`. Dieser Workflow läuft ausschließlich auf dem dedizierten Test-Branch oder nach manuellem Aufruf. Er lädt nur den fertig geprüften statischen Build als GitHub Pages-Artefakt hoch und veröffentlicht ihn mit `actions/deploy-pages@v4`, sofern **Settings → Pages → Build and deployment → Source: GitHub Actions** für das Repository eingestellt ist.

Die Website gehört zu einem öffentlichen GitHub-Repository. Kein Upload echter Personal-, Firmen-, Produktions- oder Zugangsdaten. Keine Freigabe anderer Github Actions, kein automatischer Merge nach main, kein Railway-Änderungsauftrag und kein Zugriff auf die echte Supabase-Instanz.

## Erweiterung v0.20.18 – Rollen- und Maschinenablauf (#58)

Startseite `index.html` ist eine Rollenwahl für ausdrücklich fiktive Testpersonen und eine Master-Demoperspektive. Sie simuliert die Navigation und Rechteansicht, ist aber **keine echte Anmeldung**; jeder öffentliche Besucher kann die Rolle wechseln. Es werden weder Passwörter noch reale Auth-Nutzer erzeugt.

| Rolle | Person (fiktiv) | Start und Vorführumfang |
| --- | --- | --- |
| Administration | FABOS Administration | `app.html?role=admin`: volle alte FABOS-Testoberfläche; Link zur Maschinenbelegung |
| Abteilungsleiter Fräserei | Jonas Lindner | `fertigung.html?role=lead`: eigene Fräsmaschinen, Betriebsaufträge, **Freigabebereich**; über „Weitere Module“ `app.html?role=lead` mit simuliertem `department_admin`-Datensatz |
| Mitarbeiter Fräserei | Mila Fischer | `fertigung.html?role=employee`: zugeordnete Fräsmaschinen und OP-Details, ohne Freigabebereich |

Die zwei Personen stammen bereits aus dem Fantasiemitarbeiterbestand (40 Namen) und erhalten jetzt eine eindeutig gebundene fiktive Browser-Identität. Keine zusätzlichen produktiven Mitarbeiterkonten.

Die Maschinenübersicht zeigt 19 vorhandene fiktiv belegte Maschinen, wenn „Alle Abteilungen“ gewählt ist (als Mitarbeiter/Leitung nur Fräserei). Eine Maschine öffnet den laufenden oder nächsten geplanten Arbeitsgang mit BA-Nummer, Artikel, Stückzahl, Person, Rüst- und Zykluszeit, Anweisung, Fertigungsstand sowie nächstem Auftrag. Die aus v0.20.17 stammenden **32 Test-BA und 176 OP** sind in der Auftragsfolge erkennbar. Es werden keine neuen echten Aufträge angelegt.

Im Freigabebereich werden offene Erststück-/Rüstfreigaben aus offenen Fräserei-OP abgeleitet. „Freigeben (Simulation)“ ändert nur den lokalen Zustand der geöffneten Seite – keine gemeinsamen/persistenten Freigaben und keine echte Autorisierung. Der Mitarbeiter sieht die Freigabe-Navigation nicht. Diese Präsentationsgrenze darf nicht mit tatsächlicher serverseitiger Berechtigung verwechselt werden.

Die neue Ansicht verwendet eine eigene responsive Demo-Stylesheet-Datei `css/pages-demo.css`; sie ändert die 33 historisch freigegebenen Produktions-Screenshotreferenzen nicht. Abnahme: `node --test tests/pages/pages-demo.test.js` (Admin plus Leiter/Mitarbeiter, Desktop und 390-px-Smartphone). Kein Merge oder Deployment mit diesem Auftrag.

Der bisherige öffentliche Link `schichtplan_mvp.html` wird im Pages-Build als einfache HTML-Weiterleitung auf `index.html` mit der Rollenwahl erhalten (ohne automatische Anmeldung).

## Erweiterung v0.20.19 – benutzbarer synthetischer Arbeitsraum (#60)

Die acht vorhandenen Abteilungen **Dreherei, Fräserei, Verzahnerei, Entgraterei, Härterei, Schleiferei, Qualitätssicherung und Montage** bleiben bestehen. Zu jeder Abteilung gehören **genau sechs Maschinen (48 insgesamt)**; alle Maschinen besitzen mindestens einen fiktiven, zugeordneten Arbeitsgang. Die Testbasis enthält **50 Fantasiemitarbeiter (plus einen getrennten Demo-Administrator)**, **50 initiale BA** mit mehr als 250 verketteten OP, zugeordneten Maschinen, Mitarbeitern, Artikelbezeichnungen, Material, Arbeitsanweisung, Rüst- und Zykluszeiten.

**Master-Demoperspektive:** Auf der öffentlichen Rollenwahl „Master-Zugang“ öffnen. URL `fertigung.html?role=master`. Der Master sieht sämtliche acht Abteilungen und kann filtern, Maschinen-/BA-Details aufrufen, alle Freigaben simulieren, fiktive Mitarbeiter abteilungsweise anzeigen, neu anlegen und (mit Ausnahme bestehender Abteilungsleiter) umteilen. Die fiktiven Abteilungsleitungen jeder Abteilung und jeder Testmitarbeiter können ebenfalls über Dropdowns auf der Rollenwahl geöffnet werden.

**Benutzbare Abläufe im Produktionscockpit:** Betriebsauftrag mit Artikel, Material, Stückzahl, Maschine, Person und Arbeitsanweisung anlegen; eine passende Maschine öffnen; Arbeitsgang starten, wenn Vorgänger-OP abgeschlossen und die Maschine nicht bereits belegt ist; Gutteile buchen bis zur Sollmenge; Arbeitsgang abschließen; Mitarbeitende für offene OP umteilen; Erststück-/Rüstfreigabe simulieren. Das Auftragsverzeichnis verlinkt die zugeordneten Maschinen. Die Testdatensätze sind wiederholt reproduzierbar.

**Speicherung:** Die Bedienaktionen werden nur für diese Browsertab-Testsitzung im `sessionStorage` des Browsers behalten, auch nach Seiten-Neuladen. Beim „Testdaten zurücksetzen“ werden sie verworfen. Andere Tabs, Personen, Geräte und die vollständige ältere `app.html`-Ansicht haben **keine synchronisierten Änderungen**; die ursprünglichen 50 BA sind dort als Fixture enthalten, neu im Cockpit erzeugte BA nicht. **Kein echtes Benutzerkonto, kein passwortgeschützter Master, keine gemeinsame Datenbank und keine echte Freigabeberechtigung.** Öffentliche Rollenwahl/URL-Parameter erlauben jeder Person den Master-Blick auf erfundene Daten. Ein echter sicherer Masterzugang und zentral persistente, vollständig integrierte FABOS-Geschäftsprozesse setzen ein separates Backend und fachliche Abnahme voraus.

**Sicherheitsgrenzen:** Keine Produktionsdaten, keine produktive Supabase-Auth/RLS/DB-Veränderung, kein direkter main-Push, keine Änderung der freigegebenen Snapshotbilder. Entwicklungsbranch und Draft-PR; der neue isolierte QA-Workflow ist nicht deployend und legt geprüfte Screenshots sowie ein statisches Build als herunterladbares Artefakt ab. Live-Veröffentlichung benötigt gesonderte Entscheidung und ggf. Administrationskonfiguration der GitHub Pages-Umgebung.
