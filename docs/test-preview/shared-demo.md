# FABOS v0.20.21 – isolierter gemeinsamer Testarbeitsraum

## Zweck und Abgrenzung

Diese **separate** Testfassung stellt ausschließlich frei erfundene Mitarbeiter, Maschinen und Aufträge bereit. Sie arbeitet **nicht** mit produktiven Supabase-Projekten, realer Mitarbeiter-Authentifizierung, realen Produktionsbuchungen oder echten Auftragsdaten. Hauptanwendung, Schema, RLS, Auth-Nutzer und die 33 freigegebenen Referenzbilder bleiben unverändert.

**Ausgangsdaten:** acht Abteilungen (Dreherei, Fräserei, Verzahnerei, Entgraterei, Härterei, Schleiferei, Qualitätssicherung, Montage), 50 fiktive Mitarbeiter, genau sechs Maschinen je Abteilung (48), 50 fiktive BA und 277 OP, zusätzlich eine getrennte Master-Demopersona. Jede Maschine besitzt mindestens eine OP. 20 fiktive Werkzeuge sind in der älteren Administrations-Preview vorhanden.

## Login und serverseitige Rechte

Der Server verwendet die bestehenden **drei getrennt konfigurierten, langen Testpasswörter** (JSON unter `HUMBEL_PREVIEW_PASSWORDS`), zeitlich begrenzte HttpOnly-/SameSite-Cookies und eine strikt geprüfte `Origin` bei Schreibaktionen.

- **Master** (`admin`): alle acht Abteilungen, Maschinen, BA, Mitarbeiter und Freigaben; Anlegen neuer fiktiver BA/Mitarbeiter, Mitarbeiterwechsel, OP-Zuweisungen, Rücksetzen.
- **Abteilungsleitung Fräserei** (`leader`): nur Fräserei, Auftragserstellung/OP-Umbuchung und Freigaben innerhalb dieser Abteilung.
- **Testmitarbeiter Fräserei** (`employee`, Mila Fischer): Maschinen/BA der eigenen Abteilung einsehen, ausschließlich **selbst zugeteilte** OPs starten, Gutteile buchen und abschließen; keine Freigaben und keine Personalverwaltung.

Rollen ergeben sich aus der **serverseitig bestätigten Sitzung**, nicht aus URL-Abfrageparametern, lokalen Browserrollen oder versteckten UI-Elementen. Jede Änderung wird vom Server geprüft. Produktionsautorisierung ist von der Demo vollständig getrennt.

## Gemeinsame Speicherung

Die API `/api/state` liefert rollenbezogene Ansichten; `/api/action` validiert Änderungen und speichert ein versioniertes, JSON-basiertes, ausschließlich synthetisches Testdatenpaket atomar im konfigurierten `FABOS_DEMO_STATE_PATH`. Der Zustand bleibt über Reload und Serverneustart erhalten, wenn ein persistentes Volume existiert. Eine Master-Aktion kann die Testdaten bewusst auf den fiktiven Grundzustand zurücksetzen. Sitzungen verfallen nach einer Stunde bzw. bei Neustart.

Der Server startet im **Onlinebetrieb absichtlich nicht**, solange kein eigenständiger persistenter Datenträger unter `/data` gemountet ist und `FABOS_DEMO_STATE_PATH` nicht unter `/data/` liegt. Ein temporäres Container-Dateisystem wird **niemals** als dauerhafte Speicherung ausgegeben.

Der bestehende Railway-Dienst `fabos-demo-web` im Environment `demo` zeigt derzeit eine ältere Vorschau. Keine Umstellung/Deployment ohne:
1. Prüfung der gesamten CI, insbesondere `Functional Browsers`, `Visual Regression`, `QA Integrity`, `Quality Gate`, `e2e`, zusätzliche Auth-/Persistenz-/Browsertests.
2. Ausdrückliche Review-/Merge- und Deployment-Freigabe gemäß `AGENTS.md`, `docs/DEVELOPMENT.md`.
3. Zustimmung zu möglichen Volumenkosten, Bereitstellung eines separaten Railway-Volumes mit Mount `/data`, Setzen von `FABOS_DEMO_STATE_PATH=/data/fabos-workspace.json`.
4. Konfiguration des dedizierten Startbefehls `npm run preview:shared` und des öffentlichen HTTPS-Origin; Kontrolle der Domain, Login-Sitzung, serverseitigen Rechte sowie persistenter Daten nach einem Neustart.

`npm run preview:shared` startet lokal mit einem explizit gesetzten JSON der drei mindestens 24 Zeichen langen Testpasswörter und einem absoluten, **nicht produktiven** Dateipfad. Testzugangsdaten nicht in GitHub committen oder weitergeben.

## Was noch nicht vollständig integriert ist

Dieser Kandidat bietet eine **gemeinsame, serverseitig kontrollierte Produktions-Testoberfläche**, jedoch kein fertiges, produktives FABOS-ERP: Die breitere historische `app.html` (Schichtplanung, Werkzeuge, übrige Module) ist ein **getrennter, nicht synchronisierter Browser-Mock** und nur für Master als zusätzliche Vorschau erreichbar. Neue BA aus der gemeinsamen Fertigungsansicht werden nicht automatisch in diese Alt-Ansicht synchronisiert. Es gibt aktuell drei Demo-Login-Identitäten, **keine 50 separaten echten Benutzerkonten**. Die Weboberfläche kann noch keinen frei wählbaren Mitarbeiter authentifiziert impersonieren; weitere Testkonten erfordern eine gesonderte Session-/Berechtigungserweiterung. Produktionsreife, zentrale Vollintegration und Echtdatenübernahme wurden **nicht** freigegeben oder nachgewiesen.

## Wiederholbare Prüfungen

```sh
npm ci --ignore-scripts
node scripts/pages/build.js
node --test tests/pages/shared-demo.test.js
node --test tests/pages/shared-demo-browser.test.js
node --test tests/pages/pages-demo.test.js
```

Der nicht deployende Workflow `.github/workflows/fabos-shared-demo-check.yml` deckt die isolierten Server-/Rollen-/Neustart- und Chromium Desktop-/390px-Journeys ab. Grüne fokussierte Tests ersetzen nicht das vollständige GitHub-Release-Gate.
