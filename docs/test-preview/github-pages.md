# FABOS – GitHub Pages Online-Testversion

**Quellbasis:** v0.20.16, PR #55, einschließlich der gestapelten FABOS-Design- und Entwicklungscommits ab main. Veröffentlichung aus einem **separaten Test-Branch**; kein produktiver Merge oder Supabase-Rollout.

## Ziel und UI

Die Anwendung wird aus derselben `schichtplan_mvp.html` und denselben CSS-/JS-Modulen gebaut. Auf dem veröffentlichten GitHub-Pages-Auftritt erscheinen weder der bisherige gelbe Entwicklungs-/Demo-Banner noch ergänzende technische Hinweiskarten. Beim Öffnen werden automatisch die fiktiven Produktionsdaten geladen, damit die eigentliche FABOS-Oberfläche direkt sichtbar ist.

Daten: acht Produktionsabteilungen, 40 Fantasiemitarbeiter sowie ein fiktiver Administrator, 19 Maschinen, 32 Betriebsaufträge mit 176 logisch verketteten OP, 20 Werkzeuge. Die gesonderte Seite `maschinenbelegung.html` enthält Maschinenfolge und OP-Anweisungen. Ein normaler Menülink im FABOS-Kopf führt dorthin.

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
