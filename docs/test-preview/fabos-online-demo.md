# FABOS – isolierte HTTPS-Demo (v0.20.15)

Die bestehende lokale Vorschau aus Issue #40 und die synthetischen Fertigungsdaten aus PR #53 bleiben unverändert. Die zusätzliche Online-Freigabe setzt eine eigene, getrennte Hosting-Umgebung voraus. Es werden **keine Produktionsdaten** verwendet, und die produktive Supabase-Datenbank wird weder angesprochen noch verändert.

## Sicherheitsgrenzen

- `HUMBEL_PREVIEW_ONLINE=1` aktiviert den separaten Hostingmodus. Ohne diese Variable läuft `npm run preview:test` weiter ausschließlich auf `127.0.0.1`.
- Online zwingend: `HUMBEL_PREVIEW_DEMO=1`, `HUMBEL_PREVIEW_ORIGIN=https://<freigegebener-hostname>` (exakter HTTPS-Origin, kein Pfad, Port oder Query), `HUMBEL_PREVIEW_PASSWORDS` mit separaten zufälligen Passwörtern von mindestens 24 Zeichen je Konto.
- Die Online-Demo prüft Host, Origin bei schreibenden Requests und `X-Forwarded-Proto=https` am TLS-Reverse-Proxy. Public TLS muss am geprüften Railway-Router terminiert werden, direkter Backend-Zugang darf nicht außerhalb des Routers möglich sein.
- Session-Cookies sind `HttpOnly; SameSite=Strict; Secure` im Online-Modus und verfallen nach einer Stunde. Der offene Healthcheck `/healthz` enthält lediglich `ok`.
- `/documentation/` wird im Online-Modus blockiert, damit kein zusätzlicher anonym erreichbarer Dokumentationszugang besteht.
- `/demo-plan` bleibt passwortgeschützt. Demo-Daten sind fiktiv und statisch zum Bezugszeitpunkt 16.11.2026 09:00 MEZ; Maschinenbelegungen werden nicht automatisch fortgeschrieben.
- Kein echtes Mehrbenutzersystem, kein produktiver SSO-Ersatz, keine produktive RLS-/Auth-Abnahme, keine sichere Kollaboration über Legacy-Browser-Fixtures. Pilotzugang nur mit ausdrücklich akzeptierten Einschränkungen.

## Getrennte Bereitstellung

Railway-Projekt: `FABOS Synthetic Demo`, Umgebung `demo` (nicht produktives Hauptprojekt). GitHub-Quelle: `butzi124-ship-it/schichtplan-mvp`, Entwicklungsbranch `feature/fabos-secure-online-preview-v0.20.15`. Nach final grüner CI soll die exakt geprüfte Commit-SHA für die Demo gepinnt werden. Keine automatische PR-Vorschau, kein Merge nach `main`, kein Deployment des produktiven Workflows.

**Vor der Linkfreigabe nachweisen:** abgetrennte Variablen, überprüfte TLS-Domain, gültige Testpasswörter, anonymer Aufruf → Login statt App/Daten, erfolgreicher rollenbezogener Testlogin, Zugriff auf Maschinenfolgeplan, Logout/Session-Widerruf, Desktop/Smartphone und gekennzeichnete Demo-Inhalte. Eine neue Railway-Domain darf erst nach diesem Nachweis als benutzbarer Demo-Link bezeichnet werden.

## Betrieb

`npm ci --include=dev --ignore-scripts` installiert die PGlite-Abhängigkeit für die isolierten Test-RPCs. Startkommando `npm run preview:test`, Port aus der Railway-`PORT`-Variable. Der Server bindet im Online-Modus an `0.0.0.0` und vertraut ausschließlich dem fest konfigurierten öffentlichen Host. Secrets ausschließlich in Railway-Variablen setzen, niemals in GitHub-Dateien, Tests, CI-Protokolle oder Chat-Nachrichten einchecken.

Jeder Serverneustart setzt den Demo-Datenzustand zurück. Vorschau-Änderungen sind keine echten Maschinenmeldungen. Auftrags-/Maschinendaten wurden ausschließlich synthetisch erzeugt.

## Abnahme

`node --test tests/documentation/online-preview.test.js`, `npm run test:documentation` und alle unveränderten CI-/QA-/Design-/E2E-Gates. Bei fehlerhaftem TLS, fehlendem Login-Schutz oder rotem CI bleibt der externe Zugang gesperrt. Die hier beschriebene Hostingvorbereitung ersetzt keine produktive Freigabe.
