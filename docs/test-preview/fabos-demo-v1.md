# FABOS: synthetische Produktionsdemo v1 (v0.20.14)

**Keine Echtdaten und keine produktive Aktivierung.** Dieser Demo-Datensatz ist von der echten Supabase-Datenbank vollständig getrennt. Er ist nur im lokalen Loopback-Vorschausystem verfügbar und standardmäßig **deaktiviert**.

## Demoumfang (Bezugszeit 16.11.2026, 09:00 Uhr CET)

- Acht festgelegte Startabteilungen: Dreherei, Fräserei, Verzahnerei, Entgraterei, Härterei, Schleiferei, Qualitätssicherung, Montage
- Genau 40 erfundene Personen, je fünf pro Abteilung; keine echten Beschäftigtendaten oder produktiven Auth-Nutzer
- 19 Maschinen; die namentlich bekannten Bezeichnungen **Maschine 136**, **Maschine 50**, **Hermle C62** und **Hermle C22** sind nur Identifikationsbeispiele, keine verifizierten aktuellen Maschinenbelegungen. Alle anderen Maschinen tragen ausdrücklich `DEMO` im Namen
- 32 eigenständige Demo-BA, acht fiktive Artikel und 176 sequenzielle OPs mit technischen OP-Beschreibungen
- Pro OP: fachlich passende Abteilung/Maschine, Mitarbeiter-Zuordnung, Rüstminuten, Zykluszeit je Teil, fertigungsabhängiger Start/Endtermin, Fertig-/Laufend-/Wartestatus und erreichte Gutstückmenge
- Ressourcenkonflikte: Jede Maschinenbelegung und jeder Vorgänger-Nachfolger-Übergang wird validiert; keine parallelen Maschinenbelegungen im Modell
- Zeiten: deterministisches vereinfachtes Werktags-Kapazitätsfenster 06:00–22:00 Uhr CET, ohne Feiertage, Abwesenheiten, Pausen und tatsächliche mannlose Laufzeiten. Das ist eine illustrative Annahme, keine Personal-/Maschinenoptimierung durch PlanOS.

## Nutzung der isolierten Vorschau

Das bestehende `npm run preview:test` startet **ohne** Demodaten wie bisher. Zur ausdrücklichen Aktivierung ausschließlich im isolierten lokalen Testsystem:

```bash
HUMBEL_PREVIEW_DEMO=1 npm run preview:test
```

Die bisherigen Vorschau-Testkennwörter (mindestens 24 zufällige Zeichen je aktiviertem Konto) müssen wie zuvor separat über `HUMBEL_PREVIEW_PASSWORDS` oder die ignorierte lokale Datei bereitgestellt werden. Keine Kennwörter ins Repository einchecken. Der Server lauscht unverändert nur an `127.0.0.1`, ist also **kein Internet-Demo-Deployment**.

Nach Anmeldung zeigt die Standard-App unter Produktion die simulierten laufenden Aufträge an Maschinen sowie abgeschlossene BA. Die gelbe Vorschauleiste verlinkt zusätzlich **„Maschinenbelegung und OP-Vorschau (Demodaten)“** (`/demo-plan`): Diese passwortgeschützte Nur-Lese-Ansicht zeigt je Maschine den laufenden und den nächsten vorgemerkten OP sowie die komplette Arbeitsplanreihenfolge aller 32 BA mit Sollzeiten.

**Wichtiges Datenmodell-Limit:** Die alte `production_orders`-Ansicht kennt nur Laufend/Pausiert/Abgeschlossen/Abgebrochen. Vorgemerkte Folge-OPs erscheinen daher bewusst im getrennten lesbaren Demo-Plan und **nicht** als fiktive produktive Warteschlangenfunktion der bestehenden App. Die tatsächliche OP-RPC-/PGlite-Fixture bleibt unverändert; die Legacy-Mengen sind weiterhin browserlokale Testdaten und nicht für mehrere Benutzer gemeinsam persistent. Start/Zwischenstand basieren auf einem eingefrorenen Demotermin; die Aufträge ändern ihren Status nicht automatisch durch Zeitablauf.

## Reproduzierbarkeit und Schutz

- Quelle: `scripts/preview/fabos-demo-data.js` mit `makeDemoData()` und `validate()`.
- Browser-Brücke und reine HTML-Planansicht: `scripts/preview/fabos-demo-adapter.js`.
- Der Schalter `HUMBEL_PREVIEW_DEMO` wird ausschließlich im privaten Preview-Server ausgewertet; normale Anwendung und Playwright-Regressionen verwenden unverändert ihre alten Fixtures.
- Automatische Prüfung: `node --test tests/documentation/demo-preview.test.js` sowie die regulären Projekttests.
- Nur isolierte Testkennwörter, keine echten Auftragsnummern, keine echte Maschinenkapazität oder vertrauliche Fertigungszeichnung.
- **Kein Merge, kein Deployment, keine DB-/Schema-/RLS-/Auth-Änderung und keine Produktivfreigabe.** Ein Chef-Zugang über das Internet braucht separat abgesichertes HTTPS/Authentifizierung, eine ausdrücklich freigegebene Testumgebung und die Integrationsprüfung.
